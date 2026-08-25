// ============================================================================
//  Adaptateur transporteur : FEDEX (Shipment Detail CSV, export portail FedEx)
//  1 fichier CSV brut par mois ("0g000e48hq_..." export, 1 ligne = 1 colis/
//  envoi), en-tetes FRANCAIS (confirme sur le CSV reel de juin 2026, malgre le
//  meme prefixe de nommage qu'un export anglais vu ailleurs -- variable selon
//  le mois, colonnes resolues PAR NOM, jamais par position).
//
//  Filtrage "Droits & Taxes" (process manuel documente par l'utilisateur,
//  2026-08-19) : une facture dont TOUTES les lignes ont un montant de fret
//  (Frais de transport) nul et un montant "Droits de douane et taxes" non nul
//  est un ajustement douanier pur, PAS un envoi facturable -- ses lignes sont
//  exclues du fichier livre (reproduit "si une facture correspond a Droit &
//  Taxes il faut supprimer les lignes correspondantes dans Shipment Detail").
//
//  Detection "facture d'un autre mois" (decouverte 2026-08-19, facture reelle
//  634313590 de juin 2026 : 392 lignes, expeditions 11-27/05/2026, fret total
//  22384EUR, ABSENTE du fichier livre reel bien que datee 01/06/2026) :
//  RESOLUE 2026-08-20 -- aucune regle par date n'etant fiable (confirme sur
//  plusieurs cas reels contradictoires), la regle retenue est : une facture
//  du CSV brut sans PDF correspondant parmi les PDF portail fournis par
//  l'utilisateur est exclue automatiquement (le pole transport ne recoit le
//  PDF qu'une fois la facture reellement due ce mois-ci). Validee exacte sur
//  juin 2026 : exclut 634313590 et 634364158 (memes 2 factures absentes du
//  fichier reel), Fret recalcule = 45515,83EUR = fichier livre exactement
//  (3607/3607 lignes). Ne s'applique que si au moins 1 PDF est fourni --
//  sinon aucune facture n'est exclue par cette regle (juste Droits&Taxes).
//
//  Montant (Shipment Detail!A, sert au TCD/Bilan factures/Bilan clients) :
//  formule modele =BK+BN = 'Frais de transport' + 'Remise' (deja negative).
//  Mode d'envoi (Shipment Detail!E) : =IF(Service="FedEx Priority","ST","FICP")
//  -- confirme exact sur 3607/3607 lignes reelles de juin 2026.
//  Zone (Import ERP!J) : XLOOKUP pays (ISO2 destinataire) -> colonne "FICP"
//  de l'onglet Zoning (jamais IE/RE, quasi vides) -- valide contre le CSV ERP
//  reel livre (Finlande -> Zoning FICP='T' -> CSV Zone='T').
//  TVA (Import ERP!N) : 20% si pays destinataire dans la liste UE (config,
//  27 pays), 0% sinon -- reconfirme caractere pour caractere sur le fichier
//  reel de juin 2026 (formule teste le PAYS, PAS le mode FEDEX FRANCE/
//  INTERNATIONAL, malgre ce que la video d'avril semblait suggerer via de
//  simples valeurs observees).
//  Colis volumineux (Import ERP!R) : forfait fixe 10EUR si "Dimmed Length
//  (cm)" > 60cm.
//  E/P (Import ERP!H, via TCD!F) : XLOOKUP tracking -> Shipment Detail!D
//  ("entreprise"->E, sinon P) -- colonne D reconfirmee 100% LITERALE (jamais
//  formule) sur les 3607 lignes reelles de juin -- alimentee via l'export WMS
//  partage "AAAA MM - Export expeditions_brut.xlsx" (meme mecanisme deja
//  utilise par d'autres transporteurs, cf. core/exportBrut.js), repli "P" par
//  defaut si tracking absent de l'export (avec alerte).
// ============================================================================
const path = require('path');
const pdfParse = require('pdf-parse');
const { num, round2, roundUp1 } = require('../../core/csv');
const { validate } = require('../../core/validate');
const { findBrutFiles, readBrutRows, epParTrackingFromExport } = require('../../core/exportBrut');
const cfg = require('./config.json');

function normKey(s) {
  // Apostrophes typographiques ('U+2019) frequentes dans les en-tetes CSV FedEx
  // ("Numero de suivi de l'envoi") -- normalisees en apostrophe simple avant comparaison.
  return String(s || '')
    .normalize('NFD').replace(/[̀-ͯ]/g, '')
    .replace(/[‘’]/g, "'")
    .toLowerCase().replace(/[.,]/g, '').replace(/\s+/g, ' ').trim();
}

function colIndexByName(header, name) {
  const target = normKey(name);
  return header.findIndex((h) => normKey(h) === target);
}

/** Lit un CSV FedEx (export portail, en-tete FR, separateur virgule, encodage UTF-8 --
 * confirme sur les 2 CSV bruts reels de juin 2026, tous deux en francais malgre le meme
 * prefixe "0g000e48hq_" observe sur un export anglais d'un autre mois dans la video). */
function readFedexCsv(p) {
  const fs = require('fs');
  const buf = fs.readFileSync(p);
  let text = buf.toString('utf8');
  if (text.charCodeAt(0) === 0xfeff || text.slice(0, 3) === '﻿') text = text.slice(1);
  const rows = [];
  let row = [];
  let field = '';
  let inQuotes = false;
  for (let i = 0; i < text.length; i++) {
    const c = text[i];
    if (inQuotes) {
      if (c === '"') {
        if (text[i + 1] === '"') { field += '"'; i++; } else inQuotes = false;
      } else field += c;
    } else if (c === '"') inQuotes = true;
    else if (c === ',') { row.push(field); field = ''; }
    else if (c === '\r') { /* skip */ }
    else if (c === '\n') { row.push(field); rows.push(row); row = []; field = ''; }
    else field += c;
  }
  if (field !== '' || row.length) { row.push(field); rows.push(row); }
  const header = (rows[0] || []).map((h) => String(h || '').trim());
  const data = rows.slice(1).filter((r) => r.some((v) => v !== ''));
  return { file: path.basename(p), header, rows: data };
}

/** Date "mm/jj/aaaa" -> {y,m,d} (format US du CSV FedEx, confirme sur le brut reel). */
function parseDateUs(s) {
  const m = /^(\d{1,2})\/(\d{1,2})\/(\d{4})$/.exec(String(s || '').trim());
  if (!m) return null;
  return { y: Number(m[3]), m: Number(m[1]), d: Number(m[2]) };
}

/** PDF FedEx (facture "Facture de frais de transport") : bloc labels/valeurs colle par
 * pdf-parse ("No de Client:\nNo de Facture:\n...\nMontant du\n*****1234\n634323394\n...368,95
 * EUR") + ligne "Total dû EUR <montant>" (recapitulatif final, redondant avec "Montant du" mais
 * plus simple a capturer isolement). Confirme exact sur les 9 PDF reels de juin 2026 (1 a 195
 * pages). */
/** Bordereaux (1 par envoi/tracking) : bloc "<taxable>\n<date envoi>DD/MM/YYYY<tracking>\n
 * <non taxable>\n<sous-total EUR>" (pdf-parse insere un retour a la ligne entre chaque cellule,
 * contrairement a pypdf cote Python qui aplatit tout sur 1 ligne -- pattern adapte en
 * consequence, mais memes 4 groupes). Sert au controle "bordereau facture par FedEx mais
 * absent du CSV Shipment Detail" (2026-08-20, cas reel juillet : envoi de novembre 2025
 * facture par FedEx mais jamais reporte dans le CSV, ecart 83,81EUR sur facture 634393398). */
function extractBordereaux(text) {
  const pat = /([\d.,]+)\s*\n(\d{2}\/\d{2}\/\d{4})(\d{9,15})\s*\n([\d.,]+)\s*\n([\d.,]+)/g;
  const bordereaux = [];
  let m;
  while ((m = pat.exec(text)) !== null) {
    bordereaux.push({
      tracking: m[3],
      dateEnvoi: m[2],
      sousTotal: num(m[5].replace(/\./g, '').replace(',', '.')),
    });
  }
  return bordereaux;
}

/** Supplements "surplus" (Zones eloignees/Colis volumineux/Plus-value BtoC) jamais presents
 * dans le CSV Shipment Detail -- UNIQUEMENT visibles dans le detail texte des bordereaux PDF,
 * a extraire par mots-cles (mail pole transport "Oubli facturation surplus Fedex janvier a mai
 * 2026", 2026-08-20, 7919,64EUR de surplus non factures sur 5 mois faute d'automatisation) :
 *   - "zone" (generique) / "Frais de traitement des importations aux États-Unis" (precise le
 *     2026-08-25, "s" final parfois tronque dans le PDF) -> Zones eloignees
 *   - "Charge pour dépassement de dimension" / "Supplément pour manutention supplémentaire
 *     :poids" (precise le 2026-08-25) -> Colis volumineux (S'ADDITIONNE a la regle Excel
 *     existante longueur>60cm=10EUR ET entre les 2 libelles PDF eux-memes -- confirme aupres
 *     du pole transport 2026-08-20/25 : ce sont des mecanismes independants, un tracking peut
 *     cumuler plusieurs montants a la fois, ex. tracking 873870075714 : 55EUR "dépassement de
 *     dimension" + 18EUR "manutention supplémentaire :poids" sur le MEME bordereau = 73EUR)
 *   - "Plusieurs pièces" / "Demand Surcharge" -> Plus-value BtoC
 * Chaque bordereau (1 par tracking) est delimite par son debut ("<total>\n<date>DD/MM/YYYY
 * <tracking>") jusqu'au bordereau suivant -- le montant precede son libelle, parfois avec 1
 * ligne de texte intermediaire (ex. "3,00\nExpéditeurDestinataire\nPlusieurs pièces"). */
function extractSupplements(text) {
  const trackPat = /([\d.,]+)\s*\n(\d{2}\/\d{2}\/\d{4})(\d{9,15})/g;
  const starts = [];
  let m;
  while ((m = trackPat.exec(text)) !== null) starts.push({ idx: m.index, tracking: m[3] });
  const kwPats = {
    // Zones eloignees : mot-cle generique "zone" (insensible a la casse, demande pole
    // transport 2026-08-24) + libelle precis "Frais de traitement des importations aux
    // Etats-Unis" (precise par le pole transport 2026-08-25 -- ne contient pas "zone", donc
    // pas capte par le 1er pattern, doit rester en complement). BUG TROUVE 2026-08-25 (verifie
    // sur bordereaux reels 382198649394/382408298943, PDF FEDEX_634373761_781542172.pdf) : le
    // PDF tronque le libelle a "...aux États-Uni" (SANS le "s" final) -- "s" rendu optionnel.
    ZonesEloignees: [
      /([\d.,]+)\s*\n[^\n]{0,60}\n?[^\n]*zone[^\n]*/i,
      /([\d.,]+)\s*\n[^\n]{0,60}\n?Frais de traitement des importations aux [ÉE]tats-Unis?/i,
    ],
    // Colis volumineux (PDF) : "Charge pour dépassement de dimension" + "Supplément pour
    // manutention supplémentaire :poids" (precise par le pole transport 2026-08-25, verifie
    // sur bordereau reel 873870075714, PDF FEDEX_634374150_200720433.pdf -- 55EUR dimension +
    // 18EUR poids sur le MEME bordereau, 2 lignes distinctes qui s'additionnent).
    ColisVolumineuxPdf: [
      /([\d.,]+)\s*\n[^\n]{0,60}\n?Charge pour d[ée]passement de dimension/,
      /([\d.,]+)\s*\n[^\n]{0,60}\n?Suppl[ée]ment pour manutention suppl[ée]mentaire\s*:?\s*poids/i,
    ],
    PlusValueB2C: [
      /([\d.,]+)\s*\n[^\n]{0,60}\n?Plusieurs pi[èe]ces/,
      /([\d.,]+)\s*\n[^\n]{0,60}\n?Demand Surcharge/,
    ],
  };
  const result = [];
  for (let i = 0; i < starts.length; i++) {
    const nextIdx = starts[i + 1] ? starts[i + 1].idx : text.length;
    const block = text.slice(starts[i].idx, nextIdx);
    for (const [poste, pats] of Object.entries(kwPats)) {
      for (const pat of pats) {
        const mm = pat.exec(block);
        if (mm) result.push({ tracking: starts[i].tracking, poste, montant: num(mm[1].replace(/\./g, '').replace(',', '.')) });
      }
    }
  }
  return result;
}

async function extractFedexPdfInfo(pdfPath) {
  const buf = require('fs').readFileSync(pdfPath);
  const { text } = await pdfParse(buf);
  const mInv = /No de Client\s*:\s*\nNo de Facture\s*:\s*\nDate de la facture\s*:\s*\nDate d.[eé]ch[eé]ance\s*\nMontant d[uû]\s*\n\*+\d+\s*\n(\d+)\s*\n/.exec(text);
  const mTotal = /Total d[uû]\s*EUR\s*([\d.,]+)/.exec(text);
  return {
    file: path.basename(pdfPath),
    numeroFacture: mInv ? mInv[1] : null,
    totalDu: mTotal ? num(mTotal[1].replace(/\./g, '').replace(',', '.')) : null,
    bordereaux: extractBordereaux(text),
    supplements: extractSupplements(text),
  };
}

async function process(files) {
  const csvPaths = files.csv || [];
  if (!csvPaths.length) throw new Error('Aucun fichier fourni (attendu : Shipment detail FedEx, CSV export portail).');

  const warnings = [];
  const infos = [];
  const brutes = csvPaths.map(readFedexCsv);

  // PDF lus TOT (contenu, pas le nom de fichier -- multer renomme les fichiers uploades en
  // identifiants aleatoires sans extension/nom d'origine, cf. server.js "upload.any()" ->
  // f.path, PAS f.originalname -- bug constate 2026-08-20 : le nom "FEDEX_<invoice>_<compte>.pdf"
  // n'existe QUE sur le disque de l'utilisateur avant upload, jamais recu tel quel par le
  // carrier). Reutilise plus bas pour la reconciliation "Total dû" (pas de 2e lecture).
  const pdfPaths = files.pdf || [];
  const pdfs = [];
  const pdfInvoiceNums = new Set();
  for (const p of pdfPaths) {
    try {
      const r = await extractFedexPdfInfo(p);
      if (r.totalDu != null) pdfs.push(r);
      else warnings.push(`PDF ${path.basename(p)} : "Total dû" introuvable, ignoré pour la réconciliation.`);
      if (r.numeroFacture) pdfInvoiceNums.add(r.numeroFacture);
      else warnings.push(`PDF ${path.basename(p)} : numéro de facture introuvable dans le contenu — ignoré pour le contrôle "facture sans PDF".`);
    } catch (e) {
      warnings.push(`PDF ${path.basename(p)} : lecture impossible (${e.message}).`);
    }
  }

  const header = brutes[0].header;
  const iCompte = colIndexByName(header, 'Compte du payeur');
  const iService = colIndexByName(header, 'Description du service');
  const iDateEnvoi = colIndexByName(header, "Date d'envoi (mm/jj/aaaa)");
  const iTracking = colIndexByName(header, "Numéro de suivi de l'envoi");
  // BUG TROUVE 2026-08-19 (finaliseur Python, ecart Fret 52472,41EUR vs 45515,83EUR reel) :
  // les colonnes "... en USD" du CSV brut NE SONT PAS celles utilisees par le classeur reel
  // -- le vrai "Montant" (Shipment Detail!A=BK+BN) reference les colonnes EUR "Devise de
  // facturation des frais de transport/de la remise" (noms trompeurs : ce sont des MONTANTS
  // EUR, pas des codes devise -- le vrai code devise est dans la colonne suivante "Code de la
  // devise de facturation"). Reconfirme sur un cas reel (tracking 381547998902) : USD
  // 539.82-458.09=81.73 vs EUR (colonnes ci-dessous) 465.03-394.62=70.41 = valeur reelle du
  // classeur modele.
  const iFret = colIndexByName(header, "Devise de facturation des frais de transport de l'envoi");
  const iDivers = colIndexByName(header, "Devise de facturation des frais divers de l'envoi");
  const iDroitsTaxes = colIndexByName(header, "Devise de facturation des droits de douane et taxes de l'envoi");
  const iRemise = colIndexByName(header, "Devise de facturation de la remise de l'envoi");
  const iColis = colIndexByName(header, "Colis dans l'envoi");
  const iPoids = colIndexByName(header, "Poids nominal de l'envoi (livres)");
  const iPaysDest = colIndexByName(header, 'Pays/Territoire du destinataire');
  const iLongueur = colIndexByName(header, 'Longueur volumétrique (cm)');
  const iInvoiceNum = colIndexByName(header, 'Numéro de facture');
  const iInvoiceDate = colIndexByName(header, 'Date de facturation (mm/jj/aaaa)');
  const iMoisFacturation = colIndexByName(header, 'Mois de facturation (aaaamm)');

  const requis = { iTracking, iFret, iRemise, iInvoiceNum, iPaysDest };
  const manquantes = Object.entries(requis).filter(([, v]) => v < 0).map(([k]) => k);
  if (manquantes.length) {
    throw new Error(`Colonne(s) attendue(s) introuvable(s) dans le CSV FedEx : ${manquantes.join(', ')} — vérifier l'en-tête du fichier reçu.`);
  }

  // Controle "bordereau facture par FedEx mais absent du CSV" (2026-08-20, cas reel juillet :
  // facture 634393398, ecart 83,81EUR -- un envoi de NOVEMBRE 2025 apparaissait dans le PDF
  // mais n'a jamais ete reporte dans le CSV Shipment Detail du mois). Compare TOUS les
  // trackings du PDF (bordereau par bordereau) a TOUS les trackings presents dans le CSV brut
  // (peu importe si ensuite exclus par une regle -- Droits&Taxes/sans PDF -- le but est de
  // detecter un envoi JAMAIS VU DU TOUT dans le brut, pas un envoi filtre a posteriori).
  const csvTrackings = new Set();
  for (const f of brutes) {
    for (const r of f.rows) {
      const t = String(r[iTracking] || '').trim();
      if (t) csvTrackings.add(t);
    }
  }
  for (const p of pdfs) {
    const manquants = (p.bordereaux || []).filter((b) => !csvTrackings.has(b.tracking));
    if (manquants.length) {
      const totalManquant = round2(manquants.reduce((s, b) => s + (b.sousTotal || 0), 0));
      const detail = manquants.map((b) => `${b.tracking} (envoi du ${b.dateEnvoi}, ${b.sousTotal.toFixed(2)} EUR)`).join(', ');
      warnings.push(`PDF ${p.file}${p.numeroFacture ? ` (n° ${p.numeroFacture})` : ''} : ${manquants.length} bordereau(x) facturé(s) par FedEx mais absent(s) du CSV Shipment Detail — ${detail} — total HT manquant ≈ ${totalManquant.toFixed(2)} EUR (probable envoi d'un mois antérieur jamais reporté, à vérifier avec le pôle transport).`);
    }
  }

  // Mois cible = mois majoritaire (colonne "Mois de facturation (aaaamm)" si presente, sinon
  // "Date de facturation" -- meme piege deja rencontre sur Chronopost/TNT : quelques lignes
  // residuelles d'un autre mois peuvent se glisser dans l'export).
  const comptageMois = new Map();
  for (const f of brutes) {
    for (const r of f.rows) {
      let mois = iMoisFacturation >= 0 ? String(r[iMoisFacturation] || '').trim() : '';
      if (!/^\d{6}$/.test(mois) && iInvoiceDate >= 0) {
        const d = parseDateUs(r[iInvoiceDate]);
        if (d) mois = `${d.y}${String(d.m).padStart(2, '0')}`;
      }
      if (/^\d{6}$/.test(mois)) comptageMois.set(mois, (comptageMois.get(mois) || 0) + 1);
    }
  }
  let moisCible = null;
  let dateValidite = '';
  if (comptageMois.size) {
    [moisCible] = [...comptageMois.entries()].sort((a, b) => b[1] - a[1])[0];
    const m = /^(\d{4})(\d{2})$/.exec(moisCible);
    if (m) dateValidite = `01/${m[2]}/${m[1]}`;
    if (comptageMois.size > 1) {
      const detail = [...comptageMois.entries()].sort((a, b) => b[1] - a[1]).map(([k, v]) => `${k}: ${v} ligne(s)`).join(', ');
      infos.push(`Plusieurs mois détectés dans le CSV reçu (${detail}) — mois retenu : ${moisCible} (majoritaire).`);
    }
  }

  // Regroupement par facture (Numero de facture) : detection Droits&Taxes pure (exclue) +
  // detection "facture d'un autre mois" (alerte seulement, cf. docstring en tete).
  const parFacture = new Map(); // numFacture -> { fretTotal, taxeTotal, lignes: [], moisEnvoiCount: Map }
  for (const f of brutes) {
    for (const r of f.rows) {
      const numFacture = String(r[iInvoiceNum] || '').trim();
      if (!numFacture) continue;
      if (!parFacture.has(numFacture)) parFacture.set(numFacture, { fretTotal: 0, taxeTotal: 0, lignes: [], moisEnvoi: new Map() });
      const fac = parFacture.get(numFacture);
      const fret = iFret >= 0 ? num(r[iFret]) : 0;
      const taxe = iDroitsTaxes >= 0 ? num(r[iDroitsTaxes]) : 0;
      fac.fretTotal = round2(fac.fretTotal + fret);
      fac.taxeTotal = round2(fac.taxeTotal + taxe);
      fac.lignes.push(r);
      const dEnvoi = iDateEnvoi >= 0 ? parseDateUs(r[iDateEnvoi]) : null;
      if (dEnvoi) {
        const key = `${dEnvoi.y}${String(dEnvoi.m).padStart(2, '0')}`;
        fac.moisEnvoi.set(key, (fac.moisEnvoi.get(key) || 0) + 1);
      }
    }
  }

  // ATTENTION — question ouverte non tranchee (2026-08-19, cf. memoire fedex_carrier_construit.md) :
  // sur juin 2026 reel, le fichier livre au client (2026_06_Fedex_Import.csv, 3607 lignes,
  // Fret total 45515,83EUR) EXCLUT 3 factures du CSV brut (120792009 Droits&Taxes pur,
  // 634313590 392 lignes 100% mai, 634364158 1 ligne datee juillet) -- mais AUCUNE regle
  // simple (ni "100% hors mois cible", ni "majorite hors mois cible") ne reproduit exactement
  // cette exclusion : 634333315 (100% mai, 1 ligne) EST incluse dans le fichier reel, alors
  // que 634364158 (ligne datee juin) est exclue. Inclure TOUT (sauf Droits&Taxes) surestime le
  // Fret de +12394,91EUR (~27%) par rapport a juin. Repli actuel : n'exclure QUE les factures
  // 100% Droits&Taxes (seule regle certaine, confirmee par l'utilisateur) + ALERTER sur toute
  // facture partiellement/totalement hors mois cible SANS l'exclure -- le pole transport doit
  // confirmer facture par facture avant validation finale (question en attente).
  // Facture absente des PDF portail fournis (regle utilisateur 2026-08-20) : si au moins un
  // PDF est fourni, un numero de facture du CSV brut qui n'apparait dans AUCUN nom de fichier
  // PDF fourni est tres probablement deja facture un mois precedent (cf. question ouverte
  // ci-dessus, confirme sur le cas reel 634313590 de juin 2026 : absente des 9 PDF fournis,
  // absente aussi du fichier livre reel) -- exclue automatiquement, comme la regle Droits&Taxes.
  // Ne s'applique QUE si des PDF sont fournis (sinon rien a comparer, cf. commentaire en tete).
  const facturesDroitsTaxes = [];
  const facturesSansPdf = [];
  const lignesExclues = new Set();
  for (const [numFacture, fac] of parFacture) {
    if (fac.fretTotal === 0 && fac.taxeTotal !== 0) {
      facturesDroitsTaxes.push(numFacture);
      for (const r of fac.lignes) lignesExclues.add(r);
    } else if (pdfInvoiceNums.size && !pdfInvoiceNums.has(numFacture)) {
      facturesSansPdf.push(numFacture);
      for (const r of fac.lignes) lignesExclues.add(r);
    } else if (!pdfInvoiceNums.has(numFacture) && moisCible && fac.moisEnvoi.size && !fac.moisEnvoi.has(moisCible)) {
      // BUG TROUVE 2026-08-20 (facture reelle 634364158 de juillet 2026, PDF fourni existant) :
      // cette alerte ne doit se declencher QUE si la facture n'a PAS de PDF fourni -- la
      // presence d'un PDF reel est une preuve plus forte que la simple date d'expedition des
      // lignes (elle prouve que FedEx a bien facture cette expedition CE MOIS-CI, quelle que
      // soit la date d'envoi d'origine) -- confirme utilisateur 2026-08-20 "son pdf existe et
      // donc pas encore facture" (comprendre : deja bel et bien facture ce mois-ci puisque son
      // PDF existe, pas un cas a signaler). Avant ce fix, la condition ne verifiait PAS
      // pdfInvoiceNums, seulement le "else if" (donc l'absence dans facturesSansPdf, qui ne
      // suffit pas a prouver la presence reelle du PDF si pdfInvoiceNums est vide).
      const detail = [...fac.moisEnvoi.entries()].sort((a, b) => b[1] - a[1]).map(([k, v]) => `${k}: ${v}`).join(', ');
      warnings.push(`Facture ${numFacture} : AUCUNE ligne datée dans le mois cible ${moisCible} (dates réelles : ${detail}) — probable facture déjà traitée un mois précédent/suivant, à confirmer avec le pôle transport avant facturation (non exclue automatiquement, cf. commentaire en tête de fichier).`);
    }
  }
  if (facturesDroitsTaxes.length) {
    infos.push(`Facture(s) 100% Droits & Taxes (aucun fret, ajustement douanier pur) exclue(s) du fichier livré : ${facturesDroitsTaxes.join(', ')} (${facturesDroitsTaxes.reduce((s, nf) => s + parFacture.get(nf).lignes.length, 0)} ligne(s)).`);
  }
  if (facturesSansPdf.length) {
    infos.push(`Facture(s) absente(s) des PDF portail fournis (probablement déjà facturée(s) un mois précédent) exclue(s) du fichier livré : ${facturesSansPdf.join(', ')} (${facturesSansPdf.reduce((s, nf) => s + parFacture.get(nf).lignes.length, 0)} ligne(s)) — à confirmer avec le pôle transport.`);
  }

  // Export brut WMS (E/P) : m et m-1, meme mecanisme que Geodis/DPD/Lettres (core/exportBrut.js).
  const appRoot = path.resolve(__dirname, '../../..');
  const brutPaths = moisCible ? findBrutFiles(moisCible.length === 6 ? `${moisCible.slice(0, 4)}_${moisCible.slice(4)}` : moisCible, appRoot) : [];
  const epMap = brutPaths.length ? epParTrackingFromExport(readBrutRows(brutPaths)) : new Map();
  if (!brutPaths.length) warnings.push("Export WMS 'expéditions_brut' introuvable pour ce mois (E/P) — toutes les lignes sans correspondance seront classées 'P' par défaut.");

  // Supplements "surplus" (Zones eloignees/Colis volumineux PDF/Plus-value BtoC) : map par
  // tracking, cumulee sur tous les PDF fournis (cf. extractSupplements). Colis volumineux
  // PDF s'ADDITIONNE a la regle Excel existante (longueur>60cm=10EUR, confirme pole transport
  // 2026-08-20) -- pas un remplacement.
  const supplementsParTracking = new Map(); // tracking -> { ZonesEloignees, ColisVolumineuxPdf, PlusValueB2C }
  for (const p of pdfs) {
    for (const s of p.supplements || []) {
      if (!supplementsParTracking.has(s.tracking)) supplementsParTracking.set(s.tracking, { ZonesEloignees: 0, ColisVolumineuxPdf: 0, PlusValueB2C: 0 });
      const acc = supplementsParTracking.get(s.tracking);
      acc[s.poste] = round2(acc[s.poste] + s.montant);
    }
  }
  if (supplementsParTracking.size) {
    const totaux = { ZonesEloignees: 0, ColisVolumineuxPdf: 0, PlusValueB2C: 0 };
    for (const acc of supplementsParTracking.values()) {
      totaux.ZonesEloignees = round2(totaux.ZonesEloignees + acc.ZonesEloignees);
      totaux.ColisVolumineuxPdf = round2(totaux.ColisVolumineuxPdf + acc.ColisVolumineuxPdf);
      totaux.PlusValueB2C = round2(totaux.PlusValueB2C + acc.PlusValueB2C);
    }
    infos.push(`Suppléments extraits des PDF (mots-clés, cf. mail pôle transport 2026-08-20) : Zones éloignées ${totaux.ZonesEloignees.toFixed(2)} EUR, Colis volumineux (PDF, s'ajoute aux 10 EUR calculés) ${totaux.ColisVolumineuxPdf.toFixed(2)} EUR, Plus-value BtoC ${totaux.PlusValueB2C.toFixed(2)} EUR — sur ${supplementsParTracking.size} tracking(s).`);
  }

  const recs = [];
  let nEpDefaut = 0;
  for (const f of brutes) {
    for (const r of f.rows) {
      if (lignesExclues.has(r)) continue;
      const tracking = String(r[iTracking] || '').trim();
      if (!tracking) continue;
      const fret = iFret >= 0 ? num(r[iFret]) : 0;
      const divers = iDivers >= 0 ? num(r[iDivers]) : 0;
      const remise = iRemise >= 0 ? num(r[iRemise]) : 0;
      const droitsTaxes = iDroitsTaxes >= 0 ? num(r[iDroitsTaxes]) : 0;
      const montant = round2(fret + remise); // Shipment Detail!A = BK+BN (Fret + Remise)
      const pays = iPaysDest >= 0 ? String(r[iPaysDest] || '').trim().toUpperCase() : '';
      const service = iService >= 0 ? String(r[iService] || '').trim() : '';
      const mode = service === cfg.service_st ? 'ST' : 'FICP';
      const longueur = iLongueur >= 0 ? num(r[iLongueur]) : 0;
      // Colis volumineux = regle Excel (longueur>60cm=10EUR) + supplement PDF "Charge pour
      // dépassement de dimension" (s'ADDITIONNENT, 2 mecanismes independants confirmes par le
      // pole transport 2026-08-20 -- cf. commentaire supplementsParTracking plus haut).
      const supp = supplementsParTracking.get(tracking);
      const colisVolumineux = round2((longueur > cfg.colis_volumineux_seuil_cm ? cfg.colis_volumineux_montant : 0) + (supp ? supp.ColisVolumineuxPdf : 0));
      let ep = epMap.get(tracking);
      if (!ep) { ep = 'P'; nEpDefaut++; }
      const zoneFicp = cfg.zoning[pays];
      const zone = zoneFicp === undefined ? 'inconnu' : (zoneFicp || '');
      // Import ERP!M formule modele : =IF(J="France","ST",XLOOKUP(tracking,...,'Shipment Detail'!E:E))
      // -- J=Zone (PAS Pays) : Zone='France' UNIQUEMENT pour FR (cf. config.zoning.FR="France").
      const modeFinal = zone === 'France' ? 'ST' : mode;
      recs.push({
        tracking, compte: iCompte >= 0 ? String(r[iCompte] || '').trim() : '',
        montant, poids: iPoids >= 0 ? num(r[iPoids]) : 0,
        colis: iColis >= 0 ? num(r[iColis]) : 1,
        pays, zone, mode: modeFinal, colisVolumineux, ep,
        droitsTaxes: droitsTaxes || 0,
        divers,
        // Zones eloignees / plus-value BtoC : UNIQUEMENT via extraction PDF (jamais dans le
        // CSV, cf. mail pole transport 2026-08-20) -- 0 si aucun PDF fourni ou rien trouve.
        zonesEloignees: supp ? supp.ZonesEloignees : 0,
        plusValueB2C: supp ? supp.PlusValueB2C : 0,
        tva: cfg.tva_pays_ue.includes(pays) ? cfg.champs_fixes.tva_taux : 0,
      });
    }
  }
  if (nEpDefaut) infos.push(`${nEpDefaut} ligne(s) sans correspondance dans l'export WMS (E/P) — classée(s) 'P' par défaut.`);

  // Poids livres FedEx -> kg (ARRONDI.SUP, formule modele TCD!E = ROUNDUP(poids*0.453592,1)).
  const LB_TO_KG = 0.453592;
  const importRows = recs.map((rec) => ({
    // Import ERP!A formule modele : =IF(OR(M="FICP",M="IE",M="RE",M="International",M="Europe"),
    // "FEDEX INTERNATIONAL","FEDEX FRANCE") -- M='Mode d'envoi' (deja calcule ci-dessus, rec.mode).
    Transporteur: (rec.mode === 'FICP') ? 'FEDEX INTERNATIONAL' : 'FEDEX FRANCE',
    DateValidite: dateValidite || '',
    Ref1: '', Ref2: '', IdClient: '',
    Tracking: rec.tracking, Nom: '',
    EP: rec.ep, Pays: rec.pays, Zone: rec.zone,
    NbrColis: rec.colis, Poids: roundUp1(rec.poids * LB_TO_KG),
    Mode: rec.mode, TVA: rec.tva,
    DroitsTaxes: rec.droitsTaxes || 0, Assurance: 0,
    ZonesEloignees: rec.zonesEloignees || 0, ColisVolumineux: rec.colisVolumineux,
    Adresses: 0, Fret: rec.montant, PlusValueB2C: rec.plusValueB2C || 0, TaxeGasoil: '', NbColis: '',
  }));

  const { alerts, infos: valInfos } = validate(importRows);
  infos.push(...valInfos);

  // Reconciliation par facture (Bilan factures : Total TTC compare au PDF, PAS le Fret HT).
  const totalFretHt = round2(importRows.reduce((s, r) => s + (r.Fret || 0), 0));
  infos.push(`Total Fret HT (toutes lignes) = ${totalFretHt.toFixed(2)} EUR.`);
  for (const p of pdfs) {
    infos.push(`Facture PDF ${p.file}${p.numeroFacture ? ` (n° ${p.numeroFacture})` : ''} : Total dû = ${p.totalDu.toFixed(2)} EUR.`);
  }

  const controle = {
    'Colis volumineux': round2(importRows.reduce((s, r) => s + (r.ColisVolumineux || 0), 0)),
    'Frêt': totalFretHt,
    'Droits et taxes': round2(importRows.reduce((s, r) => s + (r.DroitsTaxes || 0), 0)),
    'Zones éloignées': round2(importRows.reduce((s, r) => s + (r.ZonesEloignees || 0), 0)),
    'plus-value BtoC': round2(importRows.reduce((s, r) => s + (r.PlusValueB2C || 0), 0)),
  };

  return {
    header,
    rows: recs, recs, importRows, controle, warnings, alerts, infos,
    posteKeys: ['Colis volumineux', 'Frêt', 'Droits et taxes', 'Zones éloignées', 'plus-value BtoC'], cfg, pdfs,
    sheetNames: { raw: 'Shipment Detail', import: 'Import ERP' },
    period: moisCible ? `${moisCible.slice(0, 4)}_${moisCible.slice(4)}` : 'export',
  };
}

/** Args du finaliseur (csv + brut du mois/mois-1 + pdf) -- meme pattern que Delivengo
 * (computeFinalizerArgs), le finaliseur Python recalcule lui-meme E/P via l'export WMS
 * plutot que de dependre d'un CSV intermediaire (cf. finaliser_fedex.py). */
function computeFinalizerArgs(files, period, appRoot) {
  const brut = period ? findBrutFiles(period, appRoot) : [];
  return ['--csv', ...(files.csv || []), '--brut', ...brut, '--pdf', ...(files.pdf || [])];
}

module.exports = {
  id: 'fedex',
  name: 'FedEx',
  status: 'ready',
  taxeGasoil: "Aucune (pas de surcharge carburant repercutee au client dans le fichier livre -- colonne Gazole/'V' de Import ERP toujours vide, confirme sur juin 2026).",
  method: "1 CSV brut par mois (export portail FedEx, en-tetes FR, colonnes resolues par nom). Montant = Frais de transport + Remise. Mode d'envoi = 'ST' si Service='FedEx Priority' sinon 'FICP'. Zone = lookup pays -> Zoning!FICP. TVA = 20% si pays destinataire dans la liste UE (config), sinon 0%. Colis volumineux = 10EUR forfait si longueur dimensionnee > 60cm. E/P via export WMS brut m/m-1 (repli 'P'). Factures 100% Droits&Taxes exclues. Reconciliation PDF sur 'Total du'.",
  inputs: [
    { key: 'csv', label: 'Shipment detail (CSV export portail FedEx)', accept: '.csv', multiple: true, required: true },
    { key: 'pdf', label: 'Facture(s) PDF FedEx (contrôle du "Total dû")', accept: '.pdf', multiple: true, required: false },
  ],
  outputNaming: { workbook: '{period}_Facture Fedex', import: '{period}_Fedex_Import' },
  // Le fichier import CSV/XLSX est reconstruit depuis les valeurs REELLEMENT calculees par
  // Excel dans l'onglet "Import ERP" du classeur genere (au lieu des importRows calcules a
  // part en JS) -- remontee pole transport 2026-08-24, cf. server.js et finaliser_fedex.py.
  importFromWorkbook: true,
  finalizer: {
    script: '../automatisation/finaliser_fedex.py',
    template: '../Transporteurs/Fedex/2026_06_Facture Fedex.xlsx',
    buildArgs: (files, period, appRoot) => computeFinalizerArgs(files, period, appRoot),
  },
  process,
};
