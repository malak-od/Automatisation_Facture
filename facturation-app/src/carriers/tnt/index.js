// ============================================================================
//  Adaptateur transporteur : TNT (facture FedEx Express FR SAS "services TNT")
//  1 fichier Excel brut par mois (feuille "DET_FAC", 1 ligne = 1 evenement de
//  facturation, PAS 1 ligne = 1 colis) : plusieurs types de ligne distingues
//  par "Ligne Facture" (colonne F) :
//   - "COL"  : 1 ligne par colis (Montant HT toujours 0 -- porte le POIDS reel)
//   - "ENV"  : 1 ligne par ENVOI (id colis vide, Montant HT = vrai cout de
//     transport de CET envoi -- un envoi peut contenir PLUSIEURS colis/COL)
//   - "CPC"/"CPE" : supplements (code complement = RS/TMC/ZEL/LS/FACTOR_FOR/...)
//   - "EXP"  : 1 ligne par compte expediteur, code complement = SURCH_CARB
//     (surcharge carburant, montant GLOBAL du mois, PAS par tracking/envoi)
//
//  Mecanisme de rattachement tracking <-> envoi (reproduit du classeur modele,
//  onglet "Recherche tracking" = TCD sur numero envoi + id colis, puis
//  'Facture TNT'!C = RIGHT(LOOKUP(numeroEnvoi, ..., idColis), 16)) : cette
//  formule ne depend QUE du "numero envoi" de la ligne, JAMAIS de son propre
//  "id colis" -- donc TOUTES les lignes d'un meme envoi (la ligne "ENV" ET
//  chaque ligne COL/CPC/CPE de CHAQUE colis de cet envoi) resolvent vers LE
//  MEME tracking (Excel LOOKUP sur cle dupliquee renvoie le DERNIER match
//  d'une plage triee -> le tracking ALPHABETIQUEMENT LE PLUS GRAND de
//  l'envoi). Pour un envoi de N colis, TOUS les montants (ENV + RS/TMC/etc.
//  de CHAQUE colis) sont donc concentres sur ce seul tracking -- les N-1
//  autres colis de l'envoi n'apparaissent PAS DU TOUT dans le fichier livre,
//  meme si le PDF/releve d'operations, lui, detaille chaque colis separement
//  (N° BT distinct par destinataire). REPRODUIT FIDELEMENT ici (decision
//  utilisateur explicite 2026-08-18, "faites comme le fichier de base fait"
//  -- pas une correction/repartition de notre cru, meme si ca semble
//  contre-intuitif face au PDF).
//
//  Taxe Gasoil (SURCH_CARB) : montant GLOBAL par compte expediteur (PAS
//  ventile par tracking dans le brut, contrairement a Chronopost dont le pool
//  gazole est reparti au prorata du fret) -- jamais repercute au client dans
//  le fichier livre (meme principe que Chronopost). Sert UNIQUEMENT a la
//  reconciliation PDF : somme(colonnes O->U, tous les postes ERP en euros
//  hors Gazole) + somme(surcharges carburant PDF, "taux officiel") doit
//  egaler le "TOTAL GENERAL" HT de la facture PDF (regle confirmee
//  utilisateur 2026-08-18, verifiee empiriquement sur juin 2026 : 3634,30€ de
//  sous-totaux journaliers + 422,52€ + 1,92€ de surcharge = 4058,74€ TOTAL
//  GENERAL). Le PDF affiche 2 taux ("taux officiel" ET "taux reel") -- SEUL
//  le taux officiel fait foi (consigne utilisateur explicite), le taux reel
//  n'est qu'une info affichee par TNT, jamais utilise ici.
//
//  Tracking (id colis) : TOUJOURS du texte pur (16 chiffres, prefixe "'" dans
//  le brut Excel) -- NE JAMAIS convertir en nombre a aucune etape (perte de
//  precision/zeros de tete au-dela de 15 chiffres significatifs -> le
//  tracking ne correspond plus a rien dans l'ERP, consigne utilisateur
//  explicite 2026-08-18).
// ============================================================================
const XLSX = require('xlsx');
const path = require('path');
const pdfParse = require('pdf-parse');
const { num, round2, roundUp1 } = require('../../core/csv');
const { validate } = require('../../core/validate');
const cfg = require('./config.json');

const POSTE_KEYS = ['Assurance', 'Colis volumineux', 'Frêt', 'Gazole', 'Zones éloignées'];

function normKey(s) {
  return String(s || '')
    .normalize('NFD').replace(/[̀-ͯ]/g, '')
    .toLowerCase().replace(/[.,]/g, '').replace(/\s+/g, ' ').trim();
}

function colIndexByName(header, name) {
  const target = normKey(name);
  return header.findIndex((h) => normKey(h) === target);
}

function isXlsx(p) {
  const fs = require('fs');
  try {
    const buf = fs.readFileSync(p, { encoding: null, flag: 'r' });
    return buf.length > 4 && buf[0] === 0x50 && buf[1] === 0x4b; // "PK"
  } catch (e) {
    return false;
  }
}

/** Tracking (id colis) : retire le prefixe "'" (texte force Excel) et les espaces --
 * TOUJOURS retourne une chaine, jamais convertie en nombre (cf. docstring en tete). */
function cleanTracking(raw) {
  return String(raw || '').replace(/^'/, '').trim();
}

/** Nombre au format FR avec milliers en point (ex. "4.058,74") -- num() du core ne retire
 * QUE les espaces, un point de milliers le ferait tronquer par parseFloat (4.058,74 ->
 * "4.058.74" -> 4.058 apres le 1er remplacement de virgule). Retire les points AVANT
 * d'appeler num() -- specifique au format d'affichage PDF TNT (colis/poids n'ont jamais ce
 * probleme dans le brut Excel, deja des nombres natifs). */
function numMilliers(x) {
  return num(String(x || '').replace(/\./g, ''));
}

/** Lit le fichier brut TNT : feuille "DET_FAC" (ou 1re feuille), en-tete en ligne 1. */
function readBrut(p) {
  const wb = XLSX.readFile(p, { cellDates: false, raw: true });
  const sheetName = wb.SheetNames.includes('DET_FAC') ? 'DET_FAC' : wb.SheetNames[0];
  const ws = wb.Sheets[sheetName];
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: '' });
  const header = (rows[0] || []).map((h) => String(h || '').trim());
  // FACMAG passe par Excel (constate le 2026-11-10 : CSV ouvert puis enregistre en .xlsx) : les
  // numeros de suivi et d'envoi y sont arrondis par Excel (7520240000000000 au lieu de
  // 7520236028138204), donc inutilisables -- message clair plutot qu'un "0 ligne" silencieux.
  if (colIndexByName(header, 'Bon de transport') >= 0) {
    throw new Error('Fichier FACMAG enregistré au format Excel : les numéros de suivi y sont arrondis par Excel et ne sont plus utilisables. Déposez le fichier FACMAG.CSV tel que téléchargé chez TNT (sans l\'ouvrir dans Excel), ou le détail facture Excel (feuille DET_FAC).');
  }
  const data = rows.slice(1).filter((r) => r.some((v) => v !== ''));
  return { file: path.basename(p), header, rows: data };
}

// ---------------------------------------------------------------------------------------------
// FACMAG (export CSV TNT, decision utilisateur 2026-11-10) : autre format de la meme facture --
// 1 ligne par colis ou par envoi, frais en COLONNES (SFU, TMC, FACTOR_FOR, CNM...), au lieu du
// "detail facture" (feuille DET_FAC, 1 ligne par evenement ENV/COL/CPC/CPE/EXP). Converti ici
// au format DET_FAC, ecrit dans un .xlsx a cote du fichier recu : le reste du carrier ET le
// finaliseur Excel (via files.facture, cf. process) le lisent comme un detail facture, sans autre
// changement. Correspondance (verifiee sur octobre 2026 : total 1 690,64 EUR = TOTAL GENERAL PDF) :
//  - ligne SANS "Bon de transport" (envoi de plusieurs colis) ou colis unique ("Nb colis dans
//    l'envoi" = 1) -> ligne ENV (montant = "Prix h.t. du transport", poids de la ligne)
//  - ligne AVEC "Bon de transport" -> ligne COL (id colis = bon de transport, poids du colis)
//  - chaque colonne de frais non nulle -> ligne CPC (colis) / CPE (envoi), code complement = nom
//    de la colonne ; SAUF "SFU" (surcharge carburant) -> cumulee en 1 ligne EXP/SURCH_CARB par
//    compte, comme dans le detail facture (jamais repercutee au client)
// Champs texte au format du detail facture recu en septembre 2026 (apostrophe de tete sur
// compte/id colis/departement/code postal, numero envoi prefixe "E").
// ---------------------------------------------------------------------------------------------
const ENTETE_DET_FAC = ['Mois de Facturation', 'Indicatif Facture', 'Numero de Facture', 'Compte Expediteur', 'Client Facturé',
  'Ligne Facture', 'id colis', 'date Expedition', 'numero envoi', 'code Produit', 'poids colis', 'code complement', 'montant HT',
  'raison sociale expéditeur', 'adresse 1', 'adresse 2', 'pays', 'departement', 'code postal', 'ville',
  'raison sociale destinataire', 'adresse 1', 'adresse 2', 'pays', 'departement', 'code postal', 'ville'];
const CODE_CARBURANT_FACMAG = 'SFU';

function premiereLigneTexte(p) {
  const fs = require('fs');
  const fd = fs.openSync(p, 'r');
  try {
    const buf = Buffer.alloc(4096);
    const n = fs.readSync(fd, buf, 0, buf.length, 0);
    return buf.subarray(0, n).toString('latin1').split(/\r?\n/)[0];
  } finally {
    fs.closeSync(fd);
  }
}

function estFacmagCsv(p) {
  if (isXlsx(p)) return false;
  const l = premiereLigneTexte(p);
  return /Bon de transport/i.test(l) && /envoi/i.test(l);
}

/** CSV a guillemets (les montants FACMAG "00000001,00" contiennent le separateur virgule). */
function lireCsvGuillemets(texte, sep) {
  const lignes = [];
  let ligne = [];
  let champ = '';
  let entre = false;
  for (let i = 0; i < texte.length; i++) {
    const c = texte[i];
    if (entre) {
      if (c !== '"') champ += c;
      else if (texte[i + 1] === '"') { champ += '"'; i++; } else entre = false;
    } else if (c === '"') entre = true;
    else if (c === sep) { ligne.push(champ); champ = ''; } else if (c === '\r' || c === '\n') {
      if (c === '\r' && texte[i + 1] === '\n') i++;
      ligne.push(champ); lignes.push(ligne); ligne = []; champ = '';
    } else champ += c;
  }
  if (champ !== '' || ligne.length) { ligne.push(champ); lignes.push(ligne); }
  return lignes;
}

function convertirFacmag(p, warnings, infos) {
  const texte = require('fs').readFileSync(p).toString('latin1');
  const l0 = texte.split(/\r?\n/)[0];
  const sep = (l0.match(/;/g) || []).length > (l0.match(/,/g) || []).length ? ';' : ',';
  const lignes = lireCsvGuillemets(texte, sep);
  const h = lignes[0].map((x) => String(x || '').trim());
  const col = (nom) => colIndexByName(h, nom);
  const i = {
    mois: col('Année mois'), facture: col('N° facture'), compte: col('N° compte expéditeur'), date: col('Date expédition colis'),
    envoi: col('N° envoi'), fa: col('Facture(F) ou avoir(A)'), compteFact: col('N° compte facturé'), bt: col('Bon de transport'),
    poids: col('Poids'), service: col('Service'), nbColis: col("Nb colis dans l'envoi"), montant: col('Montant h.t. Total'),
    transport: col('Prix h.t. du transport'), options: col('Montant h.t.options et surcharges Total'),
  };
  const manquantes = Object.entries(i).filter(([, v]) => v < 0).map(([k]) => k);
  if (manquantes.length) throw new Error(`Fichier FACMAG : colonne(s) introuvable(s) (${manquantes.join(', ')}).`);
  const adr = ['Raison sociale expéditeur', 'Adresse 1 expéditeur', 'Adresse 2 expéditeur', 'Pays expéditeur', 'Département expéditeur', 'CP expéditeur', 'Ville expéditeur',
    'Raison Sociale destinataire', 'Adresse 1 destinataire', 'Adresse 2 destinataire', 'Pays destinataire', 'Département destinataire', 'CP destinataire', 'Ville destinataire'].map(col);
  // Colonnes de frais : toutes celles apres "Montant h.t.options et surcharges Total".
  const codes = h.map((nom, j) => [j, nom]).filter(([j, nom]) => j > i.options && nom);
  const donnees = lignes.slice(1).filter((r) => r.length > i.options);

  const abime = donnees.find((r) => /[eE]\+\d+$|,/.test(String(r[i.bt] || '').trim()) || /[eE]\+\d+$|,/.test(String(r[i.envoi] || '').trim()));
  if (abime) {
    throw new Error(`Fichier FACMAG abîmé : il a été ouvert puis enregistré dans Excel, qui a arrondi les numéros de suivi/d'envoi (ex. "${String(abime[i.bt] || abime[i.envoi]).trim()}"). Déposez le fichier FACMAG.CSV tel que téléchargé chez TNT, sans l'ouvrir dans Excel.`);
  }

  const vide16 = `'${' '.repeat(16)}`;
  const base = (r) => {
    const facture = String(r[i.facture] || '').trim();
    return [Number(String(r[i.mois]).trim()) || String(r[i.mois]).trim(), Number(facture.slice(0, 2)) || '', Number(facture.slice(2)) || facture,
      `'${r[i.compte]}`, `'${r[i.compteFact]}`];
  };
  const adresses = (r) => adr.map((j, k) => ([4, 5, 11, 12].includes(k) ? `'${r[j] ?? ''}` : (r[j] ?? '')));
  const ligneDet = (r, type, idColis, codeProduit, poids, code, montant) => [...base(r), type, idColis ? `'${idColis}` : vide16,
    Number(String(r[i.date]).trim()) || 0, `E${String(r[i.envoi]).trim()}`, codeProduit, poids, (code || '').padEnd(10), montant, ...adresses(r)];

  const parEnvoi = new Map(); // numero envoi -> { env: {r, poids, transport} | null, lignes: [] }
  const carburantParCompte = new Map();
  let nAvoir = 0;
  const incoherentes = [];
  for (const r of donnees) {
    const envoiNum = String(r[i.envoi] || '').trim();
    if (!envoiNum) continue;
    if (!parEnvoi.has(envoiNum)) parEnvoi.set(envoiNum, { env: null, lignes: [] });
    const e = parEnvoi.get(envoiNum);
    const bt = String(r[i.bt] || '').trim();
    const poids = num(r[i.poids]);
    const transport = num(r[i.transport]);
    const service = String(r[i.service] || '').trim();
    if (String(r[i.fa] || '').trim().toUpperCase() === 'A') nAvoir++;
    if (!bt || num(r[i.nbColis]) >= 1 || transport !== 0) {
      if (e.env) { e.env.transport = round2(e.env.transport + transport); e.env.poids = Math.max(e.env.poids, poids); } else e.env = { r, poids, transport, service };
    }
    if (bt) e.lignes.push(ligneDet(r, 'COL', bt, service, poids, '', 0));
    let sommeFrais = 0;
    for (const [j, code] of codes) {
      const m = num(r[j]);
      if (!m) continue;
      sommeFrais = round2(sommeFrais + m);
      if (code.toUpperCase() === CODE_CARBURANT_FACMAG) {
        const compte = String(r[i.compte] || '');
        carburantParCompte.set(compte, { r, montant: round2(((carburantParCompte.get(compte) || {}).montant || 0) + m) });
        continue;
      }
      e.lignes.push(ligneDet(r, bt ? 'CPC' : 'CPE', bt, '   ', bt ? poids : 0, code, m));
    }
    if (Math.abs(sommeFrais - num(r[i.options])) > 0.005 || Math.abs(round2(transport + num(r[i.options])) - num(r[i.montant])) > 0.005) {
      incoherentes.push(bt || envoiNum);
    }
  }
  const sortie = [];
  let nColis = 0;
  for (const e of parEnvoi.values()) {
    if (e.env) sortie.push(ligneDet(e.env.r, 'ENV', '', e.env.service, e.env.poids, '', e.env.transport));
    for (const l of e.lignes) { if (l[5] === 'COL') nColis++; sortie.push(l); }
  }
  let carburant = 0;
  for (const [compte, { r, montant }] of carburantParCompte) {
    carburant = round2(carburant + montant);
    sortie.push([...base(r).slice(0, 3), `'${compte}`, `'${compte}`, 'EXP', vide16, 0, 'E00000000000000', '   ', 0, 'SURCH_CARB', montant, ...ENTETE_DET_FAC.slice(13).map(() => '')]);
  }
  if (incoherentes.length) {
    warnings.push(`Fichier FACMAG : ${incoherentes.length} ligne(s) dont le total ne correspond pas au transport + frais (colonne de frais inconnue ?) — à vérifier : ${incoherentes.slice(0, 5).join(', ')}${incoherentes.length > 5 ? '…' : ''}`);
  }
  if (nAvoir) warnings.push(`Fichier FACMAG : ${nAvoir} ligne(s) d'avoir ("A") — montants repris tels quels, à vérifier.`);

  const ws = XLSX.utils.aoa_to_sheet([ENTETE_DET_FAC, ...sortie]);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'DET_FAC');
  const chemin = `${p}_det_fac.xlsx`;
  XLSX.writeFile(wb, chemin);
  infos.push(`Fichier FACMAG (CSV TNT) converti au format "détail facture" : ${donnees.length} ligne(s) FACMAG -> ${sortie.length} ligne(s), ${parEnvoi.size} envoi(s), ${nColis} colis, surcharge carburant (colonne ${CODE_CARBURANT_FACMAG}) ${carburant.toFixed(2)} EUR.`);
  return chemin;
}

/** PDF TNT (facture FedEx "services TNT") : 1 bloc par compte expediteur --
 * "<compte> : Surcharge Carburant (taux officiel X,XX %)  ...  Y,YY% taux reel  Montant"
 * -- SEUL le taux officiel fait foi (consigne utilisateur), le taux reel est ignore.
 * "TOTAL GENERAL" (colonne Montant) = total HT de la facture (sous-totaux journaliers
 * + toutes les surcharges carburant confondues). */
async function extractTntPdfInfo(pdfPath) {
  const buf = require('fs').readFileSync(pdfPath);
  const { text } = await pdfParse(buf);
  const surcharges = [];
  const reSurch = /(\d{8})\s*:\s*Surcharge\s*Carburant\s*\(taux\s*officiel\s*([\d,]+)\s*%\)[\s\S]*?([\d,]+)\s*EUR|(\d{8})\s*:\s*Surcharge\s*Carburant\s*\(taux\s*officiel\s*([\d,]+)\s*%\)/g;
  // Format observe (pdf-parse aplatit les colonnes en texte lineaire) : le montant en
  // euros de la ligne de surcharge suit immediatement le "taux reel", pas toujours capture
  // par une regex simple sur 1 seule passe -- on capture ligne par ligne le bloc complet.
  const reBloc = /(\d{8})\s*:\s*Surcharge\s*Carburant\s*\(taux\s*officiel\s*([\d,]+)\s*%\)\s*([\d,]+)\s*%\s*taux\s*r[ée]el\s*([\d,]+)/g;
  let m;
  while ((m = reBloc.exec(text)) !== null) {
    surcharges.push({
      compte: m[1],
      tauxOfficiel: num(m[2]) / 100,
      montant: numMilliers(m[4]),
    });
  }
  // "TOTAL GENERAL693 516 5.042,994.058,74   EUR" -- colis/envois/Poids/Montant tous colles
  // sans separateur (pdf-parse aplatit les colonnes du tableau) : Poids et Montant sont 2
  // nombres a virgule COLLES l'un a l'autre (ex. "5.042,99" + "4.058,74") -- capture les 2
  // groupes decimaux consecutifs juste avant "EUR", le MONTANT est le 2e (le plus proche
  // de "EUR").
  const mTotal = /TOTAL\s*GENERAL[\s\S]{0,40}?[\d]{1,3}(?:\.\d{3})*,\d{2}([\d]{1,3}(?:\.\d{3})*,\d{2})\s*EUR/.exec(text);
  const mNumero = /FACTURE\s*N[°\s]*(\d[\d\s]*\d)/.exec(text);
  return {
    file: path.basename(pdfPath),
    numeroFacture: mNumero ? mNumero[1].replace(/\s/g, '') : null,
    totalGeneral: mTotal ? numMilliers(mTotal[1].replace(/\s/g, '')) : null,
    surcharges,
  };
}

function categoriePour(codeComplement) {
  const c = String(codeComplement || '').trim();
  if (!c) return 'Frêt';
  return cfg.categories[c] || null;
}

async function process(files) {
  const warnings = [];
  const infos = [];
  // FACMAG (CSV TNT) : converti au format detail facture ; le chemin du fichier converti
  // REMPLACE le fichier recu dans files.facture, pour que le finaliseur Excel (buildArgs, appele
  // par le serveur apres process() avec ce meme objet files) lise exactement le meme contenu.
  files.facture = (files.facture || []).map((p) => (estFacmagCsv(p) ? convertirFacmag(p, warnings, infos) : p));
  const facturePaths = files.facture.filter(isXlsx);
  if (!facturePaths.length) throw new Error('Fichier TNT non reconnu : déposez le détail facture Excel (feuille DET_FAC) ou le fichier FACMAG.CSV téléchargé chez TNT.');

  const brutes = facturePaths.map(readBrut);

  const comptageMois = new Map(); // "AAAAMM" -> nb lignes
  const lignesParEnvoi = new Map(); // numeroEnvoi -> { montantEnv, lignesComplement: [{codeComplement, montantHt}], trackings: Set }
  const surchargeParCompte = new Map(); // compteExpediteur -> montantHt (SURCH_CARB, EXP)

  for (const f of brutes) {
    const iMois = colIndexByName(f.header, 'Mois de Facturation');
    const iLigneFacture = colIndexByName(f.header, 'Ligne Facture');
    const iIdColis = colIndexByName(f.header, 'id colis');
    const iNumeroEnvoi = colIndexByName(f.header, 'numero envoi');
    const iPoidsColis = colIndexByName(f.header, 'poids colis');
    const iCodeComplement = colIndexByName(f.header, 'code complement');
    const iMontantHt = colIndexByName(f.header, 'montant HT');
    const iCompteExpediteur = colIndexByName(f.header, 'Compte Expediteur');
    if (iMontantHt < 0 || iNumeroEnvoi < 0) {
      warnings.push(`${f.file} : colonne(s) attendue(s) introuvable(s) (numero envoi/montant HT) — fichier ignoré.`);
      continue;
    }

    for (const r of f.rows) {
      const mois = String(r[iMois] || '').trim();
      if (mois) comptageMois.set(mois, (comptageMois.get(mois) || 0) + 1);

      const ligneFacture = String(r[iLigneFacture] || '').trim().toUpperCase();
      const montantHt = num(r[iMontantHt]);
      const compteExpediteur = iCompteExpediteur >= 0 ? String(r[iCompteExpediteur] || '').replace(/^'/, '').trim() : '';
      const codeComplement = iCodeComplement >= 0 ? String(r[iCodeComplement] || '').trim() : '';

      // Ligne globale "surcharge carburant" (EXP/SURCH_CARB) : PAS liee a un envoi/tracking,
      // exclue du regroupement par envoi -- utilisee uniquement pour info/reconciliation.
      if (ligneFacture === 'EXP' && codeComplement === 'SURCH_CARB') {
        surchargeParCompte.set(compteExpediteur, round2((surchargeParCompte.get(compteExpediteur) || 0) + montantHt));
        continue;
      }

      const numeroEnvoi = String(r[iNumeroEnvoi] || '').trim();
      if (!numeroEnvoi) continue;
      const tracking = iIdColis >= 0 ? cleanTracking(r[iIdColis]) : '';
      const poids = iPoidsColis >= 0 ? num(r[iPoidsColis]) : 0;

      if (!lignesParEnvoi.has(numeroEnvoi)) lignesParEnvoi.set(numeroEnvoi, { trackings: new Set(), poidsParTracking: new Map(), lignesComplement: [] });
      const envoi = lignesParEnvoi.get(numeroEnvoi);
      if (tracking) {
        envoi.trackings.add(tracking);
        // Poids : MAX par tracking (reproduit le TCD 'Nb colis et poids' natif, dataField=
        // "Max. de poids colis", source = toutes les lignes "Facture TNT" -- pas seulement COL).
        envoi.poidsParTracking.set(tracking, Math.max(envoi.poidsParTracking.get(tracking) || 0, poids));
      }
      if (ligneFacture === 'ENV') {
        envoi.montantEnv = montantHt; // 1 seule ligne ENV par numero envoi, confirme empiriquement
      } else if (ligneFacture !== 'COL') {
        // CPC/CPE (RS/TMC/ZEL/LS/FACTOR_FOR/REC/...) ET autres types rares (ex. "M", constate
        // sur juillet 2026 : ligne d'ajustement/correction manuelle, code complement VIDE,
        // Montant HT reel NEGATIF -117,41€, "numero envoi" factice partage avec les lignes
        // EXP/SURCH_CARB) -- BUG TROUVE 2026-08-19 : le filtre `else if (codeComplement)`
        // ignorait silencieusement toute ligne a code complement VIDE hors ENV/COL, perdant
        // ce montant reel. categoriePour('') retourne deja "Frêt" par defaut (comme pour COL/
        // ENV) -- pas besoin de filtrer sur codeComplement, seulement d'exclure "COL" (montant
        // toujours 0, aucune info a perdre) qui alimente uniquement le poids via
        // 'poidsParTracking' ci-dessus. Rattachees au tracking de LEUR PROPRE ligne quand il
        // existe (contrairement a "ENV", ces lignes ONT generalement deja un id colis) --
        // sinon (comme "M") vont au dernier tracking de l'envoi, ou restent non affectees si
        // l'envoi n'a aucun tracking (cf. nEnvoiSansTracking plus bas).
        envoi.lignesComplement.push({ tracking, codeComplement, montantHt });
      }
    }
  }

  let dateValidite = null;
  if (comptageMois.size) {
    const [moisMajoritaire] = [...comptageMois.entries()].sort((a, b) => b[1] - a[1])[0];
    const m = /^(\d{4})(\d{2})$/.exec(moisMajoritaire);
    if (m) dateValidite = `01/${m[2]}/${m[1]}`;
    if (comptageMois.size > 1) {
      const detail = [...comptageMois.entries()].sort((a, b) => b[1] - a[1]).map(([k, v]) => `${k}: ${v} ligne(s)`).join(', ');
      infos.push(`Plusieurs mois détectés dans le fichier reçu (${detail}) — mois retenu : ${moisMajoritaire} (majoritaire).`);
    }
  }

  // Reclassement par tracking (id colis) : cumul de tous les postes (Frêt via ENV, +
  // supplements CPC/CPE de CE tracking, PAS ceux des autres trackings du meme envoi).
  const postesParTracking = new Map(); // tracking -> { [poste]: montant }
  const nbTmcParTracking = new Map();
  function ajouterPoste(tracking, poste, montant) {
    if (!tracking || !poste) return;
    if (!postesParTracking.has(tracking)) postesParTracking.set(tracking, Object.fromEntries(POSTE_KEYS.map((k) => [k, 0])));
    const p = postesParTracking.get(tracking);
    p[poste] = round2((p[poste] || 0) + montant);
  }

  let nEnvoiSansTracking = 0;
  for (const [numeroEnvoi, envoi] of lignesParEnvoi) {
    const trackings = [...envoi.trackings].sort(); // ordre alphabetique = ordre de tri du TCD modele
    // TOUS les montants de l'envoi (ENV + le RS/TMC/etc. de CHAQUE colis de cet envoi, meme
    // ceux qui ont leur propre id colis) sont attribues au tracking ALPHABETIQUEMENT LE PLUS
    // GRAND de l'envoi -- reproduit fidelement le classeur fait a la main : la colonne
    // "Tracking" de 'Facture TNT' (=RIGHT(LOOKUP(numeroEnvoi,...),16)) ne depend QUE du
    // "numero envoi" de la ligne, JAMAIS de son propre "id colis" -- donc TOUTES les lignes
    // d'un meme envoi (ENV et chaque colis/supplement) resolvent vers CE MEME tracking. Pour
    // un envoi de N colis, les N-1 autres n'apparaissent PAS du tout dans le fichier livre
    // (verifie via le TCD reel : G335='4543814314751861'=80,23€ = somme de TOUT l'envoi de 11
    // colis, PAS seulement son propre montant). Reproduit ICI a l'identique (decision
    // utilisateur 2026-08-18, "faites comme le fichier de base fait" -- meme si le PDF/releve
    // d'operations detaille chaque colis separement, le fichier livre a l'ERP ne le fait pas).
    if (!trackings.length) {
      // Envoi sans aucun colis/tracking reel (ex. "numero envoi" factice E00000000000000,
      // partage par les lignes EXP/SURCH_CARB ET par des lignes rares comme "M" -- ajustement/
      // correction manuelle constatee sur juillet 2026, montant HT reel -117,41€) : le montant
      // ENV et/ou les lignes complement de cet envoi ne peuvent etre affectes a aucun tracking
      // -> alerte explicite (perte silencieuse sinon, cf. bug 2026-08-19).
      const montantNonAffecte = round2((envoi.montantEnv || 0) + envoi.lignesComplement.reduce((s, l) => s + (l.montantHt || 0), 0));
      if (montantNonAffecte !== 0) {
        nEnvoiSansTracking++;
        warnings.push(`Envoi ${numeroEnvoi} : ${montantNonAffecte.toFixed(2)} EUR sans tracking (id colis) associé — montant non affecté au fichier livré.`);
      }
      continue;
    }
    const cible = trackings[trackings.length - 1];
    if (envoi.montantEnv != null) ajouterPoste(cible, 'Frêt', envoi.montantEnv);
    for (const l of envoi.lignesComplement) {
      const categorie = categoriePour(l.codeComplement);
      if (!categorie) { warnings.push(`Envoi ${numeroEnvoi} : code complément "${l.codeComplement}" absent de la table "Catégories" — non reclassé.`); continue; }
      if (categorie === 'Gazole') continue; // SURCH_CARB rattache a un envoi (rare/inattendu) : jamais reporte, cf. docstring
      if (l.codeComplement.toUpperCase() === 'TMC') nbTmcParTracking.set(cible, (nbTmcParTracking.get(cible) || 0) + 1);
      ajouterPoste(cible, categorie, l.montantHt);
    }
  }
  if (nEnvoiSansTracking) infos.push(`${nEnvoiSansTracking} envoi(s) avec un montant non nul mais sans tracking (id colis) associé — détail dans les avertissements ci-dessous.`);

  const poidsParTrackingGlobal = new Map();
  for (const envoi of lignesParEnvoi.values()) {
    for (const [tracking, poids] of envoi.poidsParTracking) {
      poidsParTrackingGlobal.set(tracking, Math.max(poidsParTrackingGlobal.get(tracking) || 0, poids));
    }
  }

  const recs = [];
  for (const [tracking, postes] of postesParTracking) {
    // Nbr Colis = 1 (colis de base) + 1 par ligne "TMC" rattachee a ce tracking (formule
    // modele : LOOKUP(tracking,'Nb colis et poids'!A:A,B:B)+1, B = Somme de montant HT pour
    // code complement="TMC", confirme = toujours 1 par ligne TMC sur juin 2026).
    const nbrColis = 1 + (nbTmcParTracking.get(tracking) || 0);
    recs.push({
      tracking, postes,
      poids: poidsParTrackingGlobal.get(tracking) || 0,
      nbrColis,
    });
  }
  recs.sort((a, b) => (a.tracking < b.tracking ? -1 : a.tracking > b.tracking ? 1 : 0));

  const importRows = recs.map((rec) => ({
    Transporteur: cfg.champs_fixes.Transporteur,
    DateValidite: dateValidite || '',
    Ref1: '', Ref2: '', IdClient: '',
    Tracking: rec.tracking, Nom: '',
    EP: cfg.champs_fixes['E/P'], Pays: cfg.champs_fixes.Pays, Zone: cfg.champs_fixes.Zone,
    NbrColis: rec.nbrColis, Poids: roundUp1(rec.poids),
    Mode: cfg.champs_fixes.Mode, TVA: cfg.champs_fixes.tva,
    DroitsTaxes: 0, Assurance: rec.postes.Assurance || 0,
    ZonesEloignees: rec.postes['Zones éloignées'] || 0,
    ColisVolumineux: rec.postes['Colis volumineux'] || 0,
    Adresses: 0,
    // Le gazole (SURCH_CARB) N'EST PAS repercute au client dans le fichier livre (meme
    // principe que Chronopost) -- jamais ecrit ici, sert uniquement a la reconciliation PDF.
    Fret: rec.postes['Frêt'] || 0, PlusValueB2C: 0, TaxeGasoil: '', NbColis: '',
  }));

  const { alerts, infos: valInfos } = validate(importRows);
  infos.push(...valInfos);

  // Reconciliation PDF : somme des postes ERP en euros (colonnes O->U du fichier import,
  // ici DroitsTaxes+Assurance+ZonesEloignees+ColisVolumineux+Adresses+Fret+PlusValueB2C,
  // TOUJOURS hors Gazole) + somme des surcharges carburant (montant PDF, "taux officiel")
  // doit egaler le "TOTAL GENERAL" HT de la facture (regle confirmee utilisateur 2026-08-18).
  const totalPostesHorsGazole = round2(importRows.reduce((s, r) =>
    s + (r.DroitsTaxes || 0) + (r.Assurance || 0) + (r.ZonesEloignees || 0)
      + (r.ColisVolumineux || 0) + (r.Adresses || 0) + (r.Fret || 0) + (r.PlusValueB2C || 0), 0));
  const totalSurcharges = round2([...surchargeParCompte.values()].reduce((s, v) => s + v, 0));
  infos.push(`Somme des postes hors Gazole (colonnes O→U) = ${totalPostesHorsGazole.toFixed(2)} EUR. Surcharge carburant (SURCH_CARB, brut) = ${totalSurcharges.toFixed(2)} EUR. Total attendu (hors TVA) = ${round2(totalPostesHorsGazole + totalSurcharges).toFixed(2)} EUR.`);

  const pdfPaths = files.pdf || [];
  const pdfs = [];
  for (const p of pdfPaths) {
    try {
      const r = await extractTntPdfInfo(p);
      if (r && r.totalGeneral != null) pdfs.push(r);
      else warnings.push(`PDF ${path.basename(p)} : "TOTAL GENERAL" introuvable, ignoré pour la réconciliation.`);
    } catch (e) {
      warnings.push(`PDF ${path.basename(p)} : lecture impossible (${e.message}).`);
    }
  }
  for (const p of pdfs) {
    // Taux officiel du PDF remonte en info (consigne utilisateur : c'est le taux a retenir,
    // PAS le "taux reel" affiché a cote) -- ne sert pas au calcul (SURCH_CARB brut du fichier
    // recu fait deja foi pour le montant), uniquement a la verification/tracabilite.
    for (const s of p.surcharges) {
      infos.push(`PDF ${p.file} : compte ${s.compte} — Surcharge Carburant taux officiel ${(s.tauxOfficiel * 100).toFixed(2)}% (${s.montant.toFixed(2)} EUR).`);
    }
    const calcule = round2(totalPostesHorsGazole + totalSurcharges);
    const ecart = round2(p.totalGeneral - calcule);
    const statut = Math.abs(ecart) >= 0.05 ? 'A VERIFIER' : 'OK';
    infos.push(`Facture PDF ${p.file}${p.numeroFacture ? ` (${p.numeroFacture})` : ''} : TOTAL GENERAL = ${p.totalGeneral.toFixed(2)} EUR, calculé = ${calcule.toFixed(2)} EUR (écart ${ecart >= 0 ? '+' : ''}${ecart.toFixed(2)} EUR -> ${statut})`);
  }

  const controle = {
    Assurance: round2(importRows.reduce((s, r) => s + (r.Assurance || 0), 0)),
    'Colis volumineux': round2(importRows.reduce((s, r) => s + (r.ColisVolumineux || 0), 0)),
    'Zones éloignées': round2(importRows.reduce((s, r) => s + (r.ZonesEloignees || 0), 0)),
    'Frêt': round2(importRows.reduce((s, r) => s + (r.Fret || 0), 0)),
    Gazole: totalSurcharges,
  };

  return {
    header: ['Mois de Facturation', 'Ligne Facture', 'id colis', 'numero envoi', 'code complement', 'montant HT'],
    rows: recs, recs, importRows, controle, warnings, alerts, infos,
    posteKeys: POSTE_KEYS, cfg, pdfs,
    sheetNames: { raw: 'Facture TNT', import: 'Import csv' },
    period: dateValidite ? `${dateValidite.slice(6)}_${dateValidite.slice(3, 5)}` : 'export',
  };
}

module.exports = {
  id: 'tnt',
  name: 'TNT',
  status: 'ready',
  taxeGasoil: "Reelle (surcharge carburant, code complement SURCH_CARB), montant GLOBAL par compte expediteur (pas ventilee par tracking) -- PDF affiche 'taux officiel' ET 'taux reel', SEUL le taux officiel fait foi. Jamais repercutee au client dans le fichier livre, sert a la reconciliation (somme postes + surcharge = TOTAL GENERAL PDF).",
  method: "1 fichier Excel brut (feuille DET_FAC, 1 ligne = 1 evenement, pas 1 colis ; ou export FACMAG.CSV, converti au meme format) : lignes 'ENV' portent le montant de transport reel de l'ENVOI (peut regrouper plusieurs colis 'COL') -- attribue au tracking alphabetiquement le plus grand de l'envoi (reproduit LOOKUP/cle dupliquee du modele). Supplements (CPC/CPE, code complement RS/TMC/ZEL/...) reclasses en postes ERP via la table Categories. Tracking JAMAIS converti en nombre. Reconciliation PDF : somme des postes + surcharge carburant (taux officiel) = TOTAL GENERAL.",
  inputs: [
    // .xlsb accepte en plus de .xlsx/.csv : le fichier recu en juillet 2026 etait au format
    // binaire Excel (.xlsb), pas .xlsx -- deja gere par le carrier (XLSX.readFile le lit
    // nativement) et le finaliseur (conversion via Excel COM, cf. memoire tnt_carrier_construit.md).
    // FACMAG.CSV (export CSV TNT, 2026-11-10) accepte aussi : converti au format detail facture
    // (cf. convertirFacmag).
    { key: 'facture', label: 'Facture TNT reçue (détail facture Excel, feuille DET_FAC, ou FACMAG.CSV)', accept: '.xlsx,.xlsb,.csv', multiple: false, required: true },
    { key: 'pdf', label: 'Facture(s) PDF TNT (contrôle du total et du taux officiel)', accept: '.pdf', multiple: true, required: false },
  ],
  outputNaming: { workbook: '{period}_Facture TNT', import: '{period}_TNT_Import' },
  finalizer: {
    script: '../automatisation/finaliser_tnt.py',
    template: '../Transporteurs/TNT/2026_06_Facture TNT.xlsx',
    buildArgs: (files) => [...(files.facture || []), '--pdf', ...(files.pdf || [])],
  },
  process,
};
