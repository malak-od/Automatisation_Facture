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
  const data = rows.slice(1).filter((r) => r.some((v) => v !== ''));
  return { file: path.basename(p), header, rows: data };
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
  const facturePaths = (files.facture || []).filter(isXlsx);
  if (!facturePaths.length) throw new Error('Aucun fichier fourni (attendu : facture TNT reçue, xlsx).');

  const warnings = [];
  const infos = [];
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
  method: "1 fichier Excel brut (feuille DET_FAC, 1 ligne = 1 evenement, pas 1 colis) : lignes 'ENV' portent le montant de transport reel de l'ENVOI (peut regrouper plusieurs colis 'COL') -- attribue au tracking alphabetiquement le plus grand de l'envoi (reproduit LOOKUP/cle dupliquee du modele). Supplements (CPC/CPE, code complement RS/TMC/ZEL/...) reclasses en postes ERP via la table Categories. Tracking JAMAIS converti en nombre. Reconciliation PDF : somme des postes + surcharge carburant (taux officiel) = TOTAL GENERAL.",
  inputs: [
    // .xlsb accepte en plus de .xlsx/.csv : le fichier recu en juillet 2026 etait au format
    // binaire Excel (.xlsb), pas .xlsx -- deja gere par le carrier (XLSX.readFile le lit
    // nativement) et le finaliseur (conversion via Excel COM, cf. memoire tnt_carrier_construit.md).
    { key: 'facture', label: 'Facture TNT reçue (xlsx/xlsb, feuille DET_FAC)', accept: '.xlsx,.xlsb,.csv', multiple: false, required: true },
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
