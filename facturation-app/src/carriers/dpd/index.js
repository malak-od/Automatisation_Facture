// ============================================================================
//  Adaptateur transporteur : DPD
//  Plusieurs fichiers "complément_facture..." (1 par n° facture/sous-compte DPD),
//  CSV ou XLSX (memes colonnes) -- reclassement = somme de colonnes nommees
//  precises en postes ERP, mapping decode des FORMULES du classeur fait a la
//  main (2026_06_Facture DPD.xlsx, onglet "Facture DPD"), positions de colonnes
//  NON stables (meme piege que Geodis) -> reclassement TOUJOURS par nom.
//  Poids repli 3e niveau : export expeditions du mois/mois-1 (comme GLS/Geodis),
//  le portail myDPD (4e niveau, vu en video) reste manuel.
// ============================================================================
const XLSX = require('xlsx');
const pdfParse = require('pdf-parse');
const { readCsv, num, roundUp1, round2 } = require('../../core/csv');
const { validate } = require('../../core/validate');
const { readBrutRows, poidsParTrackingFromExport } = require('../../core/exportBrut');
const path = require('path');
const cfg = require('./config.json');

const POSTE_KEYS_BRUT = ['DroitsTaxes', 'Assurance', 'ZonesEloignees', 'ColisVolumineux', 'Adresses', 'Fret', 'BtoC', 'Retour', 'Gazole'];
const POSTE_KEYS = ['DroitsTaxes', 'Assurance', 'ZonesEloignees', 'ColisVolumineux', 'Adresses', 'Fret', 'PlusValueB2C', 'Gazole'];

/** Normalise pour comparaison : accents + ponctuation traites comme equivalents,
 * comme Geodis (variantes cosmetiques vues d'un mois a l'autre chez les transporteurs). */
function normKey(s) {
  return String(s || '')
    .normalize('NFD').replace(/[̀-ͯ]/g, '')
    .toLowerCase().replace(/[.,]/g, '').replace(/\s+/g, ' ').trim();
}

function idx(header, name) {
  const target = normKey(name);
  return header.findIndex((h) => h && normKey(h) === target);
}

function readRowsXlsx(p) {
  const wb = XLSX.readFile(p, { cellDates: true });
  const ws = wb.Sheets[wb.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: '' });
  return { header: rows[0].map((h) => String(h || '').trim()), rows: rows.slice(1).filter((r) => r.some((v) => v !== '')) };
}

function readRowsCsv(p) {
  const { header, rows } = readCsv(p, ';', 'latin1');
  return { header: header.map((h) => String(h || '').trim()), rows: rows.filter((r) => r.some((v) => v !== '')) };
}

function isXlsx(p) {
  // Fichiers uploades (multer) arrivent sans extension -> detection par contenu (signature ZIP).
  const fs = require('fs');
  try {
    const buf = fs.readFileSync(p, { encoding: null, flag: 'r' });
    return buf.length > 4 && buf[0] === 0x50 && buf[1] === 0x4b; // "PK"
  } catch (e) {
    return false;
  }
}

function firstDayOfMonth(v) {
  // BUG TROUVE 2026-09-01 (meme piege identifie/corrige cote Delivengo) : XLSX.readFile
  // ({cellDates:true}) peut deriver de quelques heures lors de la conversion serial Excel ->
  // Date JS (ex. minuit UTC devient 21:59:39 la veille) -- .getMonth()/.getFullYear() en
  // heure LOCALE peuvent alors faire basculer un 1er-du-mois sur le mois precedent. Fix :
  // arrondir au jour le plus proche en UTC avant d'en extraire mois/annee.
  if (v instanceof Date) {
    const rounded = new Date(Math.round(v.getTime() / 86400000) * 86400000);
    return `01/${String(rounded.getUTCMonth() + 1).padStart(2, '0')}/${rounded.getUTCFullYear()}`;
  }
  const s = String(v || '').trim();
  let m = /^(\d{4})-(\d{2})-(\d{2})/.exec(s);
  if (m) return `01/${m[2]}/${m[1]}`;
  m = /^(\d{2})\/(\d{2})\/(\d{4})/.exec(s);
  if (m) return `01/${m[2]}/${m[3]}`;
  m = /^(\d{2})\.(\d{2})\.(\d{4})$/.exec(s); // format CSV brut DPD : "30.06.2026"
  if (m) return `01/${m[2]}/${m[3]}`;
  return null;
}

// PDF DPD ("Complement de facture" ou "FACTURE" principale -- la mise en page
// CHANGE d'un mois a l'autre, confirme juin vs juillet 2026, donc extraction
// volontairement minimale/tolerante plutot que collee a un format precis) :
// - Cle de rapprochement = 4 DERNIERS CHIFFRES DU N° DE CLIENT
//   ("N° de client : 021-009836" -> "9836"), PAS le n° de facture -- correspond
//   exactement a "No de compte" cote fichier DPD (confirme, meme valeur des
//   2 cotes ; le n° de facture n'a aucun rapport avec No de compte).
// - Montant = premier montant HT trouve sous un des libelles connus ("Total
//   EUR H.T", "TOTAL FACTURE HT"), recherche tolerante aux espaces/sauts de
//   ligne entre le libelle et le nombre (varie selon le mois).
const DPD_TOTAL_LABELS = [/Total\s*EUR\s*H\.?T\s*[\r\n]*\s*([\d\s]+[.,]\d{2})/i, /TOTAL\s*FACTURE\s*HT\s*[\r\n]*\s*([\d\s]+[.,]\d{2})/i];
async function extractDpdPdfTotal(pdfPath) {
  const buf = require('fs').readFileSync(pdfPath);
  const { text } = await pdfParse(buf);
  const mClient = /N°\s*de\s*client\s*:\s*([\d-]+)/i.exec(text);
  if (!mClient) return null;
  const last4 = mClient[1].replace(/\D/g, '').slice(-4);
  let totalHt = null;
  for (const re of DPD_TOTAL_LABELS) {
    const m = re.exec(text);
    if (m) { totalHt = num(m[1].replace(/\s/g, '')); break; }
  }
  if (totalHt == null) return null;
  return { file: path.basename(pdfPath), last4, totalHt };
}

async function process(files) {
  const paths = files.csv || [];
  if (!paths.length) throw new Error('Aucun fichier fourni (attendu : complément_facture DPD, CSV ou XLSX).');

  let header = null;
  const rows = [];
  for (const p of paths) {
    const r = isXlsx(p) ? readRowsXlsx(p) : readRowsCsv(p);
    header = r.header; // suppose meme structure pour tous les fichiers d'un meme mois
    rows.push(...r.rows);
  }

  const iTracking = idx(header, cfg.tracking_col);
  const iDate = idx(header, cfg.date_facture_col);
  const iPays = idx(header, cfg.pays_col);
  const iPoids = idx(header, cfg.poids_col);
  const iFraisDossier = idx(header, cfg.frais_dossier_source_col);
  const iNoCompte = idx(header, 'No de compte');
  const iDest = idx(header, 'Nom destinataire');
  const iNbColis = idx(header, 'Nombre de colis');
  const iRef = idx(header, 'Votre référence 1');

  const posteCols = {};
  for (const k of POSTE_KEYS_BRUT) posteCols[k] = (cfg.postes_from_columns[k] || []).map((name) => idx(header, name));

  const warnings = [];
  let dateValidite = null;

  // 1. Deux notions distinctes de "frais de dossier", confirmees dans les formules
  //    du classeur original :
  //    - "Frais dossier reel" (poste M, = colonne brute 'Frais de tenue de compte'
  //      TELLE QUELLE, deja repartie par DPD dans son propre fichier au sein d'une
  //      meme facture) -> entre dans "Total GO" (E = SUM(G:M)+SUM(A:C)), utilise
  //      pour la reconciliation PDF (Bilan factures : Somme de Total GO).
  //    - "Frais dossier modifie" (poste N, =ROUNDUP(SUM(M:M)/COUNT(M:M),2)) : le
  //      total de TOUT le mois reparti egalement sur toutes les lignes -> utilise
  //      UNIQUEMENT dans Import ERP.Fret (='Facture DPD'!I+K+N), PAS dans Total GO.
  const totalFraisTenueCompte = iFraisDossier >= 0
    ? round2(rows.reduce((s, r) => s + num(r[iFraisDossier]), 0))
    : 0;
  const fraisDossierModifieParLigne = rows.length ? Math.ceil((totalFraisTenueCompte / rows.length) * 100) / 100 : 0;

  const recsBrut = rows.map((r, i) => {
    const tracking = String(r[iTracking] ?? '').trim();
    if (!dateValidite) dateValidite = firstDayOfMonth(r[iDate]);

    const postes = {};
    for (const k of POSTE_KEYS_BRUT) {
      let total = 0;
      for (const ci of posteCols[k]) if (ci >= 0) total += num(r[ci]);
      postes[k] = round2(total);
    }
    const fraisDossierReel = iFraisDossier >= 0 ? num(r[iFraisDossier]) : 0;

    const noCompte = iNoCompte >= 0 ? String(r[iNoCompte] ?? '').trim() : '';
    const horsGo = round2(POSTE_KEYS_BRUT.filter((k) => k !== 'Gazole').reduce((s, k) => s + postes[k], 0));
    const avecGo = round2(horsGo + postes.Gazole);
    return {
      tracking,
      comref: iRef >= 0 ? String(r[iRef] ?? '').trim() : '',
      dest: iDest >= 0 ? String(r[iDest] ?? '').trim() : '',
      pays: iPays >= 0 ? String(r[iPays] ?? '').trim() : '',
      colis: iNbColis >= 0 ? num(r[iNbColis]) || 1 : 1,
      poidsBrut: iPoids >= 0 ? num(r[iPoids]) : 0,
      postes,
      horsGo, avecGo,
      fraisDossierReel,
      fraisDossierModifie: fraisDossierModifieParLigne,
      last4: noCompte.replace(/\D/g, '').slice(-4),
      raw: r,
    };
  }).filter((rec) => rec.tracking);

  // 2. Poids : brut si present, sinon Export (upload manuel OBLIGATOIRE, demande utilisateur
  //    2026-08-25) ; le 4e niveau (portail myDPD) reste manuel/hors app.
  const period = dateValidite ? `${dateValidite.slice(6)}_${dateValidite.slice(3, 5)}` : null;
  const brutPaths = files.brut || [];
  const poidsExportBrut = poidsParTrackingFromExport(readBrutRows(brutPaths));

  const infos = [];
  for (const rec of recsBrut) {
    if (rec.poidsBrut > 0) { rec.poids = rec.poidsBrut; continue; }
    if (poidsExportBrut.has(rec.tracking)) {
      rec.poids = poidsExportBrut.get(rec.tracking);
      infos.push(`${rec.tracking} : poids introuvable dans le fichier DPD -> repris de l'export expeditions (INFO_POIDSRETENU = ${rec.poids} kg)`);
    } else {
      rec.poids = 0.1;
      infos.push(`${rec.tracking} : poids introuvable (ni champ poids, ni export expeditions) -> plancher 0,1 applique, a verifier (dernier recours : portail myDPD for Business)`);
    }
  }

  // 3. Zone/TVA par pays (table Zoning du classeur fait a la main)
  const recs = recsBrut.map((rec) => {
    const zt = cfg.zoning[rec.pays] || null;
    const zone = zt ? String(zt[0]) : 'inconnu';
    const tva = zt ? zt[1] : '';
    if (!zt) warnings.push(`${rec.tracking} : pays '${rec.pays}' absent de la table de zoning DPD`);
    return { ...rec, zone, tva };
  });

  // 4. Import ERP : Fret = Facture DPD.Fret + Facture DPD.Retour + FraisDossierModifie
  //    (formule du classeur original -- Retour et Frais dossier n'ont pas de colonne
  //    ERP dediee, ils se fondent dans Fret au moment de l'import final).
  //    Regle FACTURATION EXCEL.pdf p.8 : si ce Fret ne contient QUE les frais fixes
  //    (frais de dossier 0,10 seul, ou 0,10 + Participation Surete 0,86 + Contribution
  //    Logistique 0,27 = 1,23), le deplacer en Adresses plutot que le laisser en Fret
  //    (ligne "vide" de transport reel -> ne doit pas ressortir en Fret).
  const recsAvecFret = recs.map((rec) => {
    let fret = round2(rec.postes.Fret + rec.postes.Retour + rec.fraisDossierModifie);
    let adresses = rec.postes.Adresses;
    if (Math.abs(fret - 0.10) < 0.005 || Math.abs(fret - 1.23) < 0.005) {
      adresses = round2(adresses + fret);
      fret = 0;
    }
    return { ...rec, fretFinal: fret, adressesFinal: adresses };
  });

  const importRows = recsAvecFret.map((rec) => ({
    Transporteur: cfg.champs_fixes.Transporteur,
    DateValidite: dateValidite || '',
    Ref1: '', Ref2: '', IdClient: '',
    Tracking: rec.tracking, Nom: '',
    EP: cfg.champs_fixes['E/P'], Pays: rec.pays, Zone: rec.zone,
    NbrColis: 1, Poids: roundUp1(rec.poids),
    Mode: cfg.champs_fixes['mode envoi'], TVA: rec.tva,
    DroitsTaxes: rec.postes.DroitsTaxes, Assurance: rec.postes.Assurance,
    ZonesEloignees: rec.postes.ZonesEloignees, ColisVolumineux: rec.postes.ColisVolumineux,
    Adresses: rec.adressesFinal,
    Fret: rec.fretFinal,
    PlusValueB2C: rec.postes.BtoC, TaxeGasoil: rec.postes.Gazole,
    NbColis: '',
  }));

  const controle = {};
  for (const k of POSTE_KEYS) {
    if (k === 'Fret') {
      controle.Fret = round2(recsAvecFret.reduce((s, r) => s + r.fretFinal, 0));
    } else if (k === 'Adresses') {
      controle.Adresses = round2(recsAvecFret.reduce((s, r) => s + r.adressesFinal, 0));
    } else if (k === 'PlusValueB2C') {
      controle.PlusValueB2C = round2(recs.reduce((s, r) => s + r.postes.BtoC, 0));
    } else {
      controle[k] = round2(recs.reduce((s, r) => s + (r.postes[k] || 0), 0));
    }
  }

  if (totalFraisTenueCompte) infos.push(`Frais de tenue de compte DPD (${totalFraisTenueCompte.toFixed(2)} EUR au total, reparti sur ${rows.length} ligne(s) : ${fraisDossierModifieParLigne.toFixed(2)} EUR/ligne) inclus dans Fret -> comparer au Bilan factures / factures PDF DPD.`);

  // Reconciliation PDF (optionnelle) : montant HT de chaque PDF DPD, a comparer
  // au "Total GO" (formule classeur : SUM(G:M)+SUM(A:C), donc SANS Retour(K) ni
  // FraisDossierModifie(N) -- avec FraisDossierReel(M) brut) pour ce meme
  // compte, rapproche par les 4 derniers chiffres du N° DE CLIENT du PDF
  // (= "No de compte" cote fichier DPD -- PAS le n° de facture, aucun rapport).
  const totalGoParCompte = new Map(); // last4 (No de compte) -> somme
  for (const rec of recsBrut) {
    const total = round2(
      rec.postes.DroitsTaxes + rec.postes.Assurance + rec.postes.ZonesEloignees
      + rec.postes.ColisVolumineux + rec.postes.Adresses + rec.postes.Fret
      + rec.postes.BtoC + rec.postes.Retour + rec.postes.Gazole + rec.fraisDossierReel
    );
    totalGoParCompte.set(rec.last4, round2((totalGoParCompte.get(rec.last4) || 0) + total));
  }

  const pdfPaths = files.pdf || [];
  const pdfs = [];
  for (const p of pdfPaths) {
    try {
      const r = await extractDpdPdfTotal(p);
      if (r) pdfs.push(r);
      else warnings.push(`PDF ${path.basename(p)} : "N° de client" ou montant HT introuvable, ignore pour la reconciliation.`);
    } catch (e) {
      warnings.push(`PDF ${path.basename(p)} : lecture impossible (${e.message}).`);
    }
  }
  for (const p of pdfs) {
    const calcule = totalGoParCompte.get(p.last4);
    if (calcule == null) {
      warnings.push(`Facture PDF ${p.file} (client se terminant par ${p.last4}) : aucune ligne correspondante trouvee dans les fichiers DPD traites.`);
      continue;
    }
    const ecart = round2(calcule - p.totalHt);
    infos.push(`Facture PDF ${p.file} (client ${p.last4}) : Montant HT = ${p.totalHt.toFixed(2)} EUR, calcule = ${calcule.toFixed(2)} EUR (ecart ${ecart >= 0 ? '+' : ''}${ecart.toFixed(2)} EUR${Math.abs(ecart) < 0.05 ? ' -> OK' : ' -> A VERIFIER'})`);
  }

  const { alerts, infos: validateInfos } = validate(importRows);

  // Feuille brute (repli exceljs, pas de finaliseur Excel COM pour DPD) : postes
  // internes (DroitsTaxes, ZonesEloignees...) different des libelles par defaut
  // de writeWorkbook() -> colonnes/valeurs explicites, comme GLS.
  const rawCompCols = ['Droits et taxes', 'Zones éloignées', 'Colis volumineux', 'Adresses', 'Fret', 'BtoC', 'Retour', 'Frais dossier réel', 'Frais dossier modifié', 'Gazole'];
  const rawCompRow = (rec) => [rec.postes.DroitsTaxes, rec.postes.ZonesEloignees, rec.postes.ColisVolumineux, rec.postes.Adresses, rec.postes.Fret, rec.postes.BtoC, rec.postes.Retour, rec.fraisDossierReel, rec.fraisDossierModifie, rec.postes.Gazole];

  return {
    header, rows, recs, importRows, controle, warnings, alerts,
    infos: [...infos, ...validateInfos],
    posteKeys: POSTE_KEYS, gazoleKey: 'Gazole', cfg, pdfs,
    sheetNames: { raw: 'Facture DPD', import: 'Import ERP' },
    rawCompCols, rawCompRow,
    period: period || 'export',
  };
}

module.exports = {
  id: 'dpd',
  name: 'DPD',
  status: 'ready',
  taxeGasoil: 'Calculee ligne a ligne (colonnes Taxe Consolidation + Indexation gasoil + Indexation kerosene) ; taux publie par DPD sur chaque facture (variable mensuel).',
  method: "Fichiers 'complement_facture...' DPD (CSV ou XLSX, un par n° facture/sous-compte), colonnes tres larges. ATTENTION : les colonnes se decalent d'un mois a l'autre (meme piege que Geodis), reclassement fait par NOM de colonne exclusivement. 'Frais dossier modifie' = total mensuel des frais de tenue de compte, reparti egalement (arrondi sup.) sur toutes les lignes, inclus dans Fret. Si Fret = 0,10 (frais dossier seul) ou 1,23 (+ Participation Surete 0,86 + Contribution Logistique 0,27), deplace en Adresses (FACTURATION EXCEL.pdf p.8). Poids : champ brut, sinon repli automatique sur l'export expeditions du mois/mois-1 (INFO_POIDSRETENU), sinon plancher 0,1 (dernier recours manuel : portail myDPD for Business). Zone/TVA par pays (table Zoning). PDF optionnel : compare le montant HT de chaque PDF (libelle variable selon le mois : 'Total EUR H.T' ou 'TOTAL FACTURE HT') au total calcule pour ce meme compte, rapproche par les 4 derniers chiffres du N° de client (= No de compte, pas le n° de facture).",
  inputs: [
    { key: 'csv', label: 'Fichiers complément_facture DPD (CSV ou XLSX)', accept: '.csv,.xlsx', multiple: true, required: true },
    { key: 'pdf', label: 'Facture(s) PDF DPD (contrôle du total par facture)', accept: '.pdf', multiple: true, required: false },
    { key: 'brut', label: 'Export des expéditions brutes (et mois précédent si dispo)', accept: '.xlsx,.xls', multiple: true, required: true },
  ],
  // Classeur = CLONE FIDELE du fichier fait a la main (Excel COM/Python), comme Geodis/Kuehne/Delivengo.
  outputNaming: { workbook: '{period}_Facture DPD', import: '{period}_DPD_Import' },
  finalizer: {
    script: '../automatisation/finaliser_dpd.py',
    template: '../Transporteurs/DPD/2026_06_Facture DPD.xlsx',
    buildArgs: (files) => [...(files.csv || []), '--pdf', ...(files.pdf || [])],
  },
  process,
};
