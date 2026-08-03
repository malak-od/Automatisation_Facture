// ============================================================================
//  Adaptateur transporteur : KUEHNE+NAGEL  (implemente, valide 187/187)
// ============================================================================
const path = require('path');
const { readCsv, num, colIndex, roundUp1, round2 } = require('../../core/csv');
const { reclasser } = require('../../core/reclass');
const { buildDeptLookup, zone } = require('../../core/zone');
const { validate } = require('../../core/validate');
const cfg = require('./config.json');

const POSTE_KEYS = ['Droits et taxes', 'Assurance', 'Zones eloignees', 'Colis volumineux', 'Adresses', 'Fret', 'Plus value B2C', 'Gazole'];

/** Ligne "hors grille" : affretement ou navette -> zone volontairement vide (pas de grille tarifaire). */
function isHorsGrille(rec) {
  const hg = cfg.hors_grille || {};
  const nav = hg.navette_destinataire_contient && rec.dest.toUpperCase().includes(hg.navette_destinataire_contient);
  return rec.metier === hg.metier_affretement || !!nav;
}

/** Transforme une fiche (rec) en ligne d'import ERP. */
function toImportRow(rec, dateValidite) {
  const p = rec.postes;
  const horsUe = new Set(cfg.tva.pays_hors_ue);
  return {
    _horsGrille: isHorsGrille(rec), // interne (non exporte) : sert a la validation
    _metier: rec.metier,
    Transporteur: cfg.champs_fixes.Transporteur,
    DateValidite: dateValidite,
    Ref1: rec.comref, Ref2: '', IdClient: '',
    Tracking: rec.tracking, Nom: rec.dest,
    EP: cfg.champs_fixes['E/P'], Pays: rec.pays, Zone: rec.zone,
    // Poids : arrondi SUPERIEUR (jamais au plus proche) -> voir FACTURATION EXCEL.pdf p.1
    NbrColis: Math.round(rec.colis), Poids: roundUp1(rec.poids),
    Mode: cfg.champs_fixes['mode envoi'],
    TVA: horsUe.has(rec.pays) ? cfg.tva.hors_ue : cfg.tva.defaut,
    DroitsTaxes: p['Droits et taxes'], Assurance: p.Assurance,
    ZonesEloignees: p['Zones eloignees'], ColisVolumineux: p['Colis volumineux'],
    Adresses: p.Adresses, Fret: p.Fret, PlusValueB2C: p['Plus value B2C'],
    TaxeGasoil: '', NbColis: '',
  };
}

/**
 * Traite les fichiers Kuehne d'un mois.
 * @param {object} files { csv:[paths], pdf:[paths] }
 * @returns resultat complet (importRows, controle, alerts, warnings, recs, header, pdfs...)
 */
function process(files) {
  const csvPaths = files.csv || [];
  if (!csvPaths.length) throw new Error('Aucun fichier CSV fourni (attendu : FcCSV*.csv).');

  // 1. Lecture + concatenation des CSV (facture transport + facture evenements)
  let header = null;
  const rows = [];
  for (const p of csvPaths) {
    const { header: h, rows: rs } = readCsv(p, cfg.csv.sep);
    header = h;
    rows.push(...rs.map((r) => (r.length < h.length ? r.concat(Array(h.length - r.length).fill('')) : r)));
  }

  // 2. Indices + lookup dept
  const iEdi = colIndex(header, cfg.tracking.primary);
  const iCom = colIndex(header, cfg.tracking.fallback);
  const iDest = colIndex(header, cfg.identite.destinataire);
  const iPays = colIndex(header, cfg.identite.pays_dest);
  const iPoids = colIndex(header, cfg.identite.poids);
  const iColis = colIndex(header, cfg.identite.colis);
  const iMetier = colIndex(header, cfg.identite.metier);
  const { de, dd } = buildDeptLookup(rows, header, cfg);

  // 3. Une fiche par ligne brute (reclassement + identite + zone)
  const warnings = [];
  const controle = {};
  const recs = rows.map((r) => {
    const postes = reclasser(r, header, cfg, warnings);
    for (const k of POSTE_KEYS) controle[k] = round2((controle[k] || 0) + postes[k]);
    const edi = (r[iEdi] || '').trim();
    const track = edi || (r[iCom] || '').trim();
    const pays = (r[iPays] || '').trim();
    return {
      tracking: track, comref: (r[iCom] || '').trim(), dest: (r[iDest] || '').trim(), pays,
      poids: num(r[iPoids]), colis: num(r[iColis]), metier: (r[iMetier] || '').trim(), postes,
      horsGo: round2(POSTE_KEYS.filter((k) => k !== 'Gazole').reduce((s, k) => s + postes[k], 0)),
      avecGo: round2(POSTE_KEYS.reduce((s, k) => s + postes[k], 0)),
      zone: track ? zone(track, pays, de, dd) : '',
      raw: r,
    };
  });

  // 4. Import = 1 ligne par fiche AVEC tracking (exclusion tracking vide), aucun regroupement
  const period = derivePeriod(rows, header);
  // DateValidite = 1er jour du mois de la periode reelle (deduite du CSV), jamais une
  // constante figee en config -> evite d'importer sous le mauvais mois si on oublie de la mettre a jour.
  const dateValidite = periodToDate(period) || cfg.champs_fixes.date_validite;
  const importRows = recs.filter((rec) => rec.tracking).map((rec) => toImportRow(rec, dateValidite));
  const { alerts, infos } = validate(importRows);

  return { header, rows, recs, importRows, controle, warnings, alerts, infos, posteKeys: POSTE_KEYS, cfg, period, sheetNames: { raw: 'Fichier Kuehne', import: 'Kuehne_Import' } };
}

/** Periode "AAAA_MM" depuis la 1ere Date facture des CSV (fallback : date de validite tarif). */
function derivePeriod(rows, header) {
  const i = colIndex(header, 'Date facture');
  for (const r of rows) {
    const m = (r[i] || '').trim().match(/^(\d{2})\/(\d{2})\/(\d{4})$/); // JJ/MM/AAAA
    if (m) return `${m[3]}_${m[2]}`;
  }
  const dv = (cfg.champs_fixes.date_validite || '').match(/(\d{2})\/(\d{2})\/(\d{4})/);
  return dv ? `${dv[3]}_${dv[2]}` : 'export';
}

/** "AAAA_MM" -> "01/MM/AAAA" (format ERP "Date validite tarif"). */
function periodToDate(period) {
  const m = /^(\d{4})_(\d{2})$/.exec(period);
  return m ? `01/${m[2]}/${m[1]}` : null;
}

module.exports = {
  id: 'kuehne',
  name: 'Kuehne + Nagel',
  status: 'ready',
  viticolis: true,
  taxeGasoil: "Lue dans le fichier Excel envoye par K&N (ni site, ni ERP).",
  // Nomenclature des sorties (comme le fichier fait a la main). {period} = AAAA_MM.
  outputNaming: { workbook: '{period}_Facture Kuehne', import: '{period}_Kuehne_Import' },
  // Classeur de controle = CLONE FIDELE du fichier fait a la main (feuilles/formules/TCD/mise en forme)
  // via Excel COM (Python). Le modele est le fichier existant, rafraichi avec les nouvelles donnees.
  finalizer: {
    script: '../automatisation/finaliser_kuehne.py',
    template: '../Transporteurs/Kuehne/2026_06_Facture Kuehne.xlsx',
    buildArgs: (files) => ['--csv', ...(files.csv || []), '--pdf', ...(files.pdf || [])],
  },
  method: "Flux EDI. K&N fournit 2 CSV (facture transport + facture evenements) et 2 PDF. Reclassement ~130 colonnes T_* -> 8 postes, 1 ligne CSV = 1 ligne import (pas de regroupement), gazole refacture a part.",
  inputs: [
    { key: 'csv', label: 'Fichiers CSV K&N (FcCSV*.csv)', accept: '.csv', multiple: true, required: true, match: 'FcCSV' },
    { key: 'pdf', label: 'Factures PDF (IMAGE*.pdf) — pour reconciliation', accept: '.pdf', multiple: true, required: false },
  ],
  process,
};
