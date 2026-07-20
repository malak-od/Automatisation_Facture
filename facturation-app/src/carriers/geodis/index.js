// ============================================================================
//  Adaptateur transporteur : GEODIS
//  1 ligne brute = 1 expedition (colonnes tres larges, ~100 colonnes nommees,
//  une colonne par type de frais). Reclassement = somme de colonnes nommees
//  precises en 8 postes ERP -- mapping decode des FORMULES du classeur fait a
//  la main (2026_06_Facture Geodis.xlsx), les positions de colonnes se
//  decalent d'un mois a l'autre (confirme avril vs juin) donc on cherche
//  TOUJOURS par nom, jamais par position.
// ============================================================================
const XLSX = require('xlsx');
const { num, roundUp1, round2 } = require('../../core/csv');
const { validate } = require('../../core/validate');
const cfg = require('./config.json');

const POSTE_KEYS = ['DroitsTaxes', 'Assurance', 'ZonesEloignees', 'ColisVolumineux', 'Adresses', 'Fret', 'PlusValueB2C', 'Gazole'];

function readRows(path) {
  const wb = XLSX.readFile(path, { raw: true });
  const ws = wb.Sheets[wb.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: '' });
  const header = (rows[0] || []).map((h) => String(h || '').trim());
  return { header, rows: rows.slice(1).filter((r) => r && r.length && r.some((v) => v !== '')) };
}

function idx(header, name) {
  const i = header.indexOf(name);
  if (i >= 0) return i;
  return header.findIndex((h) => h && h.includes(name));
}

function firstDayOfMonth(v) {
  // "Date piece" peut arriver en texte JJ/MM/AAAA ou AAAA-MM-JJ, ou en date Excel (via xlsx: objet Date si cellFormula/cellDates, sinon numero serie)
  if (v instanceof Date) return `01/${String(v.getMonth() + 1).padStart(2, '0')}/${v.getFullYear()}`;
  const s = String(v || '').trim();
  let m = /^(\d{4})-(\d{2})-(\d{2})/.exec(s);
  if (m) return `01/${m[2]}/${m[1]}`;
  m = /^(\d{2})\/(\d{2})\/(\d{4})/.exec(s);
  if (m) return `01/${m[2]}/${m[3]}`;
  return null;
}

function process(files) {
  const paths = files.facture || [];
  if (!paths.length) throw new Error('Aucun fichier fourni (attendu : export Facture Geodis, xlsx ou csv).');

  let header = null;
  const rows = [];
  for (const p of paths) {
    const r = readRows(p);
    header = r.header;
    rows.push(...r.rows);
  }

  const iRecepisse = idx(header, cfg.recepisse_col);
  const iRef1 = idx(header, cfg.ref1_col);
  const iDate = idx(header, cfg.date_piece_col);
  const iDest = idx(header, cfg.dest_nom_col);
  const iPays = idx(header, cfg.pays_col);
  const iNbColis = idx(header, cfg.nb_colis_col);
  const iPoids = idx(header, cfg.poids_col);
  const posteCols = {};
  for (const k of POSTE_KEYS) posteCols[k] = (cfg.postes_from_columns[k] || []).map((name) => idx(header, name));

  const warnings = [];
  let dateValidite = null;

  const recs = rows.map((r, i) => {
    const recepisseRaw = String(r[iRecepisse] ?? '').trim();
    const tracking = recepisseRaw.slice(-8); // RIGHT(...,8), comme le classeur
    if (!dateValidite) dateValidite = firstDayOfMonth(r[iDate]);

    const postes = {};
    for (const k of POSTE_KEYS) {
      let total = 0;
      for (const ci of posteCols[k]) if (ci >= 0) total += num(r[ci]);
      postes[k] = round2(total);
    }

    return {
      tracking, comref: iRef1 >= 0 ? String(r[iRef1] ?? '').trim() : '',
      dest: iDest >= 0 ? String(r[iDest] ?? '').trim() : '',
      pays: iPays >= 0 ? String(r[iPays] ?? '').trim() : '',
      colis: iNbColis >= 0 ? num(r[iNbColis]) : 1,
      poids: iPoids >= 0 ? num(r[iPoids]) : 0,
      zone: cfg.champs_fixes.zone,
      postes,
      horsGo: round2(POSTE_KEYS.filter((k) => k !== 'Gazole').reduce((s, k) => s + postes[k], 0)),
      avecGo: round2(POSTE_KEYS.reduce((s, k) => s + postes[k], 0)),
      raw: r,
    };
  }).filter((rec) => rec.tracking);

  const importRows = recs.map((rec) => ({
    Transporteur: cfg.champs_fixes.Transporteur,
    DateValidite: dateValidite || '',
    Ref1: rec.comref, Ref2: '', IdClient: '',
    Tracking: rec.tracking, Nom: rec.dest,
    EP: cfg.champs_fixes['E/P'], Pays: rec.pays, Zone: rec.zone,
    NbrColis: Math.round(rec.colis), Poids: roundUp1(rec.poids),
    Mode: cfg.champs_fixes['mode envoi'], TVA: cfg.champs_fixes.tva,
    DroitsTaxes: rec.postes.DroitsTaxes, Assurance: rec.postes.Assurance,
    ZonesEloignees: rec.postes.ZonesEloignees, ColisVolumineux: rec.postes.ColisVolumineux,
    Adresses: rec.postes.Adresses, Fret: rec.postes.Fret,
    PlusValueB2C: rec.postes.PlusValueB2C, TaxeGasoil: rec.postes.Gazole,
    NbColis: '',
  }));

  const controle = {};
  for (const k of POSTE_KEYS) controle[k] = round2(recs.reduce((s, r) => s + (r.postes[k] || 0), 0));

  const { alerts, infos } = validate(importRows);
  const fraisFixe = cfg.frais_tenue_compte_mensuel || 0;
  if (fraisFixe) infos.push(`Frais de tenue de compte GEODIS (${fraisFixe.toFixed(2)} EUR/mois, sans tracking) a ajouter manuellement -> total facture attendu = ${round2(POSTE_KEYS.reduce((s, k) => s + (controle[k] || 0), 0) + fraisFixe)} EUR HT`);

  return {
    header, rows, recs, importRows, controle, warnings, alerts, infos,
    posteKeys: POSTE_KEYS, gazoleKey: 'Gazole', cfg,
    period: dateValidite ? `${dateValidite.slice(6)}_${dateValidite.slice(3, 5)}` : 'export',
  };
}

module.exports = {
  id: 'geodis',
  name: 'Geodis',
  status: 'ready',
  viticolis: true,
  taxeGasoil: 'Calculee ligne a ligne (colonne Surcharge Carburant), pas besoin du PDF',
  method: "Fichier Facture Geodis (xlsx/csv, ~100 colonnes). ATTENTION : les colonnes se decalent d'un mois a l'autre (confirme), reclassement fait par NOM de colonne exclusivement. Poids = Poids origine (kg). E/P non fiable cote Geodis -> laisse vide (alerte 'a verifier'). Frais de tenue de compte mensuel fixe (sans tracking) remonte en info, a ajouter a la main.",
  inputs: [
    { key: 'facture', label: 'Facture Geodis (xlsx/csv)', accept: '.xlsx,.csv', multiple: true, required: true },
  ],
  process,
};
