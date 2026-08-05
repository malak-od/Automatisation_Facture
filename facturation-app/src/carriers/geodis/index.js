// ============================================================================
//  Adaptateur transporteur : GEODIS
//  1 ligne brute = 1 expedition (colonnes tres larges, ~100 colonnes nommees,
//  une colonne par type de frais). Reclassement = somme de colonnes nommees
//  precises en 8 postes ERP -- mapping decode des FORMULES du classeur fait a
//  la main (2026_06_Facture Geodis.xlsx), les positions de colonnes se
//  decalent d'un mois a l'autre (confirme avril vs juin) donc on cherche
//  TOUJOURS par nom, jamais par position.
// ============================================================================
const path = require('path');
const XLSX = require('xlsx');
const pdfParse = require('pdf-parse');
const { num, roundUp1, round2 } = require('../../core/csv');
const { validate } = require('../../core/validate');
const cfg = require('./config.json');

const POSTE_KEYS = ['DroitsTaxes', 'Assurance', 'ZonesEloignees', 'ColisVolumineux', 'Adresses', 'Fret', 'PlusValueB2C', 'Gazole'];

// La ligne d'en-tete n'est pas forcement la ligne 1 : le fichier "brut" Geodis
// l'a en ligne 1, mais le meme detail reexporte en "rapport client" (ex.
// Facturation_client.xlsx) a 2-3 lignes de titre au-dessus (voir capture
// utilisateur : 0 ligne generee car la ligne 1 ne contenait pas les en-tetes
// attendus). On cherche donc, sur TOUTES les feuilles, la 1ere ligne qui
// contient une colonne reconnaissable (recepisse_col), plutot que de supposer
// une position fixe.
function findHeaderRow(rows) {
  const marker = cfg.recepisse_col.toLowerCase();
  for (let i = 0; i < Math.min(rows.length, 15); i++) {
    const row = rows[i] || [];
    if (row.some((v) => String(v || '').toLowerCase().includes(marker))) return i;
  }
  return -1;
}

function readRows(path) {
  // cellDates:true -> les cellules "Date" (ex. Facturation_client.xlsx) reviennent en
  // objet Date JS, pas en numero de serie Excel (sinon firstDayOfMonth() ne matche rien
  // et DateValidite reste vide sur TOUTES les lignes -> "DATE manquante" partout).
  const wb = XLSX.readFile(path, { raw: true, cellDates: true });
  for (const sheetName of wb.SheetNames) {
    const ws = wb.Sheets[sheetName];
    const rows = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: '' });
    const hIdx = findHeaderRow(rows);
    if (hIdx === -1) continue; // pas la bonne feuille (ex. "Recapitulatif")
    const header = rows[hIdx].map((h) => String(h || '').trim());
    const dataRows = rows.slice(hIdx + 1).filter((r) => r && r.length && r.some((v) => v !== ''));
    return { header, rows: dataRows };
  }
  throw new Error(`Format de fichier non reconnu (colonne "${cfg.recepisse_col}" introuvable dans aucune feuille) : ce n'est probablement pas un export Geodis valide.`);
}

// Normalise pour comparaison : accents + "," / "." traites comme equivalents.
// Vu entre exports Geodis (ex. "RV tel," vs "RV tel." dans Facturation_client.xlsx,
// qui a fait perdre 5 colonnes de Fret en silence -> ~5,70EUR sur ~10 lignes).
function normKey(s) {
  return String(s || '')
    .normalize('NFD').replace(/[̀-ͯ]/g, '') // accents
    .toLowerCase().replace(/[.,]/g, '').replace(/\s+/g, ' ').trim();
}

function idx(header, name) {
  const i = header.indexOf(name);
  if (i >= 0) return i;
  const target = normKey(name);
  const exact = header.findIndex((h) => h && normKey(h) === target);
  if (exact >= 0) return exact;
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

// Facture PDF Geodis : "Montant Taxable HT|TVA 20,00%|Montant Exonere|Montant TTC..."
// (labels colles) suivi des montants "X,XX EUR" dans le meme ordre. On prend
// les 3 premiers pour Taxable/TVA/TTC. "Frais de gestion de compte" est aussi
// une ligne dediee (montant fixe mensuel, sans tracking, cf. config).
async function extractGeodisPdfTotal(pdfPath) {
  const buf = require('fs').readFileSync(pdfPath);
  const { text } = await pdfParse(buf);
  const iTaxable = text.indexOf('Montant Taxable HT');
  if (iTaxable === -1) return null;
  const after = text.slice(iTaxable);
  const amounts = [...after.matchAll(/([\d\s]+,\d{2})\s*EUR/g)].slice(0, 3).map((m) => num(m[1].replace(/\s/g, '')));
  if (!amounts.length) return null;
  const fraisMatch = /Frais de gestion de compte\s*([\d\s]+,\d{2})/.exec(text);
  return {
    file: path.basename(pdfPath),
    taxable: amounts[0],
    tva: amounts[1] ?? null,
    ttc: amounts[2] ?? null,
    fraisGestion: fraisMatch ? num(fraisMatch[1]) : null,
  };
}

async function process(files) {
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

  // Reconciliation PDF (optionnelle) : "Montant Taxable HT" de la facture Geodis,
  // a comparer au total calcule + frais de tenue de compte (poste sans tracking).
  const pdfPaths = files.pdf || [];
  const pdfs = [];
  for (const p of pdfPaths) {
    try {
      const r = await extractGeodisPdfTotal(p);
      if (r) pdfs.push(r);
      else warnings.push(`PDF ${path.basename(p)} : "Montant Taxable HT" introuvable, ignore pour la reconciliation.`);
    } catch (e) {
      warnings.push(`PDF ${path.basename(p)} : lecture impossible (${e.message}).`);
    }
  }
  const fraisFixe = (pdfs.find((p) => p.fraisGestion != null) || {}).fraisGestion ?? cfg.frais_tenue_compte_mensuel ?? 0;

  const { alerts, infos } = validate(importRows);
  const totalCalcule = round2(POSTE_KEYS.reduce((s, k) => s + (controle[k] || 0), 0) + fraisFixe);
  if (fraisFixe) infos.push(`Frais de tenue de compte GEODIS (${fraisFixe.toFixed(2)} EUR/mois, sans tracking) a ajouter manuellement -> total facture attendu = ${totalCalcule.toFixed(2)} EUR HT`);
  for (const p of pdfs) {
    const ecart = round2(totalCalcule - p.taxable);
    infos.push(`Facture PDF ${p.file} : Montant Taxable HT = ${p.taxable.toFixed(2)} EUR, calcule + frais fixe = ${totalCalcule.toFixed(2)} EUR (ecart ${ecart >= 0 ? '+' : ''}${ecart.toFixed(2)} EUR${Math.abs(ecart) < 0.05 ? ' -> OK' : ' -> A VERIFIER'})`);
  }

  return {
    header, rows, recs, importRows, controle, warnings, alerts, infos,
    posteKeys: POSTE_KEYS, gazoleKey: 'Gazole', cfg, pdfs,
    sheetNames: { raw: 'Factures Geodis', import: 'Import CSV' },
    period: dateValidite ? `${dateValidite.slice(6)}_${dateValidite.slice(3, 5)}` : 'export',
  };
}

module.exports = {
  id: 'geodis',
  name: 'Geodis',
  status: 'ready',
  viticolis: true,
  taxeGasoil: 'Calculee ligne a ligne (colonne Surcharge Carburant) ; le PDF sert a controler le total, pas a la calculer',
  method: "Fichier Facture Geodis (xlsx/csv, ~100 colonnes). ATTENTION : les colonnes se decalent d'un mois a l'autre (confirme), reclassement fait par NOM de colonne exclusivement. Poids = Poids origine (kg). E/P non fiable cote Geodis -> laisse vide (alerte 'a verifier'). Frais de tenue de compte mensuel fixe (sans tracking) remonte en info, a ajouter a la main. PDF optionnel : compare le 'Montant Taxable HT' au total calcule (+ frais fixe).",
  inputs: [
    { key: 'facture', label: 'Facture Geodis (xlsx/csv)', accept: '.xlsx,.csv', multiple: true, required: true },
    { key: 'pdf', label: 'Facture(s) PDF Geodis (controle du total)', accept: '.pdf', multiple: true, required: false },
  ],
  // Classeur = CLONE FIDELE du fichier fait a la main (Excel COM/Python), comme Kuehne/Delivengo.
  outputNaming: { workbook: '{period}_Facture Geodis', import: '{period}_Geodis_Import' },
  finalizer: {
    script: '../automatisation/finaliser_geodis.py',
    template: '../Transporteurs/Geodis/2026_06_Facture Geodis.xlsx',
    buildArgs: (files) => [...(files.facture || []), '--pdf', ...(files.pdf || [])],
  },
  process,
};
