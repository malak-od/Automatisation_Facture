// ============================================================================
//  Adaptateur transporteur : DELIVENGO (lettres suivies, LPPAQ)
//  Le classeur final "AAAA_MM_Delivengo_LPPAQ.xlsx" est un CLONE FIDELE du
//  fichier fait a la main (feuilles Pays + Fichier import, formules, couleurs),
//  regenere par le finaliseur Python (Excel COM) a partir de l'export du suivi.
//
//  Comme Kuehne : produit AUSSI un import.csv/import.xlsx a cote du classeur
//  (memes 23 colonnes IMPORT_COLUMNS que tous les transporteurs, l'ERP impose
//  une norme unique) -- mais contrairement a Kuehne, tout le calcul (poids,
//  zone, TVA via XLOOKUP sur la feuille "Pays") vit dans les FORMULES Excel du
//  finaliseur Python, pas dans du JS independant. On fait donc tourner le
//  finaliseur une fois ici (vers un fichier temporaire), on relit l'onglet
//  "Fichier import" une fois calcule pour construire importRows, et on
//  transmet ce classeur deja pret a server.js (prebuiltWorkbookPath) pour
//  eviter de refaire tourner Excel COM une 2e fois pour le meme resultat.
// ============================================================================
const os = require('os');
const path = require('path');
const fs = require('fs');
const { execFile } = require('child_process');
const execFileAsync = require('util').promisify(execFile);
const XLSX = require('xlsx');
const { round2 } = require('../../core/csv');
const { IMPORT_COLUMNS } = require('../../core/importSchema');

// positions des colonnes utiles dans l'export du suivi (index 0-based)
const COL = { remise: 7, suivi: 9, nom: 11, pays: 16, statut: 24, poids: 30 };

const POSTE_KEYS = ['DroitsTaxes', 'Assurance', 'ZonesEloignees', 'ColisVolumineux', 'Adresses', 'Fret', 'PlusValueB2C', 'TaxeGasoil'];

// En-tetes de la feuille "Fichier import" du classeur -> cle IMPORT_COLUMNS.
// "Pays" apparait 2x dans la feuille (code, colonne J, et nom, colonne X en
// bout de tableau) : indexOf() prend le PREMIER (J, le bon -- code pays ERP).
const HEADER_TO_KEY = {
  ' Transporteur ': 'Transporteur', 'Date validité tarif': 'DateValidite',
  'Réf.1': 'Ref1', 'Réf. 2': 'Ref2', 'Numéro de suivi': 'Tracking',
  'Destinataire nom': 'Nom', 'E / P': 'EP', 'Pays': 'Pays', ' Zone ': 'Zone',
  ' Nbr Colis ': 'NbrColis', ' Poids ': 'Poids', 'mode envoi': 'Mode', ' TVA ': 'TVA',
  ' Droits et taxes ': 'DroitsTaxes', ' Assurance ': 'Assurance',
  ' Zones éloignées ': 'ZonesEloignees', ' Colis volumineux ': 'ColisVolumineux',
  ' Adresses ': 'Adresses', ' Frêt ': 'Fret', ' plus-value BtoC ': 'PlusValueB2C',
  'Taxe Gasoil': 'TaxeGasoil',
};

const EXCEL_EPOCH_MS = Date.UTC(1899, 11, 30);
function excelSerialToDateStr(v) {
  if (v == null || v === '') return '';
  if (v instanceof Date) return `01/${String(v.getMonth() + 1).padStart(2, '0')}/${v.getFullYear()}`;
  if (typeof v === 'number') {
    const d = new Date(EXCEL_EPOCH_MS + v * 86400000);
    return `01/${String(d.getUTCMonth() + 1).padStart(2, '0')}/${d.getUTCFullYear()}`;
  }
  return String(v).trim();
}

// Colonnes de calcul intermediaire dans le CLASSEUR (servent a M = MAX(P,Q)) : elles y
// restent remplies (poids brut WMS / poids export du suivi), mais le fichier import ERP
// separe (23 colonnes standard) les remet a 0 -- ce ne sont pas des postes de facturation.
const ZERO_IN_IMPORT = new Set(['DroitsTaxes', 'Assurance']);

/** Relit la feuille "Fichier import" du classeur DEJA FINALISE (formules calculees,
 * purge appliquee) -> construit importRows au format standard IMPORT_COLUMNS. */
function readFichierImport(xlsxPath) {
  const wb = XLSX.readFile(xlsxPath, { cellDates: true });
  const ws = wb.Sheets['Fichier import'];
  if (!ws) throw new Error('Classeur Delivengo finalisé : onglet "Fichier import" introuvable.');
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: '' });
  const header = rows[0].map((h) => String(h || ''));
  const colFor = {};
  for (const [label, key] of Object.entries(HEADER_TO_KEY)) colFor[key] = header.indexOf(label);

  const importRows = [];
  for (const r of rows.slice(1)) {
    if (!r || !r.length || !String(r[colFor.Tracking] ?? '').trim()) continue;
    const o = {};
    for (const c of IMPORT_COLUMNS) {
      if (ZERO_IN_IMPORT.has(c.key)) { o[c.key] = 0; continue; }
      const ci = colFor[c.key];
      let v = ci >= 0 ? r[ci] : '';
      if (c.key === 'DateValidite') v = excelSerialToDateStr(v);
      else if (c.num) v = v === '' || v == null ? 0 : Number(v) || 0;
      else v = v == null ? '' : String(v);
      o[c.key] = v;
    }
    importRows.push(o);
  }
  return importRows;
}

/** Args du finaliseur (template + export + brut du mois/mois-1) : reutilise a la
 * fois par process() (calcul interne) et par finalizer.buildArgs (repli). */
function computeFinalizerArgs(files, period, appRoot) {
  const brutDir = path.resolve(appRoot, '../automatisation');
  const [y, m] = String(period || '').split('_').map(Number);
  const findBrut = (yy, mm) => {
    if (!Number.isFinite(yy) || !fs.existsSync(brutDir)) return null;
    const tag = `${yy} ${String(mm).padStart(2, '0')}`; // ex. "2026 06"
    const f = fs.readdirSync(brutDir).find((x) => x.startsWith(tag) && /brut/i.test(x) && /\.xlsx?$/i.test(x));
    return f ? path.join(brutDir, f) : null;
  };
  const brut = [findBrut(y, m), findBrut(m === 1 ? y - 1 : y, m === 1 ? 12 : m - 1)].filter(Boolean);
  return ['--export', ...(files.export || []), '--brut', ...brut];
}

async function process(files) {
  const p = (files.export || [])[0];
  if (!p) throw new Error('Aucun export du suivi Delivengo fourni (.xls).');
  const wb = XLSX.readFile(p);
  const ws = wb.Sheets[wb.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, raw: false }).slice(1).filter((r) => r && r.length && r[COL.suivi]);

  // periode = 1ere date de remise (JJ/MM/AAAA) -> AAAA_MM
  let period = 'export';
  for (const r of rows) {
    const m = String(r[COL.remise] || '').match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
    if (m) { period = `${m[3]}_${m[2]}`; break; }
  }

  // controles simples (le detail/formules sont dans le classeur genere)
  const alerts = [];
  rows.forEach((r, i) => {
    if (!String(r[COL.pays] || '').trim()) alerts.push(`L${i + 2} ${r[COL.suivi]}: pays destinataire manquant`);
    if (!String(r[COL.suivi] || '').trim()) alerts.push(`L${i + 2}: numero de suivi manquant`);
  });

  // Fait tourner le finaliseur (Excel COM) une fois, vers un fichier temporaire,
  // pour en extraire importRows -- server.js reprendra ce MEME fichier deja
  // calcule comme classeur livre (prebuiltWorkbookPath), pas la peine de relancer
  // Excel une 2e fois pour le meme resultat.
  const infos = [];
  let importRows = [];
  const tmpWb = path.join(os.tmpdir(), `delivengo_${Date.now()}.xlsx`);
  try {
    const scriptAbs = path.resolve(__dirname, '../../../..', 'automatisation', 'finaliser_delivengo.py');
    const templateAbs = path.resolve(__dirname, '../../../..', 'Transporteurs', 'Delivengo', '2026_06_Delivengo_LPPAQ.xlsx');
    const args = computeFinalizerArgs(files, period, path.resolve(__dirname, '../../..'));
    const { stdout } = await execFileAsync('python', [scriptAbs, templateAbs, tmpWb, ...args], { windowsHide: true, maxBuffer: 20 * 1024 * 1024 });
    for (const line of String(stdout || '').split(/\r?\n/)) {
      const m2 = line.match(/^INFO_POIDS_MANQUANT:(.+)$/);
      if (m2) infos.push(`Poids introuvable dans les exports bruts : ${m2[1].trim()} (plancher 0,15 appliqué — à vérifier/saisir à la main)`);
    }
    importRows = readFichierImport(tmpWb);
  } catch (e) {
    const stderr = (e.stderr && e.stderr.toString) ? e.stderr.toString() : String(e.stderr || e.message || '');
    alerts.push(`Finaliseur Delivengo KO : ${stderr.split(/\r?\n/).slice(0, 4).join(' | ') || e.message} (import ERP non généré, classeur en repli générique)`);
  }

  const controle = {};
  for (const k of POSTE_KEYS) controle[k] = round2(importRows.reduce((s, o) => s + (Number(o[k]) || 0), 0));

  return {
    lignes: rows.length, period,
    importRows, controle, warnings: [], alerts, infos, posteKeys: POSTE_KEYS, gazoleKey: 'TaxeGasoil',
    sheetNames: { raw: 'Pays', import: 'Fichier import' },
    prebuiltWorkbookPath: importRows.length ? tmpWb : null,
  };
}

module.exports = {
  id: 'delivengo',
  name: 'Delivengo',
  status: 'ready',
  taxeGasoil: 'Pas de taxe gasoil (frêt = 1,0).',
  method: "Export du suivi Delivengo (.xls). Le classeur final (feuille 'Fichier import') est un clone du fichier fait a la main : A=Date remise, F=Statut, G=N° suivi, H=Destinataire, X=Pays ; J/K/M/O/W = formules (table Pays). P (Droits et taxes) et Q (Assurance) ne sont que des colonnes de calcul intermediaire (poids export brut recherchex / poids Delivengo propre /1000, servent a M=MAX(P,Q)) : videes en fin de traitement, comme Statut et Taxe Gasoil si 0. Import ERP (23 colonnes standard) relu depuis ce meme classeur une fois calcule.",
  inputs: [
    { key: 'export', label: 'Export du suivi Delivengo (.xls)', accept: '.xls,.xlsx', multiple: false, required: true },
  ],
  outputNaming: { workbook: '{period}_Delivengo_LPPAQ', import: '{period}_Delivengo_Import' },
  finalizer: {
    script: '../automatisation/finaliser_delivengo.py',
    template: '../Transporteurs/Delivengo/2026_06_Delivengo_LPPAQ.xlsx',
    // Repli SEULEMENT si prebuiltWorkbookPath (produit dans process()) est absent/perdu.
    buildArgs: (files, period, appRoot) => computeFinalizerArgs(files, period, appRoot),
  },
  process,
};
