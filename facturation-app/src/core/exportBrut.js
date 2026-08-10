// ============================================================================
//  Export WMS "AAAA MM - Export expeditions_brut.xlsx" : donnee partagee en
//  arriere-plan (dossier Automatisation/), pas un document que l'utilisateur
//  doit fournir a chaque traitement (deja le cas pour Delivengo/Lettres).
// ============================================================================
const fs = require('fs');
const path = require('path');
const XLSX = require('xlsx');

/** Cherche 'AAAA MM*brut*.xlsx' dans Automatisation/ pour l'annee/mois donnes. */
function findBrutFile(brutDir, year, month) {
  if (!Number.isFinite(year) || !fs.existsSync(brutDir)) return null;
  const tag = `${year} ${String(month).padStart(2, '0')}`; // ex. "2026 06"
  const f = fs.readdirSync(brutDir).find((x) => x.startsWith(tag) && /brut/i.test(x) && /\.xlsx?$/i.test(x));
  return f ? path.join(brutDir, f) : null;
}

/** Chemins des exports bruts du mois 'AAAA_MM' et de M-1, dans <appRoot>/../Automatisation. */
function findBrutFiles(period, appRoot) {
  const brutDir = path.resolve(appRoot, '../Automatisation');
  const [y, m] = String(period || '').split('_').map(Number);
  const cur = findBrutFile(brutDir, y, m);
  const prev = findBrutFile(brutDir, m === 1 ? y - 1 : y, m === 1 ? 12 : m - 1);
  return [cur, prev].filter(Boolean);
}

/** Chemins de TOUS les exports bruts disponibles dans <appRoot>/../Automatisation, tous mois
 * confondus (pas seulement mois courant/M-1) -- necessaire pour Lettres : le poids d'une
 * expedition peut n'etre renseigne par le WMS que dans un export ANTERIEUR au mois de
 * l'expedition elle-meme (constate : 3 trackings a poids=0 dans l'export de juin ont un vrai
 * poids dans l'export de mai). A utiliser avec parcimonie (cout : lit chaque fichier trouve). */
function findAllBrutFiles(appRoot) {
  const brutDir = path.resolve(appRoot, '../Automatisation');
  if (!fs.existsSync(brutDir)) return [];
  return fs.readdirSync(brutDir)
    .filter((f) => /brut/i.test(f) && /\.xlsx?$/i.test(f) && !f.startsWith('~$'))
    .map((f) => path.join(brutDir, f));
}

/** Lit une ou plusieurs xlsx bruts -> lignes brutes (header exclu). */
function readBrutRows(paths) {
  const rows = [];
  for (const p of paths) {
    const wb = XLSX.readFile(p, { cellDates: true });
    const ws = wb.Sheets[wb.SheetNames[0]];
    const rs = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: '' });
    rows.push(...rs.slice(1));
  }
  return rows;
}

// Colonnes fixes de l'export brut (voir header reel : PRO_TRACKING = index 41,
// DES_PARTICULIER = index 16, INFO_POIDSRETENU = index 34, valeurs
// 'particulier'/'entreprise'/'point relais').
const COL_TRACKING = 41;
const COL_DES_PARTICULIER = 16;
const COL_POIDS_RETENU = 34;

/** Table tracking -> poids (INFO_POIDSRETENU) a partir des lignes brutes WMS deja lues
 * (findBrutFiles/readBrutRows). Sert de repli quand un transporteur ne fournit pas de
 * poids fiable dans son propre fichier (ex. GLS, DPD). */
function poidsParTrackingFromExport(rows) {
  const map = new Map();
  for (const r of rows) {
    const track = String(r[COL_TRACKING] || '').trim();
    const poids = Number(r[COL_POIDS_RETENU]);
    if (track && Number.isFinite(poids) && poids > 0 && !map.has(track)) map.set(track, poids);
  }
  return map;
}

/** DES_PARTICULIER -> E/P ERP. 'point relais' -> E (defaut le plus sur, cf. discussion metier). */
function epFromDesParticulier(v) {
  const s = String(v || '').trim().toLowerCase();
  if (s === 'particulier') return 'P';
  if (s === 'entreprise' || s === 'point relais') return 'E';
  return null;
}

/** Table tracking -> E/P (PRO_TRACKING / DES_PARTICULIER) a partir des lignes brutes WMS
 * deja lues (findBrutFiles/readBrutRows). Sert de repli quand un transporteur ne fournit
 * pas de E/P fiable dans son propre fichier (ex. Geodis). */
function epParTrackingFromExport(rows) {
  const map = new Map();
  for (const r of rows) {
    const track = String(r[COL_TRACKING] || '').trim();
    const ep = epFromDesParticulier(r[COL_DES_PARTICULIER]);
    if (track && ep && !map.has(track)) map.set(track, ep);
  }
  return map;
}

module.exports = { findBrutFiles, findAllBrutFiles, readBrutRows, epParTrackingFromExport, poidsParTrackingFromExport };
