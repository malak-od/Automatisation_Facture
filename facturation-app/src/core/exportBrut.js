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

module.exports = { findBrutFiles, readBrutRows };
