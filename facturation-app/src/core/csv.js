// Lecture des CSV transporteurs : encodage latin-1, séparateur configurable, décimale virgule.
const fs = require('fs');

/** Lit un CSV latin-1 -> { header:[...], rows:[[...],...] }. */
function readCsv(path, sep = ';') {
  const text = fs.readFileSync(path).toString('latin1');
  const lines = text.split(/\r?\n/).filter((l) => l.length > 0);
  const rows = lines.map((l) => l.split(sep));
  return { header: rows[0] || [], rows: rows.slice(1) };
}

/** Convertit un montant francais ('1 234,56', '', null) en Number (0 si vide/invalide). */
function num(x) {
  if (x == null) return 0;
  // enleve tous les espaces (normaux + insecables  ), remplace la virgule decimale
  const s = String(x).replace(/[\s ]/g, '').replace(',', '.');
  const v = parseFloat(s);
  return Number.isFinite(v) ? v : 0;
}

/** Index d'une colonne par nom : match exact, puis "contient" (tolerant aux accents casses). */
function colIndex(header, name) {
  let i = header.indexOf(name);
  if (i >= 0) return i;
  return header.findIndex((h) => h && h.includes(name));
}

const round2 = (x) => Math.round(x * 100) / 100;
const round1 = (x) => Math.round(x * 10) / 10;

module.exports = { readCsv, num, colIndex, round2, round1 };
