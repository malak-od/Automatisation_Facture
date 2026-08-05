// Ecriture des livrables : import CSV, import XLSX (valeurs seules), classeur de controle.
const fs = require('fs');
const ExcelJS = require('exceljs');
const { IMPORT_COLUMNS } = require('./importSchema');

// Postes en euros : laisses vides quand = 0 (comme l'Excel : IF(=0,"",...))
const BLANK_IF_ZERO = new Set(['DroitsTaxes', 'Assurance', 'ZonesEloignees', 'ColisVolumineux', 'Adresses', 'Fret', 'PlusValueB2C', 'TaxeGasoil']);

function cellValue(o, colDef) {
  let v = o[colDef.key];
  if (BLANK_IF_ZERO.has(colDef.key) && (v === 0 || v === '' || v == null)) return '';
  return v == null ? '' : v;
}

/** Format francais pour le CSV : 25.14 -> "25,14", 300 -> "300", 0.2 -> "0,2". */
function fmtCsv(v) {
  if (v === '' || v == null) return '';
  if (typeof v === 'number') return v.toFixed(2).replace(/0+$/, '').replace(/\.$/, '').replace('.', ',');
  return String(v);
}

/** Fichiers import (csv/xlsx) tries par tracking, pour tous les transporteurs -- le
 * CLASSEUR, lui, garde l'ordre du fichier source (ecrit a part, non affecte ici). */
function sortByTracking(importRows) {
  return [...importRows].sort((a, b) => String(a.Tracking || '').localeCompare(String(b.Tracking || ''), undefined, { numeric: true }));
}

/** Import CSV (latin-1, ';', decimale virgule). */
function writeImportCsv(importRows, filePath) {
  const rows = sortByTracking(importRows);
  const lines = [IMPORT_COLUMNS.map((c) => c.label).join(';')];
  for (const o of rows) lines.push(IMPORT_COLUMNS.map((c) => fmtCsv(cellValue(o, c))).join(';'));
  fs.writeFileSync(filePath, Buffer.from(lines.join('\r\n'), 'latin1'));
}

/** Nom de feuille Excel valide : 31 caracteres max, sans : \ / ? * [ ]. */
function sheetName(name) {
  return String(name || 'Feuille').replace(/[:\\/?*[\]]/g, '_').slice(0, 31);
}

/** Import XLSX = valeurs seules (le fichier a deposer dans l'ERP). */
async function writeImportXlsx(importRows, filePath, importSheetName = 'Import') {
  const rows = sortByTracking(importRows);
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet(sheetName(importSheetName));
  ws.addRow(IMPORT_COLUMNS.map((c) => c.label));
  for (const o of rows) {
    ws.addRow(IMPORT_COLUMNS.map((c) => {
      const v = cellValue(o, c);
      return c.num ? (v === '' ? null : v) : (v === '' ? null : String(v));
    }));
  }
  await wb.xlsx.writeFile(filePath);
}

// ---- Mise en forme (largeurs reprises du fichier fait a la main + couleurs) ----
const COLOR = {
  blue: 'FF305496',      // en-tete : colonnes brutes / identite
  green: 'FF548235',     // en-tete : colonnes calculees / postes
  teal: 'FF2E75B6',      // en-tete TCD
  white: 'FFFFFFFF',
  okFill: 'FFC6EFCE', okText: 'FF006100', // ECART OK (vert)
  totalFill: 'FFFCE4D6',                   // ligne total (peche)
};
// largeurs extraites de Kuehne/2026_06_Facture Kuehne.xlsx
const W_IMPORT = [16, 19, 20.6, 8.4, 10.7, 21.7, 7.6, 7.4, 12.9, 13.1, 12.7, 9.6, 13.9, 7.7, 17.1, 13.6, 18.7, 19.7, 12.4, 12, 18.3, 23, 10];
const W_FK = [11.4, 20.1, 11.4, 15.9, 11.4, 10, 13.9, 14.6, 11.4, 10, 10, 10];
const W_TCD = [20, 16, 42, 6, 14, 13, 11, 15, 15, 11, 11, 14, 10, 14, 13, 9, 10];

/** Applique largeurs + en-tete colore + formats nombre + volet fige. */
function styleHeader(ws, { widths = [], headerColor = COLOR.blue, accentCols = [], accentColor = COLOR.green, moneyCols = [] }) {
  widths.forEach((w, i) => { if (w) ws.getColumn(i + 1).width = w; });
  const hr = ws.getRow(1);
  hr.height = 28;
  hr.eachCell((cell, col) => {
    cell.font = { bold: true, color: { argb: COLOR.white }, size: 10 };
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: accentCols.includes(col) ? accentColor : headerColor } };
    cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
    cell.border = { bottom: { style: 'thin', color: { argb: 'FFBFBFBF' } } };
  });
  moneyCols.forEach((c) => { ws.getColumn(c).numFmt = '#,##0.00'; });
  ws.views = [{ state: 'frozen', ySplit: 1 }];
}

/** Feuille Fichier Kuehne : largeurs auto (voir le contenu) + colonnes colorees (pas l'en-tete).
 *  A->L = vert clair (donnees calculees) ; R & S (Expedition, Ref EDI) = bleu clair. */
function styleFichierKuehne(ws) {
  const GREEN = 'FFE2EFDA';
  const BLUE = 'FFDDEBF7';
  // largeurs "en dur" : ajustees au contenu de chaque colonne (min 8, max 45)
  ws.columns.forEach((col) => {
    let max = 8;
    col.eachCell({ includeEmpty: false }, (cell) => {
      const len = String(cell.value == null ? '' : cell.value).length;
      if (len > max) max = len;
    });
    col.width = Math.min(Math.max(max + 1, 8), 45);
  });
  ws.getRow(1).font = { bold: true };
  ws.getRow(1).height = 22;
  ws.views = [{ state: 'frozen', ySplit: 1 }];
  [3, 4, 5, 6, 7, 8, 9, 10, 11, 12].forEach((c) => { ws.getColumn(c).numFmt = '#,##0.00'; });
  // Couleur DANS les colonnes (toutes les lignes), pas juste l'en-tete
  const fill = (argb) => ({ type: 'pattern', pattern: 'solid', fgColor: { argb } });
  for (let r = 1; r <= ws.rowCount; r++) {
    const row = ws.getRow(r);
    for (let c = 1; c <= 12; c++) row.getCell(c).fill = fill(GREEN); // A -> L
    row.getCell(18).fill = fill(BLUE); // R = Expedition
    row.getCell(19).fill = fill(BLUE); // S = Ref EDI
  }
}

/** Classeur de controle : Fichier Kuehne + TCD (statique) + Kuehne_Import + Reconciliation. */
async function writeWorkbook(result, filePath) {
  const { header, recs, importRows, controle, alerts, warnings, pdfs, posteKeys } = result;
  // Cle du poste "taxe gasoil" : declaree explicitement (gazoleKey) ou deduite par nom connu.
  // Ne PAS supposer 'Gazole' en dur (Kuehne) : DPD/GLS utilisent 'TaxeGasoil' -> sinon le total
  // "hors gazole" de la Reconciliation inclut a tort la taxe gasoil (poste jamais exclu).
  const gazoleKey = result.gazoleKey || (posteKeys.includes('Gazole') ? 'Gazole' : posteKeys.includes('TaxeGasoil') ? 'TaxeGasoil' : null);
  // Noms d'onglets = ceux du transporteur traite (pas "Kuehne" en dur) : soit fournis
  // explicitement par l'adaptateur (result.sheetNames, pour matcher pile le classeur fait
  // a la main), soit derives du nom du transporteur.
  const carrierName = result.carrierName || 'Transporteur';
  const sn = result.sheetNames || {};
  const rawSheetName = sn.raw || `Facture ${carrierName}`;
  const importSheetName = sn.import || `${carrierName}_Import`;
  const wb = new ExcelJS.Workbook();

  // Feuille brute : colonnes de controle + donnees du transporteur. Par defaut = structure
  // Kuehne (le fait-main de reference pour ce format) ; un transporteur avec une structure
  // differente (ex. GLS : pas de Droits et taxes/Assurance/Plus value B2C, mais Zones
  // eloignees/Colis volumineux/Adresses/Fret/Gazole) fournit son propre result.rawCompCols
  // (liste de libelles) + result.rawCompRow(rec) (valeurs correspondantes).
  const fk = wb.addWorksheet(sheetName(rawSheetName));
  const z = (v) => (v === 0 ? null : v); // cellule VIDE si pas de valeur (au lieu de 0,00)
  const compCols = result.rawCompCols || ['ID Client', 'Tracking', 'Total hors GO', 'Total + GO', 'Droits et taxes', 'Assurance', 'Zones eloignees', 'Colis volumineux', 'Adresses', 'Fret', 'Plus value B2C', 'Gazole'];
  const compRow = result.rawCompRow || ((rec) => {
    const p = rec.postes;
    // Vide si 0 : Droits et taxes -> Adresses. On GARDE les 0 pour Plus value B2C et Gazole.
    return [null, rec.tracking, rec.horsGo, rec.avecGo,
      z(p['Droits et taxes']), z(p.Assurance), z(p['Zones eloignees']), z(p['Colis volumineux']), z(p.Adresses),
      p.Fret, p['Plus value B2C'], p.Gazole];
  });
  fk.addRow(compCols.concat(header));
  for (const rec of recs) {
    fk.addRow(compRow(rec).concat(rec.raw.map((v) => (v === '' ? null : v))));
  }

  // Feuille TCD : consolidation par tracking (statique ; version vivante via Excel COM cote Python)
  const tcd = wb.addWorksheet('TCD');
  tcd.addRow(['Tracking', 'ComRef', 'Destinataire', 'Pays', 'Zone', ...posteKeys, 'Total hors GO', 'Total + GO', 'Qte colis', 'Qte poids']);
  const agg = new Map();
  for (const rec of recs) {
    let a = agg.get(rec.tracking);
    if (!a) { a = { comref: rec.comref, dest: rec.dest, pays: rec.pays, zone: rec.zone, postes: {}, hg: 0, ag: 0, colis: 0, poids: 0 }; agg.set(rec.tracking, a); }
    for (const k of posteKeys) a.postes[k] = (a.postes[k] || 0) + rec.postes[k];
    a.hg += rec.horsGo; a.ag += rec.avecGo; a.colis += rec.colis; a.poids += rec.poids;
  }
  const r2 = (x) => Math.round(x * 100) / 100;
  for (const [k, a] of agg) tcd.addRow([k, a.comref, a.dest, a.pays, a.zone, ...posteKeys.map((p) => r2(a.postes[p])), r2(a.hg), r2(a.ag), Math.round(a.colis), Math.round(a.poids * 10) / 10]);

  // Feuille import : valeurs seules
  const imp = wb.addWorksheet(sheetName(importSheetName));
  imp.addRow(IMPORT_COLUMNS.map((c) => c.label));
  for (const o of importRows) imp.addRow(IMPORT_COLUMNS.map((c) => { const v = cellValue(o, c); return c.num ? (v === '' ? null : v) : (v === '' ? null : String(v)); }));

  // Feuille Reconciliation : calcul vs PDF
  const rec = wb.addWorksheet('Reconciliation');
  rec.addRow(['RECONCILIATION']);
  rec.addRow([]); rec.addRow(['Poste calcule', 'Montant EUR']);
  for (const k of posteKeys) rec.addRow([k, r2(controle[k] || 0)]);
  const hg = r2(posteKeys.filter((k) => k !== gazoleKey).reduce((s, k) => s + (controle[k] || 0), 0));
  const totalHt = r2(posteKeys.reduce((s, k) => s + (controle[k] || 0), 0));
  rec.addRow(['Total hors gazole', hg]); rec.addRow(['Gazole', r2((gazoleKey && controle[gazoleKey]) || 0)]);
  rec.addRow(['TOTAL HT calcule', totalHt]); rec.addRow([]);
  if (pdfs && pdfs.length) {
    rec.addRow(['Facture PDF', 'Taxable', 'TVA', 'TTC']);
    let sTax = 0;
    for (const p of pdfs) { rec.addRow([p.file.slice(0, 40), p.taxable, p.tva, p.ttc]); sTax += p.taxable; }
    rec.addRow(['Somme taxable PDF', r2(sTax)]);
    rec.addRow(['ECART calcul - PDF', r2(totalHt - sTax), Math.abs(totalHt - sTax) < 0.05 ? 'OK' : '!!! ECART']);
  }
  rec.addRow([]); rec.addRow([`Alertes de validation : ${alerts.length}`]);
  for (const a of alerts) rec.addRow([a]);
  if (warnings && warnings.length) { rec.addRow([`Alertes reclassement : ${warnings.length}`]); for (const w of warnings) rec.addRow([w]); }

  // --- Mise en forme visuelle (comme le fichier fait a la main) ---
  // Fichier Kuehne : largeurs auto + colonnes colorees (A-L vert, R/S bleu)
  styleFichierKuehne(fk);
  // TCD : en-tete teal, postes + totaux (col 6-15) en accent
  styleHeader(tcd, { widths: W_TCD, headerColor: COLOR.teal, accentCols: [6, 7, 8, 9, 10, 11, 12, 13, 14, 15], moneyCols: [6, 7, 8, 9, 10, 11, 12, 13, 14, 15] });
  // Kuehne_Import : 7 postes (col 15-21) en accent vert
  styleHeader(imp, { widths: W_IMPORT, accentCols: [15, 16, 17, 18, 19, 20, 21], moneyCols: [15, 16, 17, 18, 19, 20, 21] });
  imp.getColumn(12).numFmt = '0.0'; // Poids

  // Reconciliation : titre + surlignage TOTAL / ECART
  rec.getColumn(1).width = 40; rec.getColumn(2).width = 14; rec.getColumn(3).width = 10; rec.getColumn(4).width = 12;
  rec.getRow(1).font = { bold: true, size: 14, color: { argb: COLOR.blue } };
  rec.eachRow((row) => {
    const label = row.getCell(1).value;
    if (typeof label !== 'string') return;
    if (label.startsWith('ECART')) [1, 2, 3].forEach((c) => { row.getCell(c).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COLOR.okFill } }; row.getCell(c).font = { bold: true, color: { argb: COLOR.okText } }; });
    else if (label.startsWith('TOTAL HT')) [1, 2].forEach((c) => { row.getCell(c).font = { bold: true }; row.getCell(c).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COLOR.totalFill } }; });
    else if (label === 'Poste calcule' || label.startsWith('Facture PDF')) [1, 2, 3, 4].forEach((c) => { row.getCell(c).font = { bold: true, color: { argb: COLOR.white } }; row.getCell(c).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COLOR.blue } }; });
  });

  await wb.xlsx.writeFile(filePath);
  return { totalHt };
}

module.exports = { writeImportCsv, writeImportXlsx, writeWorkbook };
