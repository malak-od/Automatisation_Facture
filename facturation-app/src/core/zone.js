// Calcul de la zone de tarification : reproduit le XLOOKUP + concatenation d'Excel.
const { colIndex } = require('./csv');

/** Construit le lookup dept indexe par 'Ref EDI' (echoue -> zone 0France00, comme Excel). */
function buildDeptLookup(rows, header, cfg) {
  const iEdi = colIndex(header, cfg.tracking.primary);
  const iDe = colIndex(header, cfg.identite.dept_exp);
  const iDd = colIndex(header, cfg.identite.dept_dest);
  const de = {};
  const dd = {};
  for (const r of rows) {
    const e = (r[iEdi] || '').trim();
    if (e) {
      if (!(e in de)) de[e] = (r[iDe] || '').trim();
      if (!(e in dd)) dd[e] = (r[iDd] || '').trim();
    }
  }
  return { de, dd };
}

/** Zone = deptExp & (FR?France:etranger) & pad2(deptDest). Seul le dept DEST est pade. */
function zone(track, pays, de, dd) {
  let dexp = de[track] || '0';
  let ddest = dd[track] || '0';
  if (/^\d+$/.test(dexp)) dexp = String(parseInt(dexp, 10));
  if (/^\d+$/.test(ddest)) ddest = String(parseInt(ddest, 10)).padStart(2, '0');
  return `${dexp}${pays === 'FR' ? 'France' : 'étranger'}${ddest}`;
}

module.exports = { buildDeptLookup, zone };
