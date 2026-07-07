// Reclassement generique (le coeur) : somme les colonnes tarifaires en 8 postes, selon la config.
const { num, colIndex, round2 } = require('./csv');

/**
 * Transforme une ligne brute en 8 postes.
 * @param {string[]} row      ligne CSV
 * @param {string[]} header   en-tetes CSV
 * @param {object}   cfg      config transporteur (postes_from_columns, montant_avec_frais...)
 * @param {string[]} warnings accumulateur d'alertes (codes non classes)
 * @returns {object} postes { 'Droits et taxes':..., ..., 'Gazole':... }
 */
function reclasser(row, header, cfg, warnings) {
  const postes = {};
  const used = new Set();
  for (const [poste, cols] of Object.entries(cfg.postes_from_columns)) {
    let total = 0;
    for (const cname of cols) {
      const i = colIndex(header, cname);
      if (i >= 0) { used.add(cname); total += num(row[i]); }
    }
    postes[poste] = round2(total);
  }
  const horsGo = round2(Object.values(postes).reduce((a, b) => a + b, 0));
  const avec = num(row[colIndex(header, cfg.montant_avec_frais)]);
  postes.Gazole = round2(avec - horsGo); // poste residuel

  // garde-fou : un code T_* non classe (et != gazole) qui porte un montant fuiterait dans le Gazole
  for (let i = 0; i < header.length; i++) {
    const name = header[i];
    if (name && name.startsWith('T_') && !used.has(name) && name !== cfg.colonne_gazole) {
      const v = num(row[i]);
      if (v !== 0) warnings.push(`Code tarifaire NON classe '${name}'=${v} (fuite dans Gazole)`);
    }
  }
  return postes;
}

module.exports = { reclasser };
