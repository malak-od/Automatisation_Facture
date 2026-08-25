// ============================================================================
//  REGISTRE DES TRANSPORTEURS
//  - Kuehne : implemente.
//  - Les autres (issus des docx de process) : planifies, avec leur methode
//    documentee et les documents attendus. process() lance une erreur "a venir".
//  Pour activer un transporteur : creer src/carriers/<id>/index.js (cf. _template)
//  et remplacer l'entree planifiee par un require(...).
// ============================================================================
const kuehne = require('./carriers/kuehne');
const delivengo = require('./carriers/delivengo');
const dpd = require('./carriers/dpd');
const gls = require('./carriers/gls');
const geodis = require('./carriers/geodis');
const mondialRelay = require('./carriers/mondial_relay');
const lettres = require('./carriers/lettres');
const bls = require('./carriers/bls');
const chronopost = require('./carriers/chronopost');
const colissimo = require('./carriers/colissimo');
const tnt = require('./carriers/tnt');
const fedex = require('./carriers/fedex');
const ups = require('./carriers/ups');

function planned(meta) {
  return {
    status: 'planned',
    inputs: [],
    process() {
      throw new Error(`Transporteur "${meta.name}" pas encore implemente. Methode connue : ${meta.method}`);
    },
    ...meta,
  };
}

// Transporteurs planifies (source : FACTURATION EXCEL.docx / Notes.docx / Check.docx)
const PLANNED = [
];

const carriers = [kuehne, delivengo, dpd, gls, geodis, mondialRelay, lettres, bls, chronopost, colissimo, tnt, fedex, ups, ...PLANNED.map(planned)];
const byId = Object.fromEntries(carriers.map((c) => [c.id, c]));

/** Metadonnees publiques (pour l'UI), sans exposer process(). */
function list() {
  return carriers.map((c) => ({
    id: c.id, name: c.name, status: c.status, viticolis: !!c.viticolis,
    taxeGasoil: c.taxeGasoil, method: c.method, inputs: c.inputs || [],
  }));
}

function get(id) {
  const c = byId[id];
  if (!c) throw new Error(`Transporteur inconnu : ${id}`);
  return c;
}

module.exports = { list, get, carriers };
