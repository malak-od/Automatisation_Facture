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
  { id: 'ups', name: 'UPS', viticolis: true, taxeGasoil: 'Filles Viticolis',
    method: "Telecharger les factures du Billing (sauf 0EUR et notes de credit, filtrer le compte 80X7Y5), CSV, trier trackings, reclasser. E/P via export brut m/m-1/m-2. Colis 1Z79 -> demande d'avoir.",
    inputs: [{ key: 'csv', label: 'Factures Billing UPS (CSV/Excel)', accept: '.csv,.xlsx', multiple: true, required: true }] },
  { id: 'colissimo', name: 'Colissimo', taxeGasoil: 'Site La Poste (coeff. energetique)',
    method: "CSV prestation Colis + douanes, trier, ajouter les postes en fin de colonne, MAJ TCD, comparer PDF HT + indemnisation.",
    inputs: [{ key: 'csv', label: 'CSV prestation Colis + douanes', accept: '.csv', multiple: true, required: true }] },
  { id: 'fedex', name: 'FedEx', taxeGasoil: 'Facture PDF (TG France/International)',
    method: "Ouvrir CSV shipment detail, remplacer . par ,, E/P via export brut m/m-1, verifier FICP (IE/RE), MAJ TCD, gerer droits & taxes.",
    inputs: [{ key: 'csv', label: 'Shipment detail (CSV)', accept: '.csv', multiple: true, required: true },
             { key: 'pdf', label: 'Facture PDF', accept: '.pdf', multiple: true, required: false }] },
  { id: 'tnt', name: 'TNT', taxeGasoil: 'Facture PDF (taux officiel)',
    method: "Etendre colonnes, ajouter les postes, MAJ TCD, NE PAS convertir les trackings en nombre, verifier tarifs via surcharge carburant PDF.",
    inputs: [{ key: 'facture', label: 'Facture TNT', accept: '.xlsx,.csv', multiple: false, required: true }] },
];

const carriers = [kuehne, delivengo, dpd, gls, geodis, mondialRelay, lettres, bls, chronopost, ...PLANNED.map(planned)];
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
