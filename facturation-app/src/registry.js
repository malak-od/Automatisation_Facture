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
  { id: 'geodis', name: 'Geodis', viticolis: true, taxeGasoil: 'Dans la facture PDF',
    method: "Copier la facture Geodis (attention decalage colonnes), MAJ TCD, comparer ecart PDF. Poids/colis via export affretrans (recherchex par tracking).",
    inputs: [{ key: 'facture', label: 'Facture Geodis', accept: '.xlsx,.csv', multiple: false, required: true },
             { key: 'pdf', label: 'Facture PDF (taxe gasoil)', accept: '.pdf', multiple: true, required: false }] },
  { id: 'ceva', name: 'CEVA', taxeGasoil: 'A demander',
    method: "Etendre colonnes, si #NA ajouter dans table de correspondance, ajouter les postes en fin de colonne, MAJ TCD, comparer ecart PDF.",
    inputs: [{ key: 'facture', label: 'Facture CEVA', accept: '.xlsx,.csv', multiple: false, required: true }] },
  { id: 'colissimo', name: 'Colissimo', taxeGasoil: 'Site La Poste (coeff. energetique)',
    method: "CSV prestation Colis + douanes, trier, ajouter les postes en fin de colonne, MAJ TCD, comparer PDF HT + indemnisation.",
    inputs: [{ key: 'csv', label: 'CSV prestation Colis + douanes', accept: '.csv', multiple: true, required: true }] },
  { id: 'chronopost', name: 'Chronopost', taxeGasoil: 'Site Chronopost (routier + aerien)',
    method: "Rassembler les tableaux, trier trackings, sommes eco/surete/gazole, taux gasoil routier+aerien, MAJ TCD, supprimer lignes CAP/ECO/SUR de l'import.",
    inputs: [{ key: 'csv', label: 'Tableaux Chronopost', accept: '.csv,.xlsx', multiple: true, required: true }] },
  { id: 'fedex', name: 'FedEx', taxeGasoil: 'Facture PDF (TG France/International)',
    method: "Ouvrir CSV shipment detail, remplacer . par ,, E/P via export brut m/m-1, verifier FICP (IE/RE), MAJ TCD, gerer droits & taxes.",
    inputs: [{ key: 'csv', label: 'Shipment detail (CSV)', accept: '.csv', multiple: true, required: true },
             { key: 'pdf', label: 'Facture PDF', accept: '.pdf', multiple: true, required: false }] },
  { id: 'tnt', name: 'TNT', taxeGasoil: 'Facture PDF (taux officiel)',
    method: "Etendre colonnes, ajouter les postes, MAJ TCD, NE PAS convertir les trackings en nombre, verifier tarifs via surcharge carburant PDF.",
    inputs: [{ key: 'facture', label: 'Facture TNT', accept: '.xlsx,.csv', multiple: false, required: true }] },
  { id: 'mondial_relay', name: 'Mondial Relay', taxeGasoil: 'Identique au mois precedent',
    method: "Rassembler fichiers (dossier), transformer . en ,, MAJ TCD controle xls/pdf, nombre de collectes (facture LR), modifier 24RC->24R et LCC->24R.",
    inputs: [{ key: 'csv', label: 'Fichiers Mondial Relay (dossier)', accept: '.csv,.xlsx', multiple: true, required: true }] },
  { id: 'lettres', name: 'Lettres (Suivie / Prepa / SLAACE)', taxeGasoil: 'Pas de TG',
    method: "Directement dans le CSV d'import : filtrer expedition brute par lettre, corriger transporteur si tracking = UPS/MR/GLS, 3 fichiers (SLAACE Allemagne, Suivie, Prepa).",
    inputs: [{ key: 'export', label: 'Export expedition brute', accept: '.csv,.xlsx', multiple: true, required: true }] },
  { id: 'bls', name: 'BLS', viticolis: true, taxeGasoil: 'A definir',
    method: "Pas encore demarre. Verifier que les trackings sont dans l'ERP (rapprochement fichier affretement), copier facture BLS, MAJ TCD, comparer ecart PDF.",
    inputs: [{ key: 'facture', label: 'Facture BLS', accept: '.xlsx,.csv', multiple: false, required: true }] },
];

const carriers = [kuehne, delivengo, dpd, gls, ...PLANNED.map(planned)];
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
