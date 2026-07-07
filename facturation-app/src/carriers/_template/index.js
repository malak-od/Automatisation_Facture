// ============================================================================
//  GABARIT — pour implementer un nouveau transporteur.
//  1. Copier ce dossier en src/carriers/<id>/
//  2. Creer un config.json (mapping colonnes -> postes) si le transporteur
//     a une structure "a la Kuehne", ou coder une logique dediee ci-dessous.
//  3. Dans src/registry.js, remplacer l'entree planifiee par : require('./carriers/<id>')
// ============================================================================
// const cfg = require('./config.json');
// const { readCsv } = require('../../core/csv');
// const { reclasser } = require('../../core/reclass');
// ... reutiliser le moteur commun de src/core/

function process(files) {
  // files = { <inputKey>: [chemins...] }
  // Doit renvoyer : { header, rows, recs, importRows, controle, warnings, alerts, posteKeys, cfg }
  throw new Error('Gabarit non implemente.');
}

module.exports = {
  id: 'exemple',
  name: 'Exemple',
  status: 'planned',
  taxeGasoil: 'Source de la taxe gasoil',
  method: 'Decrire la methode de traitement de ce transporteur.',
  inputs: [
    { key: 'csv', label: 'Fichier(s) source', accept: '.csv', multiple: true, required: true },
  ],
  process,
};
