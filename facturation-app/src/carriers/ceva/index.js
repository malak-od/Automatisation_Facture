// ============================================================================
//  Adaptateur transporteur : CEVA
//  1 ligne CSV = 1 expedition. Reclassement : le montant HT part dans le poste
//  associe au "Code prestation" via une table de correspondance (PDF p.4 :
//  "Si #NA rajouter correspondance dans table correspondance") ; la taxe
//  gazole est une colonne dediee ("Surcharge carburant"), pas un calcul.
//  Pas de classeur de reference disponible en local (contrairement a
//  GLS/Geodis/Mondial Relay) -> table de correspondance seed limitee,
//  a completer au fil des mois (alerte explicite si code inconnu).
// ============================================================================
const { readCsv, num, colIndex, roundUp1, round2 } = require('../../core/csv');
const { validate } = require('../../core/validate');
const cfg = require('./config.json');

function firstDayOfMonth(v) {
  // "JJ/MM/AAAA HH:MM"
  const m = /^(\d{2})\/(\d{2})\/(\d{4})/.exec(String(v || '').trim());
  return m ? `01/${m[2]}/${m[3]}` : null;
}

function process(files) {
  const csvPaths = files.csv || [];
  if (!csvPaths.length) throw new Error('Aucun fichier CSV fourni (attendu : export CEVA).');

  let header = null;
  const rows = [];
  for (const p of csvPaths) {
    const { header: h, rows: rs } = readCsv(p, cfg.csv.sep, cfg.csv.encoding);
    header = h;
    rows.push(...rs.map((r) => (r.length < h.length ? r.concat(Array(h.length - r.length).fill('')) : r)));
  }

  const iTrack = colIndex(header, cfg.cols.tracking);
  const iNom = colIndex(header, cfg.cols.nomLivraison);
  const iPays = colIndex(header, cfg.cols.paysLivraison);
  const iDate = colIndex(header, cfg.cols.dateLivraison);
  const iNbColis = colIndex(header, cfg.cols.nbColis);
  const iPoids = colIndex(header, cfg.cols.poids);
  const iMontant = colIndex(header, cfg.cols.montantHT);
  const iGazole = colIndex(header, cfg.cols.surchargeCarburant);
  const iCode = colIndex(header, cfg.cols.codePrestation);
  const iLibelle = colIndex(header, cfg.cols.libellePrestation);

  const warnings = [];
  let dateValidite = null;

  const recs = rows.map((r) => {
    const tracking = (r[iTrack] || '').trim();
    if (!dateValidite) dateValidite = firstDayOfMonth(r[iDate]);

    const code = (r[iCode] || '').trim();
    const libelle = iLibelle >= 0 ? (r[iLibelle] || '').trim() : '';
    const posteErp = cfg.correspondance_prestation[code];
    if (!posteErp) warnings.push(`Code prestation CEVA non classe '${code}' (${libelle}) [tracking ${tracking}] -> monte dans Fret par defaut, a rajouter dans la table de correspondance`);

    const montant = round2(num(r[iMontant]));
    const gazole = round2(num(r[iGazole]));
    const postes = { DroitsTaxes: 0, Assurance: 0, ZonesEloignees: 0, ColisVolumineux: 0, Adresses: 0, Fret: 0, PlusValueB2C: 0, Gazole: gazole };
    postes[posteErp || 'Fret'] = round2((postes[posteErp || 'Fret'] || 0) + montant);

    const pays = iPays >= 0 ? (r[iPays] || '').trim() : '';

    return {
      tracking, comref: '', dest: iNom >= 0 ? (r[iNom] || '').trim() : '',
      pays, zone: pays === 'FR' ? 'France' : pays,
      colis: iNbColis >= 0 ? num(r[iNbColis]) : 1,
      poids: iPoids >= 0 ? num(r[iPoids]) : 0,
      postes,
      horsGo: round2(montant),
      avecGo: round2(montant + gazole),
      raw: r,
    };
  }).filter((rec) => rec.tracking);

  const POSTE_KEYS = ['DroitsTaxes', 'Assurance', 'ZonesEloignees', 'ColisVolumineux', 'Adresses', 'Fret', 'PlusValueB2C', 'Gazole'];

  const importRows = recs.map((rec) => ({
    Transporteur: cfg.champs_fixes.Transporteur,
    DateValidite: dateValidite || '',
    Ref1: '', Ref2: '', IdClient: '',
    Tracking: rec.tracking, Nom: rec.dest,
    EP: '', Pays: rec.pays, Zone: rec.zone,
    NbrColis: Math.round(rec.colis), Poids: roundUp1(rec.poids),
    Mode: '', TVA: cfg.champs_fixes.tva,
    DroitsTaxes: rec.postes.DroitsTaxes, Assurance: rec.postes.Assurance,
    ZonesEloignees: rec.postes.ZonesEloignees, ColisVolumineux: rec.postes.ColisVolumineux,
    Adresses: rec.postes.Adresses, Fret: rec.postes.Fret,
    PlusValueB2C: rec.postes.PlusValueB2C, TaxeGasoil: rec.postes.Gazole, NbColis: '',
  }));

  const controle = {};
  for (const k of POSTE_KEYS) controle[k] = round2(recs.reduce((s, r) => s + (r.postes[k] || 0), 0));

  const { alerts, infos } = validate(importRows);
  infos.push("E/P et mode d'envoi non determines automatiquement pour CEVA (aucune regle confirmee) -> a completer a la main avant import ERP.");

  return {
    header, rows, recs, importRows, controle, warnings, alerts, infos,
    posteKeys: POSTE_KEYS, gazoleKey: 'Gazole', cfg,
    sheetNames: { raw: 'Facture CEVA', import: 'Import CSV' },
    period: dateValidite ? `${dateValidite.slice(6)}_${dateValidite.slice(3, 5)}` : 'export',
  };
}

module.exports = {
  id: 'ceva',
  name: 'CEVA',
  status: 'ready',
  taxeGasoil: "Colonne dediee 'Surcharge carburant' dans le CSV, pas de calcul",
  method: "Export CSV CEVA. Montant HT route vers un poste via 'Code prestation' -> table de correspondance (a completer au fil des mois, alerte si code inconnu -> tombe dans Fret par defaut, comme le process manuel #N/A). Zone/E-P non confirmes par une reference -> laisses simples/vides avec alerte.",
  inputs: [
    { key: 'csv', label: 'Export CEVA (CSV)', accept: '.csv', multiple: true, required: true },
  ],
  process,
};
