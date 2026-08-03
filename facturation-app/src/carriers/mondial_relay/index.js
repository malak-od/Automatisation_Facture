// ============================================================================
//  Adaptateur transporteur : MONDIAL RELAY
//  Plusieurs fichiers CSV "Annexe_..." (1 par facture/compte). 1 ligne = 1 colis.
//  Reclassement : Fret = Montant transport (poids mesure) [+ Complement si
//  EXACTEMENT 0,03] ; Zones eloignees = Complement (sauf 0 ou 0,03) ; Zone =
//  lookup Mode+Pays dans la table 'Pays' du classeur fait a la main (deja
//  normalisee, gere 24RC/LCC->24R nativement). Indexation gasoil est reelle
//  mais n'a pas de colonne dans le fichier d'import fait a la main -> geree a
//  part dans l'ERP (note terrain "identique au mois precedent"), remontee en
//  info uniquement pour le controle du total.
// ============================================================================
const { readCsv, num, colIndex, roundUp1, round2 } = require('../../core/csv');
const { validate } = require('../../core/validate');
const cfg = require('./config.json');

const POSTE_KEYS = ['ZonesEloignees', 'Fret'];

function normalizeTracking(raw) {
  const s = String(raw || '').trim();
  return s.length > 0 && s.length < 8 ? s.padStart(8, '0') : s;
}

function firstDayOfMonth(v) {
  // format observe : AAAAMMJJ (ex. "20260623")
  const m = /^(\d{4})(\d{2})(\d{2})$/.exec(String(v || '').trim());
  return m ? `01/${m[2]}/${m[1]}` : null;
}

function process(files) {
  const csvPaths = files.csv || [];
  if (!csvPaths.length) throw new Error('Aucun fichier CSV fourni (attendu : Annexe_*.csv Mondial Relay).');

  let header = null;
  const rows = [];
  for (const p of csvPaths) {
    const { header: h, rows: rs } = readCsv(p, cfg.csv.sep, cfg.csv.encoding);
    header = h;
    rows.push(...rs.map((r) => (r.length < h.length ? r.concat(Array(h.length - r.length).fill('')) : r)));
  }

  const iTrack = colIndex(header, cfg.cols.trackingBrut);
  const iNom = colIndex(header, cfg.cols.nom);
  const iPays = colIndex(header, cfg.cols.pays);
  const iMode = colIndex(header, cfg.cols.mode);
  const iNbColis = colIndex(header, cfg.cols.nbColis);
  const iPoidsAnnonce = colIndex(header, cfg.cols.poidsAnnonce);
  const iPoidsMesure = colIndex(header, cfg.cols.poidsMesure);
  const iComplement = colIndex(header, cfg.cols.complement);
  const iMontantTransport = colIndex(header, cfg.cols.montantTransport);
  const iGazole = colIndex(header, cfg.cols.gazole);
  const iDate = colIndex(header, cfg.cols.date);

  const warnings = [];
  let dateValidite = null;
  let gazoleTotal = 0;

  const recs = [];
  for (const r of rows) {
    const trackBrut = (r[iTrack] || '').trim();
    if (!trackBrut || trackBrut === '0') continue; // ligne "vide" (voir formule LEN/IF du classeur)
    if (!dateValidite) dateValidite = firstDayOfMonth(r[iDate]);

    const complement = round2(num(r[iComplement]));
    const montant = num(r[iMontantTransport]);
    const foldComplement = Math.abs(complement - 0.03) < 0.001;
    const fret = round2(foldComplement ? montant + complement : montant);
    const zonesEloignees = (complement === 0 || foldComplement) ? 0 : complement;
    const gazoleLigne = num(r[iGazole]);
    gazoleTotal += gazoleLigne;

    const pays = (r[iPays] || '').trim();
    // 24RC et LCC -> 24R (mode envoi) : sinon zones inconnues -> avaries d'import (cf. notes terrain + PDF p.9)
    const modeBrut = (r[iMode] || '').trim();
    const mode = (modeBrut === '24RC' || modeBrut === 'LCC') ? '24R' : modeBrut;
    const zoneKey = `${mode}-${pays}`;
    const zone = cfg.zone_table[zoneKey];
    if (!zone) warnings.push(`Zone inconnue pour la cle '${zoneKey}' (tracking ${trackBrut}) -> table 'Pays' a completer`);

    const poidsGr = Math.max(num(r[iPoidsAnnonce]), num(r[iPoidsMesure]));

    recs.push({
      tracking: normalizeTracking(trackBrut),
      dest: iNom >= 0 ? (r[iNom] || '').trim() : '',
      pays, zone: zone || 'zone inconnue',
      colis: iNbColis >= 0 ? num(r[iNbColis]) : 1,
      poids: poidsGr / 1000,
      mode,
      postes: { ZonesEloignees: zonesEloignees, Fret: fret },
      horsGo: round2(zonesEloignees + fret),
      avecGo: round2(zonesEloignees + fret + gazoleLigne),
      raw: r,
    });
  }

  const importRows = recs.map((rec) => ({
    Transporteur: cfg.champs_fixes.Transporteur,
    DateValidite: dateValidite || '',
    Ref1: '', Ref2: '', IdClient: '',
    Tracking: rec.tracking, Nom: rec.dest,
    EP: cfg.champs_fixes['E/P'], Pays: rec.pays, Zone: rec.zone,
    NbrColis: Math.round(rec.colis), Poids: roundUp1(rec.poids),
    Mode: rec.mode, TVA: cfg.champs_fixes.tva,
    DroitsTaxes: 0, Assurance: 0,
    ZonesEloignees: rec.postes.ZonesEloignees, ColisVolumineux: 0, Adresses: 0,
    Fret: rec.postes.Fret, PlusValueB2C: 0, TaxeGasoil: 0, NbColis: '',
  }));

  const controle = {};
  for (const k of POSTE_KEYS) controle[k] = round2(recs.reduce((s, r) => s + (r.postes[k] || 0), 0));

  const { alerts, infos } = validate(importRows);
  const totalHorsGazole = round2(POSTE_KEYS.reduce((s, k) => s + (controle[k] || 0), 0));
  infos.push(`Indexation gasoil (${round2(gazoleTotal).toFixed(2)} EUR, "identique au mois precedent" selon le process) n'est pas dans une colonne de l'import -> geree a part dans l'ERP. Total HT reel attendu (hors frais fixes/collectes) = ${round2(totalHorsGazole + gazoleTotal).toFixed(2)} EUR`);

  return {
    header, rows, recs, importRows, controle, warnings, alerts, infos,
    posteKeys: POSTE_KEYS, cfg,
    sheetNames: { raw: 'Facture Mondial Relay', import: 'Fichier import' },
    period: dateValidite ? `${dateValidite.slice(6)}_${dateValidite.slice(3, 5)}` : 'export',
  };
}

module.exports = {
  id: 'mondial_relay',
  name: 'Mondial Relay',
  status: 'ready',
  taxeGasoil: "Reelle (indexation gasoil par ligne) mais geree hors fichier d'import, cote ERP -- 'identique au mois precedent' selon le process",
  method: "Fichiers Annexe_*.csv (1 par facture, dossier). Fret = Montant transport poids mesure (+Complement si =0,03 pile) ; Zones eloignees = Complement sinon. Zone = Mode+Pays via la table de correspondance du classeur (normalise deja 24RC/LCC -> 24R).",
  inputs: [
    { key: 'csv', label: 'Fichiers Mondial Relay (dossier Annexe_*.csv)', accept: '.csv', multiple: true, required: true },
  ],
  process,
};
