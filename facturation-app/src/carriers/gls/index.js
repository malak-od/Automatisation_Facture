// ============================================================================
//  Adaptateur transporteur : GLS
//  CSV "BCF" GLS : 1 ligne = 1 charge (colis + type de charge, colonne "Hierarchie
//  produits"). Plusieurs lignes par colis (fret + surcharges) -> regroupement par
//  "Numero de colis" avant reclassement en postes ERP.
//  Mapping categorie -> poste : onglet "Categories" du classeur fait a la main
//  (2026_06_Facture GLS.xlsx), confirme par la video process (mai 2025).
// ============================================================================
const path = require('path');
const { readCsv, num, colIndex, roundUp1, round2 } = require('../../core/csv');
const { validate } = require('../../core/validate');
const { readBrutRows } = require('../../core/exportBrut');
const cfg = require('./config.json');

const POSTE_KEYS = ['ZonesEloignees', 'ColisVolumineux', 'Adresses', 'Fret', 'TaxeGasoil'];

/** Normalise une categorie (poste GLS interne) -> cle ERP, via config.poste_to_erp. */
function erpPosteFor(categorie, warnings, ligneRef) {
  const posteGls = cfg.categorie_to_poste[categorie];
  if (!posteGls) {
    warnings.push(`Categorie GLS non classee '${categorie}' (${ligneRef}) -> montant verse dans Fret`);
    return 'Fret';
  }
  return cfg.poste_to_erp[posteGls] || 'Fret';
}

/** Poids (kg) a partir du prix de base, via la grille tarifaire PA GLS (plus proche prix). */
function poidsFromPrix(prix) {
  if (!prix) return null;
  let best = null, bestDiff = Infinity;
  for (const [kg, p] of cfg.poids_prix_table) {
    const diff = Math.abs(p - prix);
    if (diff < bestDiff) { bestDiff = diff; best = kg; }
  }
  // tolerance : le prix doit etre raisonnablement proche d'un palier connu (sinon pas fiable)
  return bestDiff <= 0.05 ? best : null;
}

/** Table tracking -> poids (INFO_POIDSRETENU) a partir des lignes brutes WMS deja lues
 * (findBrutFiles/readBrutRows, core/exportBrut). Sert de 3e repli quand le CSV BCF GLS
 * ne porte le poids sur aucune ligne du colis. */
function poidsParTrackingFromExport(rows) {
  const map = new Map();
  const { tracking: iT, poids: iP } = cfg.export_brut_cols;
  for (const r of rows) {
    const track = String(r[iT] || '').trim();
    const poids = Number(r[iP]);
    if (track && Number.isFinite(poids) && poids > 0 && !map.has(track)) map.set(track, poids);
  }
  return map;
}

function firstDayOfMonth(dateStr) {
  // formats vus : "30.06.2026" ou "JJ.MM.AAAA"
  const m = /^(\d{2})\.(\d{2})\.(\d{4})$/.exec(String(dateStr || '').trim());
  return m ? `01/${m[2]}/${m[3]}` : null;
}

function process(files) {
  const csvPaths = files.csv || [];
  if (!csvPaths.length) throw new Error('Aucun fichier CSV fourni (attendu : export BCF GLS).');

  let header = null;
  const rows = [];
  for (const p of csvPaths) {
    const { header: h, rows: rs } = readCsv(p, cfg.csv.sep, cfg.csv.encoding);
    header = h;
    rows.push(...rs.map((r) => (r.length < h.length ? r.concat(Array(h.length - r.length).fill('')) : r)));
  }

  const iTracking = colIndex(header, cfg.tracking);
  const iCategorie = colIndex(header, cfg.categorie_key);
  const iPays = colIndex(header, cfg.pays_col);
  const iDest = colIndex(header, cfg.dest_nom_col);
  const iPoids = colIndex(header, cfg.poids_col);
  const iDateFacture = colIndex(header, cfg.date_facture_col);
  const iAmountsHG = cfg.amount_cols_hors_gazole.map((c) => colIndex(header, c));
  const iGazole = colIndex(header, cfg.amount_col_gazole);
  const iRef = header.findIndex((h) => h && h.startsWith('Client Référence'));

  const warnings = [];
  let dateValidite = null;

  // 1. Regroupement par tracking (plusieurs lignes de charge par colis) -- pour le calcul
  //    des postes ERP (Poids/Zone/TVA sont determines au niveau du colis, pas de la charge).
  const byTracking = new Map();
  rows.forEach((r, idx) => {
    const track = (r[iTracking] || '').trim();
    if (!track) return;
    if (!dateValidite) dateValidite = firstDayOfMonth(r[iDateFacture]);

    let g = byTracking.get(track);
    if (!g) {
      g = {
        tracking: track, pays: (r[iPays] || '').trim(), dest: (r[iDest] || '').trim(),
        ref: '', poidsBrut: 0, poidsFretPrix: null,
        postes: { ZonesEloignees: 0, ColisVolumineux: 0, Adresses: 0, Fret: 0, TaxeGasoil: 0 },
      };
      byTracking.set(track, g);
    }

    const categorie = (r[iCategorie] || '').trim();
    const posteErp = erpPosteFor(categorie, warnings, `L${idx + 2} ${track}`);
    const montantHG = iAmountsHG.reduce((s, i) => s + (i >= 0 ? num(r[i]) : 0), 0);
    g.postes[posteErp] = round2(g.postes[posteErp] + montantHG);
    g.postes.TaxeGasoil = round2(g.postes.TaxeGasoil + (iGazole >= 0 ? num(r[iGazole]) : 0));

    const poidsLigne = iPoids >= 0 ? num(r[iPoids]) : 0;
    if (poidsLigne > g.poidsBrut) g.poidsBrut = poidsLigne;
    if ((cfg.poste_to_erp[cfg.categorie_to_poste[categorie]] || 'Fret') === 'Fret' && cfg.categorie_to_poste[categorie] === 'Fret') {
      // ligne de fret "de base" (pas une surcharge) : sa valeur de base sert de repli poids->prix
      const prixLigne = num(r[iAmountsHG[0]]); // 'Valeur de Base' = 1er de la liste
      if (prixLigne > 0) g.poidsFretPrix = prixLigne;
    }
    if (!g.ref && iRef >= 0 && r[iRef]) g.ref = r[iRef].trim();
    if (!g.dest && iDest >= 0 && r[iDest]) g.dest = r[iDest].trim();
    if (!g.pays && iPays >= 0 && r[iPays]) g.pays = r[iPays].trim();
  });

  // Export du mois (upload manuel OBLIGATOIRE, demande utilisateur 2026-08-25) : repli poids
  // 3e niveau.
  const period = dateValidite ? `${dateValidite.slice(6)}_${dateValidite.slice(3, 5)}` : null;
  const brutPaths = files.brut || [];
  const poidsExportBrut = poidsParTrackingFromExport(readBrutRows(brutPaths));

  // 2. Poids final + zone/TVA par colis (poidsFinal/zone/tva partages par toutes les
  //    charges d'un meme colis, comme dans le classeur fait a la main).
  const infos = [];
  const parColis = new Map(); // tracking -> { poids, zone, tva }
  for (const g of byTracking.values()) {
    let poids = g.poidsBrut > 0 ? g.poidsBrut : poidsFromPrix(g.poidsFretPrix);
    if (poids == null && poidsExportBrut.has(g.tracking)) {
      poids = poidsExportBrut.get(g.tracking);
      infos.push(`${g.tracking} : poids introuvable dans le CSV BCF -> repris de l'export expeditions (INFO_POIDSRETENU = ${poids} kg)`);
    }
    if (poids == null) {
      infos.push(`${g.tracking} : poids introuvable (ni champ poids, ni prix reconnu dans la grille, ni export expeditions) -> plancher 0,1 applique, a verifier`);
      poids = 0.1;
    }
    const zt = cfg.zoning[g.pays] || null;
    const zone = zt ? String(zt[0]) : '';
    const tva = zt ? zt[1] : '';
    if (!zt) warnings.push(`${g.tracking} : pays '${g.pays}' absent de la table de zoning GLS`);
    parColis.set(g.tracking, { comref: g.ref, dest: g.dest, pays: g.pays, poids, zone, tva });
  }

  // 3. Un rec PAR LIGNE DE CHARGE BRUTE (pas regroupe) : fidele a la feuille "Facture GLS"
  //    du classeur fait a la main (plusieurs lignes par colis : Fret, Fret-avise, surcharges...).
  const recs = [];
  rows.forEach((r, idx) => {
    const track = (r[iTracking] || '').trim();
    if (!track) return;
    const c = parColis.get(track);
    const categorie = (r[iCategorie] || '').trim();
    const posteGls = cfg.categorie_to_poste[categorie] || 'Fret'; // libelle "poste GLS" (colonne Categorie du classeur fait a la main)
    const posteErp = erpPosteFor(categorie, [], `L${idx + 2} ${track}`); // warnings deja remontees au passage 1
    const montantHG = iAmountsHG.reduce((s, i) => s + (i >= 0 ? num(r[i]) : 0), 0);
    const gazoleLigne = iGazole >= 0 ? num(r[iGazole]) : 0;
    const postesLigne = { ZonesEloignees: 0, ColisVolumineux: 0, Adresses: 0, Fret: 0, TaxeGasoil: gazoleLigne };
    postesLigne[posteErp] = montantHG;
    recs.push({
      tracking: track, comref: c.comref, dest: c.dest, pays: c.pays, zone: c.zone,
      poids: c.poids, colis: 1, postes: postesLigne,
      horsGo: round2(montantHG), avecGo: round2(montantHG + gazoleLigne),
      raw: r, _tva: c.tva, _posteGls: posteGls,
    });
  });

  // 4. Lignes d'import ERP : 1 par colis (regroupe), comme l'ERP l'attend
  const importRows = [...parColis.entries()].map(([tracking, c]) => {
    const g = byTracking.get(tracking);
    return {
      Transporteur: cfg.champs_fixes.Transporteur,
      DateValidite: dateValidite || '',
      Ref1: c.comref, Ref2: '', IdClient: '',
      Tracking: tracking, Nom: c.dest,
      EP: cfg.champs_fixes['E/P'], Pays: c.pays, Zone: c.zone,
      NbrColis: 1, Poids: roundUp1(c.poids),
      Mode: c.zone, // cf. Import csv du fichier fait a la main : mode envoi = meme valeur que Zone
      TVA: c.tva,
      DroitsTaxes: 0, Assurance: 0,
      ZonesEloignees: g.postes.ZonesEloignees, ColisVolumineux: g.postes.ColisVolumineux,
      Adresses: g.postes.Adresses, Fret: g.postes.Fret, PlusValueB2C: 0,
      TaxeGasoil: g.postes.TaxeGasoil, NbColis: '',
    };
  });

  const controle = {};
  for (const k of POSTE_KEYS) controle[k] = round2([...byTracking.values()].reduce((s, g) => s + (g.postes[k] || 0), 0));

  const { alerts, infos: validateInfos } = validate(importRows);

  // Feuille "Facture GLS" du classeur = fidele au fichier fait a la main : Categorie
  // (poste GLS interne, ex. "Fret avise"), Total HT, Gazole, Total HT hors gazole, Poids,
  // suivies des colonnes brutes du CSV BCF (rec.raw, colle par writeWorkbook()).
  const rawCompCols = ['Catégorie', 'Total HT', 'Gazole', 'Total HT hors gazole', 'Poids'];
  const rawCompRow = (rec) => [rec._posteGls, rec.avecGo, rec.postes.TaxeGasoil, rec.horsGo, rec.poids];

  return {
    header, rows, recs, importRows, controle, warnings, alerts,
    infos: [...infos, ...validateInfos],
    posteKeys: POSTE_KEYS, gazoleKey: 'TaxeGasoil', cfg,
    sheetNames: { raw: 'Facture GLS', import: 'Import csv' },
    rawCompCols, rawCompRow,
    period: period || 'export',
  };
}

module.exports = {
  id: 'gls',
  name: 'GLS',
  status: 'ready',
  taxeGasoil: 'Calculee ligne a ligne dans le CSV (colonne SGO) ; le PDF sert a la reconciliation TTC, pas au calcul du gazole.',
  // Nomenclature des sorties (comme le fichier fait a la main). {period} = AAAA_MM.
  outputNaming: { workbook: '{period}_Facture GLS', import: '{period}_GLS_Import' },
  // Classeur de controle = CLONE FIDELE du fichier fait a la main (11 feuilles, formules,
  // 5 TCD/pivots) via Excel COM (Python). Le modele est le fichier existant, rafraichi.
  // Le PDF (facture officielle GLS) sert a la reconciliation TTC dans "Bilan factures"
  // (video process : verification visuelle TTC calcule vs "Montant T.T.C." du PDF).
  finalizer: {
    script: '../automatisation/finaliser_gls.py',
    template: '../Transporteurs/GLS/2026_06_Facture GLS.xlsx',
    buildArgs: (files) => [...(files.csv || []), '--pdf', ...(files.pdf || [])],
  },
  method: "CSV BCF GLS (1 ligne = 1 charge). Regroupement par Numero de colis, reclassement via categorie 'Hierarchie produits' (table de correspondance GLS -> poste), poids depuis le champ brut ou deduit du prix de base (grille PA GLS), en dernier repli depuis l'export expeditions du mois/mois-1 (INFO_POIDSRETENU, recupere automatiquement dans Automatisation/, pas a fournir), zone/TVA par pays.",
  inputs: [
    { key: 'csv', label: 'Export BCF GLS (CSV)', accept: '.csv', multiple: true, required: true },
    { key: 'pdf', label: 'Facture PDF GLS — pour reconciliation TTC', accept: '.pdf', multiple: true, required: false },
    { key: 'brut', label: 'Export des expéditions brutes (et mois précédent si dispo)', accept: '.xlsx,.xls', multiple: true, required: true },
  ],
  process,
};
