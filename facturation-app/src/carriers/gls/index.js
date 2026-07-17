// ============================================================================
//  Adaptateur transporteur : GLS
//  CSV "BCF" GLS : 1 ligne = 1 charge (colis + type de charge, colonne "Hierarchie
//  produits"). Plusieurs lignes par colis (fret + surcharges) -> regroupement par
//  "Numero de colis" avant reclassement en postes ERP.
//  Mapping categorie -> poste : onglet "Categories" du classeur fait a la main
//  (2026_06_Facture GLS.xlsx), confirme par la video process (mai 2025).
// ============================================================================
const { readCsv, num, colIndex, roundUp1, round2 } = require('../../core/csv');
const { validate } = require('../../core/validate');
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

  // 1. Regroupement par tracking (plusieurs lignes de charge par colis)
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

  // 2. Poids final + zone/TVA + lignes d'import
  const infos = [];
  const recs = [];
  for (const g of byTracking.values()) {
    let poids = g.poidsBrut > 0 ? g.poidsBrut : poidsFromPrix(g.poidsFretPrix);
    if (poids == null) {
      infos.push(`${g.tracking} : poids introuvable (ni champ poids, ni prix reconnu dans la grille) -> plancher 0,1 applique, a verifier`);
      poids = 0.1;
    }
    const zt = cfg.zoning[g.pays] || null;
    const zone = zt ? zt[0] : '';
    const tva = zt ? zt[1] : '';
    if (!zt) warnings.push(`${g.tracking} : pays '${g.pays}' absent de la table de zoning GLS`);

    recs.push({
      tracking: g.tracking, comref: g.ref, dest: g.dest, pays: g.pays, zone: String(zone),
      poids, colis: 1, postes: g.postes,
      horsGo: round2(POSTE_KEYS.filter((k) => k !== 'TaxeGasoil').reduce((s, k) => s + g.postes[k], 0)),
      avecGo: round2(POSTE_KEYS.reduce((s, k) => s + g.postes[k], 0)),
      raw: [],
      _tva: tva,
    });
  }

  // 3. Lignes d'import ERP
  const importRows = recs.map((rec) => ({
    Transporteur: cfg.champs_fixes.Transporteur,
    DateValidite: dateValidite || '',
    Ref1: rec.comref, Ref2: '', IdClient: '',
    Tracking: rec.tracking, Nom: rec.dest,
    EP: cfg.champs_fixes['E/P'], Pays: rec.pays, Zone: rec.zone,
    NbrColis: rec.colis, Poids: roundUp1(rec.poids),
    Mode: rec.zone, // cf. Import csv du fichier fait a la main : mode envoi = meme valeur que Zone
    TVA: rec._tva,
    DroitsTaxes: 0, Assurance: 0,
    ZonesEloignees: rec.postes.ZonesEloignees, ColisVolumineux: rec.postes.ColisVolumineux,
    Adresses: rec.postes.Adresses, Fret: rec.postes.Fret, PlusValueB2C: 0,
    TaxeGasoil: rec.postes.TaxeGasoil, NbColis: '',
  }));

  const controle = {};
  for (const k of POSTE_KEYS) controle[k] = round2(recs.reduce((s, r) => s + (r.postes[k] || 0), 0));

  const { alerts, infos: validateInfos } = validate(importRows);

  return {
    header, rows, recs, importRows, controle, warnings, alerts,
    infos: [...infos, ...validateInfos],
    posteKeys: POSTE_KEYS, gazoleKey: 'TaxeGasoil', cfg,
    period: dateValidite ? `${dateValidite.slice(6)}_${dateValidite.slice(3, 5)}` : 'export',
  };
}

module.exports = {
  id: 'gls',
  name: 'GLS',
  status: 'ready',
  taxeGasoil: 'Calculee ligne a ligne dans le CSV (colonne SGO), pas besoin du PDF',
  method: "CSV BCF GLS (1 ligne = 1 charge). Regroupement par Numero de colis, reclassement via categorie 'Hierarchie produits' (table de correspondance GLS -> poste), poids depuis le champ brut ou deduit du prix de base (grille PA GLS), zone/TVA par pays.",
  inputs: [
    { key: 'csv', label: 'Export BCF GLS (CSV)', accept: '.csv', multiple: true, required: true },
  ],
  process,
};
