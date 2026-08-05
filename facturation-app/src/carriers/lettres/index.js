// ============================================================================
//  Adaptateur transporteur : LETTRES (Suivie / Suivie Prepa / Timbre Allemagne SLAACE)
//  Source unique = fichier WMS brut "Export expeditions_brut.xlsx" (meme fichier
//  que Delivengo utilise en fallback poids), filtre sur la colonne TRANSPORTEUR.
//  Pas de classeur de reference local pour ce transporteur -> pas de finaliseur
//  Excel COM, calcul JS pur (Fret fixe par groupe, decode depuis les 2 videos
//  process + verifie sur les 2 CSV de reference deja fournis).
//
//  3 groupes -> 3 fichiers import distincts (server.js/cli.js bouclent sur
//  result.multiImports) :
//    - "Lettre Suivie" + "LETTRE-SUIVIE-HYDRATIS" (2 comptes WMS distincts,
//      meme grille tarifaire) -> LETTRE-SUIVIE, Fret=3,83€
//    - "LETTRE-SUIVIE-PREPA" -> LETTRE-SUIVIE-PREPA, Fret=0,01€
//    - "LETTRE-TIMBRE-SLAACE" (envois Allemagne) -> LETTRE-TIMBRE-SLAACE,
//      Zone="DE" fixe, mode="DOMDE", TVA=0, Fret=1,65€
//
//  PRO_TRACKING vide -> remplace par CODE_EXPE (mecanisme observe dans la
//  video : correction manuelle du tracking invalide dans le WMS La Ruche).
// ============================================================================
const XLSX = require('xlsx');
const { round2 } = require('../../core/csv');
const { validate } = require('../../core/validate');
const cfg = require('./config.json');

const POSTE_KEYS = ['Fret'];

function firstDayOfMonth(d) {
  if (!(d instanceof Date) || Number.isNaN(d.getTime())) return null;
  return `01/${String(d.getMonth() + 1).padStart(2, '0')}/${d.getFullYear()}`;
}

function readBrut(paths) {
  const rows = [];
  for (const p of paths) {
    const wb = XLSX.readFile(p, { cellDates: true });
    const ws = wb.Sheets[wb.SheetNames[0]];
    const rs = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: '' });
    rows.push(...rs.slice(1));
  }
  return rows;
}

function process(files) {
  const paths = files.export || [];
  if (!paths.length) throw new Error("Aucun fichier fourni (attendu : export WMS 'Export expeditions_brut.xlsx').");

  const c = cfg.brut_cols;
  const rows = readBrut(paths);

  // index {transporteurBrut -> cle de groupe} pour un lookup O(1) par ligne
  const groupByTransporteur = {};
  for (const [key, g] of Object.entries(cfg.groupes)) {
    for (const t of g.transporteurs_bruts) groupByTransporteur[t] = key;
  }

  let periode = null;
  const parGroupe = {}; // { cleGroupe: [ligne...] }
  for (const r of rows) {
    const transp = String(r[c.transporteur] || '').trim();
    const groupeKey = groupByTransporteur[transp];
    if (!groupeKey) continue;
    if (!periode) periode = firstDayOfMonth(r[c.dateExpe]);
    (parGroupe[groupeKey] = parGroupe[groupeKey] || []).push(r);
  }

  const multiImports = [];
  const infos = [];
  let totalFret = 0;
  let totalLignes = 0;

  for (const [groupeKey, g] of Object.entries(cfg.groupes)) {
    const grpRows = parGroupe[groupeKey] || [];
    const importRows = grpRows.map((r) => {
      const codeExpe = String(r[c.codeExpe] || '').trim();
      const trackBrut = String(r[c.tracking] || '').trim();
      const tracking = trackBrut || codeExpe; // PRO_TRACKING vide -> CODE_EXPE
      return {
        Transporteur: g.nomTransporteurErp, DateValidite: periode || '',
        Ref1: '', Ref2: '', IdClient: '',
        Tracking: tracking, Nom: String(r[c.nom] || '').trim(),
        EP: 'P', Pays: String(r[c.pays] || '').trim(), Zone: g.zone,
        NbrColis: Math.round(Number(r[c.nbColis]) || 1), Poids: Number(r[c.poids]) || 0,
        Mode: g.mode, TVA: g.tva,
        DroitsTaxes: 0, Assurance: 0, ZonesEloignees: 0, ColisVolumineux: 0, Adresses: 0,
        Fret: g.fret, PlusValueB2C: 0, TaxeGasoil: 0, NbColis: '',
      };
    });
    const fretGroupe = round2(importRows.length * g.fret);
    totalFret += fretGroupe;
    totalLignes += importRows.length;
    multiImports.push({
      key: groupeKey, name: g.outputSuffix, sheetName: g.sheetImport,
      importRows, controle: { Fret: fretGroupe }, lignes: importRows.length,
    });
    infos.push(`${g.sheetImport} : ${importRows.length} ligne(s), Frêt fixe ${g.fret.toFixed(2)} € x ${importRows.length} = ${fretGroupe.toFixed(2)} €`);
  }

  // alertes standard (validate) sur l'ensemble combine, pour un controle global coherent avec les autres transporteurs.
  // skipPoidsDecimal : le poids brut WMS (INFO_POIDSRETENU) est repris tel quel, sans
  // arrondi de palier -- verifie sur le vrai fichier de reference (58/3683 lignes, soit
  // 1,6%, ont naturellement 1 decimale max) : le Fret etant fixe par groupe (pas calcule
  // depuis une grille poids), l'arrondi ARRONDI.SUP des autres transporteurs ne s'applique pas ici.
  const allImportRows = multiImports.flatMap((m) => m.importRows);
  const { alerts } = validate(allImportRows, { skipPoidsDecimal: true });

  return {
    lignes: totalLignes, period: periode ? `${periode.slice(6)}_${periode.slice(3, 5)}` : 'export',
    importRows: allImportRows, controle: { Fret: round2(totalFret) },
    warnings: [], alerts, infos, posteKeys: POSTE_KEYS,
    multiImports,
  };
}

module.exports = {
  id: 'lettres',
  name: 'Lettres (Suivie / Prepa / SLAACE)',
  status: 'ready',
  noWorkbook: true, // pas de classeur de reference local -> pas de classeur de sortie, juste les 3 imports
  taxeGasoil: 'Pas de taxe gasoil.',
  method: "Export WMS brut ('Export expeditions_brut.xlsx'), filtre sur TRANSPORTEUR en 3 groupes : LETTRE-SUIVIE (+ HYDRATIS), LETTRE-SUIVIE-PREPA, LETTRE-TIMBRE-SLAACE (Allemagne). Fret fixe par groupe (3,83€ / 0,01€ / 1,65€). PRO_TRACKING vide -> CODE_EXPE.",
  inputs: [
    { key: 'export', label: 'Export expeditions brut (.xlsx)', accept: '.xlsx', multiple: true, required: true },
  ],
  process,
};
