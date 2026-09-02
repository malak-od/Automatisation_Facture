// ============================================================================
//  Adaptateur transporteur : LETTRES (Suivie / Suivie Prepa)
//  Source unique = fichier WMS brut "Export expeditions_brut.xlsx" (meme fichier
//  que Delivengo utilise en fallback poids), filtre sur la colonne TRANSPORTEUR.
//  Pas de classeur de reference local pour ce transporteur -> pas de finaliseur
//  Excel COM, calcul JS pur (Fret fixe par groupe, decode depuis les 2 videos
//  process + verifie sur les 2 CSV de reference deja fournis).
//
//  2 groupes -> 2 fichiers import distincts (server.js/cli.js bouclent sur
//  result.multiImports) :
//    - "Lettre Suivie" + "LETTRE-SUIVIE-HYDRATIS" (2 comptes WMS distincts,
//      meme grille tarifaire) -> LETTRE-SUIVIE, Fret=3,83€
//    - "LETTRE-SUIVIE-PREPA" -> LETTRE-SUIVIE-PREPA, Fret=0,01€
//  "LETTRE-TIMBRE-SLAACE" (Timbre Allemagne) exclu du perimetre (decision
//  utilisateur 2026-08-07) -> lignes brutes ignorees comme n'importe quelle
//  autre valeur TRANSPORTEUR non geree par ce carrier.
//
//  PRO_TRACKING vide/duplique/mal forme ("BAC***") -> remplace par CODE_EXPE
//  (mecanisme observe dans la video : correction manuelle du tracking
//  invalide dans le WMS La Ruche -- code statut 101 = duplique, 400 = mal
//  forme, colle en formule "=CODE_EXPE"). PAS de filtre sur ETAT_EXPE (decision
//  utilisateur 2026-08-10) : toutes les lignes du groupe sont retenues, que
//  l'expedition soit livree, en cours ou en preparation.
// ============================================================================
const XLSX = require('xlsx');
const path = require('path');
const { round2 } = require('../../core/csv');
const { validate } = require('../../core/validate');
const { findAllBrutFiles, readBrutRows, poidsParTrackingFromExport } = require('../../core/exportBrut');
const cfg = require('./config.json');

const POSTE_KEYS = ['Fret'];

function firstDayOfMonth(d) {
  if (!(d instanceof Date) || Number.isNaN(d.getTime())) return null;
  // BUG TROUVE 2026-09-01 (meme piege identifie/corrige cote Delivengo) : XLSX.readFile
  // ({cellDates:true}) peut deriver de quelques heures lors de la conversion serial Excel ->
  // Date JS (ex. minuit UTC devient 21:59:39 la veille) -- .getMonth()/.getFullYear() en
  // heure LOCALE peuvent alors faire basculer un 1er-du-mois sur le mois precedent. Fix :
  // arrondir au jour le plus proche en UTC avant d'en extraire mois/annee.
  const rounded = new Date(Math.round(d.getTime() / 86400000) * 86400000);
  return `01/${String(rounded.getUTCMonth() + 1).padStart(2, '0')}/${rounded.getUTCFullYear()}`;
}

/** Normalise pour comparaison de nom de colonne (accents/casse), meme logique que Geodis/DPD. */
function normKey(s) {
  return String(s || '').normalize('NFD').replace(/[̀-ͯ]/g, '').toLowerCase().trim();
}

function idx(header, name) {
  const target = normKey(name);
  return header.findIndex((h) => h && normKey(h) === target);
}

/** Formats de PRO_TRACKING valides -- prefixe '87'+chiffres (Lettre Suivie/Hydratis, confirme sur
 * le brut reel de juin 2026 : 3588 lignes en '8700...', 87 lignes en '8750...' -- meme famille de
 * tracking, PAS uniquement '8700') ou '2L0'+chiffres (Prepa, ex. '2L02365965789'). Tout le reste
 * -- vide, duplique, 'BAC'+chiffres, corrompu/tronque (type '18^', '6A07022931190'), OU un nombre
 * qui ne commence PAS par '87' (ex. '100308757660', '3561920982564') -- est invalide -> remplace
 * par CODE_EXPE (colonne A), cf. process(). Le pole transport verifie ensuite manuellement dans
 * l'ERP si ce CODE_EXPE existe reellement : trouve -> corrige (marque rouge dans son suivi) ;
 * introuvable -> laisse tel quel et signale a part (marque jaune) -- verification hors de portee
 * de l'app (pas d'acces ERP). */
const VALID_TRACKING_RE = /^(87\d+|2L0\d+)$/;

function isTrackingValide(t) {
  return VALID_TRACKING_RE.test(t);
}

/** Detection par contenu (magic bytes ZIP) : un export WMS peut arriver en .xlsx OU en .csv
 * (confirme aout 2026 : export La Ruche fourni en CSV, colonnes decalees vs le xlsx habituel --
 * meme piege que Geodis/DPD, jamais de position fixe fiable d'un export a l'autre). XLSX.readFile
 * avec FS:';' parse le CSV correctement (guillemets/champs contenant le separateur, confirme sur
 * un cas reel : adresse "Chem. de Las Mouillas -, ;" dans le brut d'aout, un split() naif casse
 * le decompte de colonnes sur 351/13263 lignes). */
function isXlsx(p) {
  const fs = require('fs');
  try {
    const buf = fs.readFileSync(p, { encoding: null, flag: 'r' });
    return buf.length > 4 && buf[0] === 0x50 && buf[1] === 0x4b; // "PK"
  } catch (e) {
    return false;
  }
}

/** Lit un export brut (xlsx ou csv ';') -> { header, rows } -- colonnes lues par NOM ensuite
 * (idx()), jamais par position fixe (les exports La Ruche se decalent d'un mois/format a
 * l'autre, meme piege deja rencontre sur Geodis/DPD). */
function readBrutFile(p) {
  const wb = isXlsx(p) ? XLSX.readFile(p, { cellDates: true }) : XLSX.readFile(p, { cellDates: true, FS: ';' });
  const ws = wb.Sheets[wb.SheetNames[0]];
  const rs = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: '' });
  const header = rs[0].map((h) => String(h || '').replace(/^﻿/, '').trim());
  return { header, rows: rs.slice(1) };
}

function process(files) {
  const paths = files.export || [];
  if (!paths.length) throw new Error("Aucun fichier fourni (attendu : export WMS 'Export expeditions_brut.xlsx').");

  let header = null;
  const rows = [];
  for (const p of paths) {
    const r = readBrutFile(p);
    header = r.header; // suppose meme structure pour tous les fichiers d'un meme appel
    rows.push(...r.rows);
  }

  // Colonnes resolues par NOM dans l'entete reelle recue (jamais par position fixe -- cf.
  // readBrutFile). cfg.brut_cols_names porte les noms de colonnes WMS attendus.
  const cn = cfg.brut_cols_names;
  const c = {};
  for (const [key, name] of Object.entries(cn)) c[key] = idx(header, name);
  const missing = Object.entries(c).filter(([, i]) => i < 0).map(([key]) => cn[key]);
  if (missing.length) throw new Error(`Colonne(s) introuvable(s) dans l'export fourni : ${missing.join(', ')}`);

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

  // Poids brut = 0 sur la ligne traitee -> repli sur TOUS les exports bruts disponibles
  // (Automatisation/*brut*.xlsx, tous mois confondus, pas seulement mois courant/M-1) : le
  // poids d'une expedition peut n'etre renseigne par le WMS que dans un export ANTERIEUR au
  // mois de l'expedition elle-meme (constate : 3 trackings a poids=0 dans l'export de juin ont
  // un vrai poids dans l'export de mai -- le WMS met parfois a jour le poids apres coup et un
  // export plus recent capture cette maj, mais aussi parfois un export PLUS ANCIEN l'a deja).
  const brutPaths = findAllBrutFiles(path.resolve(__dirname, '../../..'));
  const poidsExportBrut = poidsParTrackingFromExport(readBrutRows(brutPaths));

  const multiImports = [];
  const infos = [];
  let totalFret = 0;
  let totalLignes = 0;
  let nbPoidsRecuperes = 0;
  let nbTrackingCorriges = 0;

  for (const [groupeKey, g] of Object.entries(cfg.groupes)) {
    const grpRows = parGroupe[groupeKey] || [];
    const importRows = grpRows.map((r) => {
      const codeExpe = String(r[c.codeExpe] || '').trim();
      const trackBrut = String(r[c.tracking] || '').trim();
      // PRO_TRACKING vide OU format non reconnu (corrompu/tronque) -> CODE_EXPE (cf. isTrackingValide).
      // Signale visuellement (bleu fonce dans le XLSX, cf. writeImportXlsx) pour reclamation
      // au pole transport, comme vu dans la video process.
      const trackingCorrige = !(trackBrut && isTrackingValide(trackBrut));
      const tracking = trackingCorrige ? codeExpe : trackBrut;
      if (trackingCorrige) nbTrackingCorriges++;
      let poids = Number(r[c.poids]) || 0;
      if (!poids && poidsExportBrut.has(tracking)) {
        poids = poidsExportBrut.get(tracking);
        nbPoidsRecuperes++;
      }
      return {
        Transporteur: g.nomTransporteurErp, DateValidite: periode || '',
        Ref1: '', Ref2: '', IdClient: '',
        Tracking: tracking, Nom: '', // vide dans les 2 fichiers de reference de juin 2026 (0/3683, 0/946) -- jamais rempli malgre DES_NOM disponible dans le brut
        EP: 'P', Pays: g.pays || String(r[c.pays] || '').trim(), Zone: g.zone,
        NbrColis: 1, // fixe a 1 dans les 2 fichiers de reference (3683/3683, 946/946), meme quand INFO_NBCOLIS brut = 2 (4 cas verifies) -- un envoi Lettre = 1 ligne facturee, peu importe le nb de colis WMS
        Poids: poids,
        Mode: g.mode, TVA: g.tva,
        DroitsTaxes: 0, Assurance: 0, ZonesEloignees: 0, ColisVolumineux: 0, Adresses: 0,
        Fret: g.fret, PlusValueB2C: 0, TaxeGasoil: 0, NbColis: '',
        _flagCol: trackingCorrige ? 'Tracking' : null,
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
  if (nbTrackingCorriges) infos.push(`${nbTrackingCorriges} tracking(s) invalide(s) remplacé(s) par le N° d'expédition (signalés en bleu foncé dans le XLSX -- à réclamer au pôle transport).`);
  if (nbPoidsRecuperes) infos.push(`${nbPoidsRecuperes} poids récupéré(s) depuis un autre export brut disponible (poids absent du fichier fourni pour cette ligne).`);

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
  name: 'Lettres (Suivie / Prepa)',
  status: 'ready',
  noWorkbook: true, // pas de classeur de reference local -> pas de classeur de sortie, juste les imports
  // Poids brut WMS repris tel quel (pas de recalcul), peut avoir plus d'1 decimale. Lu par
  // server.js pour l'alerte "POIDS X a plus d'1 decimale" (2e passe de validate() sur le CSV
  // final, cf. BUG TROUVE 2026-08-31 : la 1ere passe cote carrier avait deja ce flag mais
  // etait ecrasee par la 2e passe qui l'ignorait).
  skipPoidsDecimal: true,
  taxeGasoil: 'Pas de taxe gasoil.',
  method: "Export WMS brut ('Export expeditions_brut.xlsx'), filtre sur TRANSPORTEUR en 2 groupes : LETTRE-SUIVIE (+ HYDRATIS), LETTRE-SUIVIE-PREPA. LETTRE-TIMBRE-SLAACE exclu du périmètre. Fret fixe par groupe (3,83€ / 0,01€). Expéditions pas encore parties (tracking vide + statut 'en préparation'/'en attente') exclues, pas facturées. PRO_TRACKING invalide restant (dupliqué/mal formé) -> remplacé par CODE_EXPE.",
  inputs: [
    { key: 'export', label: 'Export expeditions brut (.xlsx)', accept: '.xlsx', multiple: true, required: true },
  ],
  process,
};
