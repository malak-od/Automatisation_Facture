// ============================================================================
//  Adaptateur transporteur : COLISSIMO
//  2 fichiers CSV bruts (separateur ';', latin1/cp1252) : "Prestations au
//  colis" (26 colonnes) et "Frais de douane" (20 colonnes). Colonnes H a N
//  (N de ligne -> Produit) COMMUNES aux 2 fichiers ; le douane n'a pas de
//  Zone/Pays/Poids/dimensions (colonnes P-Z restent vides pour ces lignes),
//  mais Total HT/Taux TVA/Rubrique/Code charge (AA-AD) sont bien alignes
//  dans les 2 cas -- confirme empiriquement cellule par cellule sur le
//  fichier reel de juin 2026 (voir Documentation/
//  Video_Colissimo_1_Preparation_fichier_import.md, Partie 5).
//
//  1 ligne brute = 1 "charge" (Transport, Remise, CAE=taxe gazole,
//  Supplement...), PAS 1 colis -- un meme "N colis" (tracking) a plusieurs
//  lignes. Regroupement par tracking necessaire (comme Chronopost/BLS).
//
//  Mapping Type de charge -> poste ERP : table "table_correspondance" du
//  config.json (26 codes "Code charge", ex. CH_CAP="Taxe gazole").
//  Mapping mode envoi/zone : concat(prefixe tracking 2 lettres, code pays
//  ISO) -> table "modes_envois" (327 entrees), le code pays venant lui-meme
//  d'un lookup sur "pays" (nom complet -> ISO 2 lettres).
//
//  Taxe gasoil : PAS de taux a chercher sur le site Colissimo -- deja
//  calculee et fournie par Colissimo dans le CSV brut (ligne "CAE",
//  categorie "Taxe gazole"). Confirme au centime pres sur juin 2026.
// ============================================================================
const { readCsv, num, round2, roundUp1 } = require('../../core/csv');
const { validate } = require('../../core/validate');
const cfg = require('./config.json');

const POSTE_KEYS = ['Adresse', 'Assurance', 'Colis volumineux', 'Droits et taxes', 'Frêt', 'plus-value BtoC', 'Taxe gazole', 'Zones éloignées'];

// Index rapides construits une fois depuis config.json.
const PAYS_NOM_VERS_ISO = Object.fromEntries(cfg.pays.map((p) => [normKey(p.nom), p.code]));
const MODES_ENVOIS_PAR_CONCAT = Object.fromEntries(cfg.modes_envois.map((m) => [m.concat, m]));

function normKey(s) {
  return String(s || '').normalize('NFD').replace(/[̀-ͯ]/g, '').toUpperCase().trim();
}

/** Lit un des 2 CSV bruts Colissimo (';' latin1) -> { header, rows }. */
function readColissimoCsv(path) {
  const { header, rows } = readCsv(path, ';', 'latin1');
  return { header: header.map((h) => h.trim()), rows: rows.filter((r) => r.some((v) => v !== '')) };
}

function colIdx(header, name) {
  return header.indexOf(name);
}

async function process(files) {
  const prestaPaths = files.presta || [];
  const douanePaths = files.douane || [];
  if (!prestaPaths.length) throw new Error('Aucun fichier fourni (attendu : CSV "Prestations au colis").');

  const warnings = [];
  const infos = [];

  // Colonnes H a N COMMUNES aux 2 CSV (N de ligne, N facture, Date, Compte facture,
  // Compte deposant, Description, Produit) ; N colis, Pays Origine/Destination, Poids Kg
  // n'existent QUE dans "Prestations au colis" -- restent vides/absentes pour "Frais de
  // douane" (confirme empiriquement, pas de decalage dangereux).
  const lignes = [];
  let dateValidite = null;

  let headerRef = null; // en-tete du 1er fichier lu, sert de reference pour la feuille "Facture Colissimo"
  function parseFichier(path, isDouane) {
    const { header, rows } = readColissimoCsv(path);
    if (!headerRef) headerRef = header;
    const iNumeroLigne = colIdx(header, 'N de ligne');
    const iNumeroFacture = colIdx(header, 'N facture');
    const iDate = colIdx(header, 'Date');
    const iProduit = colIdx(header, 'Produit');
    const iNColis = colIdx(header, 'N colis');
    const iPaysOrigine = colIdx(header, 'Pays Origine');
    const iPaysDestination = colIdx(header, 'Pays Destination');
    const iPoidsKg = colIdx(header, 'Poids Kg');
    const iNaturePoids = colIdx(header, 'Nature du poids facturé');
    const iTotalHt = colIdx(header, 'Total HT');
    const iTauxTva = colIdx(header, 'Taux TVA');
    const iCodeCharge = colIdx(header, 'Code charge');
    if (iNColis < 0 || iTotalHt < 0 || iCodeCharge < 0) {
      warnings.push(`${path.split(/[\\/]/).pop()} : colonne(s) attendue(s) introuvable(s) (N colis/Total HT/Code charge) — fichier ignoré.`);
      return;
    }
    for (const r of rows) {
      const nColis = String(r[iNColis] || '').trim();
      if (!nColis) continue;
      if (!dateValidite && iDate >= 0 && r[iDate]) {
        const m = /^(\d{2})\/(\d{2})\/(\d{4})$/.exec(String(r[iDate]).trim());
        if (m) dateValidite = `01/${m[2]}/${m[3]}`;
      }
      lignes.push({
        numeroFacture: iNumeroFacture >= 0 ? String(r[iNumeroFacture] || '').trim() : '',
        produit: iProduit >= 0 ? String(r[iProduit] || '').trim() : '',
        nColis,
        // Colissimo n'a PAS de "Pays Origine"/"Pays Destination" sur les lignes de douane
        // (confirme : ces colonnes restent vides, cf. docstring en tete de fichier) --
        // paysDestination reste '' pour ces lignes, XLOOKUP renverra "" (comme le modele).
        paysDestination: iPaysDestination >= 0 ? String(r[iPaysDestination] || '').trim() : '',
        // Nature du poids facture = "M" (Massique, majoritaire) OU "V" (Volumetrique, 41
        // lignes/juin 2026) -- les deux portent un vrai poids exploitable sur la ligne
        // "Charge <prefixe>". BUG TROUVE 2026-08-14 : filtrer sur "M" seul faisait rater
        // 33 trackings/2765 (poids=0 au lieu de la vraie valeur), confirme contre le fichier
        // fait a la main de juin 2026.
        poidsKg: (!isDouane && iPoidsKg >= 0 && iNaturePoids >= 0 && ['M', 'V'].includes(String(r[iNaturePoids] || '').trim().toUpperCase())) ? num(r[iPoidsKg]) : null,
        totalHt: num(r[iTotalHt]),
        tauxTva: iTauxTva >= 0 ? num(r[iTauxTva]) : 0,
        codeCharge: String(r[iCodeCharge] || '').trim(),
        raw: r,
      });
    }
  }

  for (const p of prestaPaths) parseFichier(p, false);
  for (const p of douanePaths) parseFichier(p, true);
  if (!lignes.length) warnings.push('Aucune ligne "colis" trouvée dans les fichier(s) fournis.');

  // Colonnes A-G de "Facture Colissimo" : Pays (ISO)/Mode envoi/Zone/Concat/Prefixe/Pays/
  // Categorie -- formules exactes confirmees ('Facture Colissimo'!A2:G2, cf. config.json).
  //
  // Chaque rec correspond a 1 LIGNE BRUTE (une "charge"), pas 1 colis -- meme convention que
  // GLS/Chronopost. "postes"/"horsGo"/"avecGo"/"raw"/"comref"/"dest"/"colis" sont exiges par
  // le contrat generique de src/core/excelOut.js#writeWorkbook (repli sans finaliseur Python,
  // utilise tant qu'aucun finalizer.script n'est branche pour Colissimo) -- structure requise
  // par la feuille "Facture Colissimo" (compRow) ET la feuille "TCD" (agregation par tracking)
  // du classeur genere. BUG TROUVE 2026-08-14 (capture ecran utilisateur, "Cannot read
  // properties of undefined (reading 'Droits et taxes')") : les recs ne portaient QUE les
  // champs metier (categorie/poidsKg/...), pas ce contrat generique -- writeWorkbook plantait
  // en tentant de lire rec.postes (undefined) sur le repli par defaut Kuehne-like.
  const recs = [];
  for (const l of lignes) {
    const prefixe = l.nColis.slice(0, 2);
    const paysIso = PAYS_NOM_VERS_ISO[normKey(l.paysDestination)] || '';
    const concat = prefixe + paysIso;
    const modeEntry = MODES_ENVOIS_PAR_CONCAT[concat];
    const zone = modeEntry ? modeEntry.zone : 'zone inconnue';
    const mode = modeEntry ? modeEntry.mode_transport : 'zone inconnue';
    if (!modeEntry) warnings.push(`Colis ${l.nColis} (facture ${l.numeroFacture}) : mode envoi "${concat}" absent de "Modes envois" — zone inconnue.`);

    if (!paysIso && l.paysDestination) warnings.push(`Colis ${l.nColis} (facture ${l.numeroFacture}) : pays "${l.paysDestination}" absent de "Pays" — "Pays à créer".`);

    const categorie = cfg.table_correspondance[l.codeCharge];
    if (!categorie) warnings.push(`Colis ${l.nColis} (facture ${l.numeroFacture}) : code charge "${l.codeCharge}" absent de "Table de correspondance" — non reclassé.`);

    const postesLigne = Object.fromEntries(POSTE_KEYS.map((k) => [k, 0]));
    if (categorie && POSTE_KEYS.includes(categorie)) postesLigne[categorie] = round2(l.totalHt);

    recs.push({
      numeroFacture: l.numeroFacture, tracking: l.nColis, comref: l.numeroFacture, dest: '',
      produit: l.produit, pays: paysIso || 'Pays à créer', zone, mode,
      poids: l.poidsKg || 0, poidsKg: l.poidsKg, colis: 1,
      totalHt: l.totalHt, tauxTva: l.tauxTva, categorie: categorie || null, codeCharge: l.codeCharge,
      postes: postesLigne, horsGo: round2(l.totalHt), avecGo: round2(l.totalHt),
      raw: l.raw,
    });
  }

  // Regroupement par tracking (N colis) : plusieurs lignes "charge" par colis (Transport,
  // CAE=taxe gazole, Remise, Supplements...) -> 1 seule ligne d'import. Poids = celui de la
  // ligne "Nature du poids facturé"=M (poids reel du colis, confirme via l'onglet "Poids"
  // du modele -- TCD "Moyenne de Poids Kg" par tracking, qui revient a cette valeur unique
  // puisqu'un colis n'a qu'une seule ligne avec poids renseigne).
  const groupes = new Map();
  for (const rec of recs) {
    if (!groupes.has(rec.tracking)) groupes.set(rec.tracking, []);
    groupes.get(rec.tracking).push(rec);
  }

  // TVA : regle simple depuis le CSV brut (PAS de table pays->TVA separee) -- confirme
  // formule exacte 'Import CSV'!N2 = IF(Taux TVA CSV = 20, 0.2, 0).
  function tvaPourTaux(taux) {
    return taux === 20 ? 0.2 : 0;
  }

  const importRows = [];
  let totalGazoleCalcule = 0;
  for (const [tracking, groupe] of groupes) {
    const base = groupe.find((r) => r.poidsKg != null) || groupe[0];
    const postes = Object.fromEntries(POSTE_KEYS.map((k) => [k, 0]));
    for (const r of groupe) {
      if (r.categorie && POSTE_KEYS.includes(r.categorie)) postes[r.categorie] = round2(postes[r.categorie] + r.totalHt);
    }
    const tauxTvaRef = groupe.find((r) => r.tauxTva)?.tauxTva ?? 0;
    totalGazoleCalcule = round2(totalGazoleCalcule + (postes['Taxe gazole'] || 0));

    importRows.push({
      Transporteur: 'COLISSIMO',
      DateValidite: dateValidite || '',
      // Id client : TOUJOURS vide dans le fichier import (decision utilisateur 2026-08-12,
      // s'applique a tous les transporteurs).
      Ref1: '', Ref2: '', IdClient: '',
      Tracking: base.tracking, Nom: '',
      EP: 'P', Pays: base.pays, Zone: base.zone,
      NbrColis: 1, Poids: base.poidsKg != null ? roundUp1(base.poidsKg) : 0,
      Mode: base.mode, TVA: tvaPourTaux(tauxTvaRef),
      DroitsTaxes: postes['Droits et taxes'] || null,
      Assurance: postes.Assurance || null,
      ZonesEloignees: postes['Zones éloignées'] || null,
      ColisVolumineux: postes['Colis volumineux'] || null,
      Adresses: postes.Adresse || null,
      Fret: postes['Frêt'], PlusValueB2C: postes['plus-value BtoC'] || 0,
      // Le gazole (categorie "Taxe gazole", ligne CAE) N'EST PAS repercute au client dans le
      // fichier livre -- confirme empiriquement (2026-08-14) : la colonne "gazole" du fichier
      // fait a la main de juin 2026 est VIDE (0) sur les 2765 lignes, meme mecanisme deja
      // observe pour Chronopost/BLS. postes['Taxe gazole'] reste calcule (utilise dans
      // "controle"/Gazole pour visibilite interne) mais jamais ecrit dans TaxeGasoil.
      TaxeGasoil: '', NbColis: '',
    });
  }

  const { alerts, infos: valInfos } = validate(importRows, {});
  infos.push(...valInfos);

  // Reconciliation PDF : Total HT officiel compare a la somme du Total HT brut de toutes
  // les lignes de cette facture (avant reclassement) -- confirme egalite stricte
  // 'Bilan Factures' : Somme Total HT (TCD) = PDF HT + Indemnisation + Avoir. La ligne PDF
  // (facture_...pdf) ne distingue pas indemnisation/avoir ici -- on compare au total BRUT
  // calcule, l'utilisateur ajuste manuellement si un ecart apparait (comme le modele).
  const totalBrutParFacture = new Map();
  for (const l of lignes) {
    totalBrutParFacture.set(l.numeroFacture, round2((totalBrutParFacture.get(l.numeroFacture) || 0) + l.totalHt));
  }
  const pdfPaths = files.pdf || [];
  for (const p of pdfPaths) {
    try {
      const pdfParse = require('pdf-parse');
      const buf = require('fs').readFileSync(p);
      const { text } = await pdfParse(buf);
      const mFacture = /FACTURE N[°\s]*([A-Z0-9]+)/.exec(text);
      const mTotal = /TOTAL\s*HT\s*([\d\s]+[.,]\d{2})\s*€/.exec(text);
      if (!mFacture || !mTotal) { warnings.push(`PDF ${p.split(/[\\/]/).pop()} : numéro de facture/total introuvable, ignoré pour la réconciliation.`); continue; }
      const numeroFacture = mFacture[1];
      const totalPdf = num(mTotal[1].replace(/\s/g, ''));
      const calcule = totalBrutParFacture.get(numeroFacture);
      if (calcule == null) { warnings.push(`Facture PDF ${p.split(/[\\/]/).pop()} (${numeroFacture}) : aucune ligne correspondante trouvée dans les fichiers traités.`); continue; }
      const ecart = round2(calcule - totalPdf);
      infos.push(`Facture PDF ${p.split(/[\\/]/).pop()} (${numeroFacture}) : Total HT PDF (avant indemnisations) = ${totalPdf.toFixed(2)} EUR, calculé = ${calcule.toFixed(2)} EUR (écart ${ecart >= 0 ? '+' : ''}${ecart.toFixed(2)} EUR — les indemnisations/avoirs du PDF ne sont pas automatiquement déduits, vérifier manuellement si écart important).`);
    } catch (e) {
      warnings.push(`PDF ${p.split(/[\\/]/).pop()} : lecture impossible (${e.message}).`);
    }
  }

  // Controle : somme des "postes" par ligne brute (recs), pas des importRows (qui excluent
  // deja le gazole de TaxeGasoil, cf. plus haut) -- coherent avec le total 2165,36€ confirme
  // contre le PDF (categorie "Taxe gazole" sur toutes les lignes CAE, avant regroupement).
  const controle = {};
  for (const k of POSTE_KEYS) controle[k] = round2(recs.reduce((s, r) => s + (r.postes[k] || 0), 0));

  // rawCompCols/rawCompRow : structure Colissimo pour la feuille "Facture Colissimo" du
  // classeur genere (repli sans finaliseur Python) -- differente du defaut Kuehne-like
  // (pas de "Total hors GO"/"Total + GO" par ligne pertinent ici, on montre plutot Categorie/
  // Code charge/Total HT), meme principe que rawCompCols/rawCompRow de GLS.
  const rawCompCols = ['N Facture', 'N Colis', 'Catégorie', 'Code charge', 'Total HT'];
  const rawCompRow = (rec) => [rec.comref, rec.tracking, rec.categorie || '', rec.codeCharge || '', rec.totalHt];

  return {
    header: headerRef || [],
    rows: recs, recs, importRows, controle, warnings, alerts, infos,
    posteKeys: POSTE_KEYS, cfg, gazoleKey: 'Taxe gazole',
    rawCompCols, rawCompRow,
    sheetNames: { raw: 'Facture Colissimo', import: 'Import CSV' },
    period: dateValidite ? `${dateValidite.slice(6)}_${dateValidite.slice(3, 5)}` : 'export',
  };
}

module.exports = {
  id: 'colissimo',
  name: 'Colissimo',
  status: 'ready',
  taxeGasoil: "Déjà calculée et fournie par Colissimo dans le CSV brut (ligne 'CAE', catégorie 'Taxe gazole') -- PAS un taux à chercher sur leur site. Confirmé au centime près sur juin 2026 (2165,36€ = 'Coefficient Ajustement Energie' du PDF).",
  method: "2 fichiers CSV bruts ('Prestations au colis' + 'Frais de douane', séparateur ';'). Chaque 'N colis' (tracking) a plusieurs lignes 'charge' (Transport/CAE/Remise/Suppléments), reclassées en postes ERP via la table de correspondance (Code charge → poste). Mode envoi/zone via concat(préfixe tracking, pays ISO) → table Modes envois. Réconciliation PDF par n° de facture (Total HT, indemnisations à ajuster manuellement).",
  inputs: [
    { key: 'presta', label: "CSV 'Prestations au colis'", accept: '.csv', multiple: true, required: true },
    { key: 'douane', label: "CSV 'Frais de douane'", accept: '.csv', multiple: true, required: false },
    { key: 'pdf', label: 'Facture(s) PDF Colissimo (contrôle du total)', accept: '.pdf', multiple: true, required: false },
  ],
  outputNaming: { workbook: '{period}_Facture Colissimo', import: '{period}_Colissimo_Import' },
  finalizer: {
    script: '../automatisation/finaliser_colissimo.py',
    template: '../Transporteurs/Colissimo/2026_06_Facture Colissimo.xlsx',
    buildArgs: (files, period) => [
      '--presta', ...(files.presta || []),
      '--douane', ...(files.douane || []),
      '--pdf', ...(files.pdf || []),
      ...(period ? ['--period', period] : []),
    ],
  },
  process,
};
