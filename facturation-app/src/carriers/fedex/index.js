// ============================================================================
//  Adaptateur transporteur : FEDEX (Shipment Detail CSV, export portail FedEx)
//  1 fichier CSV brut par mois ("0g000e48hq_..." export, 1 ligne = 1 colis/
//  envoi), en-tetes FRANCAIS (confirme sur le CSV reel de juin 2026, malgre le
//  meme prefixe de nommage qu'un export anglais vu ailleurs -- variable selon
//  le mois, colonnes resolues PAR NOM, jamais par position).
//
//  Filtrage "Droits & Taxes" (process manuel documente par l'utilisateur,
//  2026-08-19) : une facture dont TOUTES les lignes ont un montant de fret
//  (Frais de transport) nul et un montant "Droits de douane et taxes" non nul
//  est un ajustement douanier pur, PAS un envoi facturable -- ses lignes sont
//  exclues du fichier livre (reproduit "si une facture correspond a Droit &
//  Taxes il faut supprimer les lignes correspondantes dans Shipment Detail").
//
//  Detection "facture d'un autre mois" (decouverte 2026-08-19, facture reelle
//  634313590 de juin 2026 : 392 lignes, expeditions 11-27/05/2026, fret total
//  22384EUR, ABSENTE du fichier livre reel bien que datee 01/06/2026) : le
//  CSV brut peut contenir des factures deja traitees le mois precedent (le
//  numero de facture seul ne suffit pas a trancher, decision utilisateur
//  explicite 2026-08-19 "meme si facture/expedie dans un autre mois, garde-
//  les si factures au mois cible" -- donc PAS de filtre automatique par date
//  d'expedition). On se contente d'ALERTER (pas d'exclure) si une facture a
//  une part anormalement elevee de lignes dont la date d'envoi tombe hors du
//  mois cible, pour laisser le pole transport trancher au cas par cas.
//
//  Montant (Shipment Detail!A, sert au TCD/Bilan factures/Bilan clients) :
//  formule modele =BK+BN = 'Frais de transport' + 'Remise' (deja negative).
//  Mode d'envoi (Shipment Detail!E) : =IF(Service="FedEx Priority","ST","FICP")
//  -- confirme exact sur 3607/3607 lignes reelles de juin 2026.
//  Zone (Import ERP!J) : XLOOKUP pays (ISO2 destinataire) -> colonne "FICP"
//  de l'onglet Zoning (jamais IE/RE, quasi vides) -- valide contre le CSV ERP
//  reel livre (Finlande -> Zoning FICP='T' -> CSV Zone='T').
//  TVA (Import ERP!N) : 20% si pays destinataire dans la liste UE (config,
//  27 pays), 0% sinon -- reconfirme caractere pour caractere sur le fichier
//  reel de juin 2026 (formule teste le PAYS, PAS le mode FEDEX FRANCE/
//  INTERNATIONAL, malgre ce que la video d'avril semblait suggerer via de
//  simples valeurs observees).
//  Colis volumineux (Import ERP!R) : forfait fixe 10EUR si "Dimmed Length
//  (cm)" > 60cm.
//  E/P (Import ERP!H, via TCD!F) : XLOOKUP tracking -> Shipment Detail!D
//  ("entreprise"->E, sinon P) -- colonne D reconfirmee 100% LITERALE (jamais
//  formule) sur les 3607 lignes reelles de juin -- alimentee via l'export WMS
//  partage "AAAA MM - Export expeditions_brut.xlsx" (meme mecanisme deja
//  utilise par d'autres transporteurs, cf. core/exportBrut.js), repli "P" par
//  defaut si tracking absent de l'export (avec alerte).
// ============================================================================
const path = require('path');
const pdfParse = require('pdf-parse');
const { num, round2, roundUp1 } = require('../../core/csv');
const { validate } = require('../../core/validate');
const { findBrutFiles, readBrutRows, epParTrackingFromExport } = require('../../core/exportBrut');
const cfg = require('./config.json');

function normKey(s) {
  // Apostrophes typographiques ('U+2019) frequentes dans les en-tetes CSV FedEx
  // ("Numero de suivi de l'envoi") -- normalisees en apostrophe simple avant comparaison.
  return String(s || '')
    .normalize('NFD').replace(/[̀-ͯ]/g, '')
    .replace(/[‘’]/g, "'")
    .toLowerCase().replace(/[.,]/g, '').replace(/\s+/g, ' ').trim();
}

function colIndexByName(header, name) {
  const target = normKey(name);
  return header.findIndex((h) => normKey(h) === target);
}

/** Lit un CSV FedEx (export portail, en-tete FR, separateur virgule, encodage UTF-8 --
 * confirme sur les 2 CSV bruts reels de juin 2026, tous deux en francais malgre le meme
 * prefixe "0g000e48hq_" observe sur un export anglais d'un autre mois dans la video). */
function readFedexCsv(p) {
  const fs = require('fs');
  const buf = fs.readFileSync(p);
  let text = buf.toString('utf8');
  if (text.charCodeAt(0) === 0xfeff || text.slice(0, 3) === '﻿') text = text.slice(1);
  const rows = [];
  let row = [];
  let field = '';
  let inQuotes = false;
  for (let i = 0; i < text.length; i++) {
    const c = text[i];
    if (inQuotes) {
      if (c === '"') {
        if (text[i + 1] === '"') { field += '"'; i++; } else inQuotes = false;
      } else field += c;
    } else if (c === '"') inQuotes = true;
    else if (c === ',') { row.push(field); field = ''; }
    else if (c === '\r') { /* skip */ }
    else if (c === '\n') { row.push(field); rows.push(row); row = []; field = ''; }
    else field += c;
  }
  if (field !== '' || row.length) { row.push(field); rows.push(row); }
  const header = (rows[0] || []).map((h) => String(h || '').trim());
  const data = rows.slice(1).filter((r) => r.some((v) => v !== ''));
  return { file: path.basename(p), header, rows: data };
}

/** Date "mm/jj/aaaa" -> {y,m,d} (format US du CSV FedEx, confirme sur le brut reel). */
function parseDateUs(s) {
  const m = /^(\d{1,2})\/(\d{1,2})\/(\d{4})$/.exec(String(s || '').trim());
  if (!m) return null;
  return { y: Number(m[3]), m: Number(m[1]), d: Number(m[2]) };
}

/** PDF FedEx (facture "Facture de frais de transport") : bloc labels/valeurs colle par
 * pdf-parse ("No de Client:\nNo de Facture:\n...\nMontant du\n*****1234\n634323394\n...368,95
 * EUR") + ligne "Total dû EUR <montant>" (recapitulatif final, redondant avec "Montant du" mais
 * plus simple a capturer isolement). Confirme exact sur les 9 PDF reels de juin 2026 (1 a 195
 * pages). */
async function extractFedexPdfInfo(pdfPath) {
  const buf = require('fs').readFileSync(pdfPath);
  const { text } = await pdfParse(buf);
  const mInv = /No de Client\s*:\s*\nNo de Facture\s*:\s*\nDate de la facture\s*:\s*\nDate d.[eé]ch[eé]ance\s*\nMontant d[uû]\s*\n\*+\d+\s*\n(\d+)\s*\n/.exec(text);
  const mTotal = /Total d[uû]\s*EUR\s*([\d.,]+)/.exec(text);
  return {
    file: path.basename(pdfPath),
    numeroFacture: mInv ? mInv[1] : null,
    totalDu: mTotal ? num(mTotal[1].replace(/\./g, '').replace(',', '.')) : null,
  };
}

async function process(files) {
  const csvPaths = files.csv || [];
  if (!csvPaths.length) throw new Error('Aucun fichier fourni (attendu : Shipment detail FedEx, CSV export portail).');

  const warnings = [];
  const infos = [];
  const brutes = csvPaths.map(readFedexCsv);

  const header = brutes[0].header;
  const iCompte = colIndexByName(header, 'Compte du payeur');
  const iService = colIndexByName(header, 'Description du service');
  const iDateEnvoi = colIndexByName(header, "Date d'envoi (mm/jj/aaaa)");
  const iTracking = colIndexByName(header, "Numéro de suivi de l'envoi");
  const iFret = colIndexByName(header, "Montant des frais de transport de l'envoi en USD");
  const iDivers = colIndexByName(header, "Shipment Miscellaneous Charge USD");
  const iDroitsTaxes = colIndexByName(header, "Droits de douane et taxes de l'envoi en USD");
  const iRemise = colIndexByName(header, "Montant de la remise de l'envoi en USD");
  const iColis = colIndexByName(header, "Colis dans l'envoi");
  const iPoids = colIndexByName(header, "Poids nominal de l'envoi (livres)");
  const iPaysDest = colIndexByName(header, 'Pays/Territoire du destinataire');
  const iLongueur = colIndexByName(header, 'Longueur volumétrique (cm)');
  const iInvoiceNum = colIndexByName(header, 'Numéro de facture');
  const iInvoiceDate = colIndexByName(header, 'Date de facturation (mm/jj/aaaa)');
  const iMoisFacturation = colIndexByName(header, 'Mois de facturation (aaaamm)');

  const requis = { iTracking, iFret, iRemise, iInvoiceNum, iPaysDest };
  const manquantes = Object.entries(requis).filter(([, v]) => v < 0).map(([k]) => k);
  if (manquantes.length) {
    throw new Error(`Colonne(s) attendue(s) introuvable(s) dans le CSV FedEx : ${manquantes.join(', ')} — vérifier l'en-tête du fichier reçu.`);
  }

  // Mois cible = mois majoritaire (colonne "Mois de facturation (aaaamm)" si presente, sinon
  // "Date de facturation" -- meme piege deja rencontre sur Chronopost/TNT : quelques lignes
  // residuelles d'un autre mois peuvent se glisser dans l'export).
  const comptageMois = new Map();
  for (const f of brutes) {
    for (const r of f.rows) {
      let mois = iMoisFacturation >= 0 ? String(r[iMoisFacturation] || '').trim() : '';
      if (!/^\d{6}$/.test(mois) && iInvoiceDate >= 0) {
        const d = parseDateUs(r[iInvoiceDate]);
        if (d) mois = `${d.y}${String(d.m).padStart(2, '0')}`;
      }
      if (/^\d{6}$/.test(mois)) comptageMois.set(mois, (comptageMois.get(mois) || 0) + 1);
    }
  }
  let moisCible = null;
  let dateValidite = '';
  if (comptageMois.size) {
    [moisCible] = [...comptageMois.entries()].sort((a, b) => b[1] - a[1])[0];
    const m = /^(\d{4})(\d{2})$/.exec(moisCible);
    if (m) dateValidite = `01/${m[2]}/${m[1]}`;
    if (comptageMois.size > 1) {
      const detail = [...comptageMois.entries()].sort((a, b) => b[1] - a[1]).map(([k, v]) => `${k}: ${v} ligne(s)`).join(', ');
      infos.push(`Plusieurs mois détectés dans le CSV reçu (${detail}) — mois retenu : ${moisCible} (majoritaire).`);
    }
  }

  // Regroupement par facture (Numero de facture) : detection Droits&Taxes pure (exclue) +
  // detection "facture d'un autre mois" (alerte seulement, cf. docstring en tete).
  const parFacture = new Map(); // numFacture -> { fretTotal, taxeTotal, lignes: [], moisEnvoiCount: Map }
  for (const f of brutes) {
    for (const r of f.rows) {
      const numFacture = String(r[iInvoiceNum] || '').trim();
      if (!numFacture) continue;
      if (!parFacture.has(numFacture)) parFacture.set(numFacture, { fretTotal: 0, taxeTotal: 0, lignes: [], moisEnvoi: new Map() });
      const fac = parFacture.get(numFacture);
      const fret = iFret >= 0 ? num(r[iFret]) : 0;
      const taxe = iDroitsTaxes >= 0 ? num(r[iDroitsTaxes]) : 0;
      fac.fretTotal = round2(fac.fretTotal + fret);
      fac.taxeTotal = round2(fac.taxeTotal + taxe);
      fac.lignes.push(r);
      const dEnvoi = iDateEnvoi >= 0 ? parseDateUs(r[iDateEnvoi]) : null;
      if (dEnvoi) {
        const key = `${dEnvoi.y}${String(dEnvoi.m).padStart(2, '0')}`;
        fac.moisEnvoi.set(key, (fac.moisEnvoi.get(key) || 0) + 1);
      }
    }
  }

  // ATTENTION — question ouverte non tranchee (2026-08-19, cf. memoire fedex_carrier_construit.md) :
  // sur juin 2026 reel, le fichier livre au client (2026_06_Fedex_Import.csv, 3607 lignes,
  // Fret total 45515,83EUR) EXCLUT 3 factures du CSV brut (120792009 Droits&Taxes pur,
  // 634313590 392 lignes 100% mai, 634364158 1 ligne datee juillet) -- mais AUCUNE regle
  // simple (ni "100% hors mois cible", ni "majorite hors mois cible") ne reproduit exactement
  // cette exclusion : 634333315 (100% mai, 1 ligne) EST incluse dans le fichier reel, alors
  // que 634364158 (ligne datee juin) est exclue. Inclure TOUT (sauf Droits&Taxes) surestime le
  // Fret de +12394,91EUR (~27%) par rapport a juin. Repli actuel : n'exclure QUE les factures
  // 100% Droits&Taxes (seule regle certaine, confirmee par l'utilisateur) + ALERTER sur toute
  // facture partiellement/totalement hors mois cible SANS l'exclure -- le pole transport doit
  // confirmer facture par facture avant validation finale (question en attente).
  const facturesDroitsTaxes = [];
  const lignesExclues = new Set();
  for (const [numFacture, fac] of parFacture) {
    if (fac.fretTotal === 0 && fac.taxeTotal !== 0) {
      facturesDroitsTaxes.push(numFacture);
      for (const r of fac.lignes) lignesExclues.add(r);
    } else if (moisCible && fac.moisEnvoi.size && !fac.moisEnvoi.has(moisCible)) {
      const detail = [...fac.moisEnvoi.entries()].sort((a, b) => b[1] - a[1]).map(([k, v]) => `${k}: ${v}`).join(', ');
      warnings.push(`Facture ${numFacture} : AUCUNE ligne datée dans le mois cible ${moisCible} (dates réelles : ${detail}) — probable facture déjà traitée un mois précédent/suivant, à confirmer avec le pôle transport avant facturation (non exclue automatiquement, cf. commentaire en tête de fichier).`);
    }
  }
  if (facturesDroitsTaxes.length) {
    infos.push(`Facture(s) 100% Droits & Taxes (aucun fret, ajustement douanier pur) exclue(s) du fichier livré : ${facturesDroitsTaxes.join(', ')} (${lignesExclues.size} ligne(s)).`);
  }

  // Export brut WMS (E/P) : m et m-1, meme mecanisme que Geodis/DPD/Lettres (core/exportBrut.js).
  const appRoot = path.resolve(__dirname, '../../..');
  const brutPaths = moisCible ? findBrutFiles(moisCible.length === 6 ? `${moisCible.slice(0, 4)}_${moisCible.slice(4)}` : moisCible, appRoot) : [];
  const epMap = brutPaths.length ? epParTrackingFromExport(readBrutRows(brutPaths)) : new Map();
  if (!brutPaths.length) warnings.push("Export WMS 'expéditions_brut' introuvable pour ce mois (E/P) — toutes les lignes sans correspondance seront classées 'P' par défaut.");

  const recs = [];
  let nEpDefaut = 0;
  for (const f of brutes) {
    for (const r of f.rows) {
      if (lignesExclues.has(r)) continue;
      const tracking = String(r[iTracking] || '').trim();
      if (!tracking) continue;
      const fret = iFret >= 0 ? num(r[iFret]) : 0;
      const divers = iDivers >= 0 ? num(r[iDivers]) : 0;
      const remise = iRemise >= 0 ? num(r[iRemise]) : 0;
      const droitsTaxes = iDroitsTaxes >= 0 ? num(r[iDroitsTaxes]) : 0;
      const montant = round2(fret + remise); // Shipment Detail!A = BK+BN (Fret + Remise)
      const pays = iPaysDest >= 0 ? String(r[iPaysDest] || '').trim().toUpperCase() : '';
      const service = iService >= 0 ? String(r[iService] || '').trim() : '';
      const mode = service === cfg.service_st ? 'ST' : 'FICP';
      const longueur = iLongueur >= 0 ? num(r[iLongueur]) : 0;
      const colisVolumineux = longueur > cfg.colis_volumineux_seuil_cm ? cfg.colis_volumineux_montant : 0;
      let ep = epMap.get(tracking);
      if (!ep) { ep = 'P'; nEpDefaut++; }
      const zoneFicp = cfg.zoning[pays];
      const zone = zoneFicp === undefined ? 'inconnu' : (zoneFicp || '');
      // Import ERP!M formule modele : =IF(J="France","ST",XLOOKUP(tracking,...,'Shipment Detail'!E:E))
      // -- J=Zone (PAS Pays) : Zone='France' UNIQUEMENT pour FR (cf. config.zoning.FR="France").
      const modeFinal = zone === 'France' ? 'ST' : mode;
      recs.push({
        tracking, compte: iCompte >= 0 ? String(r[iCompte] || '').trim() : '',
        montant, poids: iPoids >= 0 ? num(r[iPoids]) : 0,
        colis: iColis >= 0 ? num(r[iColis]) : 1,
        pays, zone, mode: modeFinal, colisVolumineux, ep,
        droitsTaxes: droitsTaxes || 0,
        divers,
        tva: cfg.tva_pays_ue.includes(pays) ? cfg.champs_fixes.tva_taux : 0,
      });
    }
  }
  if (nEpDefaut) infos.push(`${nEpDefaut} ligne(s) sans correspondance dans l'export WMS (E/P) — classée(s) 'P' par défaut.`);

  // Poids livres FedEx -> kg (ARRONDI.SUP, formule modele TCD!E = ROUNDUP(poids*0.453592,1)).
  const LB_TO_KG = 0.453592;
  const importRows = recs.map((rec) => ({
    // Import ERP!A formule modele : =IF(OR(M="FICP",M="IE",M="RE",M="International",M="Europe"),
    // "FEDEX INTERNATIONAL","FEDEX FRANCE") -- M='Mode d'envoi' (deja calcule ci-dessus, rec.mode).
    Transporteur: (rec.mode === 'FICP') ? 'FEDEX INTERNATIONAL' : 'FEDEX FRANCE',
    DateValidite: dateValidite || '',
    Ref1: '', Ref2: '', IdClient: '',
    Tracking: rec.tracking, Nom: '',
    EP: rec.ep, Pays: rec.pays, Zone: rec.zone,
    NbrColis: rec.colis, Poids: roundUp1(rec.poids * LB_TO_KG),
    Mode: rec.mode, TVA: rec.tva,
    DroitsTaxes: rec.droitsTaxes || 0, Assurance: 0,
    ZonesEloignees: 0, ColisVolumineux: rec.colisVolumineux,
    Adresses: 0, Fret: rec.montant, PlusValueB2C: 0, TaxeGasoil: '', NbColis: '',
  }));

  const { alerts, infos: valInfos } = validate(importRows);
  infos.push(...valInfos);

  const pdfPaths = files.pdf || [];
  const pdfs = [];
  for (const p of pdfPaths) {
    try {
      const r = await extractFedexPdfInfo(p);
      if (r && r.totalDu != null) pdfs.push(r);
      else warnings.push(`PDF ${path.basename(p)} : "Total dû" introuvable, ignoré pour la réconciliation.`);
    } catch (e) {
      warnings.push(`PDF ${path.basename(p)} : lecture impossible (${e.message}).`);
    }
  }
  // Reconciliation par facture (Bilan factures : Total TTC compare au PDF, PAS le Fret HT).
  const totalFretHt = round2(importRows.reduce((s, r) => s + (r.Fret || 0), 0));
  infos.push(`Total Fret HT (toutes lignes) = ${totalFretHt.toFixed(2)} EUR.`);
  for (const p of pdfs) {
    infos.push(`Facture PDF ${p.file}${p.numeroFacture ? ` (n° ${p.numeroFacture})` : ''} : Total dû = ${p.totalDu.toFixed(2)} EUR.`);
  }

  const controle = {
    'Colis volumineux': round2(importRows.reduce((s, r) => s + (r.ColisVolumineux || 0), 0)),
    'Frêt': totalFretHt,
    'Droits et taxes': round2(importRows.reduce((s, r) => s + (r.DroitsTaxes || 0), 0)),
  };

  return {
    header,
    rows: recs, recs, importRows, controle, warnings, alerts, infos,
    posteKeys: ['Colis volumineux', 'Frêt', 'Droits et taxes'], cfg, pdfs,
    sheetNames: { raw: 'Shipment Detail', import: 'Import ERP' },
    period: moisCible ? `${moisCible.slice(0, 4)}_${moisCible.slice(4)}` : 'export',
  };
}

module.exports = {
  id: 'fedex',
  name: 'FedEx',
  status: 'ready',
  taxeGasoil: "Aucune (pas de surcharge carburant repercutee au client dans le fichier livre -- colonne Gazole/'V' de Import ERP toujours vide, confirme sur juin 2026).",
  method: "1 CSV brut par mois (export portail FedEx, en-tetes FR, colonnes resolues par nom). Montant = Frais de transport + Remise. Mode d'envoi = 'ST' si Service='FedEx Priority' sinon 'FICP'. Zone = lookup pays -> Zoning!FICP. TVA = 20% si pays destinataire dans la liste UE (config), sinon 0%. Colis volumineux = 10EUR forfait si longueur dimensionnee > 60cm. E/P via export WMS brut m/m-1 (repli 'P'). Factures 100% Droits&Taxes exclues. Reconciliation PDF sur 'Total du'.",
  inputs: [
    { key: 'csv', label: 'Shipment detail (CSV export portail FedEx)', accept: '.csv', multiple: true, required: true },
    { key: 'pdf', label: 'Facture(s) PDF FedEx (contrôle du "Total dû")', accept: '.pdf', multiple: true, required: false },
  ],
  outputNaming: { workbook: '{period}_Facture Fedex', import: '{period}_Fedex_Import' },
  finalizer: {
    script: '../automatisation/finaliser_fedex.py',
    template: '../Transporteurs/Fedex/2026_06_Facture Fedex.xlsx',
    buildArgs: (files) => [...(files.csv || []), '--pdf', ...(files.pdf || [])],
  },
  process,
};
