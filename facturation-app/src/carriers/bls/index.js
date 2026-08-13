// ============================================================================
//  Adaptateur transporteur : BLS
//  Source UNIQUE = la facture PDF recue (le xlsx brut equivalent contient la
//  MEME donnee mais en mise en page facture avec cellules fusionnees non
//  alignees entre l'en-tete et les lignes de donnees -- source d'erreurs de
//  parsing constatee en pratique -- decision utilisateur 2026-08-11 : ne
//  parser QUE le PDF, format texte lineaire plus simple a fiabiliser).
//
//  1 ligne logique = 1 "Dossier", identifiee dans le texte PDF par le motif
//  "DD/MM/AAAA NNNNNNNN ..." (date + dossier sur 8 chiffres). Le Libelle/
//  Montant peuvent etre sur la MEME ligne texte ou sur la ligne SUIVANTE
//  (retour a la ligne PDF, ex. "V/Réf ..." ou juste un espace) -- fusionnes
//  en un seul item. Une ligne "Dossier" SANS montant du tout (ex. juillet
//  2026 : Dossier 13301, "Vannes") est un montant reellement NUL (verifie en
//  comparant la somme des montants trouves au Total HT officiel de la
//  facture : la somme des 26 autres lignes egale deja EXACTEMENT le total),
//  PAS une ligne a ignorer.
//
//  "Id client"/"Nbr Colis"/"Poids" : completes via un export CSV optionnel du portail
//  "AffreTrans" (upload manuel, cle de jointure "Récépissé" = "Dossier" -- confirme sur
//  juin 2026, 13/15 lignes retrouvees). Sans ce fichier (ou pour les Dossiers absents de
//  l'export), ces champs restent vides/a 0 avec une alerte, a completer manuellement comme
//  avant (RECHERCHEX cote pole transport).
// ============================================================================
const pdfParse = require('pdf-parse');
const path = require('path');
const fs = require('fs');
const { num, round2 } = require('../../core/csv');
const { validate } = require('../../core/validate');
const cfg = require('./config.json');

/** Export CSV du portail "AffreTrans" (';' separe, UTF-8 avec BOM -- different du latin1
 * habituel des autres transporteurs). Cle de jointure : colonne "Récépissé" = "Dossier"
 * BLS (confirme sur juin 2026 : 13/15 lignes retrouvees directement -- 2 dossiers restent
 * sans correspondance, l'un absent de l'export, l'autre present mais sans Récépissé
 * renseigne -- PAS de rapprochement par deduction/montant, trop risque). Fournit "ID
 * Client" (deja le bon format numerique, ex. '8041'), "Poids (kg)", "Nb palettes".
 * Filtre implicitement sur Transporteur='BLS' (les autres lignes n'ont de toute facon pas
 * de Dossier BLS en Récépissé). */
function parseAffretementCsv(csvPath) {
  const text = fs.readFileSync(csvPath).toString('utf8').replace(/^﻿/, '');
  const lines = text.split(/\r?\n/).filter((l) => l.length > 0);
  if (!lines.length) return new Map();
  const header = lines[0].split(';').map((h) => h.trim());
  const iRecepisse = header.indexOf('Récépissé');
  const iIdClient = header.indexOf('ID Client');
  const iPoids = header.indexOf('Poids (kg)');
  const iNbPalettes = header.indexOf('Nb palettes');
  const iTransporteur = header.indexOf('Transporteur');
  if (iRecepisse < 0) return new Map();

  const map = new Map();
  for (const line of lines.slice(1)) {
    const cols = line.split(';');
    if (iTransporteur >= 0 && (cols[iTransporteur] || '').trim().toUpperCase() !== 'BLS') continue;
    const recepisse = (cols[iRecepisse] || '').trim();
    if (!recepisse || map.has(recepisse)) continue;
    map.set(recepisse, {
      idClient: iIdClient >= 0 ? (cols[iIdClient] || '').trim() : '',
      poids: iPoids >= 0 ? num(cols[iPoids]) : 0,
      nbPalettes: iNbPalettes >= 0 ? num(cols[iNbPalettes]) : 0,
    });
  }
  return map;
}

/** pdf-parse (pdfjs-dist) decode certains glyphes accentues de la police du PDF BLS avec un
 * mauvais mapping (constate empiriquement en comparant a une extraction pypdf/Python,
 * fiable elle) -- ensemble FINI de 7 caracteres mal rendus, toujours les memes, remplaces
 * ici par leur vrai caractere francais. Necessaire car le libelle est ecrit tel quel dans
 * le classeur Excel final (contrairement a un simple usage regex sur des chiffres). */
const MOJIBAKE_MAP = { 'È': 'é', 'Ë': 'è', 'Í': 'ê', 'Ó': 'î', '‡': 'à', 'Ä': '€', '∞': '°', '‚': 'â' };
function fixMojibake(s) {
  let out = String(s || '');
  for (const [bad, good] of Object.entries(MOJIBAKE_MAP)) out = out.split(bad).join(good);
  return out;
}

/** Rendu de page PDF custom : pdf-parse (par defaut) colle les items de texte SANS espace
 * quand le PDF ne code pas d'espace explicite entre eux (constate sur les factures BLS :
 * "01/06/202600012118De 21..." -- date/dossier/libelle tous colles). Reconstruit un espace
 * a partir de l'ECART REEL entre la fin d'un item (x + largeur) et le debut du suivant sur
 * la meme ligne (meme Y) -- redonne un texte "01/06/2026 00012118 De 21..." exploitable par
 * des regex classiques, sans dependre du rendu par defaut de la lib. */
function pagerenderWithSpacing(pageData) {
  return pageData.getTextContent().then((textContent) => {
    let lastY = null, lastX = null, lastWidth = 0, text = '';
    for (const item of textContent.items) {
      const x = item.transform[4];
      const y = item.transform[5];
      if (lastY != null && Math.abs(lastY - y) < 0.5) {
        if (x - (lastX + lastWidth) > 1) text += ' ';
      } else if (lastY != null) {
        text += '\n';
      }
      text += item.str;
      lastY = y; lastX = x; lastWidth = item.width;
    }
    return text;
  });
}

/** Nombre "français" toléré (espaces insécables/normaux comme séparateur de milliers,
 * virgule OU point comme décimale) -> Number, ou null si pas un montant. */
function parseMontant(s) {
  const t = String(s || '').trim().replace(/[\s ]/g, '');
  if (!/^\d+([.,]\d{1,2})?$/.test(t)) return null;
  return num(t.replace(',', '.'));
}

/** Découpe une ligne de données PDF en { date, dossier, reste } si elle commence par
 * "DD/MM/AAAA NNNNNNNN" (dossier sur 8 chiffres, tel qu'imprimé dans le PDF). */
function matchDateDossier(line) {
  const m = /^(\d{2}\/\d{2}\/\d{4})\s+0*(\d+)\s*(.*)$/.exec(line.trim());
  if (!m) return null;
  return { date: m[1], dossier: m[2], reste: m[3] };
}

/** Extrait libelle/unite/quantite/montant/codeTva depuis la fin d'une ligne de type
 * "De X A Y [M QQ.QQ] MMM.MM C" -- unite+quantite optionnelles (absentes sur certaines
 * lignes, confirme sur juin 2026 : "...Aubagne 1 300.00 1" sans "M"/quantite ; et sur
 * juillet 2026 : "...Bussy-Saint-Georges 195.00 1" idem), montant+codeTva optionnels
 * (ligne "Dossier" sans aucun montant -> objet retourne avec montant=null). */
function parseResteLigne(reste) {
  const t = reste.trim();
  // Montant + CodeTva en fin de ligne : "MMM.MM C" (C = 1 chiffre, code TVA).
  const mEnd = /^(.*?)(?:\s+([\d ]+[.,]\d{1,2})\s+(\d+))?$/.exec(t);
  let libelle = t, montant = null, codeTva = '';
  if (mEnd && mEnd[2] != null) {
    libelle = mEnd[1].trim();
    montant = parseMontant(mEnd[2]);
    codeTva = mEnd[3];
  }
  // Unite ('M') + Quantite en fin de libelle restant : "... M QQ.QQ"
  let unite = '', quantite = 0;
  const mUnite = /^(.*?)\s+([A-Z])\s+([\d ]+[.,]\d{1,2})$/.exec(libelle);
  if (mUnite) {
    libelle = mUnite[1].trim();
    unite = mUnite[2];
    quantite = parseMontant(mUnite[3]) || 0;
  }
  return { libelle, unite, quantite, montant, codeTva };
}

/** Numero de facture BLS : bloc "N° Client"+7 chiffres+date+folio colles/espaces
 * ensemble dans l'en-tete PDF (ex. "01LARUCH2601856 31/07/2026 1" -> facture
 * "2601856"), confirme identique sur juin ET juillet 2026. */
function extractNFacture(text) {
  const m = /[A-Z0-9]{5,}(\d{7})\s*\d{2}\/\d{2}\/\d{4}\s*\d+/.exec(text);
  return m ? m[1] : null;
}

/** Total HT officiel : 1re des 3 lignes-nombres isolees juste avant "...EUR" en fin de
 * facture (Total HT / Montant TVA / Net a payer), confirme sur juin ET juillet 2026. */
function extractTotalHt(text) {
  const m = /^([\d ]+[.,]\d{1,2})\s*$\n^([\d ]+[.,]\d{1,2})\s*$\n^([\d ]+[.,]\d{1,2})\s*EUR\s*$/m.exec(text);
  return m ? parseMontant(m[1]) : null;
}

/** Parse le tableau de lignes "Dossier" d'une facture PDF BLS -- source UNIQUE (cf.
 * docstring d'en-tete). Chaque ligne de données commence par "DD/MM/AAAA NNNNNNNN" ; le
 * reste de la ligne (libelle+montant) peut être sur cette même ligne texte OU splitté sur
 * la ligne suivante (retour à la ligne PDF, motif "V/Réf ..." ou libellé vide) -- fusionné
 * en un seul item. Une ligne "Dossier" sans AUCUN montant trouvable (ni sur sa propre
 * ligne, ni sur la suivante) est conservée avec montant=0 (pas ignorée, cf. docstring).
 */
async function parseFacturePdf(pdfPath) {
  const buf = require('fs').readFileSync(pdfPath);
  const { text: rawText } = await pdfParse(buf, { pagerender: pagerenderWithSpacing });
  const text = fixMojibake(rawText);

  const nFacture = extractNFacture(text);
  const totalHt = extractTotalHt(text);

  const lines = text.split(/\r?\n/);
  const startIdx = lines.findIndex((l) => l.trim() === 'FACTURE');
  const endIdx = lines.findIndex((l) => /^Escompte pour r/i.test(l.trim()));
  const dataLines = lines.slice(startIdx >= 0 ? startIdx + 1 : 0, endIdx >= 0 ? endIdx : lines.length);

  const items = [];
  let pending = null; // { date, dossier, libellePrefix } en attente d'un montant sur la ligne suivante
  for (const rawLine of dataLines) {
    const line = rawLine.trim();
    if (!line) continue;
    const head = matchDateDossier(line);

    if (head) {
      // Nouvelle ligne "Dossier" : si elle porte deja un montant, item complet immediat.
      // Sinon (reste vide ou "V/Réf ...") -> en attente de la ligne suivante.
      if (pending) { items.push({ ...pending, libelle: pending.libellePrefix, unite: '', quantite: 0, montant: 0, codeTva: '' }); }
      const parsed = parseResteLigne(head.reste);
      if (parsed.montant != null) {
        items.push({ date: head.date, dossier: head.dossier, libelle: parsed.libelle, unite: parsed.unite, quantite: parsed.quantite, montant: parsed.montant, codeTva: parsed.codeTva });
        pending = null;
      } else {
        pending = { date: head.date, dossier: head.dossier, libellePrefix: parsed.libelle };
      }
      continue;
    }

    if (pending) {
      // Ligne de suite (pas de date/dossier en tete) : contient le vrai libelle + montant.
      const parsed = parseResteLigne(line);
      const libelle = pending.libellePrefix ? `${pending.libellePrefix}\n${parsed.libelle}` : parsed.libelle;
      items.push({
        date: pending.date, dossier: pending.dossier, libelle,
        unite: parsed.unite, quantite: parsed.quantite,
        montant: parsed.montant != null ? parsed.montant : 0,
        codeTva: parsed.codeTva,
      });
      pending = null;
      continue;
    }
    // Ligne non rattachable (bloc adresse/entete deja hors dataLines normalement) -> ignorée.
  }
  if (pending) items.push({ ...pending, libelle: pending.libellePrefix, unite: '', quantite: 0, montant: 0, codeTva: '' });

  return { file: path.basename(pdfPath), nFacture, totalHt, items };
}

function firstDayOfMonth(dateStr) {
  // format "DD/MM/AAAA"
  const m = /^(\d{2})\/(\d{2})\/(\d{4})$/.exec(dateStr || '');
  return m ? `01/${m[2]}/${m[3]}` : null;
}

async function process(files) {
  const pdfPaths = files.pdf || files.facture || [];
  if (!pdfPaths.length) throw new Error('Aucun fichier fourni (attendu : facture BLS reçue, PDF).');

  const affretementPaths = files.affretement || [];
  if (!affretementPaths.length) {
    throw new Error('Export "AffreTrans" manquant (obligatoire) : sans lui, "Id client"/"Nbr Colis"/"Poids" ne peuvent pas être complétés (aucune autre source disponible pour BLS).');
  }

  const warnings = [];
  const infos = [];
  let affretementMap = new Map();
  for (const p of affretementPaths) {
    try {
      const m = parseAffretementCsv(p);
      for (const [k, v] of m) affretementMap.set(k, v);
    } catch (e) {
      warnings.push(`Export affrètement ${path.basename(p)} : lecture impossible (${e.message}).`);
    }
  }

  const factures = [];
  for (const p of pdfPaths) {
    try {
      factures.push(await parseFacturePdf(p));
    } catch (e) {
      warnings.push(`Facture PDF ${path.basename(p)} : lecture impossible (${e.message}).`);
    }
  }

  let dateValidite = null;
  const recs = [];
  for (const f of factures) {
    for (const it of f.items) {
      if (!dateValidite) {
        const d = firstDayOfMonth(it.date);
        if (d) dateValidite = d;
      }
      recs.push({
        facture: f.nFacture || '',
        dossier: String(it.dossier || '').trim(),
        libelle: it.libelle,
        montantHt: round2(it.montant || 0),
        // Navette = trajet INTERNE entre les 2 sites "21 Longvic" <-> "21 Créancey" (pas
        // seulement les lignes explicitement libellees "navette" : confirme sur juillet 2026
        // que "V/Réf ALLER/RETOUR DU 08/06 De 21 Longvic A 21 Créancey", "V/Réf MAJE 09/07 De
        // 21 Longvic A 21 Créancey" et "De 21 Longvic A 21 Créancey" -- sans le mot "navette"
        // -- sont le MEME trajet interne que les lignes "V/Réf NAVETTE ..." -- decision
        // utilisateur 2026-08-12). Cout REEL paye a BLS (pas 0€, contrairement aux lignes
        // "sans montant" ci-dessous) mais PAS refacturable au client -> reste dans le
        // classeur/reconciliation PDF (montant reel paye), exclue de importRows uniquement.
        isNavette: /21\s*longvic/i.test(it.libelle || '') && /21\s*cr[ée]ancey/i.test(it.libelle || ''),
      });
    }
  }
  if (!recs.length) warnings.push('Aucune ligne "Dossier" trouvée dans les facture(s) fournies.');

  // Ligne "sans montant" : Poids ET toutes les colonnes O->W (Droits et taxes, Assurance,
  // Zones eloignees, Colis volumineux, Adresses, Fret, plus-value BtoC, Gazole, Nb Colis) a
  // zero -> rien a facturer nulle part, supprimee de l'import (regle demandee par
  // l'utilisateur 2026-08-11, meme principe que Mondial Relay). Pour BLS specifiquement,
  // Poids n'a AUCUNE source disponible (pas dans l'export expeditions_brut -- BLS est de
  // l'affretement palette, pas gere par le WMS colis -- confirme le 2026-08-11) donc TOUJOURS
  // 0 : la regle retombe en pratique sur Fret=0 (seule colonne O->W jamais remplie pour BLS).
  let nbLignesSansMontantSupprimees = 0;
  const recsFiltres = recs.filter((rec) => {
    if (rec.montantHt === 0) { nbLignesSansMontantSupprimees++; return false; }
    return true;
  });

  // Poids/Nbr Colis par defaut pour les lignes navette dans l'import (decision
  // utilisateur 2026-08-12) : ces trajets n'ont pas de poids/colis reels mesures (pas
  // de correspondance AffreTrans standard, trajet interne entre sites) -> valeurs
  // fixes forfaitaires.
  const POIDS_NAVETTE_DEFAUT = 13200;
  const NBRCOLIS_NAVETTE_DEFAUT = 33;

  let nbAffretementTrouve = 0;
  const nbNavetteExclues = recsFiltres.filter((rec) => rec.isNavette).length;
  const montantNavetteExclu = round2(recsFiltres.filter((rec) => rec.isNavette).reduce((s, rec) => s + rec.montantHt, 0));
  const importRows = recsFiltres.map((rec) => {
    const aff = affretementMap.get(rec.dossier);
    // Navette : poids/colis forfaitaires (POIDS_NAVETTE_DEFAUT), pas de jointure
    // AffreTrans attendue -> ne compte pas dans nbAffretementTrouve/nbSansAffretement
    // (pas une vraie lacune a signaler comme pour les autres lignes).
    if (aff && !rec.isNavette) nbAffretementTrouve++;
    return {
      Transporteur: cfg.champs_fixes.Transporteur,
      DateValidite: dateValidite || '',
      // Id client : TOUJOURS vide dans le fichier import (decision utilisateur 2026-08-12,
      // s'applique a tous les transporteurs) -- aff.idClient reste dispo pour d'autres
      // usages internes mais n'est plus reporte ici.
      Ref1: '', Ref2: '', IdClient: '',
      // Le libelle peut porter un retour a la ligne (motif "V/Ref...\nDe X A Y", cf.
      // parseFacturePdf) -- remplace par un espace pour l'affichage/export (CSV/XLSX ne
      // gerent pas le wrap text comme la cellule Excel du classeur clone).
      Tracking: rec.dossier, Nom: rec.libelle.replace(/\n/g, ' '),
      EP: cfg.champs_fixes['E/P'], Pays: cfg.champs_fixes.Pays, Zone: cfg.champs_fixes.Zone,
      NbrColis: rec.isNavette ? NBRCOLIS_NAVETTE_DEFAUT : (aff ? Math.round(aff.nbPalettes) : 0),
      Poids: rec.isNavette ? POIDS_NAVETTE_DEFAUT : (aff ? aff.poids : 0),
      Mode: cfg.champs_fixes.Mode, TVA: cfg.champs_fixes.tva,
      DroitsTaxes: 0, Assurance: 0,
      ZonesEloignees: 0, ColisVolumineux: 0, Adresses: 0,
      // Navette : PAS refacturable au client -> Fret=0 dans l'import, meme si un montant
      // reel a ete paye a BLS (rec.montantHt, garde pour la reconciliation PDF ci-dessous).
      Fret: rec.isNavette ? 0 : rec.montantHt, PlusValueB2C: 0, TaxeGasoil: '', NbColis: '',
      _isNavette: rec.isNavette, // interne, retire avant retour -- exclu de validate() ci-dessous
    };
  });

  // Navette : COLIS=0/FRET=0 sont des consequences ATTENDUES de la regle ci-dessus (deja
  // expliquees en infos), pas de vraies anomalies -> exclues de validate() pour ne pas
  // generer 2 alertes de bruit par ligne navette (COLIS=0 + FRET=0).
  const { alerts, infos: valInfos } = validate(importRows.filter((o) => !o._isNavette), { skipPoidsDecimal: true });
  for (const o of importRows) delete o._isNavette;
  infos.push(...valInfos);
  const nbHorsNavette = importRows.length - nbNavetteExclues;
  const nbSansAffretement = nbHorsNavette - nbAffretementTrouve;
  infos.push(`Export affrètement : "Id client"/"Nbr Colis"/"Poids" complétés pour ${nbAffretementTrouve}/${nbHorsNavette} ligne(s) (jointure sur "Récépissé" = "Dossier", hors navette).`);
  if (nbSansAffretement) infos.push(`${nbSansAffretement} ligne(s) sans correspondance dans l'export affrètement — "Id client"/"Nbr Colis"/"Poids" à compléter manuellement.`);
  if (nbLignesSansMontantSupprimees) infos.push(`${nbLignesSansMontantSupprimees} ligne(s) supprimée(s) (Poids et Frêt tous à 0 — rien à facturer sur cette ligne).`);
  if (nbNavetteExclues) infos.push(`${nbNavetteExclues} ligne(s) "navette" (trajet interne, non refacturable au client, ${montantNavetteExclu.toFixed(2)} EUR payés à BLS) — Frêt=0 dans l'import ERP, Poids=${POIDS_NAVETTE_DEFAUT} (forfaitaire).`);

  // Reconciliation : Total HT officiel (extrait du PDF, bloc de synthese) compare au total
  // calcule (somme des montants par ligne "Dossier") -- meme fichier, meme source, donc
  // un ecart ici signale un bug de parsing, pas une divergence metier normale.
  const totalParFacture = new Map();
  for (const rec of recs) {
    if (!rec.facture) continue;
    totalParFacture.set(rec.facture, round2((totalParFacture.get(rec.facture) || 0) + rec.montantHt));
  }
  for (const f of factures) {
    if (f.totalHt == null) { warnings.push(`Facture PDF ${f.file} : Total HT introuvable dans le PDF, réconciliation ignorée.`); continue; }
    const calcule = totalParFacture.get(f.nFacture);
    if (calcule == null) { warnings.push(`Facture PDF ${f.file} (${f.nFacture}) : aucune ligne "Dossier" trouvée.`); continue; }
    const ecart = round2(f.totalHt - calcule);
    const statut = Math.abs(ecart) >= 0.05 ? 'A VERIFIER' : 'OK';
    infos.push(`Facture PDF ${f.file} (${f.nFacture}) : Total HT = ${f.totalHt.toFixed(2)} EUR, calculé = ${calcule.toFixed(2)} EUR (écart ${ecart >= 0 ? '+' : ''}${ecart.toFixed(2)} EUR -> ${statut})`);
  }

  // controle (total affiche a l'ecran) = total FACTURABLE au client, base sur importRows
  // (Fret=0 pour les lignes navette, deja integre ci-dessus) -- distinct du total
  // reconcilie avec le PDF ci-dessus, qui lui inclut la navette (montant reellement paye
  // au transporteur, rec.montantHt).
  const controle = { Fret: round2(importRows.reduce((s, r) => s + (r.Fret || 0), 0)) };

  return {
    // Classeur ("Factures BLS") = recsFiltres, PAS recs : la/les ligne(s) a montant 0€
    // (rien a facturer, cf. plus haut) sont retirees aussi du classeur, pas seulement de
    // l'import ERP (decision utilisateur 2026-08-12). Navette, elle, reste bien presente
    // (recsFiltres inclut encore les lignes isNavette, exclues uniquement de importRows).
    header: ['Date Prestation', 'Dossier', 'Libellé', 'Unité', 'Quantité', 'Montant H.T.', 'Code Tva'],
    rows: recsFiltres, recs: recsFiltres, importRows, controle, warnings, alerts, infos,
    posteKeys: ['Fret'], cfg, factures,
    sheetNames: { raw: 'Factures BLS', import: 'Import CSV' },
    period: dateValidite ? `${dateValidite.slice(6)}_${dateValidite.slice(3, 5)}` : 'export',
  };
}

module.exports = {
  id: 'bls',
  name: 'BLS',
  status: 'ready',
  viticolis: true,
  taxeGasoil: "A definir -- aucune colonne/mecanisme de taxe gasoil identifie dans le classeur fait a la main ni dans les 2 videos process pour ce transporteur.",
  method: "Bien vérifier que tous les trackings sont déjà rentrés dans l'ERP pour les différentes expé (faire le rapprochement avec le fichier affrètement) pour éviter les avaries d'imports.",
  inputs: [
    { key: 'pdf', label: 'Facture(s) PDF BLS', accept: '.pdf', multiple: true, required: true },
    { key: 'affretement', label: 'Export "AffreTrans" (CSV, pour Id client/Nbr Colis/Poids)', accept: '.csv', multiple: false, required: true },
  ],
  // Classeur = CLONE FIDELE du fichier fait a la main (Excel COM/Python), comme DPD/Geodis/GLS/Mondial Relay.
  outputNaming: { workbook: '{period}_Facture BLS', import: '{period}_BLS_Import' },
  finalizer: {
    script: '../automatisation/finaliser_bls.py',
    template: '../Transporteurs/BLS/2026_06_Facture BLS.xlsx',
    buildArgs: (files) => ['--pdf', ...(files.pdf || []), ...(files.affretement && files.affretement.length ? ['--affretement', ...files.affretement] : [])],
  },
  process,
};
