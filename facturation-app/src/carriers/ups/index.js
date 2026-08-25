// ============================================================================
//  Adaptateur transporteur : UPS (export billing.ups.com, CSV multi-fichiers)
//  Plusieurs fichiers CSV bruts par mois (1 par facture, "Invoice_<numero>_
//  <date>.csv"), SANS EN-TETE (250 colonnes, format natif UPS Billing) --
//  colonnes resolues PAR POSITION (jamais par nom, aucun en-tete disponible),
//  confirme sur les fichiers reels de juin 2026. Rassemblage multi-fichiers =
//  Power Query "Obtenir des donnees > A partir d'un dossier" cote humain,
//  reproduit ici par simple concatenation de tous les CSV fournis.
//
//  Decalage de position CONFIRME empiriquement (2026-08-20) : le classeur
//  modele colle ces donnees a partir de la colonne E de l'onglet "Facture
//  UPS" (A-D = calculees : Clients/Montant assurance/Mode envoi/Categorie) --
//  donc position CSV brut N correspond a colonne Facture UPS (N+4). Verifie
//  sur plusieurs colonnes cle : CSV pos 21 (Numero de suivi) = Facture UPS
//  col 25 (Y) ; CSV pos 44/45/46 (Code de classe/Code de description/
//  Description) = Facture UPS col 48/49/50 (AV/AW/AX) ; CSV pos 82 (Pays
//  destinataire, EN-TETE NATIF UPS TROMPEUR "Ville de l'acheteur") = Facture
//  UPS col 86 (CH).
//
//  2 "transporteurs" distincts (UPS vs UPS_COD, meme transporteur physique,
//  grilles/comptes differents) : determine par les 6 caracteres du COMPTE
//  extraits du TRACKING lui-meme (positions 3-8, ex. tracking
//  "1ZJ40E826836318856" -> compte "J40E82"), PAS par une colonne dediee --
//  formule modele exacte : XLOOKUP(RIGHT(LEFT(tracking,8),6),'Comptes UPS'!
//  A:A,B:B).
//
//  Categorie (poste ERP) : cascade EXACTE du modele (Facture UPS!D) --
//  1) NB.SI('ST SV'!Q:Q,description)<>0 -> "Adresse"
//  2) NB.SI('ST SV'!D:D,description)<>0 -> "plus-value BtoC"
//  3) codeClasse="FRT" -> "Fret"
//  4) codeClasse="TAX" -> "TVA"
//  5) sinon XLOOKUP(codeDescription,'Charge.CHG_CODE'!A:C) -- "CODE INCONNU"
//     est une VRAIE categorie du modele reel (239/472 codes), PAS une
//     erreur -- jamais reclassee, reste telle quelle dans le fichier livre.
//  NOTE : la vraie cascade modele teste 'ST SV'!Q:Q et D:D sur la colonne
//  DESCRIPTION (AX), mais ces colonnes Q/D de 'ST SV' n'ont jamais ete
//  vues remplies dans l'investigation (onglet ST SV = uniquement A/B =
//  description->ST/SV) -- traite ici comme TOUJOURS VIDE (jamais "Adresse"/
//  "plus-value BtoC" par cette voie), la vraie source de "plus-value BtoC"
//  etant le code "RES" (categories.RES) confirme dans Charge.CHG_CODE.
//
//  Montant assurance (Facture UPS!B) : =IF(codeDescription="EVS",valeurDeclaree,0)
//  -- code "EVS" = "Excess Value Service", confirme categories.EVS="Assurance".
//
//  Mode envoi (Facture UPS!C) : =IF(codeClasse="FRT",XLOOKUP(description,'ST
//  SV'!A:A,B:B),0) -- ST=Standard/SV=Saver, table config.service_vers_mode.
//
//  Zone : FOURNIE NATIVEMENT par UPS dans son export brut (pas calculee cote
//  carrier) -- confirme empiriquement (colonne "Zone" native UPS, valeur
//  directe). Poids/Nombre de colis/Montant assurance : agreges par tracking
//  via MAX (reproduit le TCD 'zone colis poids assurance', DataFields "Max
//  de ..."). Montant net (poste ERP) : agrege par tracking x categorie via
//  SOMME (reproduit le TCD 'TCD', RowField=tracking, ColField=Categorie,
//  DataField=Somme de Montant net).
//
//  E/P : =IF(declarationERP="particulier","P",IF(plusValueBtoC="","E","P"))
//  -- priorite absolue a la declaration ERP interne (export brut WMS m/m-1),
//  sinon la presence d'une plus-value BtoC (poste RES/EVS) force "P" (indice
//  indirect qu'UPS a traite l'envoi comme entreprise). Point relais->entreprise
//  deja gere par epParTrackingFromExport (core/exportBrut.js, meme regle que
//  Geodis/DPD/FedEx).
//
//  Poids UPS vs UPS_COD (formule modele 'zone colis poids assurance'!J) :
//  =IF(nbColis>3,ARRONDISUP(poids,0),IF(poids<10,ARRONDISUP(poids,1),
//  ARRONDISUP(poids,0))) pour UPS_COD ; =ARRONDISUP(poids,0) pour UPS.
//
//  Colis volumineux : bareme par palier sur le poids (formule modele exacte,
//  Fichier import!R) : <3kg=3€, <15kg=15€, <50kg=35€, <100kg=59€, <150kg=
//  177€, sinon ARRONDISUP(poids/59,0)*59€.
//
//  Assurance : =MAX(10,ARRONDISUP(2%*valeurDeclaree,2)) si valeur declaree>0,
//  sinon vide (formule modele exacte).
//
//  Zones eloignees : forfait fixe 40€ si poste "Droits et taxes" TCD non vide
//  (formule modele exacte, Fichier import!T -- pas un vrai calcul de zone
//  eloignee, un forfait declenche par la presence de droits/taxes douaniers).
//
//  Colis 1Z79 (regle FACTURATION EXCEL.docx, confirmee 2026-08-20) : colis
//  viticulteur retourne chez La Ruche (pas un vrai envoi client) -- EXCLU de
//  l'import, a signaler pour demande d'avoir (1x/mois ou trimestre).
//
//  E/P : uniquement via export WMS brut m/m-1 (core/exportBrut.js, meme
//  mecanisme que FedEx/Geodis/DPD) -- pas de fallback SI(X="";"entreprise";
//  "particulier") code ici (le docx le mentionne comme DERNIER recours
//  manuel si l'export brut ne suffit pas, laisse a la verification humaine
//  via un warning explicite plutot qu'invente une regle non verifiable).
// ============================================================================
const path = require('path');
const { num, round2, roundUp1 } = require('../../core/csv');
const { validate } = require('../../core/validate');
const { findBrutFiles, readBrutRows, epParTrackingFromExport } = require('../../core/exportBrut');
const cfg = require('./config.json');

// Positions CSV brut (1-based, converties en 0-based) -- decalage CONFIRME -4 vs Facture UPS
// (Facture UPS colle les donnees brutes a partir de la colonne E, A-D = calculees) : position
// CSV brut N (1-based) = colonne Facture UPS (N+4). Formule uniforme : <col Facture UPS> - 4 - 1.
const COL = {
  numeroCompte: 7 - 4 - 1,     // Facture UPS G (7) "Numero de compte"
  dateFacture: 9 - 4 - 1,      // Facture UPS I (9) "Date de la facture"
  numeroFacture: 10 - 4 - 1,   // Facture UPS J (10) "Numero de facture"
  montantFacture: 15 - 4 - 1,  // Facture UPS O (15) "Montant" (montant facture globale, repete)
  ref1: 20 - 4 - 1,            // Facture UPS T (20) "Numero de reference 1 de l'envoi"
  ref2: 21 - 4 - 1,            // Facture UPS U (21) "Numero de reference 2 de l'envoi"
  numeroSuivi: 25 - 4 - 1,     // Facture UPS Y (25) "Numero de suivi" -- confirme empiriquement
  nombreColis: 23 - 4 - 1,     // Facture UPS W (23) "Nombre de colis" -- natif export UPS brut
  poidsFacture: 33 - 4 - 1,    // Facture UPS AG (33) "Poids facture"
  zone: 38 - 4 - 1,            // Facture UPS AL (38) "Zone" -- NATIVE UPS, jamais calculee
  codeClasse: 48 - 4 - 1,      // Facture UPS AV (48) "Code de classification des frais"
  codeDescription: 49 - 4 - 1, // Facture UPS AW (49) "Code de description des frais"
  description: 50 - 4 - 1,     // Facture UPS AX (50) "Description des frais"
  valeurBase: 53 - 4 - 1,      // Facture UPS BA (53) "Valeur de base"
  montantNet: 57 - 4 - 1,      // Facture UPS BE (57) "Montant net"
  pays: 86 - 4 - 1,            // Facture UPS CH (86) "Pays" (en-tete natif UPS trompeur "Ville de l'acheteur")
};

function normalizeCompte(c) {
  return String(c || '').replace(/^'/, '').trim().toUpperCase();
}

/** Lit un CSV UPS brut (Billing, SANS EN-TETE, separateur virgule, latin1 -- confirme sur
 * les fichiers reels de juin 2026). Gestion de guillemets minimale (valeurs UPS n'en
 * contiennent jamais dans l'echantillon reel, mais reste tolerant). */
function readUpsCsv(p) {
  const fs = require('fs');
  const buf = fs.readFileSync(p);
  const text = buf.toString('latin1');
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
  return rows.filter((r) => r.some((v) => v !== ''));
}

/** Categorie ERP d'une ligne (cascade exacte du modele, cf. docstring en tete). */
function categoriePour(codeDescription) {
  const key = String(codeDescription || '').trim().toUpperCase();
  if (!key) return null;
  return cfg.categories[key] || 'CODE INCONNU';
}

/** Poids UPS vs UPS_COD (formule modele 'zone colis poids assurance'!I/J). */
function poidsArrondi(transporteur, nbColis, poids) {
  if (transporteur === 'UPS_COD') {
    if (nbColis > 3) return Math.ceil(poids);
    if (poids < 10) return Math.ceil(poids * 10) / 10;
    return Math.ceil(poids);
  }
  return Math.ceil(poids);
}

/** Colis volumineux : bareme par palier SUR LE MONTANT REEL du poste "Colis volumineux"
 * facture par UPS (formule modele exacte : IF(TCD!H=0,"",IF(TCD!H<3,3,...)) -- TCD!H est le
 * poste ERP montant, PAS le poids du colis, malgre le nom trompeur des variables -- BUG
 * TROUVE 2026-08-20 lors de la construction initiale : j'avais applique ce bareme sur le
 * POIDS de CHAQUE tracking, ce qui remplissait a tort "Colis volumineux" sur des lignes sans
 * aucune charge UPS reelle de ce type (Fret=0, DroitsTaxes rempli mais pas de vrai "colis
 * volumineux" facture) -- corrige : le bareme s'applique uniquement si le poste "Colis
 * volumineux" (categorie UPS reelle, table Charge.CHG_CODE) a un montant non nul). */
function colisVolumineuxMontant(montantReel) {
  if (!montantReel) return 0;
  if (montantReel < 3) return 3;
  if (montantReel < 15) return 15;
  if (montantReel < 50) return 35;
  if (montantReel < 100) return 59;
  if (montantReel < 150) return 177;
  return Math.ceil(montantReel / 59) * 59;
}

async function process(files, opts) {
  const csvPaths = files.csv || [];
  if (!csvPaths.length) throw new Error('Aucun fichier fourni (attendu : factures UPS Billing, CSV export portail).');

  const warnings = [];
  const infos = [];

  const allRows = [];
  for (const p of csvPaths) {
    const rows = readUpsCsv(p);
    for (const r of rows) allRows.push(r);
  }
  if (!allRows.length) throw new Error('Fichier(s) UPS vide(s) ou illisible(s).');

  // Mois cible = mois majoritaire (Date de facture, colonne 9) -- meme piege que les autres
  // carriers (Chronopost/TNT/FedEx) : quelques lignes residuelles d'un autre mois.
  const comptageMois = new Map();
  for (const r of allRows) {
    const d = String(r[COL.dateFacture] || '').trim();
    const m = /^(\d{4})-(\d{2})-\d{2}$/.exec(d);
    if (m) comptageMois.set(`${m[1]}${m[2]}`, (comptageMois.get(`${m[1]}${m[2]}`) || 0) + 1);
  }
  let moisCible = null;
  let dateValidite = '';
  // Mois choisi dans l'UI = source de verite si fourni (decision utilisateur 2026-08-20) --
  // l'auto-detection par majorite sur "Date de la facture" (colonne 9) n'est plus fiable
  // (remonte par le pole transport : "date validite non valide, decale le mois de juin au
  // lieu de juillet"). Le repli auto reste actif si le champ n'est pas envoye.
  if (opts && opts.period && opts.period.compact) {
    moisCible = opts.period.compact;
    const m = /^(\d{4})(\d{2})$/.exec(moisCible);
    if (m) dateValidite = `01/${m[2]}/${m[1]}`;
  } else if (comptageMois.size) {
    [moisCible] = [...comptageMois.entries()].sort((a, b) => b[1] - a[1])[0];
    const m = /^(\d{4})(\d{2})$/.exec(moisCible);
    if (m) dateValidite = `01/${m[2]}/${m[1]}`;
    if (comptageMois.size > 1) {
      const detail = [...comptageMois.entries()].sort((a, b) => b[1] - a[1]).map(([k, v]) => `${k}: ${v} ligne(s)`).join(', ');
      infos.push(`Plusieurs mois détectés dans les CSV reçus (${detail}) — mois retenu : ${moisCible} (majoritaire).`);
    }
  }

  // Colis 1Z79 (regle FACTURATION EXCEL.docx) : colis viticulteur retourne chez La Ruche --
  // reste dans "Facture UPS" (donnees brutes, decision utilisateur 2026-08-25 -- CONFIRME
  // contre le fichier reel fait-main de juillet 2026, qui garde bien ces lignes), EXCLU
  // uniquement du fichier import ERP, et agrege dans l'onglet "Demande avoir" (Tracking/
  // Nb colis/Montant/Cause -- Factures/Poids/Mode livraison laisses vides, jamais remplis
  // meme dans le fichier fait-main : saisie manuelle du pole transport). Detecte par prefixe
  // tracking "1Z79". BUG TROUVE 2026-08-25 (test sur 14 trackings reels de juillet 2026,
  // comparaison au fichier fait-main) : "codeClasse=FRT strict" etait FAUX -- Montant =
  // somme de TOUTES les lignes SAUF TVA (codeClasse=TAX) et Taxe gazole (codeClasse=FSC),
  // confirme exact sur 2/3 cas d'ecart reexamines (le 3e, 1Z79A7T06819992295, reste un ecart
  // residuel inexplique -- probable ajustement manuel isole du pole transport). Nb colis =
  // FIGE A 1 par tracking (valeur du fait-main sur 13/14 trackings testes, aucune formule de
  // somme sur les colonnes brutes ne colle -- decision utilisateur 2026-08-25).
  //
  // Lignes SANS Numero de suivi NI Numero de reference 1 (Ref1, colonne T) : demande
  // utilisateur 2026-08-25 -- supprimees de "Facture UPS" (jamais collees), pas seulement du
  // fichier import (contrairement au filtre "Montant=0" d'autres transporteurs qui ne touche
  // que l'import -- ici la ligne n'a AUCUNE cle d'identification exploitable, meme dans le
  // classeur brut). Ref1 vide, "." OU "null" en texte (valeurs non-renseignees frequentes cote
  // UPS -- "." confirme sur reel : 9 occurrences sur juin 2026 ; "null" texte signale par
  // l'utilisateur 2026-08-25) compte comme "pas de reference".
  const refVide = (v) => !v || /^\.+$/.test(v) || v.toLowerCase() === 'null';
  const CAUSE_1Z79 = "Nouveau compte pour le renvoi des colis depuis LR  l'avoir est censé être remis automatiquement sans avoir besoin de le demander";
  const demandesAvoir1Z79 = new Map(); // tracking -> montant
  let nbSansIdentification = 0;
  let nbMontantNetZero = 0;
  const lignesRetenues = [];
  for (const r of allRows) {
    const tracking = String(r[COL.numeroSuivi] || '').trim();
    const montantNetLigne = num(r[COL.montantNet]);
    if (/^1Z79/i.test(tracking)) {
      if (!demandesAvoir1Z79.has(tracking)) demandesAvoir1Z79.set(tracking, 0);
      const codeClasse = String(r[COL.codeClasse] || '').trim().toUpperCase();
      if (codeClasse !== 'TAX' && codeClasse !== 'FSC') {
        demandesAvoir1Z79.set(tracking, round2(demandesAvoir1Z79.get(tracking) + montantNetLigne));
      }
      // Lignes a Montant net = 0 : demande utilisateur 2026-08-25 -- supprimees de "Facture
      // UPS" MEME pour les 1Z79 (BUG TROUVE 2026-08-25 : je pensais a tort qu'elles etaient
      // gardees a 0EUR dans le fait-main -- confirme sur reel, tracking 1Z79A7T06819992295
      // et 8 autres : le fait-main a 0/49 lignes 1Z79 a montant=0, alors que le CSV brut en a
      // 9/58 -- toutes supprimees, ex. categorie "code inconnu" a 0EUR jamais dans le
      // fait-main). L'agregation "Demande avoir" ci-dessus reste faite AVANT ce filtre (elle
      // doit voir toutes les lignes, y compris montant=0, pour le calcul TAX/FSC exclus).
      if (montantNetLigne === 0) {
        nbMontantNetZero++;
        continue;
      }
      lignesRetenues.push(r);
      continue;
    }
    const ref1 = String(r[COL.ref1] || '').trim();
    if (!tracking && refVide(ref1)) {
      nbSansIdentification++;
      continue;
    }
    // Lignes a Montant net = 0 (hors 1Z79, deja traite ci-dessus) : demande utilisateur
    // 2026-08-25 -- supprimees de "Facture UPS" (pas seulement du fichier import), meme
    // principe que le filtre tracking/ref1 vide ci-dessus. CONFIRME sur reel (facture
    // 202600782885, juillet 2026) : 13513 lignes a montant=0 AVEC tracking rempli
    // expliquaient un ecart massif (+13513 lignes) entre le fichier genere et le fait-main --
    // distinct du filtre "sans identification" (9046 lignes SANS tracking, deja gere ci-dessus).
    if (montantNetLigne === 0) {
      nbMontantNetZero++;
      continue;
    }
    lignesRetenues.push(r);
  }
  if (nbMontantNetZero) {
    infos.push(`${nbMontantNetZero} ligne(s) à Montant net = 0 supprimée(s) de "Facture UPS".`);
  }
  if (demandesAvoir1Z79.size) {
    infos.push(`${demandesAvoir1Z79.size} colis en 1Z79 (retour viticulteur chez La Ruche) exclus de l'import — reportés dans l'onglet "Demande avoir" du classeur : ${[...demandesAvoir1Z79.keys()].join(', ')}.`);
  }
  if (nbSansIdentification) {
    infos.push(`${nbSansIdentification} ligne(s) sans Numéro de suivi ni Numéro de référence 1 supprimée(s) (aucune clé d'identification exploitable).`);
  }

  // Agregation par tracking : Categorie -> Montant net (reproduit le TCD "TCD"), + Max
  // Zone/Colis/Assurance/Poids (reproduit le TCD "zone colis poids assurance"). 1Z79 EXCLU ici
  // (reste dans lignesRetenues/"Facture UPS" mais jamais agrege -> jamais dans importRows,
  // cf. bloc ci-dessus qui les reporte deja separement dans demandesAvoir1Z79).
  const parTracking = new Map(); // tracking -> { postes: {cat: montant}, zone, nbColis, poids, assurance, compte, pays }
  for (const r of lignesRetenues) {
    const tracking = String(r[COL.numeroSuivi] || '').trim();
    if (!tracking || demandesAvoir1Z79.has(tracking)) continue;
    if (!parTracking.has(tracking)) {
      parTracking.set(tracking, { postes: {}, zone: 0, nbColis: 0, poids: 0, assurance: 0, compte: '', pays: '', mode: '', aLigneFret: false });
    }
    const acc = parTracking.get(tracking);

    const codeClasse = String(r[COL.codeClasse] || '').trim().toUpperCase();
    const codeDescription = String(r[COL.codeDescription] || '').trim();
    const description = String(r[COL.description] || '').trim();
    const montantNet = num(r[COL.montantNet]);
    const valeurBase = num(r[COL.valeurBase]);

    // Cascade EXACTE du modele (Facture UPS!D) : Adresse et plus-value BtoC (testees sur la
    // DESCRIPTION, colonne AX) sont prioritaires sur codeClasse=FRT/TAX -- BUG TROUVE
    // 2026-08-20 : je pensais 'ST SV'!Q:Q/D:D toujours vides, ce qui classait a tort des
    // lignes FRT "Correction d'adresse..." en Frêt au lieu d'Adresse (cf. config.json).
    let categorie;
    if (cfg.descriptions_adresse.includes(description)) categorie = 'Adresse';
    else if (cfg.descriptions_plus_value.includes(description)) categorie = 'plus-value BtoC';
    else if (codeClasse === 'FRT') categorie = 'Frêt';
    else if (codeClasse === 'TAX') categorie = 'TVA';
    else categorie = categoriePour(codeDescription);
    if (categorie) acc.postes[categorie] = round2((acc.postes[categorie] || 0) + montantNet);

    // Mode envoi : reproduit le TCD 'ST SV' + formule finale Fichier import!M =IF(COUNTIF(
    // 'ST SV'!H:H,tracking)=0,"inconnu",XLOOKUP(tracking,'ST SV'!H:H,'ST SV'!N:N)) -- N =
    // IF(compteurSV<>"","SV","ST"). "inconnu" UNIQUEMENT si le tracking n'a AUCUNE ligne
    // codeClasse="FRT" (absent du TCD) ; sinon "SV" si au moins une ligne FRT matche la table
    // service_vers_mode="SV", "ST" PAR DEFAUT (y compris si la description ne matche aucune
    // entree de la table -- BUG TROUVE 2026-08-20 : je renvoyais a tort "inconnu" pour CE cas,
    // alors que la vraie formule Excel renvoie "ST" des qu'il existe au moins une ligne FRT,
    // quelle que soit sa description).
    if (codeClasse === 'FRT') {
      acc.aLigneFret = true;
      if (cfg.service_vers_mode[description] === 'SV') acc.mode = 'SV';
    }

    if (codeDescription.toUpperCase() === 'EVS') acc.assurance = Math.max(acc.assurance, valeurBase);

    const zone = num(r[COL.zone]);
    if (zone > acc.zone) acc.zone = zone;
    // "Max de Nombre de colis" (reproduit le TCD 'zone colis poids assurance', formule
    // modele Fichier import!N = MAX(XLOOKUP(...,E:E),A) -- A='Nb Colis ERP', jamais rempli
    // ici, donc MAX se reduit au max natif du CSV brut).
    const nbColisLigne = num(r[COL.nombreColis]);
    if (nbColisLigne > acc.nbColis) acc.nbColis = nbColisLigne;
    const poidsLigne = num(r[COL.poidsFacture]);
    if (poidsLigne > acc.poids) acc.poids = poidsLigne;
    if (!acc.compte) acc.compte = normalizeCompte(tracking.slice(2, 8));
    if (!acc.pays) acc.pays = String(r[COL.pays] || '').trim().toUpperCase();
  }

  // Transporteur (UPS/UPS_COD) via compte extrait du tracking.
  for (const [tracking, acc] of parTracking) {
    acc.transporteur = cfg.comptes[acc.compte] || 'inconnu';
    if (acc.transporteur === 'inconnu') {
      warnings.push(`Tracking ${tracking} : compte "${acc.compte}" absent de la table Comptes UPS — transporteur non déterminé.`);
    }
  }

  // Export brut WMS (E/P) : m et m-1, meme mecanisme que Geodis/DPD/FedEx (core/exportBrut.js).
  const appRoot = path.resolve(__dirname, '../../..');
  const brutPaths = moisCible ? findBrutFiles(moisCible.length === 6 ? `${moisCible.slice(0, 4)}_${moisCible.slice(4)}` : moisCible, appRoot) : [];
  const epMap = brutPaths.length ? epParTrackingFromExport(readBrutRows(brutPaths)) : new Map();
  if (!brutPaths.length) warnings.push("Export WMS 'expéditions_brut' introuvable pour ce mois (E/P) — toutes les lignes sans correspondance seront classées 'P' par défaut (sauf plus-value BtoC détectée, qui force 'P').");

  const recs = [];
  let nEpDefaut = 0;
  for (const [tracking, acc] of parTracking) {
    const poidsArr = poidsArrondi(acc.transporteur, acc.nbColis, acc.poids);
    const fret = acc.postes['Frêt'] || 0;
    const droitsTaxes = acc.postes['Droits et taxes'] || 0;
    const colisVolumineux = colisVolumineuxMontant(acc.postes['Colis volumineux'] || 0);
    // Assurance : =MAX(10,ROUNDUP(2%*valeurDeclaree,2)) si valeur declaree>0.
    const assurance = acc.assurance > 0 ? Math.max(10, Math.ceil(0.02 * acc.assurance * 100) / 100) : 0;
    // Zones eloignees : forfait 40€ si le poste "Zones eloignees" REEL (categorie UPS, table
    // Charge.CHG_CODE) est present sur ce tracking -- formule modele exacte : IF(TCD!O="","",40)
    // -- TCD!O = colonne "Zones eloignees" (15e colField du TCD). BUG TROUVE 2026-08-20 : j'avais
    // a tort mappe TCD!O sur "Droits et taxes" (col reelle = TCD!I, 9e), ce qui declenchait le
    // forfait sur 8719/8719 lignes au lieu de 29/8719 dans le fichier reel de juin 2026.
    const zonesEloignees = acc.postes['Zones éloignées'] ? 40 : 0;
    const plusValueBtoC = acc.postes['plus-value BtoC'] || 0;
    // Adresses : forfait 11,50€ si le poste "Adresse" REEL est present sur ce tracking
    // (formule modele exacte : IF(TCD!F="","",11.5) -- TCD!F = colonne "Adresse", 6e colField).
    const adresses = acc.postes.Adresse ? 11.5 : 0;

    // E/P : declaration ERP prioritaire, sinon plus-value BtoC force "P" (formule modele
    // exacte, confirme video 2-1). Point relais deja gere par epParTrackingFromExport.
    let epDeclare = epMap.get(tracking);
    let ep;
    if (epDeclare === 'particulier') ep = 'P';
    else if (plusValueBtoC !== 0) ep = 'P';
    else if (epDeclare === 'entreprise') ep = 'E';
    else { ep = 'P'; nEpDefaut++; }

    // TVA : =IF(TCD!N="",0,0.2) -- teste la PRESENCE du poste "TVA" reel (categorie UPS,
    // codeClasse="TAX"), PAS une liste de pays codee en dur (BUG TROUVE 2026-08-20, corrige
    // avant tout usage en production).
    const tva = acc.postes.TVA ? cfg.champs_fixes.tva_taux : 0;

    const modeFinal = acc.aLigneFret ? (acc.mode === 'SV' ? 'SV' : 'ST') : 'inconnu';

    recs.push({
      tracking, compte: acc.compte, transporteur: acc.transporteur,
      zone: acc.zone, poids: poidsArr, nbColis: acc.nbColis || 1,
      pays: acc.pays, ep, tva, mode: modeFinal,
      fret, droitsTaxes, colisVolumineux, assurance, zonesEloignees, plusValueBtoC, adresses,
      categorieInconnue: !!acc.postes['CODE INCONNU'],
    });
  }
  if (nEpDefaut) infos.push(`${nEpDefaut} ligne(s) sans correspondance dans l'export WMS (E/P) ni plus-value BtoC détectée — classée(s) 'P' par défaut.`);

  // Trackings SANS AUCUNE charge facturable (tous postes ERP a 0 -- typiquement une ligne
  // "INF"/"Retours indelivrable" isolee, sans ligne FRT ni aucun autre poste) : EXCLUS de
  // l'import, confirme par comparaison avec le fichier reel livre (2026_06_Facture UPS.xlsx,
  // 8736/8736 lignes ont AU MOINS un poste non nul, 0 ligne totalement vide) -- reproduit la
  // meme regle transversale "jamais de fret a 0 sans autre poste" (cf. core/validate.js et
  // FACTURATION EXCEL.docx "jamais de fret a 0€").
  const recsAvecMontant = recs.filter((r) => (r.fret || r.droitsTaxes || r.assurance || r.zonesEloignees || r.colisVolumineux || r.plusValueBtoC || r.adresses));
  const nVides = recs.length - recsAvecMontant.length;
  if (nVides) infos.push(`${nVides} tracking(s) sans aucune charge facturable (ligne informative/retour indélivrable isolée, tous postes ERP à 0) exclus de l'import.`);

  const lignesCodeInconnu = recsAvecMontant.filter((r) => r.categorieInconnue);
  if (lignesCodeInconnu.length) {
    infos.push(`${lignesCodeInconnu.length} tracking(s) avec au moins une ligne "CODE INCONNU" (code de frais UPS non répertorié dans la table Charge.CHG_CODE) — catégorie normale du modèle réel, montant conservé mais non reclassé.`);
  }

  const importRows = recsAvecMontant.map((rec) => ({
    Transporteur: rec.transporteur,
    DateValidite: dateValidite || '',
    Ref1: '', Ref2: '', IdClient: '',
    Tracking: rec.tracking, Nom: '',
    // Zone : =IF(LEN(C)>2,C,IF(L="FR","France",XLOOKUP(...))) -- C=zone native UPS (meme
    // valeur que rec.zone). BUG TROUVE 2026-08-20 : Pays="FR" FORCE TOUJOURS "France" pour
    // une zone native courte (1-2 chiffres), MEME si cette zone est un nombre non-nul (ex.
    // zone=3 + Pays=FR -> "France", PAS "3") -- confirme majoritaire dans le fichier reel de
    // juin 2026 (7622/8719 lignes = "France"). Seules les zones a 3+ caracteres (codes 3
    // chiffres internationaux, ex. "706") passent telles quelles, meme si Pays=FR (rare/jamais
    // en pratique).
    EP: rec.ep, Pays: rec.pays,
    Zone: String(rec.zone || '').length > 2 ? String(rec.zone) : (rec.pays === 'FR' ? 'France' : (rec.zone ? String(rec.zone) : 'inconnu')),
    NbrColis: rec.nbColis, Poids: roundUp1(rec.poids),
    Mode: rec.mode, TVA: rec.tva,
    DroitsTaxes: rec.droitsTaxes || 0, Assurance: rec.assurance || 0,
    ZonesEloignees: rec.zonesEloignees || 0, ColisVolumineux: rec.colisVolumineux || 0,
    Adresses: rec.adresses || 0, Fret: rec.fret || 0, PlusValueB2C: rec.plusValueBtoC || 0, TaxeGasoil: '', NbColis: '',
  }));

  const { alerts, infos: valInfos } = validate(importRows);
  infos.push(...valInfos);

  const controle = {
    'Fret': round2(importRows.reduce((s, r) => s + (r.Fret || 0), 0)),
    'Droits et taxes': round2(importRows.reduce((s, r) => s + (r.DroitsTaxes || 0), 0)),
    'Assurance': round2(importRows.reduce((s, r) => s + (r.Assurance || 0), 0)),
    'Zones éloignées': round2(importRows.reduce((s, r) => s + (r.ZonesEloignees || 0), 0)),
    'Colis volumineux': round2(importRows.reduce((s, r) => s + (r.ColisVolumineux || 0), 0)),
    'plus-value BtoC': round2(importRows.reduce((s, r) => s + (r.PlusValueB2C || 0), 0)),
  };

  // Demande avoir (1Z79) : Tracking/Nb colis/Montant -- Factures/Poids/Mode livraison
  // laisses vides (saisie manuelle du pole transport, jamais remplis meme dans le fichier
  // fait-main). Reproduit cote finaliser_ups.py pour peupler l'onglet "Demande avoir" du
  // classeur (le Node n'ecrit pas ce classeur, expose ici pour tracabilite/coherence).
  const demandesAvoir = [...demandesAvoir1Z79.entries()].map(([tracking, montant]) => ({
    Tracking: tracking, Factures: '', NbColis: 1, Poids: '', Mode: '', Montant: montant, Cause: CAUSE_1Z79,
  }));

  return {
    header: [],
    rows: recs, recs, importRows, controle, warnings, alerts, infos,
    posteKeys: ['Fret', 'Droits et taxes', 'Assurance', 'Zones éloignées', 'Colis volumineux', 'plus-value BtoC'], cfg,
    sheetNames: { raw: 'Facture UPS', import: 'Fichier import' },
    period: moisCible ? `${moisCible.slice(0, 4)}_${moisCible.slice(4)}` : 'export',
    demandesAvoir,
  };
}

/** Args du finaliseur (csv + brut du mois/mois-1) -- meme pattern que FedEx/Delivengo. */
function computeFinalizerArgs(files, period, appRoot) {
  const brut = period ? findBrutFiles(period, appRoot) : [];
  return ['--csv', ...(files.csv || []), '--brut', ...brut];
}

module.exports = {
  id: 'ups',
  name: 'UPS',
  status: 'ready',
  viticolis: true,
  taxeGasoil: 'Fille Viticolis',
  method: "Fichiers CSV Billing UPS (sans en-tete, colonnes resolues par position). Transporteur UPS/UPS_COD via compte extrait du tracking. Categorie ERP via cascade code classe (FRT/TAX) puis table Charge.CHG_CODE (472 codes). Zone FOURNIE par UPS (pas calculee). Poids/Colis volumineux/Assurance via formules modele exactes (baremes/plafonds). E/P via export WMS m/m-1, plus-value BtoC force 'P'. Colis 1Z79 exclus (demande d'avoir).",
  inputs: [
    { key: 'csv', label: 'Factures UPS Billing (CSV, 1 par facture)', accept: '.csv', multiple: true, required: true },
  ],
  outputNaming: { workbook: '{period}_Facture UPS', import: '{period}_UPS_Import' },
  finalizer: {
    script: '../automatisation/finaliser_ups.py',
    template: '../Transporteurs/UPS/2026_06_Facture UPS.xlsx',
    buildArgs: (files, period, appRoot) => computeFinalizerArgs(files, period, appRoot),
  },
  process,
};
