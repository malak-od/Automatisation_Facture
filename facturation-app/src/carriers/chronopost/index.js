// ============================================================================
//  Adaptateur transporteur : CHRONOPOST
//  Le transporteur le plus riche du projet : 2 fichiers Excel bruts par mois
//  (1 par sous-compte/contrat -- standard "51291303" et 2SHOP "65481903"),
//  chacun avec plusieurs No Facture, chacun rapproche d'un PDF distinct.
//
//  Lignes "normales" (1 ligne = 1 colis, Numero LT = vrai tracking) VS lignes
//  "forfaitaires" (Numero LT = code court CAP*/ECO*/SUR*, Montant HT = un
//  forfait mensuel pour TOUTE la facture, pas un colis) : reperees par
//  PREFIXE de Numero LT, PAS par le libelle Type prestation (qui varie
//  legerement : "Surcharge Carburant Aerien" vs "Routier"). Confirme :
//  CAP*=Surcharge Carburant (I=Aerien/N=National-Routier/O=Outre-mer),
//  ECO*=Participation Eco-Responsable, SUR*=Surete colis. Le nombre de
//  triplets I/N/O PAR FACTURE varie (le contrat 2SHOP n'a jamais que N).
//
//  Le POOL "gazole" (CAP*) de chaque facture est REDISTRIBUE au prorata du
//  fret sur toutes les lignes normales de cette meme facture (mecanisme
//  confirme dans le classeur modele, colonne AB 'Facture Chronopost' :
//  =$AG$5/$AG$2*Z -- pool gazole / total fret * fret de la ligne). ECO/SUR
//  restent des POSTES CALCULES mais sont classes "Frêt" au final (feuille
//  Categories du modele, PAS un poste distinct dans l'import ERP) -- les
//  lignes forfaitaires elles-memes ne sont JAMAIS incluses dans l'import.
//
//  Mapping mode envoi/zone ERP : table "Bibliotheque transporteurs" du
//  classeur modele (cf. config.json), avec 2 regles speciales confirmees
//  par formule (TCD!U3) :
//   - produit='6B' -> zone = lookup Pays ARRIVEE dans 'Zoning 2shop'
//   - produit='6C' -> zone = lookup Pays DEPART dans 'Zoning 2shop' (table C/D)
//   - produit=17 ou 44 -> zone = "<produit>_<Zone Tarifaire>" (ex. "44_C1")
//   - sinon -> lookup direct table Bibliotheque
// ============================================================================
const XLSX = require('xlsx');
const pdfParse = require('pdf-parse');
const path = require('path');
const { num, round2, roundUp1 } = require('../../core/csv');
const { validate } = require('../../core/validate');
const cfg = require('./config.json');

const POSTE_KEYS = ['Adresse', 'Assurance', 'Colis volumineux', 'Corse', 'Droits et taxes', 'Frais facturation', 'Frêt', 'Gazole', 'Zones éloignées'];

function normKey(s) {
  return String(s || '')
    .normalize('NFD').replace(/[̀-ͯ]/g, '')
    .toLowerCase().replace(/[.,]/g, '').replace(/\s+/g, ' ').trim();
}

function colIndexByName(header, name) {
  const target = normKey(name);
  return header.findIndex((h) => normKey(h) === target);
}

function isXlsx(p) {
  const fs = require('fs');
  try {
    const buf = fs.readFileSync(p, { encoding: null, flag: 'r' });
    return buf.length > 4 && buf[0] === 0x50 && buf[1] === 0x4b; // "PK"
  } catch (e) {
    return false;
  }
}

/** Lit une facture brute Chronopost : feuille "Donnees" (ou 1re feuille), en-tete en
 * ligne 4 (index 3), donnees a partir de la ligne 5. Meme structure confirmee sur les
 * 2 sous-comptes (standard/2shop). */
function readFactureBrute(p) {
  const wb = XLSX.readFile(p, { cellDates: false });
  const ws = wb.Sheets[wb.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: '' });
  const headerRowIdx = 3; // ligne 4 (0-indexee)
  if (!rows[headerRowIdx]) throw new Error(`Fichier ${path.basename(p)} : ligne d'en-tete (4) introuvable.`);
  const header = rows[headerRowIdx].map((h) => String(h || '').trim());
  const data = rows.slice(headerRowIdx + 1).filter((r) => r.some((v) => v !== ''));
  return { file: path.basename(p), header, rows: data };
}

function excelSerialToDate(v) {
  if (typeof v === 'number') return new Date(Date.UTC(1899, 11, 30) + v * 86400000);
  const m = /^(\d{2})\/(\d{2})\/(\d{4})$/.exec(String(v || '').trim());
  if (m) return new Date(Date.UTC(+m[3], +m[2] - 1, +m[1]));
  return null;
}

/** PDF Chronopost : bloc "Compte<n>Facture<n>Date<date>" en tete/pied de page, puis
 * "TOTAL FACTURE<HT><TVA><TTC>" colle (3 montants) en fin de document -- format texte
 * structure, confirme sur les 7 PDF de juin 2026 (pas de balise <ref> comme Mondial
 * Relay/DPD, mais des libelles fixes fiables). */
async function extractChronoPdfInfo(pdfPath) {
  const buf = require('fs').readFileSync(pdfPath);
  const { text } = await pdfParse(buf);
  const mCompte = /Compte\s*(\d+)\s*Facture\s*(\d+)/.exec(text);
  const mTotal = /TOTAL\s*FACTURE\s*([\d\s]+[.,]\d{2})\s*([\d\s]+[.,]\d{2})\s*([\d\s]+[.,]\d{2})/.exec(text);
  if (!mCompte || !mTotal) return null;
  return {
    file: path.basename(pdfPath),
    compte: mCompte[1],
    facture: mCompte[2],
    totalHt: num(mTotal[1].replace(/\s/g, '')),
  };
}

function isForfaitaire(numeroLt) {
  const s = String(numeroLt || '').trim().toUpperCase();
  return /^(CAP|ECO|SUR)/.test(s);
}

function typeForfaitaire(numeroLt) {
  const s = String(numeroLt || '').trim().toUpperCase();
  if (s.startsWith('CAP')) return 'CAP';
  if (s.startsWith('ECO')) return 'ECO';
  if (s.startsWith('SUR')) return 'SUR';
  return null;
}

/** Categorie ERP pour un Type prestation donne, via la table 'categories' du config.json
 * (LOOKUP exact-match, comme la feuille 'Categories' du classeur modele). */
function categoriePour(typePrestation) {
  return cfg.categories[String(typePrestation || '').trim()] || null;
}

/** Mode envoi / zone / Transporteur ERP pour un code produit Chrono donne, via la table
 * 'bibliotheque' + les 2 regles speciales confirmees par formule (TCD!U3) : produit
 * 6B/6C -> zone via 'Zoning 2shop' (pays arrivee/depart selon le cas) ; produit 17/44
 * -> zone = "<produit>_<zoneTarifaire>". */
function mappingErp(produit, zoneTarifaire, paysArrivee, paysDepart) {
  const p = String(produit || '').trim();
  const entry = cfg.bibliotheque[p];
  if (!entry) return { modeEnvoi: 'inconnu', zone: 'inconnu', transporteur: 'inconnu' };

  let zone = entry.zone;
  if (p === '6B') {
    zone = cfg.zoning_2shop[String(paysArrivee || '').trim().toUpperCase()];
    zone = zone != null ? String(zone) : 'inconnu';
  } else if (p === '6C') {
    zone = cfg.zoning_2shop_6c[String(paysDepart || '').trim().toUpperCase()];
    zone = zone != null ? String(zone) : 'inconnu';
  } else if (p === '17' || p === '44') {
    zone = `${p}_${String(zoneTarifaire || '').trim()}`;
  }
  return { modeEnvoi: entry.modeEnvoi, zone, transporteur: entry.transporteur };
}

async function process(files) {
  const facturePaths = (files.facture || []).filter(isXlsx);
  if (!facturePaths.length) throw new Error('Aucun fichier fourni (attendu : facture(s) Chronopost reçue(s), xlsx, 1 par sous-compte).');

  const warnings = [];
  const infos = [];
  const brutes = facturePaths.map(readFactureBrute);

  // Date validite = mois MAJORITAIRE sur l'ensemble des lignes (PAS la 1re ligne trouvee) --
  // BUG TROUVE 2026-08-13 : le brut Chronopost recu pour un mois donne contient parfois
  // quelques lignes residuelles du mois precedent EN TETE de fichier (livraisons tardives
  // facturees avec la coupure du mois suivant, confirme sur facture_chronopost_51291303_
  // 202607.xlsx : 3 lignes datees du 30/06/2026 en lignes 5-6 du brut, sur 1717 lignes
  // totales dont 1714 de juillet). Prendre la 1re ligne donnait "juin" pour un fichier
  // envoye et rempli a 99,8% de juillet -> nom de sortie errone (2026_06_* au lieu de
  // 2026_07_*). Compte les occurrences par (annee, mois) sur TOUTES les lignes, retient le
  // plus frequent.
  const comptageMois = new Map(); // "AAAA-MM" -> nb lignes
  const lignesNormales = []; // { facture, sousCompte, numeroLt, typePrestation, tva, zoneTarifaire, poids, produit, montantHt, paysArrivee, paysDepart }
  const lignesForfaitaires = []; // { facture, sousCompte, type: CAP/ECO/SUR, montantHt }

  for (const f of brutes) {
    const iFacture = colIndexByName(f.header, 'No Facture');
    const iSousCompte = colIndexByName(f.header, 'Sous-compte');
    const iDate = colIndexByName(f.header, 'Date LT');
    const iPaysDepart = colIndexByName(f.header, 'Pays depart');
    const iPaysArrivee = colIndexByName(f.header, 'Pays arrivee');
    const iNumeroLt = colIndexByName(f.header, 'Numero LT');
    const iTypePrestation = colIndexByName(f.header, 'Type prestation');
    const iTva = colIndexByName(f.header, 'TVA');
    const iZoneTarifaire = colIndexByName(f.header, 'Zone Tarifaire');
    const iPoids = colIndexByName(f.header, 'Poids');
    const iProduit = colIndexByName(f.header, 'Produit');
    const iMontantHt = colIndexByName(f.header, 'Montant HT');
    if (iFacture < 0 || iNumeroLt < 0 || iMontantHt < 0) {
      warnings.push(`${f.file} : colonne(s) attendue(s) introuvable(s) (No Facture/Numero LT/Montant HT) — fichier ignoré.`);
      continue;
    }

    for (const r of f.rows) {
      const facture = String(r[iFacture] || '').trim();
      const sousCompte = String(r[iSousCompte] || '').trim();
      const numeroLt = String(r[iNumeroLt] || '').trim();
      const montantHt = num(r[iMontantHt]);
      if (!facture || !numeroLt) continue;

      const d = excelSerialToDate(r[iDate]);
      if (d) {
        const cle = `${d.getUTCFullYear()}-${String(d.getUTCMonth() + 1).padStart(2, '0')}`;
        comptageMois.set(cle, (comptageMois.get(cle) || 0) + 1);
      }

      if (isForfaitaire(numeroLt)) {
        lignesForfaitaires.push({ facture, sousCompte, type: typeForfaitaire(numeroLt), montantHt });
        continue;
      }

      lignesNormales.push({
        facture, sousCompte, numeroLt,
        typePrestation: iTypePrestation >= 0 ? String(r[iTypePrestation] || '').trim() : '',
        tva: iTva >= 0 ? String(r[iTva] || '').trim() : '',
        zoneTarifaire: iZoneTarifaire >= 0 ? String(r[iZoneTarifaire] || '').trim() : '',
        poids: iPoids >= 0 ? num(r[iPoids]) : 0,
        produit: iProduit >= 0 ? String(r[iProduit] || '').trim() : '',
        montantHt,
        paysDepart: iPaysDepart >= 0 ? String(r[iPaysDepart] || '').trim() : '',
        paysArrivee: iPaysArrivee >= 0 ? String(r[iPaysArrivee] || '').trim() : '',
      });
    }
  }
  if (!lignesNormales.length) warnings.push('Aucune ligne "colis" trouvée dans les facture(s) fournies.');

  let dateValidite = null;
  if (comptageMois.size) {
    const [cleMajoritaire] = [...comptageMois.entries()].sort((a, b) => b[1] - a[1])[0];
    const [annee, mois] = cleMajoritaire.split('-');
    dateValidite = `01/${mois}/${annee}`;
    if (comptageMois.size > 1) {
      const detail = [...comptageMois.entries()].sort((a, b) => b[1] - a[1]).map(([k, v]) => `${k}: ${v} ligne(s)`).join(', ');
      infos.push(`Plusieurs mois détectés dans les fichiers reçus (${detail}) — mois retenu : ${mois}/${annee} (majoritaire).`);
    }
  }

  // Pool gazole (CAP*) par facture, redistribue au prorata du Montant HT ("fret brut" avant
  // reclassement) de chaque ligne normale de CETTE MEME facture -- confirme formule
  // 'Facture Chronopost'!AB : =poolGazoleFacture/totalFretFacture*fretLigne.
  const poolGazoleParFacture = new Map();
  const totalFretParFacture = new Map();
  for (const l of lignesForfaitaires) {
    if (l.type !== 'CAP') continue;
    poolGazoleParFacture.set(l.facture, round2((poolGazoleParFacture.get(l.facture) || 0) + l.montantHt));
  }
  for (const l of lignesNormales) {
    totalFretParFacture.set(l.facture, (totalFretParFacture.get(l.facture) || 0) + l.montantHt);
  }

  // Reclassement de chaque ligne normale en postes ERP (via Categories), + gazole reparti.
  const recs = [];
  for (const l of lignesNormales) {
    const categorie = categoriePour(l.typePrestation);
    if (!categorie) warnings.push(`Ligne ${l.numeroLt} (facture ${l.facture}) : Type prestation "${l.typePrestation}" absent de la table "Catégories" — non reclassée.`);

    const map = mappingErp(l.produit, l.zoneTarifaire, l.paysArrivee, l.paysDepart);
    if (map.modeEnvoi === 'inconnu') warnings.push(`Ligne ${l.numeroLt} (facture ${l.facture}) : produit Chrono "${l.produit}" absent de "Bibliothèque transporteurs" — mode envoi/zone/Transporteur "inconnu".`);

    // Gazole reparti au prorata du fret de la ligne (formule modele 'Facture Chronopost'!AB
    // = $AG$5/$AG$2*Z). L'ARRONDI PAR LIGNE (necessaire : le fichier d'import a besoin d'une
    // valeur par ligne) cree une derive cumulative INEVITABLE sur le total (jusqu'a ~1,6EUR
    // par facture constate en juin 2026 -- meme phenomene que dans le classeur Excel reel,
    // dont chaque cellule est aussi affichee arrondie a 2 decimales) -- PAS un bug : la
    // reconciliation PDF fait foi (elle compare le total BRUT avant reclassement, insensible
    // a cette derive), pas le sous-total Gazole seul.
    const totalFret = totalFretParFacture.get(l.facture) || 0;
    const poolGazole = poolGazoleParFacture.get(l.facture) || 0;
    const gazoleLigne = totalFret > 0 ? (poolGazole / totalFret) * l.montantHt : 0;

    // "sureté + eco" (colonne AA du modele) : forfait FIXE ajoute a TOUTE ligne normale
    // (pas seulement Sûreté colis/Participation Eco-Responsable malgre le nom trompeur de
    // la colonne) -- 0,08€ pour les modes 2SHOP (6B/6C/5X/5Y), 0,5€ sinon. Formule modele
    // EXACTE : AA=IF(mode∈{6C,6B,5X,5Y},0.08,0.5), AC(hors gazole)=Z+AA -- donc ce forfait
    // s'ajoute au Frêt de CHAQUE ligne. BUG TROUVE 2026-08-12 (5e bug, apres Corse/forfait
    // ZE-Adresses/Frais facturation/E-P) : ce forfait n'etait PAS applique du tout, cause
    // principale de l'ecart residuel de ~1040€ sur juillet 2026 (confirme empiriquement :
    // l'ecart par tracking = EXACTEMENT 0,5€ x nombre de lignes du tracking, ex. un
    // tracking avec 2 lignes 'Type prestation' manque exactement 1,00€ = 2x0,5€).
    const modeSureteEcoBas = ['6B', '6C', '5X', '5Y'].includes(map.modeEnvoi);
    const suretéEcoLigne = modeSureteEcoBas ? 0.08 : 0.5;

    const postes = Object.fromEntries(POSTE_KEYS.map((k) => [k, 0]));
    if (categorie && POSTE_KEYS.includes(categorie)) postes[categorie] = round2((postes[categorie] || 0) + l.montantHt);
    // "Frais facturation" (Frais de gestion/Surcharge Facture Papier) n'a PAS de colonne ERP
    // dediee -- meme trou que "Corse" (cf. plus bas), fusionne dans "Frêt" par precaution
    // (comme Surete/Eco-Responsable le sont deja, cf. table Categories) -- aucune ligne
    // reelle rencontree a ce jour (juin/juillet 2026), donc non confirme empiriquement,
    // mais evite une perte silencieuse si ce cas se presente (decision utilisateur 2026-08-12).
    if (postes['Frais facturation'] > 0) {
      postes['Frêt'] = round2(postes['Frêt'] + postes['Frais facturation']);
      postes['Frais facturation'] = 0;
    }
    postes['Frêt'] = round2(postes['Frêt'] + suretéEcoLigne);
    postes.Gazole = round2((postes.Gazole || 0) + gazoleLigne);
    // Produits Chrono connus mais HORS CONTRAT (erreur de facturation cote Chronopost, pas
    // un cas metier a corriger automatiquement ici -- cf. email pole transport 2026-08-12,
    // 6 trackings juillet 2026 factures en '3Z' malgre resiliation de ce produit) : juste
    // signale pour reclamation, le mapping reste tel quel (pas de tarif corrige devine).
    if (cfg.produits_hors_contrat && cfg.produits_hors_contrat.includes(String(l.produit || '').trim())) {
      warnings.push(`Ligne ${l.numeroLt} (facture ${l.facture}) : produit Chrono "${l.produit}" HORS CONTRAT (résilié) — probable erreur de facturation Chronopost à réclamer, tarif potentiellement plus élevé qu'attendu.`);
    }

    const client = cfg.sous_comptes[l.sousCompte];
    if (!client && l.sousCompte) warnings.push(`Ligne ${l.numeroLt} (facture ${l.facture}) : sous-compte "${l.sousCompte}" absent de "Sous-comptes" — Id client vide.`);

    // "Pays arrivee" brut VIDE = envoi NATIONAL (Chronopost ne renseigne ce champ que pour
    // l'international, confirme sur les 2 fichiers bruts de juillet 2026 : codes postaux
    // francais avec Pays arrivee vide sur ~2440 lignes/4148) -- PAS une donnee manquante,
    // meme principe que la table pays_tva du modele ("Pays vide ou non liste = 20% par
    // defaut" -- pas de cle "FR" dans le modele). Defaut "FR" applique ICI (pour
    // Pays/TVA export), PAS sur l.paysArrivee brut (reste tel quel pour le lookup
    // zoning_2shop 6B/6C, qui n'a de toute facon aucune cle "FR" -- produit 6B/6C toujours
    // international, jamais affecte par ce defaut).
    const paysArriveeExport = l.paysArrivee || 'FR';
    recs.push({
      facture: l.facture, sousCompte: l.sousCompte, tracking: l.numeroLt,
      client: client || '', pays: paysArriveeExport, zone: map.zone,
      poids: l.poids, mode: map.modeEnvoi, transporteur: map.transporteur,
      tva: cfg.pays_tva[paysArriveeExport.toUpperCase()] ?? 0.2,
      postes, montantHt: l.montantHt, gazoleLigne,
      totalHorsGazole: round2(l.montantHt),
      totalAvecGazole: round2(l.montantHt + gazoleLigne),
    });
  }

  // "Zones eloignees" et "Adresses" ne sont PAS exportes au montant BRUT : le classeur
  // fait a la main applique un FORFAIT BINAIRE (formule 'Fichier import' du modele,
  // confirmee ET verifiee empiriquement sur 8 lignes reelles de juillet 2026 face au CSV
  // fait a la main) :
  //  - "Zones eloignees" export = SI((Corse + Zones eloignees brutes) > 15EUR, 29, SI(>0, 9.5, "")
  //    -- les postes "Corse" (Supplement Corse 18h) et "Zones eloignees" (Supp Zone
  //    Internationale Eloignee / Zones Difficiles d'acces / etc.) sont FUSIONNES en un
  //    SEUL forfait de sortie -- il n'y a PAS de colonne "Corse" separee dans le schema
  //    ERP standard (23 colonnes), donc "Corse" etait auparavant PERDU (jamais ecrit nulle
  //    part) avant cette correction (2026-08-12, ecart de ~960EUR trouve en comparant a la
  //    reference reelle de juillet 2026).
  //  - "Adresses" export = SI(Adresse brute > 8.5EUR, 17, SI(>0, 8.5, "") -- formule modele
  //    deja capturee (TCD!I>8.5 -> 17, sinon 8.5), seul le forfait bas (8.5) a pu etre
  //    confirme empiriquement sur les donnees reelles de juillet (aucun cas >8.5EUR
  //    rencontre), le forfait haut (17) reste donc base sur la formule modele seule.
  function forfaitZonesEloignees(montantBrut) {
    if (montantBrut <= 0) return 0;
    return montantBrut > 15 ? 29 : 9.5;
  }
  function forfaitAdresses(montantBrut) {
    if (montantBrut <= 0) return 0;
    return montantBrut > 8.5 ? 17 : 8.5;
  }

  // E/P : formule corrigee par le pole transport (2026-08-12, apres 64 avaries en juillet
  // 2026) -- SEUL le mode envoi "1S" a une grille tarifaire "Entreprise" ("E"), TOUS les
  // autres modes sont "Particulier" ("P"). Remplace le "E" fixe utilise precedemment (faux
  // pour tout mode != 1S).
  function epPourMode(mode) {
    return String(mode || '').trim() === '1S' ? 'E' : 'P';
  }

  // Lignes dont le produit Chrono n'est pas reconnu dans "Bibliotheque transporteurs"
  // (ex. produit "0" vu en juillet 2026, EXP20260727-2876730, 11 trackings) : Transporteur/
  // mode envoi/Zone tous "inconnu" -- inexploitables pour l'ERP. Confirme empiriquement
  // (2026-08-12) en comparant a la reference faite a la main de juillet 2026 : elle contient
  // exactement 3917 lignes, notre sortie en avait 3928 -- les 11 lignes en trop etaient EXACTEMENT
  // ces trackings "produit 0". Comme les lignes forfaitaires CAP/ECO/SUR, elles restent dans
  // "recs"/le classeur (reconciliation PDF sur le total BRUT, investigation) mais disparaissent
  // de l'import -- meme traitement, pas de tarif invente.
  //
  // REGROUPEMENT PAR TRACKING (2026-08-13) : un meme "Numero LT" a souvent plusieurs lignes
  // dans le brut (1 "Transport" avec le poids reel + 1..N lignes "supplement" -- ex. "Zones
  // Difficiles d'acces", poids=0 car ce n'est pas un colis separe). Le fichier fait a la main
  // n'a QU'UNE SEULE ligne par tracking (confirme : 3917 tracking uniques = 3917 lignes du
  // CSV reel de juillet 2026, poids/suppléments fusionnes sur cette unique ligne) -- avant ce
  // fix, chaque ligne du brut devenait une ligne d'import separee, produisant (a) des fausses
  // alertes "POIDS = 0" sur les lignes supplement (187 en juillet 2026) et (b) un risque de
  // double comptage NbrColis/Poids cote ERP (2 lignes = 2 colis au lieu d'1 pour le meme
  // tracking). Poids/Zone/Mode/Transporteur/Pays/TVA sont IDENTIQUES sur toutes les lignes
  // d'un meme tracking (verifie empiriquement sur juillet 2026, aucune exception) -- pris sur
  // la premiere ligne du groupe ; les postes (Frêt/Zones eloignees/etc.) et le gazole reparti
  // sont CUMULES sur le groupe.
  const recsValides = recs.filter((rec) => rec.mode !== 'inconnu');
  const groupesParTracking = new Map();
  for (const rec of recsValides) {
    if (!groupesParTracking.has(rec.tracking)) groupesParTracking.set(rec.tracking, []);
    groupesParTracking.get(rec.tracking).push(rec);
  }
  const importRows = [...groupesParTracking.values()].map((groupe) => {
    const base = groupe.find((r) => r.poids > 0) || groupe[0];
    // Poids = SOMME de toutes les lignes du groupe (pas juste celui de "base") -- reproduit
    // exactement le mecanisme du TCD Excel reel ("TCD poids", dataField=Somme de Poids,
    // verifie dans la video process 2026-08-13) : le TCD ne suppose JAMAIS que les lignes
    // supplement ont poids=0, il somme mecaniquement. Notre "base.poids" seul donnait le
    // meme resultat SEULEMENT parce que les lignes supplement observees ont toujours eu
    // poids=0 (verifie sur 3917 trackings de juillet 2026, aucune exception) -- mais c'etait
    // une supposition fragile, pas une garantie structurelle. Somme desormais pour etre aussi
    // robuste que le TCD si un mois futur a un supplement avec un poids non-nul.
    // PAS de round2 ici (contrairement aux postes en €) : roundUp1 a besoin de la precision
    // brute (poids Chronopost a 3 decimales) pour arrondir correctement au dixieme superieur
    // -- round2(4.403)=4.40 puis roundUp1(4.40)=4.4 est FAUX vs roundUp1(4.403)=4.5 direct.
    const poidsSomme = groupe.reduce((s, r) => s + (r.poids || 0), 0);
    const postes = Object.fromEntries(POSTE_KEYS.map((k) => [k, round2(groupe.reduce((s, r) => s + (r.postes[k] || 0), 0))]));
    // Le forfait "sureté + eco" (0,08€ 2SHOP / 0,5€ sinon) est ajoute AU FRET DE CHAQUE LIGNE
    // BRUTE dans "recs"/le classeur (fidele au modele Excel, colonne AA -- 1 ligne = 1 forfait,
    // confirme par formule). Mais au niveau TRACKING (import groupe), ce forfait ne doit
    // s'appliquer qu'UNE SEULE FOIS par tracking, pas une fois par ligne fusionnee -- sinon un
    // tracking avec N lignes (Transport + suppléments) se voit facturer N x le forfait au lieu
    // de 1x. BUG TROUVE 2026-08-13 (7e bug, capture ecran utilisateur "187 alertes POIDS=0" ->
    // regroupement par tracking -> ecart residuel de -0,50€ exact sur 168 trackings a 2 lignes,
    // confirme = exactement 1 forfait de trop). Retire les (N-1) forfaits en trop du Frêt cumule.
    const modeSureteEcoBas = ['6B', '6C', '5X', '5Y'].includes(base.mode);
    const suretéEcoForfait = modeSureteEcoBas ? 0.08 : 0.5;
    postes['Frêt'] = round2(postes['Frêt'] - suretéEcoForfait * (groupe.length - 1));
    return {
      Transporteur: base.transporteur,
      DateValidite: dateValidite || '',
      // Id client : TOUJOURS vide dans le fichier import (decision utilisateur 2026-08-12,
      // s'applique a tous les transporteurs) -- rec.client reste utilise pour d'autres
      // usages internes (ex. warnings sous-compte inconnu) mais n'est plus reporte ici.
      Ref1: '', Ref2: '', IdClient: '',
      Tracking: base.tracking, Nom: '',
      EP: epPourMode(base.mode), Pays: base.pays, Zone: base.zone,
      // Poids arrondi au dixieme de kg SUPERIEUR (meme convention que DPD/Geodis/GLS/Kuehne/
      // Mondial Relay, cf. roundUp1) -- confirme empiriquement (2026-08-13) contre le CSV fait
      // a la main de juillet 2026 : match exact sur 3917/3917 trackings (poids brut Chronopost
      // a 3 decimales, ex. 4,638kg -> 4,7kg exporte).
      NbrColis: 1, Poids: roundUp1(poidsSomme),
      Mode: base.mode, TVA: base.tva,
      DroitsTaxes: postes['Droits et taxes'], Assurance: postes.Assurance,
      ZonesEloignees: forfaitZonesEloignees(postes['Zones éloignées'] + postes.Corse),
      ColisVolumineux: postes['Colis volumineux'],
      Adresses: forfaitAdresses(postes.Adresse),
      // Le gazole (pool CAP* reparti au prorata) N'EST PAS repercute au client dans le fichier
      // livre -- confirme empiriquement (2026-08-12) : la colonne "gazole" du CSV reel de
      // juillet 2026 est VIDE sur les 3917 lignes, et "Frêt" reel = notre Frêt SANS le gazole
      // (matches exacts). Le gazole facture par Chronopost (visible sur le PDF/"Contrôle pdf")
      // est donc un cout absorbe en interne, pas transmis au client -- la reconciliation PDF
      // (plus bas) travaille directement sur lignesNormales/lignesForfaitaires brutes, pas
      // sur ce regroupement par tracking.
      Fret: postes['Frêt'], PlusValueB2C: 0, TaxeGasoil: '', NbColis: '',
    };
  });

  const { alerts, infos: valInfos } = validate(importRows, { skipPoidsDecimal: true });
  infos.push(...valInfos);

  // Reconciliation PDF : Total HT officiel (bloc "TOTAL FACTURE" du PDF) compare a la somme
  // des lignes normales + lignes forfaitaires de cette meme facture (Montant HT brut, avant
  // reclassement -- correspond exactement au total imprime sur la facture source).
  const totalBrutParFacture = new Map();
  for (const l of [...lignesNormales, ...lignesForfaitaires]) {
    totalBrutParFacture.set(l.facture, round2((totalBrutParFacture.get(l.facture) || 0) + l.montantHt));
  }
  const pdfPaths = files.pdf || [];
  const pdfs = [];
  for (const p of pdfPaths) {
    try {
      const r = await extractChronoPdfInfo(p);
      if (r) pdfs.push(r);
      else warnings.push(`PDF ${path.basename(p)} : total/numéro de facture introuvable, ignoré pour la réconciliation.`);
    } catch (e) {
      warnings.push(`PDF ${path.basename(p)} : lecture impossible (${e.message}).`);
    }
  }
  for (const p of pdfs) {
    const calcule = totalBrutParFacture.get(p.facture);
    if (calcule == null) { warnings.push(`Facture PDF ${p.file} (${p.facture}) : aucune ligne correspondante trouvée dans les fichiers traités.`); continue; }
    const ecart = round2(p.totalHt - calcule);
    const statut = Math.abs(ecart) >= 0.05 ? 'A VERIFIER' : 'OK';
    infos.push(`Facture PDF ${p.file} (${p.facture}) : Total HT = ${p.totalHt.toFixed(2)} EUR, calculé = ${calcule.toFixed(2)} EUR (écart ${ecart >= 0 ? '+' : ''}${ecart.toFixed(2)} EUR -> ${statut})`);
  }

  // "controle"/"totaux" (affiche UI, "Total HT") reflete le fichier LIVRE (donc APRES le
  // forfait Zones eloignees/Adresses -- decision utilisateur 2026-08-12), PAS le montant
  // brut reellement facture (celui-ci reste compare separement a chaque PDF plus haut, ou
  // un ecart est NORMAL et attendu a cause du forfait, pas un signe d'erreur).
  const controle = {
    'Droits et taxes': round2(importRows.reduce((s, r) => s + (r.DroitsTaxes || 0), 0)),
    Assurance: round2(importRows.reduce((s, r) => s + (r.Assurance || 0), 0)),
    'Zones éloignées': round2(importRows.reduce((s, r) => s + (r.ZonesEloignees || 0), 0)),
    'Colis volumineux': round2(importRows.reduce((s, r) => s + (r.ColisVolumineux || 0), 0)),
    Adresse: round2(importRows.reduce((s, r) => s + (r.Adresses || 0), 0)),
    'Frêt': round2(importRows.reduce((s, r) => s + (r.Fret || 0), 0)),
    Gazole: round2(importRows.reduce((s, r) => s + (typeof r.TaxeGasoil === 'number' ? r.TaxeGasoil : 0), 0)),
  };

  return {
    header: ['No Facture', 'Sous-compte', 'Numero LT', 'Type prestation', 'Zone Tarifaire', 'Poids', 'Produit', 'Montant HT'],
    rows: recs, recs, importRows, controle, warnings, alerts, infos,
    posteKeys: POSTE_KEYS, cfg, pdfs, lignesForfaitaires,
    sheetNames: { raw: 'Facture Chronopost', import: 'Fichier import' },
    period: dateValidite ? `${dateValidite.slice(6)}_${dateValidite.slice(3, 5)}` : 'export',
  };
}

module.exports = {
  id: 'chronopost',
  name: 'Chronopost',
  status: 'ready',
  taxeGasoil: "Reelle, imprimee en clair sur chaque facture PDF (ex. 'Surcharge Carburant Routier : 12,75% sur montant de X EUR') -- PAS une constante fixe a chercher sur le site chaque mois. Facturee en lignes forfaitaires separees (CAP*) par le transporteur, mises en pool par facture puis redistribuees au prorata du fret sur chaque ligne normale de cette meme facture.",
  method: "2 fichiers Excel bruts (1 par sous-compte, en-tête ligne 4) : les lignes 'colis' (Numero LT = tracking) sont reclassées en postes ERP via la table Catégories (Type prestation → poste). Les lignes forfaitaires (Numero LT commençant par CAP/ECO/SUR = Surcharge Carburant/Éco-Responsable/Sûreté) ne sont jamais dans l'import — leur montant Gazole (CAP) est mis en pool par facture puis redistribué au prorata du frêt sur chaque ligne colis. Mode envoi/zone via la table Bibliothèque transporteurs (+ règles spéciales 6B/6C via Zoning 2shop, 17/44 via Zone Tarifaire). Réconciliation PDF par n° de facture (Total HT).",
  inputs: [
    { key: 'facture', label: 'Facture(s) Chronopost reçue(s) (xlsx, 1 par sous-compte)', accept: '.xlsx', multiple: true, required: true },
    { key: 'pdf', label: 'Facture(s) PDF Chronopost (contrôle du total)', accept: '.pdf', multiple: true, required: false },
  ],
  outputNaming: { workbook: '{period}_Facture Chronopost', import: '{period}_Chronopost_Import' },
  finalizer: {
    script: '../automatisation/finaliser_chronopost.py',
    template: '../Transporteurs/Chronopost/2026_06_Facture Chronopost.xlsx',
    buildArgs: (files) => [...(files.facture || []), '--pdf', ...(files.pdf || [])],
  },
  process,
};
