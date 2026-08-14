// ============================================================================
//  Adaptateur transporteur : MONDIAL RELAY
//  Plusieurs fichiers CSV "Annexe_..." (1 par facture/compte). 1 ligne = 1 colis.
//  Reclassement : Fret = Montant transport poids mesure [+ Complement si
//  EXACTEMENT 0,03] ; Zones eloignees = Complement (sauf 0 ou 0,03) ; Zone =
//  lookup Mode+Pays dans la table 'Pays' du classeur fait a la main. Le mode
//  envoi '24RC'/'LCC' est normalise en '24R' A LA FOIS pour la cle de lookup
//  ET pour la valeur AFFICHEE dans le fichier d'import (regle communiquee par
//  le pole transport, 2026-08-14 : laisser le mode brut affiche cree des
//  zones "inconnues" cote ERP, qui refait son propre lookup zone a partir du
//  mode envoi affiche -> avaries d'import). Indexation gasoil est reelle mais
//  n'a pas de colonne dans le fichier d'import fait a la main -> geree a part
//  dans l'ERP (note terrain "identique au mois precedent"), remontee en info
//  uniquement pour le controle du total.
//
//  Verifie a 100% sur les montants (Fret/ZonesEloignees) contre le fichier
//  fait a la main de juin 2026 (12308/12308 lignes).
//
//  "Facture France" (compte LARUCH) : le PDF recu inclut un montant de
//  COLLECTE (transport aller vers les clients depuis l'entrepot) qui n'est
//  PAS facture aux clients finaux et n'apparait donc PAS dans les CSV
//  Annexe_* -> ecart systematique entre le total calcule et le PDF sur cette
//  seule facture (684,00EUR net constate en juin, note manuelle dans l'onglet
//  'Controle xls pdf' du classeur fait a la main, cellule S30/T30/U30
//  'Collecte'). Ne PAS chercher a faire tomber cet ecart a zero par le calcul
//  -- il est normal, la collecte doit etre ajoutee manuellement au controle.
// ============================================================================
const path = require('path');
const pdfParse = require('pdf-parse');
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

/** PDF Mondial Relay : balise <ref>...codeClient:X--...--numeroFacture:Y--...--montantHT:Z--...</ref>
 * en pied de CHAQUE page, format machine-friendly deja structure (pas de mise en page a parser)
 * -- confirme identique sur les 12 PDF de juin 2026. numeroFacture correspond exactement au nom
 * du CSV Annexe_<numeroFacture>_... (colonne 'Nr de facture' du CSV) -> cle de rapprochement directe. */
async function extractMrPdfTotal(pdfPath) {
  const buf = require('fs').readFileSync(pdfPath);
  const { text } = await pdfParse(buf);
  const m = /<ref>(.+?)<\/ref>/.exec(text);
  if (!m) return null;
  const fields = Object.fromEntries(m[1].split('--').map((kv) => kv.split(':')));
  if (!fields.numeroFacture || !fields.montantHT) return null;
  return {
    file: path.basename(pdfPath),
    numeroFacture: fields.numeroFacture,
    montantHt: num(fields.montantHT.replace(/\./g, '').replace(',', '.')),
  };
}

async function process(files) {
  const csvPaths = files.csv || [];
  if (!csvPaths.length) throw new Error('Aucun fichier CSV fourni (attendu : Annexe_*.csv Mondial Relay).');

  let header = null;
  const rows = [];
  for (const p of csvPaths) {
    const { header: h, rows: rs } = readCsv(p, cfg.csv.sep, cfg.csv.encoding);
    header = h;
    rows.push(...rs.map((r) => (r.length < h.length ? r.concat(Array(h.length - r.length).fill('')) : r)));
  }

  const iFacture = colIndex(header, cfg.cols.nrFacture);
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
    // Lignes de resume/en-tete de fichier (1 par CSV, Nom = 'LA RUCHE LOGISTIQUE FR', tracking
    // 0 ou 1, Date = '0') : gardees telles quelles dans le fichier fait a la main (verifie sur
    // juin 2026 : 13 lignes identiques, Zone='zone inconnue', tout a 0) -> ne PAS les filtrer,
    // seulement une ligne totalement vide (aucun tracking du tout) est ignoree.
    if (!trackBrut) continue;
    if (!dateValidite) {
      const d = firstDayOfMonth(r[iDate]);
      if (d) dateValidite = d; // ignore les lignes a Date='0' (resume/en-tete), sans les exclure du resultat
    }

    const complement = round2(num(r[iComplement]));
    const montant = num(r[iMontantTransport]);
    // Seuil = 0.03. ERREUR PRECEDENTE CORRIGEE (2026-08-10) : ce seuil avait ete change a tort
    // vers 0,08 en lisant les formules Q3/T3 du 'Fichier import' du modele (qui utilisent bien
    // 0,08) -- mais en recalculant les totaux sur les 2 seuils et en les comparant a
    // 'Controle xls pdf'!K16/M16 (Somme de Complement/Montant du transport, valeurs de
    // reference deja validees : 1545.20EUR / 36385.88EUR), seul 0,03 redonne ces totaux
    // exacts. Les formules Q3/T3 du modele semblent elles-memes erronees/obsoletes (ou
    // s'appliquer a un autre contexte) -- NE PAS s'y fier pour ce seuil, se fier au total
    // Controle xls pdf.
    const foldComplement = Math.abs(complement - 0.03) < 0.001;
    const fret = round2(foldComplement ? montant + complement : montant);
    const zonesEloignees = (complement === 0 || foldComplement) ? 0 : complement;
    const gazoleLigne = num(r[iGazole]);
    gazoleTotal += gazoleLigne;

    const pays = (r[iPays] || '').trim();
    // 24RC et LCC -> 24R, a la fois pour la CLE DE LOOKUP de zone ET pour la valeur de mode
    // envoi AFFICHEE dans le fichier d'import (regle communiquee par le pole transport,
    // 2026-08-14) : laisser '24RC'/'LCC' affiches cree des zones "inconnues" cote ERP (qui
    // fait son propre lookup a partir du mode envoi affiche) -> avaries d'import. Avant cette
    // correction, seule la cle de lookup interne etait normalisee (le mode BRUT restait
    // affiche tel quel, cf. ancien commentaire/limite documentee du 2026-08-10) -- desormais
    // la meme normalisation s'applique aux deux usages.
    const modeBrut = (r[iMode] || '').trim();
    const mode = (modeBrut === '24RC' || modeBrut === 'LCC') ? '24R' : modeBrut;
    const zoneKey = `${mode}-${pays}`;
    const zone = cfg.zone_table[zoneKey];
    if (!zone) warnings.push(`Zone inconnue pour la cle '${zoneKey}' (tracking ${trackBrut}) -> table 'Pays' a completer`);

    const poidsGr = Math.max(num(r[iPoidsAnnonce]), num(r[iPoidsMesure]));

    recs.push({
      facture: iFacture >= 0 ? (r[iFacture] || '').trim() : '',
      tracking: normalizeTracking(trackBrut),
      dest: iNom >= 0 ? (r[iNom] || '').trim() : '',
      pays, zone: zone || 'zone inconnue',
      colis: iNbColis >= 0 ? num(r[iNbColis]) : 1,
      poids: poidsGr / 1000,
      mode,
      postes: { ZonesEloignees: zonesEloignees, Fret: fret },
      horsGo: round2(zonesEloignees + fret),
      avecGo: round2(zonesEloignees + fret + gazoleLigne),
      gazoleLigne,
      // Non arrondi ligne a ligne (contrairement a postes.*, deja arrondis pour l'affichage
      // colonne par colonne du fichier d'import) -- necessaire pour la reconciliation PDF :
      // arrondir chaque ligne avant de sommer sur des milliers de lignes accumule une erreur
      // (jusqu'a ~1EUR par facture, constate en juin 2026) que le PDF (qui arrondit seulement
      // ses propres sous-totaux) n'a pas.
      totalBrutNonArrondi: num(r[iComplement]) + num(r[iMontantTransport]) + gazoleLigne,
      raw: r,
    });
  }

  // Ligne "sans montant" : Fret ET tous les autres postes a facturer (Zones eloignees,
  // Gazole -- Droits et taxes/Assurance/Colis volumineux/Adresses/Plus-value BtoC sont
  // TOUJOURS 0 pour Mondial Relay, pas de calcul dedie) sont a 0 -> rien a facturer nulle
  // part sur cette ligne, supprimee plutot que de generer 3 alertes en cascade sans valeur
  // metier (COLIS=0/POIDS=0/FRET=0) -- decision utilisateur 2026-08-10, verifiee sur juillet
  // 2026 : la plupart de ces lignes ont un vrai tracking/poids/destinataire, mais 'Montant du
  // transport' ET 'Complement' sont a 0 dans le CSV source lui-meme (pas juste un artefact de
  // colonnes vides ailleurs). Inclut aussi les lignes de resume/en-tete de fichier (tracking
  // normalise '00000000'/'00000001', Zone='zone inconnue') -- confirme par l'utilisateur
  // (capture d'ecran) : a supprimer systematiquement comme les autres, meme si le fichier de
  // reference de juin les gardait (choix retenu pour les traitements futurs).
  //
  // EXCLUSION SUPPLEMENTAIRE : la ligne de resume 'tracking = 1' (normalise '00000001') porte
  // le total Gazole/Semi de TOUT le fichier (pas une expedition -- Nom vide, tous les
  // identifiants vides, colonne 'Semi' = montant de la collecte, deja utilisee a part pour
  // recalculer 'Bilan clients'!F3 et 'Controle xls pdf'!S30, cf. finaliser_mondial_relay.py)
  // -> son Gazole non-nul lui faisait echapper au filtre "tout a zero" ci-dessus, mais elle ne
  // doit jamais apparaitre dans le fichier d'import (decision utilisateur 2026-08-10).
  let nbLignesSansMontantSupprimees = 0;
  const recsFiltres = recs.filter((rec) => {
    if (rec.tracking === '00000001') { nbLignesSansMontantSupprimees++; return false; }
    const toutAZero = rec.postes.ZonesEloignees === 0 && rec.postes.Fret === 0 && rec.gazoleLigne === 0;
    if (toutAZero) { nbLignesSansMontantSupprimees++; return false; }
    return true;
  });

  const importRows = recsFiltres.map((rec) => ({
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
  if (nbLignesSansMontantSupprimees) infos.push(`${nbLignesSansMontantSupprimees} ligne(s) supprimée(s) (Frêt, Zones éloignées et Gazole tous à 0 — rien à facturer sur cette ligne).`);

  // Reconciliation PDF (optionnelle) : montant HT total par facture (balise <ref> en pied de
  // chaque page PDF, cf. extractMrPdfTotal) compare au total calcule pour cette meme facture
  // (Complement + Montant transport + Indexation gasoil, SOMME NON ARRONDIE ligne a ligne --
  // cf. totalBrutNonArrondi -- puis arrondie une seule fois au total, comme le fait le PDF sur
  // ses propres sous-totaux). Rapproche par 'Nr de facture' (= nom du PDF/CSV, ex.
  // F3LARUCH2600000265), pas par tracking individuel (le PDF ne detaille pas colis par colis).
  const totalParFacture = new Map();
  for (const rec of recs) {
    const cle = rec.facture;
    if (!cle) continue;
    totalParFacture.set(cle, (totalParFacture.get(cle) || 0) + rec.totalBrutNonArrondi);
  }
  for (const [cle, v] of totalParFacture) totalParFacture.set(cle, round2(v));

  const pdfPaths = files.pdf || [];
  const pdfs = [];
  for (const p of pdfPaths) {
    try {
      const r = await extractMrPdfTotal(p);
      if (r) pdfs.push(r);
      else warnings.push(`PDF ${path.basename(p)} : balise de synthese introuvable, ignore pour la reconciliation.`);
    } catch (e) {
      warnings.push(`PDF ${path.basename(p)} : lecture impossible (${e.message}).`);
    }
  }
  for (const p of pdfs) {
    const calcule = totalParFacture.get(p.numeroFacture);
    if (calcule == null) {
      warnings.push(`Facture PDF ${p.file} (${p.numeroFacture}) : aucune ligne correspondante trouvee dans les CSV traites.`);
      continue;
    }
    const ecart = round2(p.montantHt - calcule);
    // Facture LARUCH (compte France) : peut porter un montant de COLLECTE (transport aller vers
    // les clients), jamais dans les CSV (cf. docstring d'en-tete) -> un ecart POSITIF notable
    // (>1EUR, au-dela du bruit d'arrondi centime habituel sur les autres factures) est attendu
    // sur ce compte precis, pas un signe d'erreur.
    const estCollecteProbable = p.numeroFacture.startsWith('F3LARUCH') && ecart > 1;
    let statut = 'OK';
    if (Math.abs(ecart) >= 0.05) statut = estCollecteProbable ? 'probable collecte (normal, non facturee aux clients)' : 'A VERIFIER';
    infos.push(`Facture PDF ${p.file} (${p.numeroFacture}) : Montant HT = ${p.montantHt.toFixed(2)} EUR, calcule = ${calcule.toFixed(2)} EUR (ecart ${ecart >= 0 ? '+' : ''}${ecart.toFixed(2)} EUR -> ${statut})`);
  }

  return {
    header, rows, recs, importRows, controle, warnings, alerts, infos,
    posteKeys: POSTE_KEYS, cfg, pdfs,
    sheetNames: { raw: 'Facture Mondial Relay', import: 'Fichier import' },
    period: dateValidite ? `${dateValidite.slice(6)}_${dateValidite.slice(3, 5)}` : 'export',
  };
}

module.exports = {
  id: 'mondial_relay',
  name: 'Mondial Relay',
  status: 'ready',
  taxeGasoil: "Reelle (indexation gasoil par ligne) mais geree hors fichier d'import, cote ERP -- 'identique au mois precedent' selon le process",
  method: "Fichiers Annexe_*.csv (1 par facture, dossier). Fret = Montant transport poids mesure (+Complement si =0,03 pile) ; Zones eloignees = Complement sinon. Zone = Mode+Pays via la table de correspondance du classeur (24RC/LCC normalises pour ce lookup uniquement, mode envoi affiche reste brut). Facture 'France' (LARUCH) : ecart normal avec le PDF = montant de collecte non facture aux clients, a ajouter manuellement au controle.",
  inputs: [
    { key: 'csv', label: 'Fichiers Mondial Relay (dossier Annexe_*.csv)', accept: '.csv', multiple: true, required: true },
    { key: 'pdf', label: 'Facture(s) PDF Mondial Relay (contrôle du total par facture)', accept: '.pdf', multiple: true, required: false },
  ],
  // Classeur = CLONE FIDELE du fichier fait a la main (Excel COM/Python), comme DPD/Geodis/GLS.
  outputNaming: { workbook: '{period}_Facture Mondial Relay', import: '{period}_MondialRelay_Import' },
  finalizer: {
    script: '../automatisation/finaliser_mondial_relay.py',
    template: '../Transporteurs/Mondial Relay/2026_06_Facture Mondial Relay.xlsx',
    buildArgs: (files) => [...(files.csv || []), '--pdf', ...(files.pdf || [])],
  },
  process,
};
