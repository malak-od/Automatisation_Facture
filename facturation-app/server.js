// ============================================================================
//  SERVEUR WEB — choisir un transporteur, deposer ses documents, generer l'import.
//  Lancer : npm start   puis ouvrir http://localhost:3000
// ============================================================================
require('dotenv').config(); // charge facturation-app/.env (identifiants API UPS, etc.) AVANT tout le reste
const express = require('express');
const multer = require('multer');
const fs = require('fs');
const os = require('os');
const path = require('path');
const { execFile } = require('child_process');
const execFileAsync = require('util').promisify(execFile);
const registry = require('./src/registry');
const { writeImportCsv, writeWorkbook, readImportRowsFromValuesCsv, readImportCsvFinal } = require('./src/core/excelOut');
const { validate } = require('./src/core/validate');
const { IMPORT_COLUMNS } = require('./src/core/importSchema');

const app = express();
const PORT = Number(process.env.PORT) || 4000; // Number() : sinon "4000"+1 = "40001" (concat de chaîne)
const UPLOADS = path.join(__dirname, 'uploads');
const OUTPUTS = path.join(__dirname, 'outputs');
// Dossier Telechargements Windows (demande utilisateur 2026-08-26 : le bouton "Ouvrir le
// dossier" doit pointer ici, pas sur outputs/ de l'app) -- sous-dossier par stamp (comme
// outputs/<stamp>/) pour ne pas melanger les fichiers de generations differentes.
const DOWNLOADS = path.join(os.homedir(), 'Downloads');
fs.mkdirSync(UPLOADS, { recursive: true });
fs.mkdirSync(OUTPUTS, { recursive: true });

const upload = multer({ dest: UPLOADS });

/** Message clair quand le classeur de sortie est ouvert dans Excel (EBUSY a
 * l'ecriture) : le message brut Node/Windows ne dit pas quoi faire. */
function explainFileError(e, targetPath) {
  const msg = String((e && e.stderr) || e.message || e);
  // Node (EBUSY/"resource busy") ET Python (shutil.copyfile via PermissionError, message
  // localise different selon la version/langue Windows -- "Permission denied", "Errno 13",
  // "utilisé par un autre programme"...) -- meme cause reelle (fichier de sortie ouvert dans
  // Excel au moment de la generation), 2 sources de traceback distinctes (constate sur TNT,
  // 2026-08-19 : shutil.copyfile echoue en PermissionError, pas capte par la regex d'origine
  // -> le vrai message n'etait jamais affiche, l'utilisateur voyait a la place le crash en
  // cascade du repli exceljs, "Cannot read properties of undefined (reading 'map')").
  if (/EBUSY|resource busy|being used by another process|PermissionError|Errno 13|Permission denied|utilis[ée] par un autre programme/i.test(msg)) {
    return new Error(`Le fichier "${path.basename(targetPath)}" est actuellement ouvert dans Excel (ou un autre programme) — ferme-le puis réessaie.`);
  }
  return e;
}

/** Filtre les messages d'erreur techniques (stack Node/Python, TypeError, chemins bruts...)
 * avant affichage utilisateur (audit pole transport 2026-08-26, point 3 : un message brut
 * remontait tel quel a l'ecran pour tout ce qui n'etait pas deja gere par explainFileError,
 * ex. bug Python inattendu -> illisible pour un non-developpeur). Le detail complet reste dans
 * les logs serveur (console.error) pour le dev -- seul l'AFFICHAGE change ici. Les messages
 * metier volontaires (deja rediges en francais clair par le code, ex. "Aucun fichier fourni...")
 * ne matchent aucun de ces patterns et passent donc inchanges. */
function userFacingError(e) {
  const msg = String((e && e.message) || e || '');
  const looksTechnical = /Traceback \(most recent call last\)|^(Type|Reference|Syntax|Range)Error|Cannot read propert|is not a function|is not defined|ENOENT|EACCES|undefined is not|\bat \S+:\d+:\d+|^Error: spawn|non[- ]zero exit status/i.test(msg);
  if (!looksTechnical) return msg;
  return "Une erreur inattendue est survenue pendant la génération — réessayez, et si le problème persiste, signalez-le à l'informatique avec l'heure exacte (le détail technique est conservé dans les logs du serveur).";
}

// Postes affiches dans le total UI -- colonnes O (Droits et taxes) a V (Gazole) du fichier
// import.csv (demande utilisateur 2026-08-26 : "le total ça sera la somme de toutes les lignes
// de la colonne O jusqu'à la colonne V"). Calcule sur le VRAI CSV final ecrit sur disque, pas
// sur une donnee intermediaire en memoire -- meme principe que les alertes (cf. plus bas).
const TOTAL_POSTE_KEYS = ['DroitsTaxes', 'Assurance', 'ZonesEloignees', 'ColisVolumineux', 'Adresses', 'Fret', 'PlusValueB2C', 'TaxeGasoil'];
// Libelles lisibles pour l'UI, repris de IMPORT_COLUMNS (source unique des en-tetes) plutot que
// duplique en dur ici.
const POSTE_LABELS = Object.fromEntries(IMPORT_COLUMNS.filter((c) => TOTAL_POSTE_KEYS.includes(c.key)).map((c) => [c.key, c.label.trim()]));

/** Somme les postes O->V sur des lignes d'import deja relues (readImportCsvFinal) -- TaxeGasoil
 * reste texte dans le schema (num:false) meme s'il est numerique en pratique, d'ou le parse
 * manuel commun a tous les postes plutot que de se fier a IMPORT_COLUMNS.num. */
function sumPostesRows(rows) {
  const totaux = Object.fromEntries(TOTAL_POSTE_KEYS.map((k) => [k, 0]));
  for (const r of rows || []) {
    for (const k of TOTAL_POSTE_KEYS) {
      const raw = r[k];
      const n = typeof raw === 'number' ? raw : Number(String(raw ?? '').replace(',', '.'));
      if (Number.isFinite(n)) totaux[k] += n;
    }
  }
  for (const k of TOTAL_POSTE_KEYS) totaux[k] = Math.round(totaux[k] * 100) / 100;
  return totaux;
}

app.use(express.json());
app.use(express.static(path.join(__dirname, 'public')));
app.use('/outputs', express.static(OUTPUTS));

// Liste des transporteurs (metadonnees pour l'UI)
app.get('/api/carriers', (_req, res) => res.json(registry.list()));

// Ouvre le dossier de sortie (copie dans Telechargements) dans l'Explorateur Windows (logiciel local)
app.post('/api/reveal', (req, res) => {
  const stamp = String(req.body.stamp || '').replace(/[^a-z0-9_]/gi, '');
  const dir = path.join(DOWNLOADS, stamp);
  if (!stamp || !dir.startsWith(DOWNLOADS) || !fs.existsSync(dir)) return res.status(400).json({ error: 'Dossier introuvable' });
  execFile('explorer.exe', [dir], () => {}); // explorer renvoie un code != 0 meme en cas de succes -> on ignore
  res.json({ ok: true });
});

// Traitement : recoit les fichiers + l'id transporteur, genere les livrables
app.post('/api/process', upload.any(), async (req, res) => {
  try {
    const carrier = registry.get(req.body.carrier);
    if (carrier.status !== 'ready') {
      return res.status(400).json({ error: `"${carrier.name}" pas encore implemente.`, method: carrier.method });
    }
    // regroupe les fichiers uploades par cle d'input (nom de champ)
    const files = {};
    // Noms de fichiers ORIGINAUX (avant renommage multer en hex sans extension), meme
    // regroupement que `files` -- ajoute a part (n'existait pas avant) pour ne rien casser
    // des carriers qui traitent deja `files.<key>` comme un simple tableau de chemins passe
    // tel quel a execFileAsync. Necessaire pour trier les exports "AAAA MM - Export
    // expeditions_brut.xlsx" par mois (UPS, repli Poids sur mois courant/M-1/M-2 -- l'ordre de
    // depot par l'utilisateur ne peut pas etre suppose fiable).
    const fileNames = {};
    for (const f of req.files || []) {
      (files[f.fieldname] = files[f.fieldname] || []).push(f.path);
      (fileNames[f.fieldname] = fileNames[f.fieldname] || []).push(f.originalname);
    }

    // Mois choisi dans l'UI = source de verite si fourni (decision utilisateur 2026-08-20,
    // corrige le bug UPS "date validite decale un mois" -- l'auto-detection par carrier reste
    // le repli si le champ n'est pas envoye). Le select envoie "AAAA_MM" (meme format que
    // result.period) ; on passe aussi la variante sans underscore "AAAAMM" (format interne
    // moisCible d'UPS) pour eviter que chaque carrier ne refasse la conversion.
    const periodRaw = String(req.body.period || '');
    const periodMatch = /^(\d{4})_(\d{2})$/.exec(periodRaw);
    const periodOverride = periodMatch ? { formatted: periodRaw, compact: `${periodMatch[1]}${periodMatch[2]}` } : null;
    const result = await carrier.process(files, { period: periodOverride, fileNames });
    const period = (periodOverride && periodOverride.formatted) || result.period || 'export';
    const suffix = 'test'; // suffixe figé (différencie du fichier fait à la main)
    const workbookOnly = !!carrier.workbookOnly; // ex. Delivengo : le classeur EST l'import

    // Nomenclature (comme le fichier fait a la main) + suffixe
    const naming = carrier.outputNaming || { workbook: '{period}_classeur', import: '{period}_import' };
    const wbName = `${naming.workbook.replace('{period}', period)}_${suffix}.xlsx`;
    const impCsvName = naming.import ? `${naming.import.replace('{period}', period)}_${suffix}.csv` : null;

    // outputs/<carrier>_<periode>/ (ex. kuehne_2026_06). Regenerer le meme mois ecrase le dossier.
    const stamp = `${carrier.id}_${period}`;
    const dir = path.join(OUTPUTS, stamp);
    fs.mkdirSync(dir, { recursive: true });

    // 2) Classeur = CLONE FIDELE du fichier fait a la main (Excel COM/Python) ; repli exceljs.
    // noWorkbook : transporteurs sans classeur de reference (ex. Lettres, 3 imports
    // separes calcules en JS pur) -> writeWorkbook() attend une structure recs/rows
    // absente ici, pas la peine de produire un classeur casse/vide.
    const wbPath = path.join(dir, wbName);
    let workbookMode = 'exceljs';
    // Rempli si le finaliseur a exporte les valeurs calculees de son onglet Import (cf.
    // carrier.importFromWorkbook ci-dessous) -- reste null pour tous les autres carriers.
    let importValuesCsvPath = null;
    if (carrier.noWorkbook) {
      workbookMode = 'none';
    } else if (result.prebuiltWorkbookPath && fs.existsSync(result.prebuiltWorkbookPath)) {
      // deja genere/calcule dans process() (ex. Delivengo, qui relit ce meme
      // classeur pour construire importRows) -> reprendre tel quel, pas relancer Excel.
      try {
        fs.copyFileSync(result.prebuiltWorkbookPath, wbPath);
      } catch (e) {
        throw explainFileError(e, wbPath);
      }
      try { fs.unlinkSync(result.prebuiltWorkbookPath); } catch (e) { /* ignore */ }
      workbookMode = 'clone';
    } else if (carrier.finalizer) {
      try {
        const scriptAbs = path.resolve(__dirname, carrier.finalizer.script);
        const templateAbs = path.resolve(__dirname, carrier.finalizer.template);
        const extra = carrier.finalizer.buildArgs ? carrier.finalizer.buildArgs(files, period, __dirname, fileNames) : (files.csv || []);
        const { stdout } = await execFileAsync('python', [scriptAbs, templateAbs, wbPath, ...extra], { windowsHide: true, maxBuffer: 20 * 1024 * 1024 });
        workbookMode = 'clone';
        // le finaliseur peut signaler des cas a verifier (ex. poids introuvable dans les bruts,
        // pays/mode envoi ajoutes automatiquement -- cf. finaliser_colissimo.py)
        for (const line of String(stdout || '').split(/\r?\n/)) {
          const mPoids = line.match(/^INFO_POIDS_MANQUANT:(.+)$/);
          if (mPoids) (result.infos = result.infos || []).push(`Poids introuvable dans les exports bruts : ${mPoids[1].trim()} (plancher 0,15 appliqué — à vérifier/saisir à la main)`);
          const mPays = line.match(/^AJOUT_PAYS_AUTO:(.+)$/);
          if (mPays) (result.warnings = result.warnings || []).push(`Pays ajouté automatiquement à la table "Pays" : ${mPays[1].trim()}`);
          const mMode = line.match(/^AJOUT_MODE_ENVOI_AUTO:(.+)$/);
          if (mMode) (result.warnings = result.warnings || []).push(`Mode d'envoi déduit automatiquement (à vérifier) : ${mMode[1].trim()}`);
          const mExport = line.match(/^EXPORT_IMPORT_VALEURS:(.+)$/);
          if (mExport) importValuesCsvPath = mExport[1].trim();
        }
      } catch (e) {
        console.warn('Finaliseur Excel KO :', String(e.stderr || e.message || '').slice(0, 300));
        const explained = explainFileError(e, wbPath);
        // Fichier verrouille : le repli exceljs echouerait pareil sur le meme chemin, inutile d'essayer.
        if (workbookOnly || explained !== e) throw explained;
        await writeWorkbook({ ...result, pdfs: result.pdfs || null, carrierName: carrier.name }, wbPath);
      }
    } else {
      await writeWorkbook({ ...result, pdfs: result.pdfs || null }, wbPath);
    }

    // 1) Import ERP valeurs seules (CSV UNIQUEMENT, plus de XLSX -- demande utilisateur
    // 2026-08-26 : 2 fichiers de sortie par generation, pas 3 -- classeur Excel + CSV import,
    // pour tous les transporteurs sans exception, y compris Lettres/multiImports). Genere
    // APRES le classeur (et non avant, cf. ordre historique) car carrier.importFromWorkbook
    // (Colissimo/Fedex/UPS) relit les valeurs REELLEMENT calculees par Excel dans le classeur
    // tout juste genere -- remontee pole transport 2026-08-24 : "le fichier import CSV de
    // quelques transporteurs deconne alors que la feuille Import CSV du fichier de facture est
    // correcte -> copier-coller depuis le fichier de la facture, coller en valeur". Repli sur
    // result.importRows (calcule en JS, comme avant) si le fichier de valeurs est absent
    // (finaliseur en echec, ou carrier sans ce flag).
    // multiImports : transporteurs a plusieurs fichiers import distincts (ex. Lettres :
    // Suivie / Prepa / SLAACE, fidele aux 3 fichiers du process reel) -- jamais combine avec
    // importFromWorkbook a ce jour.
    const multiDownloads = [];
    // Total UI (colonnes O->V) recalcule sur le VRAI CSV final ecrit sur disque -- initialise a
    // partir de result.controle (repli si aucun CSV n'est ecrit, ex. workbookOnly), ecrase plus
    // bas des que les fichiers sont relus.
    let totaux = Object.fromEntries(TOTAL_POSTE_KEYS.map((k) => [k, Math.round((result.controle[k] || 0) * 100) / 100]));
    if (!workbookOnly && result.multiImports) {
      const sommeMulti = Object.fromEntries(TOTAL_POSTE_KEYS.map((k) => [k, 0]));
      for (const m of result.multiImports) {
        const csvName = `${period}_${m.name}_Import_${suffix}.csv`;
        const csvPath = path.join(dir, csvName);
        writeImportCsv(m.importRows, csvPath);
        multiDownloads.push({
          key: m.key, name: m.name, lignes: m.lignes,
          csv: { url: `/outputs/${stamp}/${encodeURIComponent(csvName)}`, name: csvName },
        });
        const finalRowsM = readImportCsvFinal(csvPath);
        if (finalRowsM) {
          const t = sumPostesRows(finalRowsM);
          for (const k of TOTAL_POSTE_KEYS) sommeMulti[k] += t[k];
        }
      }
      for (const k of TOTAL_POSTE_KEYS) totaux[k] = Math.round(sommeMulti[k] * 100) / 100;
    } else if (!workbookOnly && impCsvName) {
      let rowsForImport = result.importRows;
      if (carrier.importFromWorkbook && importValuesCsvPath) {
        const fromWorkbook = readImportRowsFromValuesCsv(importValuesCsvPath);
        if (fromWorkbook && fromWorkbook.length) rowsForImport = fromWorkbook;
        try { fs.unlinkSync(importValuesCsvPath); } catch (e) { /* ignore */ }
      }
      const impCsvPath = path.join(dir, impCsvName);
      writeImportCsv(rowsForImport, impCsvPath);
      // Alertes (POIDS=0, ZONE manquante, etc.) ET total UI (colonnes O->V, "Droits et taxes" a
      // "Gazole") recalcules APRES coup en relisant le VRAI fichier CSV ecrit sur disque --
      // demande utilisateur 2026-08-26 : "il faut remonter ces alertes apres la generation
      // finale du fichier import.csv, car c'est celui qu'on importe dans l'ERP" (pas sur une
      // donnee intermediaire en memoire, meme si le contenu est cense etre identique --
      // garantit qu'aucune transformation d'ecriture, ex. arrondi/formatage fmtCsv, ne fait
      // diverger l'alerte/le total du fichier reellement livre). S'applique a tous les
      // transporteurs, pas seulement ceux avec importFromWorkbook.
      const finalRows = readImportCsvFinal(impCsvPath);
      if (finalRows) {
        result.alerts = validate(finalRows).alerts;
        totaux = sumPostesRows(finalRows);
      }
    }

    const link = (name) => ({ url: `/outputs/${stamp}/${encodeURIComponent(name)}`, name });
    const downloads = {};
    if (workbookMode !== 'none') downloads.workbook = link(wbName);
    if (!workbookOnly && impCsvName && !result.multiImports) { downloads.csv = link(impCsvName); }
    if (multiDownloads.length) downloads.multi = multiDownloads;

    // Copie des fichiers generes dans Telechargements\<stamp>\ (demande utilisateur 2026-08-26 :
    // le bouton "Ouvrir le dossier" doit pointer vers Telechargements, pas outputs/ de l'app) --
    // best-effort : ne bloque jamais la reponse si Telechargements est inaccessible/plein.
    try {
      const dlDir = path.join(DOWNLOADS, stamp);
      fs.mkdirSync(dlDir, { recursive: true });
      for (const name of fs.readdirSync(dir)) {
        fs.copyFileSync(path.join(dir, name), path.join(dlDir, name));
      }
    } catch (e) {
      console.warn('Copie vers Telechargements KO :', String(e.message || e).slice(0, 200));
    }

    res.json({
      carrier: carrier.name,
      periode: period,
      stamp,
      classeurClone: workbookMode === 'clone',
      lignes: result.lignes != null ? result.lignes : result.importRows.length,
      totaux: Object.fromEntries(TOTAL_POSTE_KEYS.map((k) => [POSTE_LABELS[k], totaux[k] || 0])),
      totalHt: Math.round(TOTAL_POSTE_KEYS.reduce((s, k) => s + (totaux[k] || 0), 0) * 100) / 100,
      warnings: result.warnings,
      alerts: result.alerts,
      infos: result.infos || [],
      downloads,
    });
  } catch (e) {
    console.error('Error in /api/process:', e && e.stack ? e.stack : e);
    res.status(500).json({ error: userFacingError(e) });
  }
});

// Démarrage robuste : si le port est occupé (app déjà lancée ?), on essaie le suivant.
function start(port, attemptsLeft = 10) {
  const server = app.listen(port, () => console.log(`Facturation transporteurs -> http://localhost:${port}`));
  server.on('error', (err) => {
    if (err.code === 'EADDRINUSE' && attemptsLeft > 0) {
      console.warn(`Port ${port} occupé (l'app tourne peut-être déjà). J'essaie le port ${port + 1}…`);
      start(port + 1, attemptsLeft - 1);
    } else if (err.code === 'EADDRINUSE') {
      console.error(`Ports ${PORT}–${port} tous occupés. Ferme l'app déjà lancée, ou choisis un port : set PORT=4000 && npm start`);
      process.exit(1);
    } else {
      throw err;
    }
  });
}
start(PORT);
