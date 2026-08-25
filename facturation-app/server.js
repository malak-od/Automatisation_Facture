// ============================================================================
//  SERVEUR WEB — choisir un transporteur, deposer ses documents, generer l'import.
//  Lancer : npm start   puis ouvrir http://localhost:3000
// ============================================================================
const express = require('express');
const multer = require('multer');
const fs = require('fs');
const path = require('path');
const { execFile } = require('child_process');
const execFileAsync = require('util').promisify(execFile);
const registry = require('./src/registry');
const { writeImportCsv, writeImportXlsx, writeWorkbook, readImportRowsFromValuesCsv } = require('./src/core/excelOut');

const app = express();
const PORT = Number(process.env.PORT) || 4000; // Number() : sinon "4000"+1 = "40001" (concat de chaîne)
const UPLOADS = path.join(__dirname, 'uploads');
const OUTPUTS = path.join(__dirname, 'outputs');
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

app.use(express.json());
app.use(express.static(path.join(__dirname, 'public')));
app.use('/outputs', express.static(OUTPUTS));

// Liste des transporteurs (metadonnees pour l'UI)
app.get('/api/carriers', (_req, res) => res.json(registry.list()));

// Ouvre le dossier de sortie dans l'Explorateur Windows (logiciel local)
app.post('/api/reveal', (req, res) => {
  const stamp = String(req.body.stamp || '').replace(/[^a-z0-9_]/gi, '');
  const dir = path.join(OUTPUTS, stamp);
  if (!stamp || !dir.startsWith(OUTPUTS) || !fs.existsSync(dir)) return res.status(400).json({ error: 'Dossier introuvable' });
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
    for (const f of req.files || []) (files[f.fieldname] = files[f.fieldname] || []).push(f.path);

    // Mois choisi dans l'UI = source de verite si fourni (decision utilisateur 2026-08-20,
    // corrige le bug UPS "date validite decale un mois" -- l'auto-detection par carrier reste
    // le repli si le champ n'est pas envoye). Le select envoie "AAAA_MM" (meme format que
    // result.period) ; on passe aussi la variante sans underscore "AAAAMM" (format interne
    // moisCible d'UPS) pour eviter que chaque carrier ne refasse la conversion.
    const periodRaw = String(req.body.period || '');
    const periodMatch = /^(\d{4})_(\d{2})$/.exec(periodRaw);
    const periodOverride = periodMatch ? { formatted: periodRaw, compact: `${periodMatch[1]}${periodMatch[2]}` } : null;
    const result = await carrier.process(files, { period: periodOverride });
    const period = (periodOverride && periodOverride.formatted) || result.period || 'export';
    const suffix = 'test'; // suffixe figé (différencie du fichier fait à la main)
    const workbookOnly = !!carrier.workbookOnly; // ex. Delivengo : le classeur EST l'import
    const posteKeys = result.posteKeys || [];

    // Nomenclature (comme le fichier fait a la main) + suffixe
    const naming = carrier.outputNaming || { workbook: '{period}_classeur', import: '{period}_import' };
    const wbName = `${naming.workbook.replace('{period}', period)}_${suffix}.xlsx`;
    const impXlsxName = naming.import ? `${naming.import.replace('{period}', period)}_${suffix}.xlsx` : null;
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
        const extra = carrier.finalizer.buildArgs ? carrier.finalizer.buildArgs(files, period, __dirname) : (files.csv || []);
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

    // 1) Import ERP valeurs seules (CSV + XLSX) — sauf transporteurs "classeur seul". Genere
    // APRES le classeur (et non avant, cf. ordre historique) car carrier.importFromWorkbook
    // (Colissimo/Fedex) relit les valeurs REELLEMENT calculees par Excel dans le classeur tout
    // juste genere -- remontee pole transport 2026-08-24 : "le fichier import CSV de quelques
    // transporteurs deconne alors que la feuille Import CSV du fichier de facture est correcte
    // -> copier-coller depuis le fichier de la facture, coller en valeur". Repli sur
    // result.importRows (calcule en JS, comme avant) si le fichier de valeurs est absent
    // (finaliseur en echec, ou carrier sans ce flag).
    // multiImports : transporteurs a plusieurs fichiers import distincts (ex. Lettres :
    // Suivie / Prepa / SLAACE, fidele aux 3 fichiers du process reel) -- jamais combine avec
    // importFromWorkbook a ce jour.
    const multiDownloads = [];
    if (!workbookOnly && result.multiImports) {
      for (const m of result.multiImports) {
        const csvName = `${period}_${m.name}_Import_${suffix}.csv`;
        const xlsxName = `${period}_${m.name}_Import_${suffix}.xlsx`;
        writeImportCsv(m.importRows, path.join(dir, csvName));
        await writeImportXlsx(m.importRows, path.join(dir, xlsxName), m.sheetName || m.name);
        multiDownloads.push({
          key: m.key, name: m.name, lignes: m.lignes,
          csv: { url: `/outputs/${stamp}/${encodeURIComponent(csvName)}`, name: csvName },
          xlsx: { url: `/outputs/${stamp}/${encodeURIComponent(xlsxName)}`, name: xlsxName },
        });
      }
    } else if (!workbookOnly && impXlsxName) {
      let rowsForImport = result.importRows;
      if (carrier.importFromWorkbook && importValuesCsvPath) {
        const fromWorkbook = readImportRowsFromValuesCsv(importValuesCsvPath);
        if (fromWorkbook && fromWorkbook.length) rowsForImport = fromWorkbook;
        try { fs.unlinkSync(importValuesCsvPath); } catch (e) { /* ignore */ }
      }
      writeImportCsv(rowsForImport, path.join(dir, impCsvName));
      await writeImportXlsx(rowsForImport, path.join(dir, impXlsxName), (result.sheetNames || {}).import || `${carrier.name}_Import`);
    }

    const link = (name) => ({ url: `/outputs/${stamp}/${encodeURIComponent(name)}`, name });
    const downloads = {};
    if (workbookMode !== 'none') downloads.workbook = link(wbName);
    if (!workbookOnly && impXlsxName && !result.multiImports) { downloads.xlsx = link(impXlsxName); downloads.csv = link(impCsvName); }
    if (multiDownloads.length) downloads.multi = multiDownloads;
    res.json({
      carrier: carrier.name,
      periode: period,
      stamp,
      classeurClone: workbookMode === 'clone',
      lignes: result.lignes != null ? result.lignes : result.importRows.length,
      totaux: Object.fromEntries(posteKeys.map((k) => [k, Math.round((result.controle[k] || 0) * 100) / 100])),
      totalHt: Math.round(posteKeys.reduce((s, k) => s + (result.controle[k] || 0), 0) * 100) / 100,
      warnings: result.warnings,
      alerts: result.alerts,
      infos: result.infos || [],
      downloads,
    });
  } catch (e) {
    console.error('Error in /api/process:', e && e.stack ? e.stack : e);
    res.status(500).json({ error: e.message });
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
