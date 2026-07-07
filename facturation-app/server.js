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
const { writeImportCsv, writeImportXlsx, writeWorkbook } = require('./src/core/excelOut');

const app = express();
const PORT = Number(process.env.PORT) || 4000; // Number() : sinon "4000"+1 = "40001" (concat de chaîne)
const UPLOADS = path.join(__dirname, 'uploads');
const OUTPUTS = path.join(__dirname, 'outputs');
fs.mkdirSync(UPLOADS, { recursive: true });
fs.mkdirSync(OUTPUTS, { recursive: true });

const upload = multer({ dest: UPLOADS });

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

    const result = carrier.process(files);
    const period = result.period || 'export';
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

    // 1) Import ERP valeurs seules (CSV + XLSX) — sauf transporteurs "classeur seul"
    if (!workbookOnly && impXlsxName) {
      writeImportCsv(result.importRows, path.join(dir, impCsvName));
      await writeImportXlsx(result.importRows, path.join(dir, impXlsxName));
    }

    // 2) Classeur = CLONE FIDELE du fichier fait a la main (Excel COM/Python) ; repli exceljs.
    const wbPath = path.join(dir, wbName);
    let workbookMode = 'exceljs';
    if (carrier.finalizer) {
      try {
        const scriptAbs = path.resolve(__dirname, carrier.finalizer.script);
        const templateAbs = path.resolve(__dirname, carrier.finalizer.template);
        const extra = carrier.finalizer.buildArgs ? carrier.finalizer.buildArgs(files, period, __dirname) : (files.csv || []);
        const { stdout } = await execFileAsync('python', [scriptAbs, templateAbs, wbPath, ...extra], { windowsHide: true, maxBuffer: 20 * 1024 * 1024 });
        workbookMode = 'clone';
        // le finaliseur peut signaler des cas a verifier (ex. poids introuvable dans les bruts)
        for (const line of String(stdout || '').split(/\r?\n/)) {
          const m = line.match(/^INFO_POIDS_MANQUANT:(.+)$/);
          if (m) (result.infos = result.infos || []).push(`Poids introuvable dans les exports bruts : ${m[1].trim()} (plancher 0,15 appliqué — à vérifier/saisir à la main)`);
        }
      } catch (e) {
        console.warn('Finaliseur Excel KO :', String(e.stderr || e.message || '').slice(0, 300));
        if (workbookOnly) throw e; // pas de repli exceljs possible (pas de schema d'import generique)
        await writeWorkbook({ ...result, pdfs: null }, wbPath);
      }
    } else {
      await writeWorkbook({ ...result, pdfs: null }, wbPath);
    }

    const link = (name) => ({ url: `/outputs/${stamp}/${encodeURIComponent(name)}`, name });
    const downloads = { workbook: link(wbName) };
    if (!workbookOnly && impXlsxName) { downloads.xlsx = link(impXlsxName); downloads.csv = link(impCsvName); }
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
