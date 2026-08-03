// ============================================================================
//  Adaptateur transporteur : DPD
//  Lance le finaliseur Python `automatisation/finaliser_dpd.py` puis importe le CSV
//  d'ERP généré et retourne `importRows` (format attendu par l'app).
// ============================================================================
const path = require('path');
const fs = require('fs');
const os = require('os');
const { execFileSync } = require('child_process');
const { readCsv, num } = require('../../core/csv');
const { IMPORT_COLUMNS } = require('../../core/importSchema');
const { validate } = require('../../core/validate');

function process(files) {
  const csvPaths = files.csv || [];
  if (!csvPaths.length) throw new Error('Aucun fichier CSV fourni (attendu : fichiers DPD).');

  // Exécute le finaliseur Python qui produit un CSV d'import ERP
  const script = path.resolve(__dirname, '../../../..', 'automatisation', 'finaliser_dpd.py');
  const out = path.join(os.tmpdir(), `dpd_import_${Date.now()}.csv`);
  const args = ['--csv', ...csvPaths, '--output', out];
  let outBuf;
  try {
    outBuf = execFileSync('python', [script, ...args], { windowsHide: true, maxBuffer: 20 * 1024 * 1024 });
  } catch (e) {
    const stderr = (e.stderr && e.stderr.toString) ? e.stderr.toString() : String(e.stderr || e.message || '');
    const stdout = (e.stdout && e.stdout.toString) ? e.stdout.toString() : String(e.stdout || '');
    // Log full outputs to server console for debugging, and return a concise message to the UI
    console.error('DPD finalizer stdout:\n', stdout);
    console.error('DPD finalizer stderr:\n', stderr);
    throw new Error(`Finaliseur DPD KO : ${stderr.split(/\r?\n/).slice(0,4).join(' | ') || e.message}`);
  }
  const pyOut = outBuf ? outBuf.toString() : '';
  if (pyOut) console.log('DPD finalizer stdout:', pyOut.split(/\r?\n/).slice(0,5).join(' | '));

  if (!fs.existsSync(out)) throw new Error('Le finaliseur DPD n\'a pas produit le fichier attendu.');

  // Lit le CSV d'import (latin1 ; ';') et mappe colonnes vers IMPORT_COLUMNS
  const { header: csvHeader, rows } = readCsv(out, ';');
  const importRows = rows.map((r) => {
    const obj = {};
    IMPORT_COLUMNS.forEach((c, i) => { obj[c.key] = r[i] != null ? r[i] : ''; });
    return obj;
  });

  // Construire des `recs` minimaux à partir des importRows pour que writeWorkbook fonctionne
  const posteKeys = ['DroitsTaxes', 'Assurance', 'ZonesEloignees', 'ColisVolumineux', 'Adresses', 'Fret', 'PlusValueB2C', 'TaxeGasoil'];
  const recs = importRows.map((ir) => {
    const postes = {};
    let hg = 0, ag = 0;
    for (const k of posteKeys) {
      const v = num(ir[k]);
      postes[k.replace(/[^A-Za-z0-9]/g, '')] = v; // keep keys safe (they already match)
      if (k !== 'TaxeGasoil') hg += v;
      ag += v;
    }
    return {
      tracking: ir.Tracking || '',
      comref: ir.Ref1 || ir.Ref2 || '',
      dest: ir.Nom || '',
      pays: ir.Pays || '',
      zone: '',
      poids: num(ir.Poids),
      colis: parseInt(ir.NbrColis || 0, 10) || 0,
      postes,
      horsGo: Math.round(hg * 100) / 100,
      avecGo: Math.round(ag * 100) / 100,
      zone: '',
      raw: IMPORT_COLUMNS.map((c) => ir[c.key] || ''),
    };
  });

  // Controle = somme des postes
  const controle = {};
  for (const k of posteKeys) controle[k] = recs.reduce((s, r) => s + (r.postes[k] || 0), 0);

  try { fs.unlinkSync(out); } catch (e) { /* ignore */ }

  const { alerts, infos } = validate(importRows);
  return {
    header: IMPORT_COLUMNS.map((c) => c.label), recs, lignes: importRows.length, importRows, controle, warnings: [], alerts, infos, posteKeys,
    sheetNames: { raw: 'Facture DPD', import: 'Import ERP' },
  };
}

module.exports = {
  id: 'dpd',
  name: 'DPD',
  status: 'ready',
  taxeGasoil: 'Site DPD — routier 13,85%',
  method: 'Rassembler les fichiers (dossier CSV). Voir automatisation/finaliser_dpd.py',
  inputs: [
    { key: 'csv', label: 'Fichiers DPD (dossier CSV)', accept: '.csv', multiple: true, required: true },
  ],
  process,
};
