#!/usr/bin/env node
// ============================================================================
//  CLI — meme moteur que l'app web, en ligne de commande.
//  Ex : node cli.js kuehne --input ../Kuehne --out ./out --check ../Kuehne/2026_06_Kuehne_Import.csv
// ============================================================================
const fs = require('fs');
const path = require('path');
const registry = require('./src/registry');
const { writeImportCsv, writeImportXlsx, writeWorkbook } = require('./src/core/excelOut');
const { num } = require('./src/core/csv');

function arg(name, def) {
  const i = process.argv.indexOf('--' + name);
  return i >= 0 ? process.argv[i + 1] : def;
}

/** Rassemble les fichiers d'entree du dossier selon les inputs declares par le transporteur. */
function gatherFiles(carrier, dir) {
  const files = fs.readdirSync(dir);
  const out = {};
  for (const inp of carrier.inputs) {
    const exts = inp.accept.split(',').map((e) => e.trim().toLowerCase());
    let matched = files.filter((f) => exts.some((e) => f.toLowerCase().endsWith(e)));
    if (inp.match) matched = matched.filter((f) => f.includes(inp.match));
    out[inp.key] = matched.map((f) => path.join(dir, f));
  }
  return out;
}

/** Auto-controle : compare l'import genere a un export ERP de reference (multiset). */
function selfCheck(importRows, refCsv) {
  const text = fs.readFileSync(refCsv).toString('latin1');
  const rows = text.split(/\r?\n/).filter((l) => l.length).slice(1).map((l) => l.split(';'));
  const f = (s) => Math.round(num(s) * 100) / 100;
  const sigRef = (r) => [r[5].trim(), r[9].trim(), Math.round(num(r[10])), f(r[11]), f(r[14]), f(r[15]), f(r[16]), f(r[17]), f(r[18]), f(r[19]), f(r[20])].join('|');
  const g = (v) => (typeof v === 'number' && v !== '' ? Math.round(v * 100) / 100 : 0);
  const sigGen = (o) => [o.Tracking, o.Zone, Math.round(o.NbrColis), Math.round(o.Poids * 100) / 100, g(o.DroitsTaxes), g(o.Assurance), g(o.ZonesEloignees), g(o.ColisVolumineux), g(o.Adresses), g(o.Fret), g(o.PlusValueB2C)].join('|');
  const count = (arr) => arr.reduce((m, s) => m.set(s, (m.get(s) || 0) + 1), new Map());
  const cr = count(rows.filter((r) => r.length > 20).map(sigRef));
  const cg = count(importRows.map(sigGen));
  let ok = 0;
  for (const [s, n] of cg) ok += Math.min(n, cr.get(s) || 0);
  const totRef = [...cr.values()].reduce((a, b) => a + b, 0);
  return { ok, totRef, totGen: importRows.length };
}

async function main() {
  const carrierId = process.argv[2];
  if (!carrierId || carrierId.startsWith('--')) {
    console.log('Transporteurs :', registry.list().map((c) => `${c.id}(${c.status})`).join(', '));
    console.log('Usage : node cli.js <transporteur> --input <dossier> [--out <dossier>] [--check <ref.csv>]');
    return;
  }
  const carrier = registry.get(carrierId);
  const dir = arg('input', '.');
  const files = gatherFiles(carrier, dir);
  console.log(`== ${carrier.name} == fichiers :`, Object.fromEntries(Object.entries(files).map(([k, v]) => [k, v.length])));

  const res = carrier.process(files);
  const totalHt = res.posteKeys.reduce((s, k) => s + (res.controle[k] || 0), 0);
  console.log(`\n${res.importRows.length} lignes d'import.  Totaux postes :`);
  for (const k of res.posteKeys) console.log(`   ${k.padEnd(18)}= ${Math.round((res.controle[k] || 0) * 100) / 100}`);
  console.log(`   ${'TOTAL HT'.padEnd(18)}= ${Math.round(totalHt * 100) / 100}`);
  console.log(`\n${res.warnings.length} alerte(s) reclassement, ${res.alerts.length} alerte(s) a verifier, ${(res.infos || []).length} info(s) affret/navette (hors grille).`);

  const ref = arg('check');
  if (ref) {
    const c = selfCheck(res.importRows, ref);
    console.log(`\n== AUTO-CONTROLE : ${c.ok}/${c.totRef} lignes identiques (genere ${c.totGen}) ==`);
  }

  const out = arg('out');
  if (out) {
    fs.mkdirSync(out, { recursive: true });
    writeImportCsv(res.importRows, path.join(out, 'import.csv'));
    await writeImportXlsx(res.importRows, path.join(out, 'import.xlsx'));
    await writeWorkbook({ ...res, pdfs: null }, path.join(out, `${carrier.id}_workbook.xlsx`));
    console.log(`\nFichiers ecrits dans ${out}/ : import.csv, import.xlsx, ${carrier.id}_workbook.xlsx`);
  }
}

main().catch((e) => { console.error('ERREUR :', e.message); process.exit(1); });
