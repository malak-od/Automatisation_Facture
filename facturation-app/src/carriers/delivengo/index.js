// ============================================================================
//  Adaptateur transporteur : DELIVENGO (lettres suivies, LPPAQ)
//  Le classeur final "AAAA_MM_Delivengo_LPPAQ.xlsx" est un CLONE FIDELE du
//  fichier fait a la main (feuilles Pays + Fichier import, formules, couleurs),
//  regenere par le finaliseur Python (Excel COM) a partir de l'export du suivi.
// ============================================================================
const XLSX = require('xlsx');

// positions des colonnes utiles dans l'export du suivi (index 0-based)
const COL = { remise: 7, suivi: 9, nom: 11, pays: 16, statut: 24, poids: 30 };

function process(files) {
  const p = (files.export || [])[0];
  if (!p) throw new Error('Aucun export du suivi Delivengo fourni (.xls).');
  const wb = XLSX.readFile(p);
  const ws = wb.Sheets[wb.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, raw: false }).slice(1).filter((r) => r && r.length && r[COL.suivi]);

  // periode = 1ere date de remise (JJ/MM/AAAA) -> AAAA_MM
  let period = 'export';
  for (const r of rows) {
    const m = String(r[COL.remise] || '').match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
    if (m) { period = `${m[3]}_${m[2]}`; break; }
  }

  // controles simples (le detail/formules sont dans le classeur genere)
  const alerts = [];
  rows.forEach((r, i) => {
    if (!String(r[COL.pays] || '').trim()) alerts.push(`L${i + 2} ${r[COL.suivi]}: pays destinataire manquant`);
    if (!String(r[COL.suivi] || '').trim()) alerts.push(`L${i + 2}: numero de suivi manquant`);
  });

  return {
    lignes: rows.length, period,
    importRows: [], controle: {}, warnings: [], alerts, infos: [], posteKeys: [],
  };
}

module.exports = {
  id: 'delivengo',
  name: 'Delivengo',
  status: 'ready',
  taxeGasoil: 'Pas de taxe gasoil (frêt = 1,0).',
  method: "Export du suivi Delivengo (.xls). Le classeur final (feuille 'Fichier import') est un clone du fichier fait a la main : A=Date remise, F=Statut, G=N° suivi, H=Destinataire, X=Pays ; J/K/M/O/W = formules (table Pays). P (Droits et taxes) et Q (Assurance) ne sont que des colonnes de calcul intermediaire (poids export brut recherchex / poids Delivengo propre /1000, servent a M=MAX(P,Q)) : videes en fin de traitement, comme Statut et Taxe Gasoil si 0.",
  inputs: [
    { key: 'export', label: 'Export du suivi Delivengo (.xls)', accept: '.xls,.xlsx', multiple: false, required: true },
  ],
  outputNaming: { workbook: '{period}_Delivengo_LPPAQ' },
  workbookOnly: true, // pas d'import ERP separe : le classeur (onglet Fichier import) EST le livrable
  finalizer: {
    script: '../automatisation/finaliser_delivengo.py',
    template: '../Transporteurs/Delivengo/2026_06_Delivengo_LPPAQ.xlsx',
    // <template> <sortie> --export <suivi.xls> --brut <brut_mois> <brut_mois-1>
    // Le brut WMS (poids réel) est un fichier partagé mensuel dans automatisation/
    // nommé « AAAA MM - Export...brut...xlsx » ; on prend le mois de la période + le précédent (repli).
    buildArgs: (files, period, appRoot) => {
      const path = require('path');
      const fs = require('fs');
      const brutDir = path.resolve(appRoot, '../automatisation');
      const [y, m] = String(period || '').split('_').map(Number);
      const findBrut = (yy, mm) => {
        if (!Number.isFinite(yy) || !fs.existsSync(brutDir)) return null;
        const tag = `${yy} ${String(mm).padStart(2, '0')}`; // ex. "2026 06"
        const f = fs.readdirSync(brutDir).find((x) => x.startsWith(tag) && /brut/i.test(x) && /\.xlsx?$/i.test(x));
        return f ? path.join(brutDir, f) : null;
      };
      const brut = [findBrut(y, m), findBrut(m === 1 ? y - 1 : y, m === 1 ? 12 : m - 1)].filter(Boolean);
      return ['--export', ...(files.export || []), '--brut', ...brut];
    },
  },
  process,
};
