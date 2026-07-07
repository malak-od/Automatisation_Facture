# Facturation Transporteurs — logiciel multi-transporteurs (Node.js)

Automatise la transformation des documents fournis par les transporteurs en **fichier d'import ERP**, avec un **classeur de contrôle** et une **réconciliation**. Conçu multi-transporteurs : **Kuehne opérationnel**, les 13 autres (issus des docs de process) sont déjà déclarés et prêts à implémenter.

## Installation

```bash
cd facturation-app
npm install
```

## Utilisation — interface web

```bash
npm start
# ouvrir http://localhost:3000
```

1. Choisir le transporteur dans la liste.
2. Déposer les documents qu'il fournit (pour Kuehne : les CSV `FcCSV*.csv`, + les PDF).
3. Cliquer **Générer** → totaux, alertes de validation, et téléchargement de :
   - `classeur_controle.xlsx` (Fichier Kuehne + TCD + Import + Réconciliation),
   - `import.xlsx` (valeurs seules → à déposer dans l'ERP),
   - `import.csv`.

## Utilisation — ligne de commande

```bash
node cli.js kuehne --input ../Kuehne --out ./out --check ../Kuehne/2026_06_Kuehne_Import.csv
```

`--check` compare la sortie à un export ERP de référence (auto-contrôle). Sur juin 2026 : **187/187 lignes identiques**.

## Architecture

```
server.js            API HTTP + sert l'UI
cli.js               même moteur en CLI
public/index.html    interface (choix transporteur → upload → génération)
src/
  registry.js        registre : kuehne (prêt) + 13 transporteurs planifiés
  core/              moteur COMMUN, générique :
    csv.js             lecture latin-1 / ';' / décimale virgule
    reclass.js         reclassement colonnes -> 8 postes (piloté par config)
    zone.js            calcul de zone (lookup dept + concat)
    validate.js        check-list automatique (poids/colis/fret/zones)
    excelOut.js        écriture CSV / XLSX / classeur (exceljs)
    importSchema.js    les 23 colonnes de l'import ERP
  carriers/
    kuehne/            adaptateur Kuehne + config.json (mapping validé)
    _template/         gabarit pour un nouveau transporteur
```

**Principe** : toute la logique métier d'un transporteur est dans son dossier (`config.json` + `index.js`). Le `core/` ne change jamais. 

## Ajouter un transporteur

1. `cp -r src/carriers/_template src/carriers/<id>`
2. Écrire son `config.json` (mapping colonnes → postes) et/ou sa logique dans `index.js`.
3. Dans `src/registry.js`, remplacer l'entrée planifiée par `require('./carriers/<id>')`.

Le transporteur déclare : `name`, `status`, `taxeGasoil`, `method`, `inputs` (documents attendus) et `process(files)`.

## Limite connue : TCD « vivant »

exceljs (comme toute lib JS) **ne crée pas de tableau croisé dynamique natif**. Le classeur Node produit un **TCD statique** (valeurs correctes). Pour un **TCD vivant** (pivot Excel qui se recalcule), utiliser le finaliseur Python fourni séparément (`../automatisation/facturation_kuehne.py --live-tcd`, via Excel COM sous Windows).

## Réconciliation PDF

Pour l'instant, la lecture des PDF (réconciliation automatique du montant taxable) est faite côté Python (`automatisation/`). Portage Node prévu (via `pdf-parse`).
