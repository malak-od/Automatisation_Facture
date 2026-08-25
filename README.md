# Automatisation Facturation Transporteurs — Documentation

*Vue d'ensemble du projet et de son déroulement, transporteur par transporteur.*

---

## 1. Le projet en une phrase

Chaque mois, chaque transporteur envoie ses données de facturation brutes (CSV/Excel/PDF) ; le projet **traduit ces données hétérogènes en un fichier d'import ERP standardisé** (23 colonnes, une ligne par expédition, ventilée en 8 postes de charge), **réconcilié avec les factures PDF officielles**, pour permettre la refacturation aux clients finaux.

## 2. Le principe général

Pour chaque transporteur, le processus suit le même schéma :

1. **Entrée** : fichier(s) brut(s) du transporteur (export CSV/XLSX du portail, ou facture PDF), propres au format de chaque transporteur — aucun n'a le même schéma de colonnes.
2. **Reclassement** : chaque ligne/charge brute est reclassée dans l'un des **8 postes ERP standard** : Frêt, Zones éloignées, Colis volumineux, Adresses, Assurance, Droits et taxes, Plus-value BtoC, Taxe gazole.
3. **Sortie** :
   - un **classeur Excel** — clone fidèle d'un fichier de référence fait à la main pour ce transporteur (mêmes onglets, mêmes formules, mêmes tableaux croisés dynamiques) ;
   - un **fichier d'import** (CSV + XLSX) au format standard ERP (23 colonnes), prêt à être injecté.
4. **Réconciliation** : le total calculé est comparé au total de la facture PDF officielle du transporteur, pour détecter tout écart avant l'import.

Le fichier d'import ERP (23 colonnes) est le **format pivot commun** à tous les transporteurs :

`Transporteur · Date validité tarif · Réf.1 · Réf.2 · Id client · N° Tracking · Nom · E/P · Pays · Zone · Nbr Colis · Poids · Mode envoi · TVA · Droits et taxes · Assurance · Zones éloignées · Colis volumineux · Adresses · Frêt · Plus-value BtoC · Gazole · Nb Colis`

## 3. Architecture technique

- **Application Node.js** (interface web) : un adaptateur par transporteur, qui lit les fichiers bruts et produit le fichier d'import.
- **Finaliseur Python** (Excel COM, sous Windows) : clone le classeur de référence du transporteur, colle les données du mois traité, recalcule les formules/tableaux croisés dynamiques natifs, produit un classeur fidèle au fichier fait à la main.
- **Réconciliation PDF** : extraction du total officiel depuis la ou les facture(s) PDF fournies, comparaison au montant calculé.

Chaque transporteur a son propre schéma de colonnes, ses propres pièges de format (encodage, décalage de colonnes d'un mois à l'autre, colonnes qui changent de position), et sa propre logique de reclassement — reconstituée par rétro-ingénierie du classeur fait à la main (lecture des formules réelles) puis validée empiriquement contre un ou plusieurs mois réels déjà livrés.

## 4. Méthodologie de reconstitution (par transporteur)

1. **Exploration** des fichiers d'entrée bruts (CSV, XLSX, PDF) et du classeur de référence fait à la main.
2. **Lecture du classeur** en deux passes : les valeurs (résultats) puis les formules (la logique métier réelle).
3. **Décodage des formules** pour reconstituer la règle de reclassement colonne brute → poste ERP.
4. **Prototype** rejouant le pipeline depuis les fichiers bruts, comparé ligne à ligne au fichier déjà livré (référence réelle).
5. **Validation** : le nombre de lignes, les montants par poste et le total réconcilié doivent correspondre exactement (ou avec un écart documenté et expliqué) au fichier de référence.

> Une règle métier n'est considérée fiable que lorsqu'elle est prouvée de deux façons indépendantes : lecture de la formule Excel **et** reproduction exacte de la sortie réelle.

## 5. Pièges récurrents (tous transporteurs confondus)

- **Encodage** : la plupart des exports transporteurs sont en latin-1, pas UTF-8.
- **Décalage de colonnes** : certains transporteurs changent l'ordre/le nombre de colonnes d'un mois à l'autre — le reclassement doit se faire par **nom de colonne**, jamais par position fixe, sauf exception documentée (ex. UPS, export sans en-tête).
- **Tableaux croisés dynamiques (PivotTable)** : le cache source d'un TCD cloné depuis un mois précédent reste souvent figé sur l'ancienne plage de données — doit être redirigé vers la vraie plage du mois traité à chaque génération.
- **Items de cache obsolètes** : un TCD peut réafficher un item (catégorie, tracking) d'un mois antérieur qui n'existe plus dans les données actuelles — à purger explicitement.
- **Poids manquant** : plusieurs transporteurs n'indiquent pas toujours le poids réel dans leur export ; un repli sur un export WMS partagé (mois courant ou précédent) est utilisé en dernier recours.
- **E/P (Entreprise/Particulier)** : rarement fiable directement depuis l'export transporteur — résolu via le même export WMS partagé quand disponible.
- **Lignes sans tracking ni référence** : à exclure systématiquement (lignes de synthèse, ajustements sans expédition associée) — jamais de simple filtre "montant nul" seul, le critère d'identification prime.

## 6. Les transporteurs

Statut : tous les transporteurs listés sont opérationnels (`ready`), sauf mention contraire.

### Kuehne+Nagel
- **Entrée** : 2 CSV (facture transport + facture « événements »), 2 PDF, taxe gazole lue dans un fichier Excel fourni par le transporteur.
- **Reclassement** : ~120 colonnes de détail tarifaire regroupées en 8 postes ; granularité 1 ligne CSV = 1 ligne import (pas de regroupement par tracking).
- **Piège clé** : les lignes de synthèse gazole ont un tracking vide — seul un tracking vide (pas le libellé de la ligne) doit servir de critère d'exclusion.

### Delivengo
- **Entrée** : export du suivi Delivengo (.xls).
- **Particularité** : pas de taxe gazole (frêt fixe).

### DPD
- **Entrée** : fichiers « complément facture » (CSV/XLSX, un par facture).
- **Piège clé** : les colonnes se décalent d'un mois à l'autre — reclassement exclusivement par nom de colonne. Repli automatique sur un export d'expéditions partagé si le poids brut est absent.

### GLS
- **Entrée** : export BCF (CSV), facture PDF pour réconciliation.
- **Reclassement** : regroupement par numéro de colis, catégorie via une table de correspondance dédiée.

### Geodis
- **Entrée** : facture Geodis (XLSX/CSV), PDF optionnel.
- **Piège clé** : colonnes qui se décalent d'un mois à l'autre (même famille que DPD) — reclassement par nom de colonne uniquement.

### Mondial Relay
- **Entrée** : fichiers « Annexe » (CSV, un par facture/dossier), PDF pour réconciliation.
- **Particularité** : une facture dédiée à la collecte (transport aller) génère un écart normal et documenté face au calcul par ligne.

### Lettres (Suivie / Prépa)
- **Entrée** : export d'expéditions partagé, filtré par groupe de transporteur.
- **Particularité** : frêt fixe par groupe, pas de taxe gazole ; les expéditions pas encore parties sont exclues.

### BLS
- **Entrée** : facture(s) PDF (source unique), export d'affrètement (CSV) pour compléter Id client/Nb colis/Poids.
- **Particularité** : trajets internes ("navette") entre deux sites, refacturés au client avec leur montant réel.

### Chronopost
- **Entrée** : 2 fichiers Excel bruts (un par sous-compte), PDF pour réconciliation.
- **Piège clé** : la taxe gazole est facturée en lignes forfaitaires séparées, mises en pool par facture puis redistribuées au prorata du frêt sur les lignes normales.

### Colissimo
- **Entrée** : 2 CSV (prestations au colis + frais de douane), PDF, fichier d'import du mois précédent (optionnel, pour compléter les charges de douane isolées).
- **Particularité** : la taxe gazole est déjà calculée et fournie directement par le transporteur dans le CSV brut.

### TNT
- **Entrée** : 1 fichier Excel brut (feuille détail facture), PDF pour réconciliation.
- **Piège clé** : une ligne d'événement peut porter le montant de transport réel de tout un envoi regroupant plusieurs colis — attribution au tracking correct nécessaire.

### FedEx
- **Entrée** : 1 CSV brut par mois (export portail), PDF pour réconciliation.
- **Particularité** : plusieurs suppléments (zones éloignées, colis volumineux) n'apparaissent que dans le texte des PDF, pas dans le CSV — extraits par recherche de mots-clés.

### UPS
- **Entrée** : CSV « Billing » (un par facture, sans en-tête — colonnes résolues par position, cas particulier parmi les transporteurs).
- **Particularité** : deux comptes distincts (standard et contre-remboursement) déterminés depuis le tracking. Les colis retournés à l'expéditeur (préfixe de tracking dédié) restent visibles dans le classeur mais sont exclus du fichier d'import et reportés automatiquement dans un onglet dédié aux demandes d'avoir. Les lignes sans aucune charge facturable sont supprimées.

## 7. Limites connues du projet

- Chaque transporteur a été validé sur un ou plusieurs mois réels, avec un niveau de confiance variable (validation humaine par l'équipe métier vs. comparaison automatique exhaustive contre un fichier déjà livré).
- Certains écarts résiduels mineurs, déjà investigués et documentés, restent volontairement non corrigés lorsqu'ils sont jugés non significatifs (montant négligeable, cas isolé).
- L'application traite un mois à la fois, en local — pas encore d'environnement de production partagé (API/base de données/déploiement centralisé).
