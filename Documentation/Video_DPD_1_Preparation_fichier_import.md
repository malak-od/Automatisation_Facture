# Script vidéo — DPD_1_Preparation fichier import.mp4

> Vidéo source : `Transporteurs/DPD/DPD_1_Preparation fichier import.mp4`
> Durée : **878 s (~14 min 38 s)**, sans audio exploitable — transcription 100% visuelle (frames extraites à
> intervalle fixe 1/7s + détection de changement de scène, ~150 frames lues).
> Mois traité dans la vidéo : **avril 2026** (`2026 04`), en comparaison avec le mois précédent (`2026 03`).
> Classeur mensuel manipulé : `2026_04_Facture DPD.xlsx` (onglets **Zoning, Bilan factures, Facture DPD,
> Import ERP, Bilan clients, Tarifs achat DPD, Suppléments DPD** — tous visibles dans la barre d'onglets en
> fin de vidéo).

Cette vidéo couvre uniquement la **préparation du fichier d'import** : récupération/consolidation des
fichiers reçus de DPD, reclassement dans `Facture DPD`, réconciliation avec les factures PDF, extraction de
l'onglet `Import ERP` vers un CSV, et contrôles qualité (poids à 0, tracking). Elle ne montre pas l'import
final dans l'ERP/SI.

---

## Partie 1 — Récupération et consolidation des fichiers source DPD (0:00 – 1:40)

1. (0:00–0:21) Excel vide ouvert (Classeur1), point de départ.
2. (0:24) Ouverture d'un explorateur de fichiers / boîte "Parcourir" Power Query, navigation vers :
   `\\192.168.5.3\Comptabilité La Ruche\$Facturation automatique\1 - Factures transporteurs +
   calc[uls]...\2026\2026 04\DPD` → sélection du sous-dossier **`excel`**.
   - Ce sous-dossier `excel` contient la version **.xlsx** de chaque fichier reçu de DPD (à la racine du
     dossier `DPD` se trouvent aussi des **.csv** portant le même nom — probablement le même contenu dans
     les deux formats, le Power Query utilisé ici travaille sur les `.xlsx`).
3. (0:28–0:40) Liste des fichiers du dossier `excel` : de nombreux fichiers nommés
   **`complément_facture02100_2604_0XXXX-30.04.2026.xlsx`** (un fichier par n° de facture/sous-compte DPD,
   ~11 fichiers visibles) → clic sur le bouton **"Combiner"** de Power Query (Données > À partir d'un
   dossier > Combiner les fichiers, feuille "Sheet1" de chaque classeur).
4. (0:49–1:31) Chargement Power Query : la requête charge et fusionne tous les fichiers du dossier en une
   table unique. Message final : **"3484 lignes chargées, 68 erreurs"**. Un nouvel onglet nommé **"excel"**
   est créé dans le classeur avec le résultat consolidé.
5. (1:28–1:36) Colonnes visibles dans la table consolidée "excel" (en-tête complet, extrait) : `Source.Name`,
   `Type (Slave Export)`, `No facture`, `Date de facture`, `Valable du`, `Valable au`, `No de compte`,
   `Sous-compte/Agence`, `Sous-compte`, `Nom`, ... puis plus loin : `Nom expéditeur`, `Rue expéditeur`,
   `Ville expéditeur`, `Code postal expéditeur`, `Code pays expéditeur`, `Nom destinataire`, `Rue
   destinataire`, ... et des colonnes de totaux : `Total poids`, `Total Valeur déclarée`, `Total Frais de
   douane`, `Total prestations transports`, `Total service`, `Dont Total Participation Sûreté`, `Total
   condition`, `Dont Frais de tenue de compte`, `Dont total frais de gazoil`, `Total Transport Taxable`,
   `Total facture soumis à TVA`, `Total facture HT`, `TVA`, `Total TTC`.
6. (1:34–1:36) Plus loin dans les colonnes : `largeur`/`hauteur`/`Poids volumétrique`/`Type poids`/`Colis
   refacturé`/`Prix transport`/`Supplément île et montagne`/`Coût de la VD`/**`Taxe Collection
   Request`**/**`Taxe fixe`**/**`Taxe Consolidation`**/**`Indexation gasoil`**/`Indexation kérosène`/
   `Supplément prédict...` — ce sont les colonnes brutes de surcharge DPD par colis (à ne pas confondre avec
   les "Total ..." qui sont des totaux de facture).
7. (1:38) Clic droit sur l'en-tête de la colonne A `Source.Name` → menu contextuel "Supprimer" (nettoyage de
   la colonne technique ajoutée par le "Combiner" Power Query, avant collage dans le classeur mensuel).

## Partie 2 — Consultation du classeur de référence et comparaison mois précédent (1:52 – 2:20)

8. (1:52–2:00) Navigation entre plusieurs classeurs Excel ouverts (visible via la barre des tâches/onglets
   de fenêtres) : le classeur mensuel en cours et un classeur de référence du mois précédent, probablement
   pour vérifier que la structure de colonnes n'a pas changé d'un mois sur l'autre.
9. (2:08–2:20) Consultation de l'onglet **"Bilan factures"** : un **TCD** (tableau croisé dynamique) avec
   en étiquettes de lignes les comptes/sous-comptes et en valeurs **"Somme de Total GO"** / **"Somme de
   Total hors GO"**. À côté, un tableau manuel intitulé **"Factures pdf"** avec les colonnes : n° de
   facture, Montant, Ecart — sert à comparer le montant HT calculé (TCD) au montant HT lu sur chaque facture
   PDF réelle de DPD.

## Partie 3 — Structure de l'onglet Facture DPD et formule "Frais de dossier" (2:20 – 3:51)

10. (2:20–3:14) Consultation de l'onglet **"Facture DPD"**. Colonnes de postes calculés visibles (partie
    gauche, en-têtes) : **Droits et taxes, Assurance, Zones éloignées, Total hors GO, Total GO, Client,
    Colis (volumineux), Adresses, Frêt, BtoC, Retour, Gazole, Frais dossier [modifié]** — puis à droite les
    colonnes brutes recollées depuis la table consolidée (No facture, Date de facture, ..., Supplément île
    et montagne, Coût de la VD, Taxe Collection Request, Taxe fixe, Taxe Consolidation, Indexation gasoil,
    Indexation kérosène, Supplément prédict...).
11. (3:14) **Barre de formule sur la cellule N3085** (colonne "Frais de dossier modifié") :
    `=ARRONDI.SUP($DI$1/NB(M:M);2)` — confirme la lecture faite en amont sur le classeur : le total fixe
    de frais de dossier (`$DI$1`) est réparti par arrondi supérieur (2 décimales) sur le nombre de lignes
    de la colonne M (`NB(M:M)`), donc chaque ligne facturée reçoit une part égale du total, arrondie au
    centime supérieur.
12. (3:30–3:51) Poursuite de la revue colonne par colonne de l'onglet Facture DPD (défilement horizontal),
    confirmant la présence de toutes les colonnes brutes DPD listées en partie 1.

## Partie 4 — Facture PDF DPD réelle : taux Indexation gasoil (4:04 – 4:40)

13. (4:04–4:15) Ouverture d'un PDF de facture DPD réelle (probablement une pièce jointe du mail DPD, hors
    Excel). Ligne visible : **"Indexation gasoil : 21,18 % de 2 559,26 EUR = 542,05"** — c'est le taux de
    surcharge carburant (gazole) appliqué par DPD sur la facture concernée pour la période. Autre ligne
    visible : **"Frais de tenue de compte : 20,00"**.
14. (4:15–4:40) Retour côté classeur : poursuite de la réconciliation dans "Bilan factures" (voir partie 5).

## Partie 5 — Réconciliation Bilan factures vs PDF (4:40 – 5:50)

15. (4:40–5:50) Boucle de saisie manuelle dans le tableau "Factures pdf" de l'onglet Bilan factures :
    l'utilisateur ouvre successivement plusieurs PDF de factures DPD réelles (une fenêtre PDF par
    compte/sous-compte), relève le montant HT total de chaque facture, puis le saisit dans la colonne
    "Montant" du tableau de contrôle en face du TCD (colonne "Somme de Total GO"). La colonne "Ecart"
    calcule la différence entre le TCD et la saisie manuelle — sert de garde-fou pour valider que le
    reclassement des colonnes brutes vers les postes (Facture DPD) reproduit bien les montants facturés
    réels par DPD.

## Partie 6 — Onglet Import ERP : structure et formules (5:54 – 7:00)

16. (5:54–6:20) Passage à l'onglet **"Import ERP"**. Structure en une ligne par colis (contrairement à
    Facture DPD qui semble mélanger lignes de détail), avec les mêmes intitulés de postes que Facture DPD
    en en-tête.
17. (6:20) **Barre de formule visible** sur une cellule de la colonne "Frêt"/T :
    `=ARRONDI('Facture DPD'!I3+'Facture DPD'!K3;2)` — le poste "Frêt" de Import ERP est calculé comme la
    somme arrondie (2 décimales) des colonnes I ("Frêt") et K ("Retour") de l'onglet Facture DPD. Confirme
    que Import ERP reconstruit ses colonnes par référence directe aux colonnes de postes de Facture DPD
    (pas aux colonnes brutes).
18. (6:35–7:00) Message d'alerte visible en rouge dans le classeur (cellule/note) : **"ATTENTION AUX POIDS A
    0"** et **"Format tracking à modifier"** — signalent deux contrôles qualité obligatoires avant export
    (traités en partie 8 et 9).

## Partie 7 — Fichier CSV d'import ERP final (7:56 – 8:50)

19. (7:56–8:30) Navigation vers le dossier `$Facturation automatique\2 - Fichiers csv import\2026\2026 04\`
    (répertoire distinct du dossier "1 - Factures transporteurs" utilisé en partie 1) et ouverture du
    fichier **`2026_04_DPD_Import.csv`**. C'est le fichier final destiné à l'import ERP, obtenu en collant
    (valeurs) le contenu de l'onglet "Import ERP" du classeur `2026_04_Facture DPD.xlsx` dans un classeur
    CSV autonome.
20. (8:30–8:50) Colonnes du CSV final (en-tête) : `Transport, Date valid[ité], Réf.1, Réf.2, Id client, N°
    Tracking, Nom, E/P, Pays, Zone, Nbr Colis, Poids, mode env[oi], TVA, Droits et [taxes], Assuranc[e],
    Zones élo[ignées], Colis volu[mineux], Adresses, Frêt, plus-valu[e], Gazole, nb colis`. Un filtre est
    appliqué sur la colonne Pays : recherche des lignes avec une valeur suspecte ("0"/inconnu) — 6 lignes
    trouvées à vérifier/corriger.

## Partie 8 — Contrôle "Poids = 0" et fallback en 3 niveaux (8:50 – 12:10)

21. (8:50–9:30) Dans le CSV `2026_04_DPD_Import`, filtre sur la colonne **Poids** : recherche des valeurs à
    **0** (ou vides) → **45 lignes trouvées sur 3420** au total. TVA = 0,2 sur ces lignes (donc ce ne sont
    pas des lignes parasites de sous-total, juste des colis avec poids manquant).
22. (9:30–10:20) **Niveau 1 de fallback poids** : pour deux tracking à poids 0 (`10214000635001` et
    `10214000743318`), l'utilisateur tape une formule **RECHERCHEX** en colonne Poids (L) :
    `=RECHERCHEX(F46;'[2026 03 - Export expéditions_brut.xlsx]exportDemandeExpedition_2026030'!$AP:$AP;
    '[2026 03 - Export expéditions_brut.xlsx]exportDemandeExpedition_2026030'!$AI:$AI;"")`
    — recherche le n° de tracking (colonne F, valeur cherchée) dans la colonne **AP** ("PRO_TRACKING") du
    classeur externe **`2026 03 - Export expéditions_brut.xlsx`** (export du **mois précédent**, onglet
    `exportDemandeExpedition_2026030`), et renvoie la valeur trouvée en face dans la colonne **AI**
    ("INFO_POI[DS]", info poids). Résultat trouvé : `0,4`.
23. (10:20–10:44) **Niveau 2 de fallback** : le même RECHERCHEX est retenté sur le classeur
    **`2026 04 - Export expéditions_brut.xlsx`** (export du **mois courant** cette fois, onglet
    `exportDemandeExpedition_2026040`) — utilisé quand le mois précédent ne donne pas de résultat.
24. (10:44) Une fois le poids trouvé et collé en valeur, l'utilisateur tape **`=ARRONDI.SUP(`** en tête de
    colonne Poids (probablement pour arrondir le poids récupéré au dixième ou à l'entier supérieur — la
    suite de la formule n'est pas visible/lisible).
25. (11:10–11:40) Poursuite du contrôle poids = 0 : d'autres groupes de lignes filtrées à poids 0, y compris
    un groupe avec **Pays = GB** (55 lignes trouvées, TVA = 0 pour celles-ci — logique car export hors UE).
26. (11:40–12:10) **Niveau 3 de fallback** (dernier recours) : pour un colis dont le poids reste introuvable
    dans les exports internes (tracking `250021589312658`, mention "Dimensions indisponibles"), l'utilisateur
    ouvre le portail web **myDPD for Business** (`dpdgroup.com/business/FR/myparcels/track_and_trace/...`)
    et effectue une recherche directe du numéro de colis dans "Suivi de colis" pour lire le poids/les infos
    de livraison directement depuis le site de DPD. Le colis affiche "0,14 kg", statut "Livré".

## Partie 9 — Contrôle qualité final sur l'onglet Facture DPD (12:15 – 14:38)

27. (12:15–13:30) Retour dans le classeur `2026_04_Facture DPD.xlsx`, onglet **Facture DPD** : défilement
    horizontal pour vérifier les colonnes de postes calculés une dernière fois — `Droits et taxes, Assu[rance],
    Zones élo[ignées], Total hors GO, Total GO, Clien[t], Coli[s], Adresses, Frêt, Bto[C], Retour, Gazole,
    Frais dossier [modifié]` avec des valeurs par ligne, ex. lignes 916-917 : Total hors GO = 15,03 €, Total
    GO = 15,86 €, Frêt = 5,03 €, Gazole = 0,83 €, **Frais dossier = 10,00 €** (ligne différente : 0,03 € /
    0,25 € pour d'autres lignes — confirme que le "Frais de dossier modifié" varie ligne à ligne selon la
    répartition ARRONDI.SUP vue en partie 3, et non une valeur fixe).
28. (13:30–14:10) Poursuite du défilement horizontal jusqu'aux colonnes d'adresse destinataire : `CP
    destinataire, Code pays destinataire, N° Colis, DPD ID, Votre référence (x3), Date expédition, Nombre
    de colis`, puis `Rue destinataire, Ville destinataire` (ex. "POUILLY EN AUXOIS", CP 21320).
29. (14:22–14:38) Dernière frame : bureau Windows visible avec plusieurs onglets Firefox ouverts en tâche de
    fond, dont **"myDPD for Business"**, **"Connexion | UPS - États-Unis"**, et plusieurs onglets "Système
    d'Information de..." — confirme que la vérification poids/tracking multi-transporteur (DPD, UPS) via
    les portails web est une pratique courante en fin de process, pas seulement un cas isolé pour DPD.

---

## Réponses aux points demandés par l'appelant

1. **Fichier(s) source(s) reçus de DPD** : plusieurs fichiers **`complément_facture02100_2604_0XXXX-
   30.04.2026`** (un par n° de facture/sous-compte DPD, ~11 fichiers pour ce mois), présents en **deux
   formats dans le dossier mensuel** : `.csv` à la racine du dossier `DPD/` et `.xlsx` dans un sous-dossier
   `DPD/excel/`. Le classeur mensuel utilise Power Query **"À partir d'un dossier" → Combiner** sur le
   sous-dossier `excel` (les `.xlsx`), ce qui charge et fusionne automatiquement tous les fichiers en une
   seule table (3484 lignes / 68 erreurs pour ce mois), avec une colonne technique `Source.Name` à
   supprimer après coup.
2. **Étapes manuelles non déductibles des formules seules** :
   - Suppression de la colonne technique `Source.Name` après le Combiner Power Query.
   - Réconciliation manuelle (Bilan factures) : ouverture de chaque PDF de facture DPD réelle et saisie du
     montant HT dans le tableau de contrôle, comparé au TCD — sert de garde-fou, pas de calcul automatique.
   - Contrôle et correction des lignes à **Poids = 0** avec fallback en 3 niveaux : (1) RECHERCHEX dans
     l'export brut expéditions du **mois précédent**, (2) RECHERCHEX dans l'export brut expéditions du
     **mois courant**, (3) consultation manuelle du portail **myDPD for Business** (tracking en ligne) en
     dernier recours. Ce pattern est le même que celui déjà documenté pour GLS (mémoire :
     `gls_poids_export_brut_fallback.md`), donc à généraliser côté carrier DPD.
   - Contrôle du **format du numéro de tracking** ("Format tracking à modifier" — notation scientifique du
     DPD ID à corriger, N° Colis à privilégier) — mentionné en alerte mais la correction elle-même n'est
     pas montrée en détail dans cette vidéo.
   - Filtre sur colonne Pays pour repérer les lignes à zone/pays incohérent (valeur "0"/inconnu) avant
     export.
3. **"Frais de dossier modifié" (colonne N)** : confirmé en formule dans la vidéo —
   `=ARRONDI.SUP($DI$1/NB(M:M);2)`. `$DI$1` est un total fixe de frais de dossier (probablement le montant
   facturé par DPD au titre des frais de dossier pour la période, visible ailleurs dans le classeur ou sur
   la facture PDF), réparti à parts égales (arrondies au centime supérieur) sur toutes les lignes de la
   colonne M. Les valeurs observées dans l'onglet Facture DPD confirment que ce montant varie ligne à ligne
   (10,00 € pour certaines lignes, 0,03 € / 0,25 € / 0,89 € pour d'autres) — cohérent avec une répartition
   par compte/sous-compte plutôt qu'un total unique pour tout le classeur (**à vérifier** : `$DI$1` semble
   référencer une cellule différente selon le bloc de lignes, donc probablement un total par compte DPD
   et non un total classeur entier — la vidéo ne montre pas explicitement la cellule DI1 ni sa formule).
4. **Onglets Zoning / Bilan factures / Bilan clients / Tarifs achat DPD / Suppléments DPD** :
   - **Bilan factures** : contient le TCD de réconciliation (Total GO / Total hors GO par compte) et le
     tableau de contrôle manuel "Factures pdf" (montant PDF vs calcul, écart) — voir partie 5.
   - **Zoning, Bilan clients, Tarifs achat DPD, Suppléments DPD** : onglets visibles dans la barre d'onglets
     du classeur en fin de vidéo (partie 9, capture des noms d'onglets), mais **leur contenu/formules ne
     sont pas ouverts ni montrés** dans cette vidéo — probablement utilisés en amont (tarifs d'achat DPD
     pour calcul de marge, zones tarifaires, liste clients) mais leur usage précis reste à vérifier dans une
     autre vidéo ou en lisant directement leurs formules dans le classeur `2026_06_Facture DPD.xlsx`.
5. **Taxe Gasoil / Gazole pour DPD** :
   - Poste "Gazole" dans Facture DPD, alimenté par la colonne brute **"Indexation gasoil"** de la table
     consolidée DPD (colonne visible dans l'export mais formule de reclassement exacte non lisible à
     l'écran).
   - Taux observé sur une facture PDF réelle DPD (avril 2026) : **21,18 % de 2 559,26 € = 542,05 €**
     — confirme que le taux d'indexation gasoil DPD est variable mois à mois et publié par DPD directement
     sur chaque facture (comme un pourcentage appliqué au montant transport, pas un montant fixe par colis).
   - Distinction **"Total GO" vs "Total hors GO"** dans Facture DPD et le TCD de Bilan factures : GO =
     Gazole/Gasoil, donc "Total GO" = montant total incluant la taxe gasoil, "Total hors GO" = montant hors
     cette taxe. Utilisé pour isoler l'impact de la surcharge carburant dans la réconciliation.

---

## Points ambigus / illisibles à faire confirmer par le pôle transport

- **Cellule `$DI$1`** (numérateur de la formule Frais de dossier modifié) : sa valeur, son origine (montant
  facturé par DPD ?) et si elle est unique pour tout le classeur ou redéfinie par bloc de lignes/compte
  ne sont pas visibles dans cette vidéo — les valeurs de Frais de dossier observées varient trop pour être
  un total classeur unique divisé par toutes les lignes ; probablement une cellule différente par bloc de
  lignes (donc par compte/sous-compte DPD). À vérifier directement dans le classeur.
- **Mapping exact colonnes brutes → postes** de l'onglet Facture DPD (ex. quelle(s) colonne(s) brute(s)
  alimentent "Assurance", "Adresses", "BtoC", "Retour", "Colis") : les en-têtes sont lisibles mais aucune
  formule de reclassement (autre que Frais de dossier et le RECHERCHEX poids) n'a été affichée dans la
  barre de formule pendant la vidéo — à confirmer en lisant directement les formules du classeur
  `2026_06_Facture DPD.xlsx`.
- **Formule complète `=ARRONDI.SUP(` tapée en partie 8, étape 24** : la vidéo montre le début de la saisie
  de la formule mais la frame suivante disponible ne montre plus la formule en cours de frappe (changement
  de scène / clic ailleurs) — le second argument (nombre de décimales) et la référence exacte ne sont pas
  confirmés à l'écran.
- **Contenu détaillé des onglets Zoning, Bilan clients, Tarifs achat DPD, Suppléments DPD** : seuls les noms
  d'onglets sont visibles (barre d'onglets en bas de l'écran, fin de vidéo) ; aucun n'est ouvert/consulté
  pendant cette vidéo — leur rôle exact dans le calcul (zones tarifaires, tarifs d'achat pour la marge,
  liste des clients/comptes, suppléments spécifiques) reste à confirmer, éventuellement dans une autre
  vidéo de la série DPD ou en lisant directement leurs formules.
- **Format d'enregistrement final du CSV `2026_04_DPD_Import`** : le fichier est déjà ouvert avec l'extension
  `.csv` au début de la partie 7, mais l'étape de sauvegarde/export proprement dite (choix du séparateur,
  encodage) n'est pas filmée dans cette vidéo — à confirmer que le séparateur est bien `;` comme pour les
  fichiers `complément_facture` sources.
- **Ligne "Type (Slave Export)"** et signification exacte des lignes de sous-total à exclure : mentionnée
  dans une analyse antérieure (`Analyse des videos de process - DPD.md`, autre série de vidéos) mais pas
  revue explicitement dans cette vidéo-ci ; le nettoyage des sous-totaux n'apparaît pas clairement filmé
  ici (la consolidation Power Query semble directement produire une table exploitable avec seulement des
  "erreurs" comptabilisées, 68 sur 3484 lignes) — à clarifier si ces 68 erreurs correspondent aux lignes de
  sous-total ou à autre chose (ex. lignes vides, format de date).
- Un document existant (`Documentation/Analyse des videos de process - DPD.md`) décrit une série différente
  de **5 vidéos** (`Process Facturation - Facture DPD_1..5.mp4`, mois mai 2025, dossier `DPD` a priori sans
  fichiers de données à l'époque) avec des détails complémentaires (réception par email, décompression
  d'archives, règle spéciale client "2SHOP" sans gazole en coût de revient). Cette vidéo-ci (`DPD_1_
  Preparation fichier import.mp4`, mois avril 2026) semble être une version plus récente/différente du
  même process, avec le dossier réseau déjà rempli de fichiers `.csv`/`.xlsx` (pas de mail/décompression
  montrés). **À confirmer avec le pôle transport lequel des deux enregistrements reflète le process actuel**
  (réception par email + décompression vs dossier réseau déjà préparé), ou si les deux coexistent selon le
  mois/l'expéditeur DPD.
