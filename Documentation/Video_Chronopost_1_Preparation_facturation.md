# Transcription vidéo — Chronopost_1_Preparation fichier facturation.mp4

Durée totale réelle : **15 min 59 s (959,4 s)** — le nom du fichier / la
consigne de départ mentionnait 9 min 19, mais `ffprobe` confirme 959,4 s.
Résolution capture d'écran Windows (1916×1028), Excel (classeur de travail
`2026_04_Facture Chronopost.xlsx`, fichiers sources `facture_chronopost_
<compte>_202604.xlsx`, CSV `2026_04_Chronopost_Import.csv`) + brefs passages
sur l'explorateur de fichiers. Aucun navigateur web/portail n'apparaît dans
cette vidéo (contrairement à ce qui était supposé dans une précédente
analyse — voir section Réconciliation point 8).

Aucune bande son exploitable : toute l'info vient de ce qui est visible à
l'écran (barre de formule, onglets, cellules, filtres, menus contextuels).
Timestamps approximatifs en `mm:ss`, dérivés de 175 frames extraites
(intervalle fixe 6 s + détection de changement de scène, dédupliquées).

Ce document **re-transcrit** la vidéo avec un focus sur la réconciliation
des notes manuscrites du pôle transport face à ce qui est réellement visible
à l'écran. Il complète (et corrige par endroits) la précédente transcription
`Documentation/Analyse des videos de process - Chronopost.md`.

---

## 0. Onglets du classeur de travail (confirmés à l'écran, ordre exact)

Barre d'onglets en bas de `2026_04_Facture Chronopost.xlsx`, dans cet ordre :

**Bibliothèque transporteurs | Catégories | TARIFS | cap à 5% | Contrôle pdf
| Zoning 2shop | Facture Chronopost | TCD poids | TCD | Fichier import |
Bilan clients | Avoir**

Note : l'onglet actif par défaut au tout début est **TARIFS** (surligné en
vert) — donc **TARIFS**, **cap à 5%**, **Bibliothèque transporteurs** et
**Zoning 2shop** ne sont QUE des onglets visibles dans la barre, ils ne
sont **jamais ouverts/cliqués** pendant cette vidéo (aucune frame ne montre
leur contenu). Leur logique interne reste donc non capturée par cette vidéo
(cf. section Points ambigus).

---

## Partie 1 — Ouverture du modèle du mois précédent et navigation fichiers (00:00 – 00:20)

1. (00:00) Écran de démarrage Excel : page "Bonjour" listant les fichiers
   récents. On y voit clairement le chemin réseau des fichiers Chronopost :
   `\\192.168.5.3\Comptabilité La Ruche\$Facturation automatique\
   1 - Factures transporteurs + calculs\2026\2026 03\` pour
   `2026_03_Facture Chronopost.xlsx`, `facture_chronopost_65481903_
   202603.xlsx`, `facture_chronopost_51291303_202603.xlsx` (mois de MARS,
   servant de modèle) et `2026\2026 04\` pour `2026_04_Facture
   Chronopost.xlsx`, `2026_04_Facture UPS.xlsx` etc. (autres transporteurs
   traités le même jour). Confirme le chemin réseau exact du classeur type
   pour tous les transporteurs : `$Facturation automatique\1 - Factures
   transporteurs + calculs\<AAAA>\<AAAA> <MM>\<Transporteur>\`.
2. (00:06) `2026_03_Facture Chronopost.xlsx` s'ouvre (écran de chargement
   noir, transition).
3. (00:10 – 00:14) Bref retour à l'explorateur de fichiers Windows : dossier
   `2026 04` listant tous les sous-dossiers transporteurs (BLS, CEVA,
   **Chronopost**, Colissimo, Delivengo, DHL, DPD, Fedex, Geodis, GLS,
   Kuehne, La Poste, Mondial Relay, TNT, UPS, UPS - 80X7Y5) — confirme que
   Chronopost est traité au sein du même process mensuel que les autres
   transporteurs de ce projet.

## Partie 2 — Vue d'ensemble "Facture Chronopost" : structure et données brutes (00:18 – 01:50)

4. (00:18 – 00:48) Le classeur de travail `2026_04_Facture Chronopost.xlsx`
   est actif sur l'onglet **Facture Chronopost**. En-têtes confirmés en
   ligne 1, colonnes A à S (partie gauche du tableau) :
   **A = ID Clients | B = No Facture | C = Sous-compte | D = Date LT |
   E = Code postal depart | F = Code postal arrivee | G = Pays depart |
   H = Pays arrivee | I = Ref Destinataire | J = Ref Expediteur |
   K = No Groupage tarifaire | L = Numero LT | M = Groupage |
   N = Type prestation | O = TVA | P = Observations | Q = Zone Tarifaire |
   R = Poids | S = Produit**.
   Le tableau contient ~1550 lignes de données consolidées (compteur de
   ligne visible jusqu'à ~1567).
5. (00:48 – 01:10) Défilement dans les lignes de données : valeurs "5Y" en
   colonne S (Produit) pour un grand nombre de lignes "Transport" ; colonne
   L (Numero LT) contient des identifiants de type `XT257517968TS`,
   `XW247350336JF`, `XF140480926FR` (formats différents selon le
   sous-compte/contrat).
6. (01:10 – 01:50) Basculement bref sur les fichiers source bruts (non
   retouchés) `facture_chronopost_51291303_202604.xlsx` et
   `facture_chronopost_65481903_202604.xlsx` — chacun a une feuille unique
   **"Données"**, titre en A1 **"Détail de facture chronopost (Période
   04/2026)"**, A2 **"Compte : 51291303 - LA RUCHE LOGISTIQUE"** (ou
   65481903), en-têtes exacts en **ligne 4** :
   **No Facture | Sous-compte | Date LT | Code postal depart | Code postal
   arrivee | Pays depart | Pays arrivee | Ref Destinataire | Ref Expediteur
   | No Groupage tarifaire | Numero LT | Groupage | Type prestation | TVA |
   Observations | Zone Tarifaire | Poids | Produit | Montant HT | Raison
   sociale** — données à partir de la ligne 5. Le sous-compte "65481903"
   contient bien les lignes des No Facture 13541639 à 13548411 environ
   (mois de mars, dans ce fichier ouvert en avril comme référence croisée).

## Partie 3 — RÉCONCILIATION 1 : libellés exacts "Type prestation" et codes "Numero LT" (01:50 – 04:20)

**C'est ici que les frames confirment/précisent les points 1 et 3 de la
consigne (préfixes CAP/ECO/SUR).**

7. (03:10 – 03:20, frame très nette sur le fichier source brut
   `facture_chronopost_65481903_202604.xlsx`, feuille Données) : la colonne
   **K = "Type prestation"** contient EXACTEMENT les libellés texte
   suivants (contrairement à l'hypothèse "CAP/ECO/SURT" qui sont en fait
   des codes de la colonne "Numero LT", PAS des valeurs de "Type
   prestation") :
   - **"Transport"**
   - **"Zones Difficiles d'accès"**
   - **"Participation Eco-Responsable"**
   - **"Sûreté colis"**
   - **"Surcharge Carburant Routier"** (et vraisemblablement "Surcharge
     Carburant Aérien" par analogie, vu plus loin dans "Facture
     Chronopost")
   Et la colonne **J = "Numero LT"** (dans le fichier source brut ; devient
   colonne **L "Numero LT"** une fois consolidé dans "Facture Chronopost")
   contient, pour CES lignes spéciales (Montant HT forfaitaire, pas de
   colis réel associé), des codes courts au lieu d'un numéro de tracking :
   **ECORN**, **SURTN**, **CAPN2** (vus dans le fichier source brut) et,
   dans le classeur consolidé "Facture Chronopost" (onglet dédié, cf.
   point 8), on voit une gamme plus large : **CAPI1, CAPI2, CAPN1, CAPN2
   (×6 lignes), CAPO1** (Surcharge Carburant, suffixes I=Aérien/N=National-
   Routier/O=Outre-mer ou Otherwise — cf. point 3 des "Points ambigus" de
   la précédente transcription, non résolu ici non plus mais motif
   confirmé visuellement), **ECORI, ECORN (×6), ECORO** (Participation
   Éco-Responsable), **SURTI, SURTN (×6), SURTO** (Sûreté colis).
8. (03:20 – 03:40, onglet **Facture Chronopost**, lignes 2 à 27) : table
   récapitulative confirmée ligne par ligne — colonne N "Type prestation"
   / colonne L "Numero LT" / colonne T "Montant HT" :
   | L (Numero LT) | N (Type prestation) | O (TVA) | T (Montant HT) |
   |---|---|---|---|
   | CAPI1 | Surcharge Carburant Aérien | I | 86,56 |
   | CAPI2 | Surcharge Carburant Routier | I | 39,40 |
   | CAPN1 | Surcharge Carburant Aérien | N | 837,31 |
   | CAPN2 (×6) | Surcharge Carburant Routier | N | 37,05 / 3,23 / 142,99 / 466,79 / 1,05 / 33,92 |
   | CAPO1 | Surcharge Carburant Aérien | O | 10,95 |
   | ECORI | Participation Eco- | I | 3,50 |
   | ECORN (×6) | Participation Eco- | N | 11,67 / 0,51 / 16,83 / 39,80 / 0,10 / 14,90 |
   | ECORO | Participation Eco- | O | 0,10 |
   | SURTI | Sûreté colis | I | 14,00 |
   | SURTN (×6) | Sûreté colis | N | 19,45 / 0,85 / 28,05 / 159,20 / 0,40 / 59,60 |
   | SURTO | Sûreté colis | O | 0,40 |
   Puis à partir de la ligne 28, No Groupage tarifaire = "999999" et Numero
   LT redevient un vrai tracking (ex. `XF140480926FR`), Type prestation =
   "Transport" ou "Supp Retour Expediteur Inter[national]" ou "Zones
   Difficiles d'accès".

   **➜ RÉCONCILIATION POINT 1 : le libellé utilisateur "CAP = taxe gazole,
   ECO = taxe ECP, SURT = taxe sûreté" est CONFIRMÉ dans son principe mais
   imprécis dans le détail. Il n'y a PAS de colonne "Type prestation"
   contenant littéralement "CAP"/"ECO"/"SUR" : ce sont des PRÉFIXES de la
   colonne "Numero LT" (L) qui servent de clé courte, alors que la colonne
   "Type prestation" (N) contient le libellé complet en français
   ("Surcharge Carburant Aérien/Routier", "Participation Eco-Responsable",
   "Sûreté colis"). Les deux colonnes sont redondantes/cohérentes pour ces
   lignes forfaitaires. Le motif de détection par préfixe est donc bien
   `CAP*` / `ECO*` / `SUR*` sur la colonne L "Numero LT" (ce préfixe est
   plus fiable pour un pattern-matching automatique que le texte complet
   de "Type prestation", qui varie légèrement — "Surcharge Carburant
   Aérien" vs "Routier"). "ECP" mentionné par l'utilisateur = coquille pour
   "Eco-Responsable" probablement, ou abréviation orale.**

## Partie 4 — RÉCONCILIATION 2/3 : colonnes AE/AF/AG, tri par colonne L, formules exactes (03:40 – 06:00)

9. (03:40 – 04:20, onglet Facture Chronopost, défilement horizontal vers
   les colonnes K à AG) : après la colonne S "Produit", on trouve :
   **T = Montant HT | U = Raison sociale | V = (colonne étroite, non
   labellée nettement, contenu "R" partiellement visible) | W = Zoning
   2shop (%) | X = Gazole % | Y = frêt | Z = sureté + eco | AA = gazole |
   AB = hors gazole | AC = Catégories | AD = Total avec G(azole) |
   AE = Total | AF = (zone de sous-totaux, libellés en colonne AE en face)
   | AG = (zone taux taxe gasoil)**.
10. (03:40) Colonne **X "Gazole %"** = **15,15%** pour TOUTES les lignes
    visibles (valeur unique du mois, reportée automatiquement sur chaque
    ligne — probablement une référence à une cellule fixe du classeur,
    formule non capturée avec certitude mais valeur constante confirmée
    visuellement sur ~30 lignes).
11. (03:45 – 04:00) Colonne **AC "Catégories"** = résultat du reclassement
    Type prestation → poste : "Gazole" (pour toutes les lignes CAP*), puis
    à partir de la ligne 12 "Frêt" pour les lignes ECOR* et SURT* (fait
    inattendu : dans ce classeur, ECO et SUR sont classées "Frêt" dans la
    colonne "Catégories" à ce stade — **incohérent avec le libellé "eco"/
    "sureté" utilisé plus loin en colonne AF/AE**, cf. point ambigu en fin
    de doc).
12. (04:00 – 04:20, cellule **AF1** sélectionnée) : valeur affichée
    `13 003,84` puis `13 064,81` (deux captures à des instants différents
    du calcul, le classeur recalculait) — un **grand total en haut de la
    zone de sous-totaux**, probablement `=SOMME(AF2:AF5)` ou équivalent
    (formule exacte non capturée pour AF1 lui-même).
13. (04:12 – 04:20, cellule **AF2** sélectionnée, libellé AE2 = "Frêt") :
    barre de formule affiche `=SOMME(T22:T30)` PUIS (recalcul/correction
    visible sur une frame suivante très proche) `=SOMME(T20:T27)` — deux
    valeurs de plage légèrement différentes captées à des instants
    successifs de la construction de la formule (l'utilisateur est en
    train d'AJUSTER la plage manuellement, probablement pour caler sur les
    bonnes lignes "Transport" après avoir trié/regroupé par colonne L).
    Résultat affiché dans une frame proche : AF2 = **10 971,73** (poste
    "Frêt").
14. (04:20, cellule **AF3** sélectionnée, libellé AE3 = "eco") : barre de
    formule **`=SOMME(T13:T21)`** — CONFIRMÉ avec certitude : la plage
    T13:T21 correspond exactement aux 9 lignes ECORI/ECORN(×6)/ECORO
    identifiées au point 8 ci-dessus. Résultat = **117,36**.
15. (04:12, variante antérieure vue sur une frame précédente) : AF3 avait
    été tapé une première fois comme `=SOMME(T2:T12)` (plage englobant
    alors TOUTES les lignes CAP* + une partie ECO*, visiblement une
    ébauche/erreur initiale) avant d'être corrigée à `T13:T21`.
16. (04:20) Cellule **AE4** = "sureté" (libellé), **AE5** = "Gazole"
    (libellé) — confirmant l'existence de labels "Frêt" / "eco" / "sureté"
    / "Gazole" en colonne AE, en face des sous-totaux AF2 à AF5.
17. Cellule **AE row plus bas** = "Gestion" et "Gazole réel" / "(avant
    remise)" apparaissent également dans la même zone (AE6/AE7 environ) —
    présence d'un concept de **"Gazole réel" vs "Gazole avant remise"**
    confirmée visuellement (correspond au point "Remise gazole" de
    l'ancienne transcription, toujours pas totalement détaillé en formule
    ici).

   **➜ RÉCONCILIATION POINT 2/3 : le principe "colonne AF = chaque label
   fait avec un tri par colonne L" est CONFIRMÉ : les sous-totaux AF2:AF5
   sont bien des `=SOMME(T<début>:T<fin>)` sur des PLAGES DE LIGNES
   CONTIGUËS, et ces plages ne sont cohérentes QUE si les lignes ont été
   préalablement triées/regroupées par la colonne L "Numero LT" (pour que
   toutes les CAP* soient contiguës, puis toutes les ECO*, puis toutes les
   SUR*, puis les Transport). Ce n'est PAS une formule dynamique du type
   `SOMMEPROD` ou `SOMME.SI(L:L;"CAP*";T:T)` — c'est une somme sur une
   plage FIXE tapée à la main après tri visuel, ce qui explique que
   l'utilisateur ait dû CORRIGER la plage en direct (T22:T30 → T20:T27,
   T2:T12 → T13:T21) : le nombre de lignes par catégorie change d'un mois
   à l'autre selon le nombre de sous-comptes/lignes forfaitaires
   facturées, donc la plage doit être réajustée à chaque fois par
   observation visuelle des frontières entre catégories après le tri.**

## Partie 5 — RÉCONCILIATION 3 (suite) : AG8/AG9 taxe gasoil (04:20 – 04:40)

18. (04:20, zone AG8:AH9 visible) :
    - **AG8 = 17,55%**, libellé associé en AH8 = **"routier"**
    - **AG9 = 29,75%**, libellé associé en AH9 = **"aérien"**

    **➜ RÉCONCILIATION POINT 3 (AG8/AG9) : CONFIRMÉ EXACTEMENT.** AG8 =
    taux de taxe gasoil routier Chronopost du mois (17,55 %), AG9 = taux
    aérien (29,75 %). Ces deux valeurs sont bien saisies à la main dans
    ces cellules fixes du classeur (probablement recopiées depuis le SI
    `si.laruche-logistique.fr` — non montré dans cette vidéo, seulement le
    résultat final dans le classeur). Notez que la colonne X "Gazole %"
    (15,15 %, point 10) est DIFFÉRENTE des taux AG8/AG9 (17,55 %/29,75 %)
    — donc au moins 3 taux gazole coexistent dans ce classeur : un taux
    "Gazole %" appliqué ligne à ligne (15,15 %, probablement le taux
    "après remise"/"réel" mixte), et les 2 taux officiels routier/aérien
    (17,55 %/29,75 %) probablement utilisés pour un calcul de contrôle ou
    de redistribution séparé. Le lien exact entre ces 3 taux (formule de
    passage) n'est pas visible dans les frames disponibles.

## Partie 6 — RÉCONCILIATION 3 (colonnes W à AC) et feuille "Catégories" (04:40 – 06:00)

19. (04:40 – 05:00, onglet **Catégories**, liste de validation visible en
    filtre colonne) : les valeurs exactes de la liste déroulante
    "Catégories" sont — **Adresse, Assurance, Colis volumineux, Corse,
    Droits et taxes, Frais facturation, Frêt, Gazole, Zones éloignées**.
    Confirme la table de la précédente transcription (mapping Type
    prestation → poste), mais rien de nouveau vu ici sur le mapping
    ligne-par-ligne détaillé (feuille défilée rapidement, contenu complet
    non capturé net).
20. Concernant l'énoncé **"Étirer les colonnes W à AC"** : d'après les
    en-têtes confirmés au point 9, ces colonnes contiennent : **W = Zoning
    2shop (taux %) | X = Gazole % (15,15%, valeur mensuelle appliquée à
    chaque ligne) | Y = frêt (montant, probablement `=SI(catégorie="Frêt";
    T;0)` ou similaire) | Z = sureté + eco (montant regroupé) |
    AA = gazole (montant, quote-part gazole de la ligne) | AB = hors
    gazole (montant total moins gazole) | AC = Catégories (résultat du
    reclassement Type prestation → poste, via RECHERCHEV/liste sur la
    feuille "Catégories")**. Ce sont donc des colonnes de calcul PAR LIGNE
    (formules) qu'il faut réétirer (recopier vers le bas) chaque mois
    quand le nombre de lignes total change — cohérent avec l'énoncé de
    l'utilisateur, mais le détail exact des formules W/Y/Z/AA/AB n'a pas
    été capturé net dans une frame en gros plan (uniquement les résultats
    numériques et les en-têtes ont pu être lus).
21. (05:00 – 05:40) Retour à l'onglet **Facture Chronopost**, poursuite du
    tri visuel colonne L (ligne 28 à ~1567) : Type prestation redevient
    "Transport" / "Zones Difficiles d'accès" / "Supp Retour Expediteur
    Inter" en flux continu, No Groupage tarifaire = "999999" pour toutes
    ces lignes normales (par opposition à "0" pour les lignes forfaitaires
    CAP/ECO/SUR).

## Partie 7 — Contrôle sur fichier source brut avant consolidation (05:40 – 06:20)

22. (05:47 – 06:20) Un nouveau classeur vide "Classeur1" est ouvert, puis
    des données sont collées dedans (visible via le Presse-papiers
    activé/Ctrl) — copie du contenu de `facture_chronopost_65481903_
    202604.xlsx` (les dernières lignes de ce fichier source, lignes 984 à
    992 environ). On y confirme, dans les toutes dernières lignes du
    fichier source (fin de période) :
    - Ligne "13586828 | 2 | 30/04/2026 | (vide) | ECORN | X | Participation
      Eco-Responsable | N | | XX | 0,000 | X | **16,83**"
    - Ligne "13586828 | 2 | 30/04/2026 | (vide) | SURTN | X | Sûreté colis
      | N | | XX | 0,000 | X | **28,05**"
    - Ligne "13586828 | 2 | 30/04/2026 | (vide) | CAPN2 | X | Surcharge
      Carburant Routier | N | 3,00% | XX | 0,000 | X | **142,99**"
    Ces 3 valeurs (16,83 / 28,05 / 142,99) correspondent numériquement à
    des valeurs déjà vues en colonne T de "Facture Chronopost" (point 8),
    confirmant qu'il s'agit bien des lignes forfaitaires du DERNIER
    sous-compte consolidé dans le classeur.

## Partie 8 — RÉCONCILIATION 4 : Contrôle PDF (05:03 – 05:20, et retour 12:27 – 12:36)

23. (05:03, onglet **Contrôle pdf**) : Tableau Croisé Dynamique (TCD) —
    colonnes **A = Étiquettes de lignes** (No Facture) | **B = Somme de
    Montant HT** | **C = pdf** (saisie manuelle) | **D = écarts**. Valeurs
    lues :
    | No Facture | Somme de Montant HT | pdf | écarts |
    |---|---|---|---|
    | 13580211 | 4 172,60 € | 4 172,60 € (4172,6 tapé) | - € |
    | 13580212 | 9,35 € | 9,35 € | - € |
    | 13580213 | 4 636,78 € | 4 636,78 € (4636,78 tapé) | - € |
    | 13586826 | 1 291,62 € | 1 291,62 € (1291,62 tapé) | - € |
    | 13586827 | 111,47 € | 111,47 € | - € |
    | 13586828 | 4 942,12 € | 4 942,12 € (4942,1 tapé, complété) | - € |
    | (vide) | — | — | - € |
    | **Total général** | **15 163,94 €** | **15 163,94 €** | **- €** |

    **➜ RÉCONCILIATION POINT 4 : CONFIRMÉ EXACTEMENT.** L'onglet "Contrôle
    pdf" (nommé ainsi, pas "Contrôle PDF" avec majuscules — la casse
    exacte de l'onglet visible dans la barre est "**Contrôle pdf**") est
    bien un TCD "Somme de Montant HT" par No Facture (source = "Facture
    Chronopost"), comparé colonne par colonne à une colonne "pdf" saisie
    À LA MAIN pour chaque No Facture (6 valeurs par mois dans cet exemple,
    correspondant aux 6 PDF `CHRONOPOST_*.pdf` réels du dossier projet :
    3 PDF pour le compte 51291303 + 3-4 pour 65481903, cohérent avec les 7
    PDF présents dans `Transporteurs/Chronopost/`). Colonne "écarts" =
    `Somme de Montant HT - pdf`, doit être 0,00 € (`- €` en affichage
    comptable = zéro) pour valider le mois. Le Total général confirmé =
    **15 163,94 €** pour avril 2026 (à ne pas confondre avec le total mai
    2025 ≈ 12 970,60 € mentionné dans la précédente transcription, qui
    concernait un autre mois).

## Partie 9 — Onglet "TCD" : structure, colonnes, et bug #VALEUR!/#N/A (06:06 – 09:00)

24. (06:06 – 06:30, onglet **TCD**) : première vue — TCD avec en-têtes de
    colonnes personnalisées calculées à côté du TCD natif : on y voit des
    valeurs **#VALEUR!** dans plusieurs cellules d'une colonne proche de
    "Total GO" / "Total avec GC" / "Total hors GC" / "ID client".
25. (06:36 – 07:00) Colonnes visibles du TCD (partie gauche) : **Numero
    LT | Produit | TVA | Zone Tarifaire | Catégories** (Adresse, Assurance,
    Colis volumineux, Corse, Droits et taxes, Frais facturation, Frêt,
    Gazole, Zones éloignées listés en filtre) puis, plus à droite, des
    colonnes "zone" / "Transporteur" / "Frêt+CAP+ECO" (nom informel,
    lecture partielle).
26. (07:00 – 07:24) Menu contextuel (clic droit sur le TCD) : options
    **"Actualiser"** (rafraîchissement du TCD) bien visibles dans le menu
    — confirme que le TCD est manuellement rafraîchi après toute
    modification de la source "Facture Chronopost".
27. (07:24 – 08:00) Colonne à droite du TCD (hors zone TCD native, colonne
    de calcul manuelle) : formule confirmée en barre de formule pour une
    cellule de la colonne **S** (libellée "Sureté + eco" dans l'en-tête
    visible) :
    ```
    =RECHERCHEX(E1267;'Facture Chronopost'!L:L;'Facture Chronopost'!Z:Z;"")
    ```
    Recherche la valeur de E1267 (probablement le Numero LT/tracking de la
    ligne du TCD) dans la colonne **L "Numero LT"** de la feuille "Facture
    Chronopost", et renvoie la valeur correspondante de la colonne **Z**
    de cette même feuille (Z = "sureté + eco", cf. point 9). 4e argument
    `""` = valeur si non trouvé.
28. (08:00 – 08:20) Formule confirmée pour la colonne **T** (libellée
    "mode envoi" ou équivalent, à droite de S) :
    ```
    =SI(NB.SI('Bibliothèque transporteurs'!C:C;F58)=0;"inconnu";
    RECHERCHEX(F58;'Bibliothèque transporteurs'!C:C;
    'Bibliothèque transporteurs'!A:A))
    ```
    Recherche F58 (Produit ou Zone Tarifaire de la ligne) dans la colonne
    **C** de la feuille "**Bibliothèque transporteurs**", renvoie la
    colonne **A** de cette même feuille (mode envoi ERP) ; si non trouvé,
    renvoie le texte "inconnu". **Confirme l'existence et le rôle de la
    feuille "Bibliothèque transporteurs" comme table de correspondance
    Produit/Zone Chrono → mode envoi ERP**, MAIS cette feuille elle-même
    n'a jamais été ouverte/affichée dans la vidéo (cf. section 0) — son
    contenu (colonnes A, B, C exactes) reste à documenter séparément.
29. (08:12 – 08:36) **Bug #N/A repéré et corrigé en direct** (confirme
    EXACTEMENT le piège technique signalé par l'utilisateur) :
    - Une colonne du TCD affiche des erreurs `#N/A` en cascade.
    - L'utilisateur clique sur une cellule contenant la valeur "17"
      (Produit, colonne S de "Facture Chronopost") : un triangle vert
      d'avertissement Excel apparaît en haut à gauche de la cellule, avec
      l'info-bulle **"Nombre stocké sous forme de texte"** — cette valeur
      "17" est donc du **TEXTE** et non un **NOMBRE**, ce qui casse la
      correspondance RECHERCHEX/RECHERCHEV avec la colonne "Bibliothèque
      transporteurs" (dont la clé de correspondance C1141/F58 etc.
      attend probablement un nombre).
    - Menu contextuel du triangle d'avertissement : option **"Convertir en
      nombre"** sélectionnée par l'utilisateur.
    - Séquence observée : sélection de toute la colonne **S "Produit"**
      de la feuille **Facture Chronopost**, clic sur l'icône
      d'avertissement, choix "Convertir en nombre" → tous les "17", "44",
      "86" etc. (codes produit Chronopost stockés en texte) deviennent des
      nombres.
    - Retour à l'onglet **TCD** : clic droit → **"Actualiser"** → les
      `#N/A` disparaissent, remplacés par les vrais modes envoi ERP.

    **➜ RÉCONCILIATION POINT 6 (piège #N/A) : CONFIRMÉ EXACTEMENT et
    précisé.** La cause exacte est que la colonne **S "Produit"** de
    l'onglet **Facture Chronopost** (contenant les codes produit Chrono
    comme "17", "44", "86", "2", "1S"...) est parfois collée/importée en
    tant que **texte** (alignement à gauche, triangle vert) au lieu de
    **nombre** (alignement à droite), ce qui casse la recherche
    RECHERCHEX/RECHERCHEV vers "Bibliothèque transporteurs" (colonne clé
    numérique). La correction manuelle : sélectionner toute la colonne S,
    utiliser le bouton d'erreur Excel "Convertir en nombre" (ou
    équivalent Données > Convertir), PUIS retourner sur l'onglet **TCD**
    et faire **clic droit > Actualiser** sur le TCD pour que les formules
    RECHERCHEX recalculent correctement. L'utilisateur a raison sur le
    principe ; la nuance est que ce n'est pas systématiquement "#N/A" à
    l'écran — la vidéo montre `#N/A` dans le TCD ET `#VALEUR!` ailleurs
    (deux symptômes différents possibles de la même cause).
30. (08:36 – 08:52, onglet TCD toujours) : colonnes **Y** et **Z**
    identifiées comme "Formule Test à vérifier" (Y) et "Vérif avec
    ancienne formule" (Z), avec une formule complexe visible dans la barre
    de formule pour une cellule Z :
    ```
    =ARRONDI.SUP(B1267;2)-M1267-L1267-N1267-J1267-K1267-I1267-Q1267
    ```
    Cette formule calcule le **Frêt par différence** : Montant HT total
    (B1267, arrondi au centime supérieur) moins toutes les colonnes de
    postes déjà connus (M, L, N, J, K, I, Q — Gazole, Adresse, Assurance,
    Corse, Droits et taxes, Zones éloignées, etc.) = résiduel = Frêt.
    C'est une méthode de calcul ALTERNATIVE/DE CONTRÔLE au calcul direct
    par somme des lignes "Transport", présente en colonne Z pour vérifier
    la cohérence avec l'ancienne formule (colonne Y "Formule Test").

   **➜ RÉCONCILIATION POINT 6 (colonnes à MAJ) : le libellé utilisateur
   "MAJ colonne R, étirer colonnes A B C S T U V W Y Z" n'a pas pu être
   confirmé colonne par colonne avec certitude à 100% dans les frames
   disponibles (le TCD défile souvent trop vite ou le zoom ne permet pas
   de lire tous les en-têtes de lettre de colonne simultanément), mais le
   PRINCIPE est cohérent avec ce qui est observé : plusieurs colonnes de
   calcul manuel juxtaposées au TCD natif (S = sureté+eco via RECHERCHEX,
   T = mode envoi ERP via RECHERCHEX, Y/Z = formules de contrôle du Frêt)
   doivent être réétirées chaque mois car le nombre de lignes du TCD
   change. Le point "Attention c'est la colonne Y qui reprend le frêt de
   l'onglet import" n'a pas pu être vérifié littéralement (colonne Y vue
   ici = "Formule Test", pas explicitement "reprend le frêt de l'onglet
   import" — à confirmer avec le pôle transport, incohérence possible
   entre onglet "TCD" et onglet "Fichier import" sur le sens de qui
   alimente qui).**

## Partie 10 — Onglet "TCD poids" (aperçu bref)

31. Aucune frame de la vidéo ne montre l'onglet **TCD poids** ouvert avec
    un contenu net et lisible (l'onglet est visible dans la barre mais son
    contenu — "MAJ colonne B, étirer colonne C" selon l'énoncé utilisateur
    — n'a pas pu être confirmé visuellement dans cette vidéo). **Point à
    reconfirmer** (cf. section finale).

## Partie 11 — Onglet "Fichier import" (classeur de travail) : formules et note 2SHOP (09:00 – 10:15)

32. (09:00 – 09:20, onglet **Fichier import** du classeur `.xlsx`, PAS le
    CSV final) : en-têtes de colonnes confirmés, dans l'ordre : **A =
    Transporteur | B = Date validité[ tarif] | C = Réf.1 | D = Réf. 2 |
    E = Id client | F = N° Tracking | G = Nom | H = E/P | I = Pays |
    J = Zone | K = Nbr Coli(s) | L = Poids | M = mode envoi | N = TVA |
    O = Droits et taxe(s) | P = Assurance | Q = Zones éloignées | R = Colis
    volu(mineux) | S = Adress(es) | T = Frêt | U = plus-value B(t?) |
    V = gazole | W = Nb colis**. Une note rouge en haut à droite du tableau
    (colonne W/X, hors zone structurée) affiche : **"2SHOP[?] : ne pas
    mettre de gazole dans les coûts de revient"**.
33. (09:12) Cellule sélectionnée : formule `=TCD!V3` — confirme que
    l'onglet **Fichier import** est bien ALIMENTÉ PAR FORMULE depuis
    l'onglet **TCD** (référence directe `=TCD!<cellule>`), et pas par
    copier-coller de valeurs à ce stade.
34. (09:24 – 09:40) Ligne 1 des données (juste sous les en-têtes) : colonne
    O "TVA" contient une formule visible :
    ```
    =SI(I3="";0,2;RECHERCHE(I3;'Pays TVA'!A:A;'Pays TVA'!B:B))
    ```
    Recherche le Pays (I3) dans la feuille "Pays TVA" ; si Pays vide,
    applique 0,2 (20%) par défaut ; sinon le taux correspondant (ex. 0 pour
    certains pays hors UE).
35. (09:40 – 10:00) Défilement des lignes 2 à 12 de "Fichier import" : ce
    sont EXACTEMENT les lignes CAP*/ECO*/SUR* (les lignes forfaitaires
    identifiées au point 8), avec Transporteur = "inconnu", Pays =
    "inconnu", Zone = "inconnu" pour beaucoup d'entre elles — confirme que
    CES lignes forfaitaires SONT ENCORE PRÉSENTES dans l'onglet classeur
    "Fichier import" à ce stade (avant l'étape finale de nettoyage sur le
    CSV, cf. Partie 13).
36. (10:00 – 10:15) Colonne **J "Zone"** contient des valeurs au format
    `<mode>_<sous-zone>`, ex. **"17_Z2"**, **"6C"** — confirme le format de
    concaténation mode-envoi_zone évoqué par l'utilisateur (point 8 de la
    consigne).
37. (10:15 – 10:15) On voit aussi la mention **"CHRONO_2SHOP"** apparaître
    comme valeur de la colonne Transporteur pour certaines lignes issues du
    sous-compte 65481903 (2ème contrat) — confirme la distinction visuelle
    entre les deux contrats "CHRONOPOST" (standard) et "CHRONO_2SHOP" dans
    le fichier final.

## Partie 12 — CSV de référence du mois précédent (10:15 – 10:35)

38. (10:15 – 10:35) Bref passage sur le fichier `2026_03_Chronopost_
    Import.csv` (mois précédent, ouvert comme référence de structure) : le
    fichier CSV NE CONTIENT PAS de lignes CAP*/ECO*/SUR* — il commence
    directement par des lignes "Transport" avec un vrai N° Tracking en
    colonne F. Confirme que la suppression des lignes forfaitaires a bien
    lieu à un moment donné entre l'onglet classeur "Fichier import"
    (Partie 11, où elles sont encore présentes) et le CSV final exporté.

## Partie 13 — RÉCONCILIATION 8 : nettoyage et export du CSV final `2026_04_Chronopost_Import.csv` (10:35 – 11:36)

39. (11:11 – 11:30, fichier `2026_04_Chronopost_Import.csv` ouvert dans
    Excel, titre de fenêtre confirmé) : sélection des lignes 2 à 12 (les
    12 lignes CAP*/ECO*/SUR*, colonne F "N° Tracking" affichant CAPI1,
    CAPI2, CAPN1, CAPN2 ×6, CAPO1, ECORI, ECORN, ECORO, SURTI, SURTN,
    SURTO — même liste qu'au point 8) via sélection de plage de lignes
    (clic sur les numéros de ligne 2 à 12 dans la marge gauche).
40. (11:14 – 11:20) Menu contextuel (clic droit sur la sélection de
    lignes) → option **"Supprimer"** avec sous-choix impliquant
    "Ligne entière" (icône/texte partiellement visible mais action
    confirmée par le résultat).

   **➜ RÉCONCILIATION POINT 8 (suppression lignes CAP/ECO/SUR) : CONFIRMÉ
   EXACTEMENT.** Les 12 lignes forfaitaires (CAP*/ECO*/SUR*) sont bien
   supprimées (lignes entières, pas juste le contenu) dans le fichier CSV
   FINAL `2026_04_Chronopost_Import.csv` — elles ne doivent effectivement
   PAS apparaître dans l'import ERP final, elles ont uniquement servi de
   source aux sous-totaux "Gazole"/"eco"/"sureté" calculés plus haut dans
   "Facture Chronopost" (colonnes AF) puis redistribués/agrégés par le TCD
   avant d'arriver dans "Fichier import"/le CSV sous forme de colonnes
   "gazole" (V) et "Adresses"(S)/"Zones éloignées"(Q) réparties sur les
   lignes de transport normales, plutôt que sous forme de lignes séparées.**

## Partie 14 — Modification de la date (11:29 – 12:00, déduit du contexte, non capté net)

41. Cette étape ("modifier date 1ère ligne + étirer colonne" — Points 3 et
    7 de la consigne, à la fois pour l'onglet "Import" du classeur ET le
    CSV final) n'a pas pu être observée comme une action isolée et nette
    dans les frames disponibles (probablement noyée dans le défilement
    rapide entre 09:00 et 11:30). On observe seulement, de façon indirecte
    :
    - Dans "Fichier import" (classeur), colonne B "Date validité[tarif]" =
      "01/04/2026" pour toutes les lignes visibles (valeur uniforme, donc
      cohérente avec "modifier la date en 1ère ligne puis étirer/recopier
      vers le bas").
    - **Anomalie détectée dans le CSV final AVANT nettoyage complet**
      (frame f_0100/f_0101, onglet `2026_04_Chronopost_Import.csv`,
      colonne B "Date validité[tarif]") : les 5 premières lignes (2 à 5,
      les lignes CAPI1/CAPI2/CAPN1/CAPN2) affichent bien "01/04/2026",
      MAIS les lignes suivantes (6 à 30 environ, ECORI et au-delà)
      affichent la valeur **"46113"** — un **NUMÉRO DE SÉRIE DE DATE EXCEL
      NON CONVERTI** (46113 correspond au 1er avril 2026 en interne Excel,
      mais affiché en format "Nombre" au lieu de "Date"). **Ce même type
      de bug (date-serial non converti) a déjà été rencontré et corrigé
      pour Geodis** (cf. mémoire projet `b147d74 Fix Geodis: ... date
      serial fix`) — pattern à surveiller/généraliser pour Chronopost
      aussi lors du codage de l'automatisation.
    - Cette anomalie de date SEMBLE corrigée plus tard dans la vidéo : les
      frames de la Partie 15/16 (après ~12:30) montrent une colonne B
      "Date valid[...]" affichant "01/04/2026" de façon uniforme sur
      toutes les lignes du CSV final nettoyé.

## Partie 15 — RÉCONCILIATION 8 (mapping mode envoi / zone) et QA finale par filtre (12:27 – 15:59)

42. (12:27 – 15:59) Longue séquence de **contrôle qualité par filtrage
    successif** sur la colonne **M "mode envoi"** du CSV final
    `2026_04_Chronopost_Import.csv` (1555 enregistrements au total) :
    l'utilisateur ouvre le filtre de la colonne M, coche/décoche des
    valeurs une par une ou par groupes, et parcourt visuellement les
    lignes filtrées pour vérifier la cohérence Zone (colonne J, valeurs
    numériques 2 à 42 ou codes 16/86/1S/44/17) ↔ mode envoi (colonne M).
    Valeurs de mode envoi confirmées présentes dans le filtre : **2, 9,
    16, 25, 26, 30, 33, 39, 42, 86, 17_Z1 à 17_Z4 (et probablement plus,
    liste tronquée par la fenêtre du filtre), 6B, 6C**.
43. (13:24 – 13:30, zoom sur des lignes filtrées "mode envoi = 6B") :
    toutes les lignes "6B" ont pour valeur de **Zone (colonne J) = "9"**
    ou Pays = "PT" (Portugal) — Transporteur = "CHRONO_2SHOP", tracking
    au format `XT327...TS`.
44. (13:42, filtre suivant "mode envoi = 6C") : lignes avec Zone (colonne
    J) = **25, 26, 30, 33, 39, 42** — Pays = "FR" (France), Transporteur =
    "CHRONO_2SHOP", tracking au format `XT413...TS`.
45. (14:00 environ, ouverture du filtre Zone en amont) : la liste des
    valeurs de zone associées à "6B" montre les cases cochées **9** (et
    listant aussi 2, 16, 25, 26, 30, 33, 39, 42 dans la liste déroulante
    complète du filtre — mais SEULE "9" est cochée pour le mode "6B" dans
    la frame observée) ; pour "6C", les zones cochées sont **25, 26, 30,
    33, 39, 42** (6 valeurs).

   **➜ RÉCONCILIATION POINT 8 (mapping mode envoi/zone) : PARTIELLEMENT
   CONFIRMÉ.** Le tableau donné par l'utilisateur ("6B/entre 1 et 24,
   6C/entre 25 et 42") est **cohérent avec la borne haute observée** (6C
   va jusqu'à 42, et les valeurs vues pour 6C sont 25/26/30/33/39/42,
   toutes ≥25 et ≤42) mais la vidéo ne montre PAS de preuve directe que
   "6B" couvre l'intégralité de 1 à 24 (seule la valeur "9" a été vue
   cochée/présente dans les données pour 6B dans cette période précise —
   il est possible que "6B" couvre bien la plage 1-24 mais que seules
   quelques valeurs de cette plage soient représentées dans les données
   réelles d'avril 2026). Le mapping "16/16, 86/86, 5X/5X, 5Y/5Y, 2/2,
   1S/1S, 44/44_C1 à 44_C4, 17/17_Z1 à 17_Z9" n'a PAS pu être vérifié
   directement dans cette vidéo au-delà de la présence de ces codes
   (16, 86, 17_Z1-Z4, 44_C4 vu en Partie 13 point via 44_C4 déjà vu au
   frame f_0102) dans les colonnes M "mode envoi" et J "Zone" du CSV final
   — mais AUCUNE frame ne montre la feuille "Bibliothèque transporteurs"
   elle-même, donc la RÈGLE/TABLE SOURCE de cette correspondance
   (feuille "Bibliothèque transporteurs", colonnes A/B/C, cf. Partie 9
   point 28) n'est **pas visible** dans cette vidéo — seul le RÉSULTAT
   (valeurs déjà calculées dans le CSV final) est observable. Il faudra
   ouvrir directement la feuille "Bibliothèque transporteurs" du classeur
   réel `2026_04_Facture Chronopost.xlsx` (fichier présent dans le
   projet) pour lire la table complète.**
46. Cette séquence de filtrage successif (12:27 à 15:59, soit près de 3.5
    minutes de la vidéo) semble être une **procédure de contrôle qualité
    manuelle** : vérifier visuellement, mode envoi par mode envoi, que
    toutes les lignes de ce mode ont une Zone cohérente et qu'aucune ligne
    n'affiche "inconnu" — mais aucune action de correction n'est observée
    pendant cette séquence (uniquement de la consultation/vérification),
    et la vidéo se termine (15:59) sur cet état de filtrage, sans étape de
    clôture/export finale visible après.

## Partie 16 — RÉCONCILIATION 9 : deux contrats standard / 2shop (transversal à toute la vidéo)

47. Confirmé de façon transversale (Parties 2, 6, 11, 13, 15) : les deux
    fichiers sources `facture_chronopost_51291303_202604.xlsx` (Sous-
    compte "0" ou "2", Transporteur final = **"CHRONOPOST"**) et
    `facture_chronopost_65481903_202604.xlsx` (Transporteur final =
    **"CHRONO_2SHOP"**) sont bien traités comme deux flux distincts qui
    convergent dans la même feuille "Facture Chronopost" consolidée (même
    tableau, mêmes colonnes) puis ressortent avec une valeur de
    Transporteur différente dans le CSV final ("CHRONOPOST" vs
    "CHRONO_2SHOP").
48. **Concernant la règle "2SHOP : ne pas mettre de gazole dans les coûts
    de revient"** : la note rouge existe bel et bien dans l'onglet
    "Fichier import" (Partie 11, point 32) mais AUCUNE formule
    conditionnelle explicite du type `SI(Transporteur="CHRONO_2SHOP";0;
    Gazole)` n'a été repérée dans les frames disponibles pour la colonne
    V "gazole" — la note semble être un RAPPEL MANUEL pour une vérification
    a posteriori plutôt qu'une règle automatisée dans le classeur observé
    à ce stade. **À creuser en code** : soit cette règle est appliquée
    plus tôt (redistribution du gazole au niveau TCD, qui pourrait exclure
    les lignes CHRONO_2SHOP du pool à redistribuer), soit elle nécessite
    une vérification manuelle systématique (risque d'erreur humaine si pas
    automatisée).

---

## Synthèse de la réconciliation (point par point, vs. notes manuscrites de l'utilisateur)

| # | Point utilisateur | Statut | Précision apportée |
|---|---|---|---|
| 1 | CAP=taxe gazole, ECO=taxe ECP, SURT=taxe sûreté | **Confirmé (imprécis)** | Ce sont des préfixes de la colonne **L "Numero LT"**, pas des valeurs de la colonne **N "Type prestation"** (qui contient le libellé complet FR). Motif : `CAP*`/`ECO*`/`SUR*` sur colonne L. |
| 2 | Colonne AF = sous-totaux par poste, tri par colonne L | **Confirmé** | AF2 "Frêt", AF3 "eco", AF4 "sureté", AF5 "Gazole" = `SOMME(T<début>:T<fin>)` sur plage FIXE (pas dynamique), recalée à la main après tri visuel colonne L. |
| 3a | AF "Somme frêt/éco/sûreté/gazole" = colonne T des lignes concernées | **Confirmé** | Ex. AF3 "eco" = `=SOMME(T13:T21)`, exactement les 9 lignes ECORI/ECORN×6/ECORO. |
| 3b | AG8 = taxe gasoil routier, AG9 = taxe gasoil aérien | **Confirmé exactement** | AG8 = 17,55% "routier", AG9 = 29,75% "aérien". Coexiste avec un 3e taux "Gazole %" (colonne X) = 15,15% appliqué ligne à ligne — lien entre les 3 taux non capturé. |
| 3c | Étirer colonnes W à AC | **Confirmé (partiel)** | W=Zoning 2shop%, X=Gazole%, Y=frêt, Z=sureté+eco, AA=gazole, AB=hors gazole, AC=Catégories. Formules exactes de W/Y/Z/AA/AB non lues net. |
| 4 | Onglet "Contrôle PDF" : MAJ TCD vs PDF montant HT | **Confirmé exactement** | Onglet nommé "**Contrôle pdf**" (minuscules). TCD Somme Montant HT par No Facture vs colonne "pdf" saisie main, écart doit = 0. Total avril 2026 = 15 163,94€. |
| 5 | Onglet "TCD poids" : MAJ colonne B, étirer C | **Non observé** | Onglet jamais ouvert net dans cette vidéo. |
| 6a | Onglet TCD : MAJ colonne R, étirer A B C S T U V W Y Z | **Partiel** | S="sureté+eco" (RECHERCHEX), T="mode envoi" (RECHERCHEX + Bibliothèque transporteurs), Y="Formule Test", Z="Vérif avec ancienne formule" (calcul Frêt résiduel). Colonnes A,B,C,U,V,W non identifiées avec certitude. R non identifiée. |
| 6b | Colonne Y reprend le frêt de l'onglet import | **Non confirmé** | Colonne Y vue ici = "Formule Test" (contrôle), pas explicitement liée à "onglet import" — à reconfirmer. |
| 6c | #N/A si colonne S en texte → convertir en nombre → réactualiser TCD | **Confirmé exactement** | Colonne **S "Produit"** de "Facture Chronopost" (codes "17","44"...) parfois stockée en texte (triangle vert "Nombre stocké sous forme de texte") → casse RECHERCHEX vers Bibliothèque transporteurs → symptôme #N/A (et aussi #VALEUR! observé) dans TCD. Correction : "Convertir en nombre" sur colonne S, puis clic droit > Actualiser sur le TCD. |
| 7 | Onglet Import (classeur) : modifier date 1ère ligne + étirer | **Confirmé indirectement** | Colonne B "Date validité tarif" uniforme "01/04/2026" observée ; anomalie de date-serial Excel (46113) détectée à un stade intermédiaire du CSV, similaire au bug déjà connu chez Geodis. |
| 8a | CSV final : modifier date + étirer | **Confirmé indirectement** | Même mécanisme que 7, corrigé avant la fin de la vidéo. |
| 8b | CSV final : supprimer lignes CAP/ECO/SUR | **Confirmé exactement** | 12 lignes forfaitaires supprimées (ligne entière) du CSV final `2026_04_Chronopost_Import.csv`. |
| 8c | Vérifier somme totale + X2 TCD > somme totale PDF | **Non observé tel quel** | La comparaison observée est plutôt "Contrôle pdf" = Somme Montant HT (TCD) **=** PDF, écart **= 0** (égalité stricte, pas d'inégalité "supérieure"). La formulation "+X2" n'a pas de correspondance visible ; possible confusion avec une autre feuille ou un contrôle non filmé. |
| 8d | Mapping mode envoi/zone (16/16, 86/86, 5X/5X, 5Y/5Y, 2/2, 1S/1S, 44/44_C1-C4, 17/17_Z1-Z9, 6B/1-24, 6C/25-42) | **Partiellement confirmé** | Codes 16, 86, 17_Z1-Z4, 44_C4, 6B, 6C bien observés dans les données finales (colonnes M "mode envoi" / J "Zone"). Table source = feuille "**Bibliothèque transporteurs**" (jamais ouverte dans la vidéo). 6C confirmé = zones 25 à 42. 6B confirmé seulement pour zone "9" dans les données réelles (pas de preuve directe de la plage complète 1-24). |
| 9 | Deux contrats standard/2shop, 2SHOP sans gazole en coût de revient | **Confirmé (partiel)** | Transporteur final = "CHRONOPOST" (compte 51291303) vs "CHRONO_2SHOP" (compte 65481903). Note manuelle rouge "2SHOP : ne pas mettre de gazole dans les coûts de revient" bien présente dans "Fichier import", mais AUCUNE formule conditionnelle automatique observée pour cette règle — semble être un contrôle manuel, pas (encore) automatisé dans le classeur. |

---

## Points ambigus / illisibles à faire confirmer par le pôle transport

1. **Feuilles jamais ouvertes dans cette vidéo** : "Bibliothèque
   transporteurs" (table mode envoi ERP), "Zoning 2shop" (zones 2shop),
   "TARIFS" et "cap à 5%" (logique de plafonnement à 5%, toujours non
   élucidée), "TCD poids" (agrégation par poids), "Bilan clients", "Avoir".
   Aucune de leur logique interne n'est visible dans cette vidéo — il
   faudra soit une 2e vidéo, soit ouvrir directement le fichier réel
   `2026_06_Facture Chronopost.xlsx` du projet pour les lire.
2. **Formules exactes des colonnes W, Y, Z, AA, AB de "Facture
   Chronopost"** (Zoning 2shop%, frêt, sureté+eco, gazole, hors gazole) :
   seuls les résultats numériques et les en-têtes ont été lus avec
   certitude ; le contenu de la barre de formule pour CES colonnes
   spécifiques (par opposition à AF qui a été bien capturé) n'a pas pu
   être lu net dans les frames disponibles.
3. **Lien exact entre les 3 taux "gazole"** observés : Gazole % (colonne
   X, 15,15%), taux routier AG8 (17,55%), taux aérien AG9 (29,75%). La
   formule qui combine ces 3 valeurs (moyenne pondérée ? remise
   spécifique ? application différenciée selon Type prestation
   Aérien/Routier de la ligne ?) n'apparaît dans aucune frame lisible.
4. **Concept "Gazole réel" vs "(avant remise)"** (labels AE6/AE7 environ,
   point 17) : présence confirmée mais formule/calcul non observé —
   correspond probablement à la "remise gazole" déjà signalée comme point
   ouvert dans la précédente transcription.
5. **Colonne AC "Catégories" classe ECO/SUR sous "Frêt"** (point 11) —
   semble en tension avec l'existence de sous-totaux distincts "eco" et
   "sureté" en colonne AF (Partie 4). Cela pourrait signifier que : (a)
   "Catégories" (AC) sert à un TOTAL "hors gazole" volontairement fusionné
   Frêt+Eco+Sûreté, alors que AF calcule des sous-totaux plus fins pour
   contrôle interne uniquement ; ou (b) c'est une incohérence/legacy dans
   le classeur. **À confirmer avec le pôle transport : la ventilation
   finale envoyée à l'ERP distingue-t-elle réellement "Sûreté" et "Eco" du
   "Frêt", ou sont-ils fusionnés dans le poste Frêt au final ?** (Cette
   question est directement pertinente pour le code de
   `facturation-app/src/carriers/`.)
6. **Formulation "somme totale + X2 du TCD supérieure à la somme totale du
   PDF"** (point 8c de la consigne) : aucune trace de ce contrôle précis
   n'a été localisée dans les frames. Le seul contrôle Total/PDF observé
   est l'égalité stricte "Contrôle pdf" (écart=0). Il est possible que ce
   contrôle "+X2" existe dans une feuille non ouverte (TCD poids ? Bilan
   clients ?) ou qu'il s'agisse d'un contrôle mental/manuel du pôle
   transport non retranscrit dans le classeur. **À clarifier
   directement avec l'utilisateur.**
7. **Plage complète du mapping "6B" = zones 1 à 24** : seule la valeur "9"
   a été vue dans les données réelles pour le mode "6B" pendant la
   séquence de QA filtrée (12:27-15:59). Les autres valeurs de la plage
   1-24 n'apparaissent peut-être simplement pas dans les données d'avril
   2026 (pas d'expédition vers ces zones ce mois-là), donc ce n'est pas
   une contradiction, juste un manque de preuve positive exhaustive.
8. **Étape "modifier date 1ère ligne + étirer colonne"** (Import classeur
   ET CSV final, points 3/7/8a) : non observée comme action isolée nette
   — seulement le résultat (colonne B uniforme "01/04/2026") et un bug de
   date-serial Excel intermédiaire (valeur "46113" non convertie,
   frame ~01:40 dans le CSV avant nettoyage complet). Pattern similaire au
   bug déjà corrigé pour Geodis (`b147d74`) — à vérifier/généraliser lors
   du codage Chronopost.
9. **Aucune connexion internet/navigateur visible dans cette vidéo** :
   contrairement à ce que suggérait la consigne initiale ("disponible sur
   leur site" pour AG8/AG9), aucune frame ne montre un navigateur web ou
   le SI `si.laruche-logistique.fr`. Les taux AG8/AG9 (17,55%/29,75%)
   apparaissent déjà pré-saisis dans le classeur au moment où la vidéo les
   montre — leur SOURCE (recherche manuelle sur le site Chronopost, ou
   consultation du SI interne) n'est pas filmée dans cette vidéo
   spécifique (peut-être montrée dans une autre vidéo Chronopost non
   fournie ici).
10. **Colonnes A, B, C, R, U, V, W de l'onglet TCD** (point 6a) : présence
    d'en-têtes non lus avec certitude suffisante pour les documenter
    précisément (zoom insuffisant sur les frames disponibles à ces
    instants). Le TCD complet (avec tous ses en-têtes) mériterait d'être
    ouvert directement dans le fichier réel pour une lecture complète
    colonne par colonne.
11. **Formule exacte de AF1** (grand total en haut de la zone de
    sous-totaux, point 12) : valeur affichée (13 003,84 puis 13 064,81)
    mais formule non capturée nette dans la barre de formule.
12. **Bornes exactes des plages AF2/AF3 pour "Frêt"** : deux valeurs
    successives vues pour AF2 (`T22:T30` puis `T20:T27`) suggèrent que la
    plage réelle finale n'est pas garantie être celle vue en dernier dans
    les frames disponibles (le classeur pourrait avoir été encore ajusté
    après la dernière frame capturée sur cette zone) — à vérifier
    directement dans le fichier réel si possible.
