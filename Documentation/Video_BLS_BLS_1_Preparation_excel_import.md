# Transcription vidéo — BLS_1_Preparation excel import.mp4

Durée totale : ~10 min 46 s (646,5 s). Résolution capture d'écran Windows,
Excel + navigateur (Edge/Chrome) + explorateur de fichiers + Outlook.
Timestamps approximatifs en `mm:ss`, dérivés des frames extraites (intervalle
fixe 5 s + détection de changement de scène).

Aucune bande son exploitable : toute l'info vient de ce qui est visible à
l'écran (barre de formule, onglets, cellules, boîtes de dialogue).

---

## Partie 1 — Repérage de la facture reçue et du modèle de mois précédent (00:00 – 01:40 env.)

1. (00:00) Écran de démarrage : dossier Windows Explorer ouvert sur un
   répertoire réseau contenant des sous-dossiers de transporteurs
   (nom visible partiellement, type `... Factures transporteurs + calculs`).
2. (00:02 – 00:25) Navigation dans l'explorateur de fichiers : ouverture d'un
   dossier `BLS`, dans une arborescence par année/mois
   (`... \2026\2026 03\BLS\` ou similaire — texte de la barre d'adresse pas
   totalement net).
3. (00:28 – 00:40) Un fichier Excel nommé sur le modèle
   `2026_03_Facture BLS.xlsx` est ouvert (mois de MARS, servant de modèle/
   référence pour le mois précédent). Le classeur possède les mêmes onglets
   que le classeur de juin déjà en référence : **Factures BLS**, **Bilan
   PDF**, **Import CSV**, **Bilan client**, **Avoir**.
4. (00:40 – 00:50) L'utilisateur (pôle transport) clique sur l'onglet
   **Bilan client** du fichier de mars. Un tableau croisé dynamique (TCD)
   apparaît avec en ligne des codes clients NUMÉRIQUES (valeurs du type
   3752, 8041, 7261, etc. — mêmes codes que ceux déjà repérés dans le
   classeur de juin) et en colonnes une agrégation de montants.
   **Ce TCD est alimenté par la colonne "ID Client" de l'onglet "Factures
   BLS"** du même classeur (source du TCD = plage Factures BLS!A:J typiquement).
5. (00:45 – 01:00) L'utilisateur retourne sur l'onglet **Factures BLS** du
   fichier de mars : on y voit que la colonne A "ID Client" contient bien
   des valeurs numériques SAISIES (pas de formule visible dans la barre de
   formule — valeur brute type "3752" affichée telle quelle en cellule A2,
   A3...). Confirme que ces codes sont bien saisis à la main dans "Factures
   BLS" (pas de RECHERCHEV/XLOOKUP visible à ce stade).
6. (01:00 – 01:40) Bascule vers Outlook (ou client mail web) : recherche
   dans la boîte de réception d'un mail contenant la facture BLS d'avril
   2026 (objet du type "Facture BLS avril 2026" ou similaire, pas totalement
   lisible). Une pièce jointe Excel/PDF est visible dans l'aperçu du mail.

## Partie 2 — Récupération et nettoyage du fichier source brut (01:40 – 04:00 env.)

7. (01:40 – 02:00) Téléchargement/ouverture de la pièce jointe du mail :
   fichier au format facture (type `2600772 (1).xlsx`), mise en page
   "facture" avec en-tête société, **N° Client : 01LARUCH**, un numéro de
   folio, et des colonnes non alignées en tableau simple (comparable au
   fichier de référence `2601458.xlsx` déjà connu). PAS un tableau plat :
   il y a des cellules fusionnées / décalées, un en-tête de plusieurs
   lignes avant les données.
8. (02:00 – 02:30) Dans ce fichier source brut, on peut lire en bas/en-tête
   un total **NET A PAYER : 12420,00 EUR** (visible en cellule ou zone de
   texte de la mise en page facture).
9. (02:10 – 02:40) Retour dans l'explorateur : `Enregistrer sous` (Save As)
   du classeur `2026_03_Facture BLS.xlsx` sous un nouveau nom
   `2026_04_Facture BLS.xlsx` (clonage du modèle du mois précédent pour
   préparer le nouveau mois). La boîte de dialogue "Enregistrer sous"
   affiche une arborescence de dossiers avec un chemin du type
   `... 1 - Factures transporteurs + calculs\2026\2026 04\BLS\`.
10. (02:40 – 03:20) Dans le nouveau fichier `2026_04_Facture BLS.xlsx`,
    onglet **Factures BLS** : la ligne d'en-tête est identique à celle du
    fichier modèle de juin déjà connu — **ID Client / n° facture / Date
    Prestation / Dossier / Libellé / Impact CO2 / Unité / Quantité / Prix
    Unitaire / Montant H.T. / Code Tva**.
11. (03:00 – 03:40) Fenêtres côte à côte : à gauche le fichier source brut
    `2600772 (1).xlsx` (ou équivalent, mise en page facture), à droite le
    classeur de travail `2026_04_Facture BLS.xlsx` onglet "Factures BLS".
    Début de la copie manuelle ligne par ligne des données depuis le fichier
    source vers le classeur de travail :
    - colonne **Dossier** (ex. 11111, 11126, 11127, 11157, 11158, 11170,
      11194, 11237, 11301, 11309, 11341, 11425, 11514, 11513, 11609) copiée
      depuis la colonne équivalente ("Folio"/"Réf" ?) du fichier source.
    - colonne **Libellé** copiée avec un texte de type trajet
      "De [ville départ] A [ville arrivée]" (ex. "De 21 Créancey A 75
      Paris", "De 21 Créancey A NL Hoofddorp-2132" pour un envoi vers les
      Pays-Bas).
    - colonne **Unité** = "M" pour toutes les lignes (constaté identique
      pour toutes les lignes visibles).
    - colonne **Quantité**, **Prix Unitaire** (partiellement visible/vide
      pour certaines lignes) et **Montant H.T.** copiées/collées depuis la
      source.
12. (03:20 – 04:00) La colonne B **"n° facture"** est remplie avec la valeur
    **2600772** répétée sur toutes les 16 lignes (numéro de facture unique
    du mois, identique à celui du fichier source `2600772 (1).xlsx`).
    Colonne C **Date Prestation** copiée aussi (dates d'avril 2026,
    01/04/2026 à 30/04/2026).

## Partie 3 — Constat et correction d'un écart de numéro de facture (04:00 – 05:10 env.)

13. (04:00) Onglet **Bilan PDF** du classeur `2026_04_Facture BLS.xlsx` :
    un TCD compare la somme du Montant H.T. (source = "Factures BLS") à une
    valeur "PDF" saisie manuellement dans une cellule dédiée. Un écart
    négatif apparaît (ex. TCD affiche "Somme de Montant HT" = 3785 alors
    que la cellule "PDF"/Total attendu contient 4975, écart -1190) —
    lecture partiellement floue sur cette frame précise, mais le principe
    (comparaison Somme calculée vs Total PDF saisi à la main) est clair.
14. (04:10 – 04:40) Retour sur "Factures BLS" : l'utilisateur remarque que
    la colonne B "n° facture" avait initialement une valeur incorrecte pour
    certaines lignes (ex. "2600513" au lieu de "2600772"), ce qui faussait
    potentiellement le filtre du TCD "Bilan PDF" ou "Bilan client" (le TCD
    utilise n° facture comme filtre de page probablement). Correction :
    sélection de la colonne B entière et collage de la valeur correcte
    "2600772" sur toutes les lignes (2 à 16).
15. (04:40 – 05:10) Retour sur **Bilan PDF** : clic droit sur le TCD >
    **Actualiser** (rafraîchissement du TCD). Après actualisation, la
    "Somme de Montant HT" calculée devient **10350** (nouvelle valeur après
    correction de n° facture). La cellule "PDF"/Total attendu est ensuite
    corrigée manuellement pour refléter le vrai total de la facture PDF
    reçue — l'utilisateur tape/colle la valeur du **TOTAL H.T.** lu sur le
    PDF officiel de la facture (probablement 10350 aussi, pour obtenir un
    écart de 0). **Mécanisme confirmé : "Bilan PDF" compare la Somme de
    Montant H.T. calculée dans "Factures BLS" (via TCD) à une valeur "Total
    H.T." saisie manuellement à partir du PDF de la facture transporteur —
    l'égalité (écart = 0) valide la cohérence du mois.**

## Partie 4 — Annotations manuelles sur cas particuliers (05:10 – 05:35 env.)

16. (05:10 – 05:20) Sur l'onglet **Import CSV** (ou "Factures BLS", à
    confirmer — la frame montre en fait l'onglet **Import CSV** actif en
    bas de l'écran), une cellule en colonne **X** (hors zone de données
    structurée, colonne annexe/commentaire libre) reçoit une note tapée
    manuellement : *"Même transport facturé 1 seul fois donc 4960€"* — en
    référence aux deux lignes Dossier 11157 et 11158 (toutes deux "De 21
    Créancey A NL Hoofddorp-2132", Montant H.T. 2480,00 chacune,
    surlignées en orange/pêche dans le tableau) : le commentaire indique
    que même si deux lignes existent avec le même trajet, il s'agit du même
    transport facturé une seule fois pour un montant total de 4960€
    (2 x 2480).
17. (05:20 – 05:35) Une seconde note est tapée dans la cellule juste
    en-dessous (ligne du Dossier 11170, Fret = 0) : *"Transport offert car
    livré en corse au lieu de l'Italie et 2 semaines pour retourner la
    marchandise"* — cas particulier annoté manuellement pour justifier un
    montant de fret à 0€ sur cette ligne (litige de livraison /
    dédommagement).

## Partie 5 — Reconstruction du fichier "Import CSV" / export CSV final (05:35 – 06:15 env.)

18. (05:35 – 05:50) Retour sur onglet **Import CSV** du classeur
    `2026_04_Facture BLS.xlsx`. En-têtes de colonnes confirmées, dans
    l'ordre (colonnes A à W environ) :
    **Transporteur (A) / Date validité tarif (B) / Réf.1 (C) / Réf.2 (D) /
    Id client (E) / N° Tracking (F) / Nom (G) / E/P (H) / Pays (I) /
    Zone (J) / Nbr Coli (K) / Poids (L) / mode envoi (M) / TVA (N) /
    Droits et tax (O) / Assurance (P) / Zones él[oignées] (Q) / Colis
    vol[umineux] (R) / Adresse (S) / Frêt (T) / plus-value Bt[?] (U) /
    Gazole (V) / Nb Colis (W, doublon ?)**.
    Ligne 1 = en-têtes, lignes 2 à 16 = les 15 lignes de données du mois
    (Dossier 11111 à 11609, cf. liste ci-dessus).
19. (05:45 – 06:00) Valeurs constatées ligne par ligne dans "Import CSV"
    pour avril 2026 :
    - Transporteur = "BLS" pour toutes les lignes.
    - Date validité tarif = "01/04/2026" pour toutes les lignes.
    - Id client (colonne E) = **VIDE** pour toutes les 15 lignes (confirmé
      à l'écran, aucune valeur saisie dans cette colonne à ce stade de la
      vidéo).
    - N° Tracking (colonne F) = les valeurs Dossier (11111, 11126,
      11127...).
    - Nom (colonne G) = libellé trajet ("De 21 Créancey A 75 Paris" etc.).
    - E/P (colonne H) = "E" pour toutes les lignes (Enlèvement/Expédition
      probablement).
    - Pays (colonne I) = "France" pour toutes les lignes.
    - Zone (colonne J) = "France" pour toutes les lignes.
    - Nbr Coli (colonne K) = valeurs numériques variables (1, 3, 1, 65, 65,
      1, 7, 7, 5, 31, 1, 1, 2, 18, 1).
    - Poids (colonne L) = valeurs numériques (400, 762, 120, 7150, 7150,
      325, 980, 1700, 1750, 12400, 228, 232, 423, 3821,5, 497).
    - mode envoi (colonne M) = "ST" pour toutes les lignes.
    - Frêt (colonne T) = montants (195, 265, 265, 2480, 2480, 0, 650, 740,
      365, 385, 215, 195, 315, 1605, 195).
    - Colonne X (hors tableau structuré) : commentaires libres évoqués au
      point 16-17 ci-dessus.
20. (06:00 – 06:15) Confirmation visuelle : les lignes Dossier 11157/11158
    (transport unique facturé 2x sur le mois) et Dossier 11170 (transport
    offert) sont surlignées en couleur (orange/pêche clair) dans le tableau
    "Import CSV" pour signaler visuellement ces cas particuliers au
    lecteur du fichier.

## Partie 6 — Ouverture du CSV import du mois précédent comme référence (06:15 – 07:30 env.)

21. (06:15 – 06:40) Ouverture d'un fichier CSV séparé (PAS un onglet du
    classeur xlsx, un vrai fichier `.csv` distinct) nommé
    `2026_03_BLS_Import.csv` (mois de mars, mois précédent) — ouvert dans
    Excel pour consultation, sert de modèle de structure du CSV final à
    produire.
22. (06:40 – 07:10) Comparaison structure du CSV de mars avec l'onglet
    "Import CSV" du classeur de travail avril — mêmes colonnes, même ordre.
    L'utilisateur navigue dans les colonnes K (Nbr Coli) et L (Poids) du
    CSV de mars : ce sont des **valeurs numériques figées** (pas de
    formule visible dans la barre de formule en sélectionnant ces cellules
    — confirme que "Nbr Colis" et "Poids" sont des valeurs saisies/collées
    "en dur", pas calculées par formule dans le fichier CSV final).
23. (07:10 – 07:30) Retour au classeur de travail `2026_04_Facture
    BLS.xlsx`, onglet "Import CSV" : sélection de la plage de données
    (lignes 2 à 16, colonnes F à W environ) puis **Copier**.

## Partie 7 — Enregistrement du nouveau fichier CSV d'import (07:30 – 08:00 env.)

24. (07:30 – 07:45) `Fichier > Enregistrer sous` d'un nouveau fichier CSV :
    boîte de dialogue Windows "Enregistrer sous" avec arborescence de
    dossiers visible :
    - Chemin racine réseau contenant deux branches distinctes :
      `2 - Fichiers csv import\2026\2026 04\` (pour le CSV) et
      `1 - Factures transporteurs + calculs\2026\2026 04\BLS\` (pour le
      classeur de calcul xlsx).
    - Nom de fichier proposé/saisi : **`2026_04_BLS_Import.csv`**
      (convention `AAAA_MM_BLS_Import.csv`, cohérente avec le nom du
      fichier de référence de mars observé au point 21).
25. (07:45 – 08:00) Le fichier est enregistré au format CSV (séparateur
    point-virgule probable, non confirmé visuellement avec certitude).
    Les données collées dans ce nouveau fichier CSV sont des **valeurs
    figées** (Collage spécial > Valeurs, pas de formules) — confirmé par
    absence de formule dans la barre de formule lors de la sélection de
    cellules du nouveau CSV.

## Partie 8 — Consultation et export depuis le portail "AffreTrans" (08:00 – 09:30 env.)

26. (08:00 – 08:20) Bascule vers un navigateur web (onglet déjà ouvert) sur
    l'URL **`planeo.laruche-logistique-france.fr/affrtrans/...`** — portail
    web nommé **"AffreTrans"** (logo "AT AffreTrans" dans le bandeau
    latéral gauche), outil interne de l'entreprise "La Ruche Logistique"
    pour la gestion des affrètements/commandes de transport. Ce n'est PAS
    un logiciel tiers du transporteur BLS lui-même : c'est l'outil de
    gestion des expéditions de "La Ruche Logistique" (l'expéditeur/client),
    utilisé pour tous les transporteurs (Géodis, BLS, Kuehne, CEVA
    apparaissent comme options de transporteur dans l'interface).
27. (08:20 – 08:50) Menu de gauche du portail AffreTrans : **Tableau de
    bord / Affrètement / Historique / Statistiques / Export / Stats /
    Annuaire / Import / Modèles Emails / Administration**. Clic sur
    **"Export / Stats"**.
28. (08:50 – 09:20) Dans la section Export/Stats, sélection d'un type
    d'export **"Export Affrètement"** avec un filtre de plage de dates
    (probablement le mois d'avril 2026, ou une période couvrant les
    enlèvements du mois). Un bouton d'export génère un fichier téléchargé
    nommé **`export_affretement_2026-05-04_1108.csv`** (export généré le
    04/05/2026 à 11h08, contenant les données d'affrètement).
29. (09:20 – 09:30) Le fichier CSV exporté `export_affretement_2026-05-04_
    1108.csv` est ouvert dans Excel. Colonnes observées incluent (ordre
    approximatif, lecture partielle) : N° affrètement, Client, **ID
    Client**, N° expé[dition], **Poids (kg)** (colonne J), **Tracking**
    (colonne K), N° facture, Montant facture, **Nb palettes** (colonne P).
    **Ce fichier CSV "export_affretement_..." est la SOURCE de vérité pour
    le rapprochement Poids/Nbr Colis/ID Client par Dossier(=Tracking).**

## Partie 9 — Formules RECHERCHEX pour Poids et Nbr Colis (09:30 – 10:15 env.)

30. (09:30 – 09:50) Retour au classeur `2026_04_Facture BLS.xlsx`, onglet
    "Import CSV" (ou classeur de travail équivalent) : dans la colonne
    **L "Poids"**, une formule est saisie/visible dans la barre de formule
    pour la cellule L2 :
    ```
    =RECHERCHEX(F2;'export_affretement_2026-05-04_1108.csv'!$K:$K;'export_affretement_2026-05-04_1108.csv'!$J:$J;"")
    ```
    Cette formule RECHERCHEX (XLOOKUP) cherche la valeur de **F2** (colonne
    "N° Tracking" = le "Dossier", ex. "11111") dans la colonne **K** du
    fichier export AffreTrans (colonne "Tracking" de l'export), et renvoie
    la valeur correspondante de la colonne **J** de ce même export (colonne
    "Poids (kg)"). Le 4e argument `""` est la valeur de retour si non
    trouvé (chaîne vide).
31. (09:50 – 10:05) Formule similaire testée/visible pour la colonne
    **K "Nbr Coli"** (ou "Nb Colis") :
    ```
    =RECHERCHEX(F2;'export_affretement_2026-05-04_1108.csv'!$K:$K;'export_affretement_2026-05-04_1108.csv'!$P:$P;"")
    ```
    Recherche toujours par "Tracking" (colonne K de l'export), mais renvoie
    cette fois la colonne **P** de l'export (colonne "Nb palettes").
    **Confirme le mécanisme : le rapprochement entre "Import CSV"/"Factures
    BLS" et l'outil d'affrètement (AffreTrans) se fait sur la clé "N°
    Tracking" = "Dossier" (côté classeur BLS) = "Tracking" (côté export
    AffreTrans, colonne K de l'export).**
32. (10:05 – 10:15) La formule RECHERCHEX est ensuite recopiée/étirée sur
    toutes les lignes du tableau (probablement via poignée de recopie ou
    Ctrl+D). Résultat visible : la colonne L "Poids" se remplit avec des
    valeurs issues de l'export AffreTrans.

## Partie 10 — Finalisation en valeurs et derniers ajustements (10:15 – 10:46)

33. (10:15 – 10:30) Sélection de la colonne L "Poids" (et K "Nbr Coli")
    remplie par formule, puis **Copier > Collage spécial > Valeurs**
    (transformation des formules RECHERCHEX en valeurs figées, cohérent
    avec le fichier CSV final observé qui ne contient QUE des valeurs, pas
    de formules).
34. (10:30 – 10:46) Dernière frame : classeur `2026_04_Facture BLS.xlsx`
    onglet "Import CSV" affichant le tableau final avec X7 sélectionné
    (cellule de commentaire vide, prête pour une éventuelle note
    supplémentaire) — fin de la vidéo sur cet état.

---

## Points ambigus / illisibles à faire confirmer par le pôle transport

1. **Écran d'ouverture (00:00-00:02)** : le nom exact du dossier réseau
   racine et l'arborescence complète ne sont pas totalement lisibles sur
   la première frame (résolution/flou). Le chemin final confirmé par la
   boîte "Enregistrer sous" (`... 1 - Factures transporteurs + calculs\
   2026\2026 04\BLS\` et `2 - Fichiers csv import\2026\2026 04\`) est
   fiable, mais le tout début de la navigation est à vérifier si besoin.
2. **Objet exact du mail Outlook (01:00-01:40)** : le texte de l'objet du
   mail contenant la facture BLS d'avril n'est pas parfaitement net à
   l'écran ; à confirmer littéralement si utile pour un futur pattern de
   détection automatique.
3. **Valeurs précises du TCD "Bilan PDF" avant/après correction
   (04:00-05:10)** : les chiffres "3785", "4975", "-1190", "10350" sont
   lus du mieux possible sur les frames disponibles mais une ou deux
   frames intermédiaires sont légèrement floues (mouvement de souris/
   défilement) — à re-vérifier en rejouant la vidéo à vitesse réduite
   sur ce passage précis si les montants exacts sont critiques.
4. **Valeur exacte "PDF"/Total attendu finale dans Bilan PDF** : on
   constate que l'écart devient 0 après correction, mais la valeur
   numérique finale exacte tapée dans la cellule "Total PDF" n'a pas pu
   être lue avec certitude à 100% (probable 10350, cohérent avec la Somme
   de Montant HT recalculée, mais à confirmer).
5. **Format exact du séparateur CSV** (point-virgule vs virgule) du fichier
   `2026_04_BLS_Import.csv` : non déterminable visuellement depuis Excel
   (Excel affiche toujours en cellules, pas en texte brut) — à vérifier en
   ouvrant le fichier CSV produit dans un éditeur de texte.
6. **Colonne "Id client" dans "Import CSV"** : confirmée VIDE tout au long
   de cette vidéo (aucune saisie observée pour cette colonne dans le
   fichier "Import CSV" ni dans le CSV final `2026_04_BLS_Import.csv`).
   **Le mécanisme de détermination du code client (colonne "ID Client" de
   l'onglet "Factures BLS", utilisée par le TCD "Bilan client") N'A PAS
   été montré dans cette vidéo au-delà du constat que les valeurs sont
   déjà présentes dans le fichier modèle de mars (point 4-5) : on ne voit
   PAS, dans cette vidéo 1, le pôle transport saisir concrètement un
   nouveau code client pour une nouvelle ligne, ni consulter un
   outil/ERP pour déterminer ce code.** Cela reste à documenter/confirmer
   séparément (peut-être visible seulement dans la vidéo 2, ou dans une
   partie de la vidéo 1 non capturée par l'échantillonnage de frames, ou
   sur un mois où le pôle transport doit réellement chercher un nouveau
   client — dans les frames disponibles, la colonne "ID Client" du
   nouveau fichier avril reste vide de bout en bout).
7. **Colonne "Réf.1" / "Réf.2" de "Import CSV"** : contenu de ces deux
   colonnes non déterminé avec certitude (semblent vides dans les frames
   consultées, mais à confirmer).
8. **Aucune mention explicite de "taxe gasoil" / "Gazole"** n'a été
   observée dans les frames de cette vidéo au-delà de la présence de la
   colonne "Gazole" (vide) dans l'en-tête de "Import CSV" — aucune formule
   ni valeur n'y est saisie pendant la vidéo. Le statut "à définir" du
   registre semble donc cohérent avec ce qui est observable ici (colonne
   présente mais non utilisée/non calculée).
9. **Zone/Pays** : toutes les lignes observées ont Pays=France, Zone=
   France, y compris la ligne vers "NL Hoofddorp" (Pays-Bas) dans le
   libellé — incohérence apparente entre le libellé du trajet (destination
   internationale) et la colonne Pays/Zone qui reste "France" partout. À
   signaler au pôle transport : est-ce normal (le "Pays" fait référence au
   pays de FACTURATION/société BLS et non au pays de livraison) ou est-ce
   une donnée à corriger manuellement selon les cas ?
10. **Formule RECHERCHEX précise** : la syntaxe complète a été lue au
    mieux sur les frames disponibles (transcrite ci-dessus), mais certains
    caractères (guillemets, points-virgules vs virgules comme séparateurs
    d'arguments) pourraient légèrement différer de la syntaxe exacte
    tapée à l'écran — à vérifier directement dans le fichier
    `2026_04_Facture BLS.xlsx` réel s'il est conservé quelque part, ou en
    rejouant la vidéo image par image sur ce passage (09:30-09:50).
11. **Origine de "Nbr Colis" recherché via colonne P "Nb palettes"** : il
    est possible que la colonne "Nbr Coli" du classeur BLS soit en réalité
    un nombre de PALETTES et non un nombre de colis au sens strict — à
    clarifier avec le pôle transport (terminologie AffreTrans vs
    terminologie facture BLS).
