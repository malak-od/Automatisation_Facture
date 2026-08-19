# Transcription vidéo — FEDEX_1_Preparation fichier import.mp4

Durée totale réelle : **11 min 55 s (714,97 s)** (`ffprobe`). Statut :
**TRANSCRIPTION TERMINÉE** (toutes les frames disponibles ont été lues et
intégrées ; voir section "Points ambigus" en fin de document pour les
zones jamais montrées dans cette vidéo, notamment les onglets Zoning et
Bilan clients).

Aucune bande son exploitable : toute l'info vient de ce qui est visible à
l'écran (barre de formule, onglets, cellules, filtres, menus contextuels).
Timestamps approximatifs en `mm:ss`, dérivés de 89 frames à intervalle fixe
(8 s) + 85 frames de détection de changement de scène, lues dans l'ordre
chronologique par lots de 10-15.

**ATTENTION MAJEURE** : la vidéo montre le traitement du mois d'**AVRIL
2026** (classeur `2026_04_Facture Fedex.xlsx`, dossier réseau
`$Facturation automatique\1 - Factures transporteurs + calculs\2026\2026
04\Fedex\`), PAS juin 2026. Les noms d'onglets et la structure générale
semblent stables (à confirmer précisément colonne par colonne), mais
certaines observations peuvent différer du fichier juin 2026 déjà inspecté
par l'appelant — signalé explicitement à chaque fois.

Croisement systématique avec :
- La mémoire projet déjà actée sur FedEx (colonnes BK-BO, écart TVA/BL,
  fichier `fedex_colonnes_bk_bo_ecart_tva.md`).
- Le classeur réel `Transporteurs/Fedex/2026_06_Facture Fedex.xlsx` déjà
  inspecté par l'appelant (6 feuilles : Zoning, Bilan factures, Shipment
  Detail, TCD, Bilan clients, Import ERP).
- Les fichiers présents dans `Transporteurs/Fedex/` : 2 CSV bruts
  (`FEDEX 06 2026.csv`, `0g000e48hq_FEDEX 06-2026 v2_20260710125723.csv`),
  `2026_06_Fedex_Import.csv`, 9 PDF de factures, `surcharge 01 07 25 au
  30 06 26.XLSX`.

Le classeur avril 2026 montré dans la vidéo a **6 onglets** dans cet ordre :
**Zoning | Bilan factures | Shipment Detail | TCD | Bilan clients | Import
ERP** — identique à l'ordre confirmé par l'appelant sur le fichier juin
2026.

---

## Statut de la transcription

**TERMINÉE.** 41 étapes numérotées réparties en 7 parties chronologiques
(ouverture réseau, Import ERP aperçu, Shipment Detail structure et
formules, fichier externe TMS "Export expéditions_brut", Bilan factures,
onglet TCD, Import ERP formules détaillées, QA finale sur le CSV
d'import). Les onglets "Zoning" et "Bilan clients" du classeur n'ont
jamais été ouverts dans cette vidéo (cf. Points ambigus).

---

## Partie 0 — Ouverture / contexte réseau (00:00 – 00:56)

1. (00:00 – 00:13) Explorateur de fichiers Windows, dossier réseau
   `Réseau > 192.168.5.3 > Comptabilité La Ruche > $Facturation automatique
   > 1 - Factures transporteurs + calculs > 2026 > 2026 04 > Fedex`.
   Contenu visible : 2 fichiers CSV (`0g000e48hq_Fedex 04-26_202605...csv`
   et `0g000e48hq_fedex 04-26_20260512084332...csv`) et 9 PDF de factures
   FedEx (`FEDEX_634243207_781542172.pdf`, `FEDEX_634243603_200720433.pdf`,
   `FEDEX_634251047_781542172.pdf`, `FEDEX_634251402_200720433.pdf`,
   `FEDEX_634259662_781542172.pdf`, `FEDEX_634260023_200720433.pdf`,
   `FEDEX_634268737_200720433.pdf`, `FEDEX_634278311_781542172.pdf`,
   `FEDEX_634278838_200720433.pdf`) — **11 éléments au total pour avril
   2026, avec 9 PDF (contrairement à juin où l'appelant observe 9 factures
   mais seulement 8 PDF)**. Noter les 2 suffixes de compte client
   différents dans les noms de PDF : `_781542172` et `_200720433` — deux
   comptes FedEx distincts facturés séparément (probablement France vs
   International, cf. Partie 2).
2. (00:13 – 00:16) Ouverture du fichier CSV
   `0g000e48hq_fedex 04-26_20260512084332...csv` dans Excel. **En-tête
   ligne 1, colonnes EN ANGLAIS**, texte brut visible dans la barre de
   formule (cellule A1, avant parsing colonnes) : *"Payer Account, Invoice
   Month (yyyymm), OPCO, Service Type, Service Description, Pay Type,
   Shipment Date (mm/dd/yyyy), Shipment Delivery Date (mm/dd/yyyy),
   Shipment Tracking Number, Postal Identification Number, Shipper Name,
   Shipper Company Name, Shipper Address, Shipper City, Shipper
   State/Province, Shipper Country/Territory, Shipper Postal Code, Shipment
   Freight Charge Amount USD, Shipment Miscellaneous Charge Amount USD,
   Shipment Duty and Tax Charge Amount USD, Shipment Discount Amount USD,
   Net Charge Amount USD, Pieces In Shipment, Shipment Rated Weight
   (Pounds), Original Weight (Pounds), Proof Of Delivery Recipient,
   Recipient Name, Recipient Company Name, Recipient Address, Recipient
   City, Recipient State/Province, Recipient Country/Territory, Recipient
   Postal Code, Reference Notes Line 1, Reference Notes Line 2, Reference
   Notes Line 3, Department Number, PO Number, Pricing Zone, Shipment DIM
   Flag (Y or N), Dimmed Height (cm), Dimmed Width (cm), Dimmed Length
   (cm), Recipient Original Address, Recipient Original City, Recipient
   Original State/Province, Recipient Original Postal Code, Recipient
   Original Country/Territory, Shipment..."* — **CE FICHIER A DES
   COLONNES EN USD ("Amount USD"), pas "Billed Currency"**, ce qui suggère
   qu'il existe DEUX jeux de colonnes de montants dans ce CSV (USD et
   "Billed Currency"/EUR), la suite de l'en-tête continuant hors-écran.
3. (00:16 – 00:29) Suite de la lecture de l'en-tête complet (toujours
   cellule A1, texte continu) : après les colonnes ci-dessus, la suite
   contient bien un **second groupe de colonnes "...Billed Currency"** :
   *"...Invoice Date (mm/dd/yyyy), Invoice Number, Master Tracking Number,
   Domestic/Intl, Package Type, Shipment Delivery Time, **Shipment Freight
   Charge Billed Currency, Shipment Miscellaneous Charge Billed Currency,
   Shipment Duty And Tax Charge Billed Currency, Shipment Discount Billed
   Currency, Net Charge Billed Currency**, Billed Currency Code, Exchange
   Rate to USD, Weight Type Code, Customer Order Number"*. **CONFIRME** que
   le fichier source brut CSV a bien les 2 groupes de colonnes de montants
   (USD ET Billed Currency/EUR) mentionnés dans la mémoire projet — les
   noms "Shipment Freight Charge Billed Currency" etc. sont bien présents
   tels quels dans le CSV source, en toutes lettres dans l'en-tête, donc
   **PAS une traduction/un renommage fait manuellement** : les en-têtes
   anglais du xlsx modèle (`Shipment Detail`) proviennent directement du
   CSV source de ce type de fichier.
4. (00:29 – 00:41) Données collées/affichées en tableau (colonnes A à W
   visibles) : Payer Account = `200720433`, Invoice Month = `202604`,
   OPCO = `Express`, Service Type = `UO`, Service Description = `FedEx
   Priority`, Pay Type = `Bill_Sender_Prepaid`, dates, Shipment Tracking
   Number (ex. `8,7021E+11` — grand nombre tronqué en notation
   scientifique), Postal Identification Number = `NA`, Shipper
   Name/Company = `HYDRATIS B2B`, Shipper Address = `VOIE COMMUNALE 4`,
   Shipper City = `CR?ANCEY` (probablement "CRÉANCEY", problème d'encodage
   visible à l'écran), Shipper State = `21`, Shipper Country = `FR`,
   Shipper Postal Code = `21320`, colonnes de montants visibles (ex.
   48.69 / 4.77 / 0.00 / -37.06 / 16.40, Pieces = 3, poids = 36.9...).
5. (00:38 – 00:44) Boîte de dialogue **"Assistant Conversion — Étape 1 sur
   3"** visible : Excel propose de convertir le texte en colonnes, "Type de
   données d'origine" = **Délimité** sélectionné (par opposition à
   "Largeur fixe"). Confirme que le fichier CSV brut est bien
   délimité (par virgule) et doit être importé/converti via l'Assistant
   Texte natif d'Excel — cohérent avec `0g000e48hq_FEDEX 06-2026
   v2_20260710125723.csv` étant nommé de façon similaire au fichier vu ici
   (`0g000e48hq_fedex 04-26_20260512084332...csv`), suggérant que **ce type
   de fichier `0g000e48hq_...` EST LA SOURCE BRUTE UTILISÉE, avant
   traitement, pour construire "Shipment Detail"** (nom de export commun à
   avril et juin, préfixe `0g000e48hq_` = identifiant d'export FedEx
   récurrent).
6. (00:44 – 00:56) Fenêtre Excel réduite/changement de focus : passage à
   l'Explorateur de fichiers, dossier `2026 04` (niveau parent, tous
   transporteurs) : liste des sous-dossiers (BLS, CEVA, Chronopost,
   Colissimo, Delivengo, DHL, DPD, **Fedex**, Geodis, GLS, Kuehne, La
   Poste, Mondial Relay, TNT, UPS, UPS - 80X7Y5) et des fichiers
   `2026_04_Facture <Transporteur>.xlsx` pour chacun. Le fichier
   `2026_04_Facture Fedex.xlsx` (1559 Ko) est sélectionné puis ouvert.

## Partie 1 — Onglet "Import ERP" (aperçu initial) et onglet "Shipment
Detail" : structure et colonnes A-E (00:56 – 02:08)

7. (00:56 – 01:10) Premier onglet visible à l'ouverture : **"Import ERP"**
   (PAS "Import CSV" comme Colissimo — confirme que pour FedEx le nom
   d'onglet exact est bien "Import ERP", cohérent avec ce que l'appelant a
   déjà observé sur juin 2026). En-têtes confirmés visuellement : **A =
   "Transporteur" | B = "Date valid[ité tarif]" | C = "Réf.1" | D = "Réf.
   2" | E = "Id client" | F = "N° Tracking" | G = "Nom" | H = "E / P" |
   I = "Pays" | J = "Zone" | K = "Nbr Colis" | L = "Poids" | M = "mode
   env[oi]" | N = "TVA" | O = "Droits et [taxes]" | P = "Assuran[ce]" |
   Q = "Zones él[oignées]" | R = "Colis vol[umineux]" | S = "Adresse[s]"
   | T = "Frêt" | U = "plus-valu[e BtoC]" | V = "Gazole" | W = "Nb colis"**
   — mapping IDENTIQUE au motif déjà vu chez Colissimo (mêmes libellés de
   colonnes O à V), confirmant un template Import ERP standard partagé
   entre transporteurs.
8. Données visibles lignes 2-19 : Transporteur alterne entre **"FEDEX
   INTERNATIONAL"** et **"FEDEX FRANCE"**. Colonne E/P alterne entre
   **"P"** et **"E"** (lettres seules ici, pas "entreprise"/"ST" — à noter,
   possible différence de représentation entre "Import ERP" (P/E) et
   "Shipment Detail" (particulier/entreprise + ST/FICP), cf. point 10-11
   ci-dessous). Colonne Pays montre des codes 2 lettres (**DK, IE, IT, PL,
   SE, FR, GR, US, CA**) et colonne Zone montre des valeurs **T, S, U, H,
   A, France** — cohérent avec un système de zonage par pays/mode. Colonne
   "mode env[oi]" montre **IE, FICP, International** comme valeurs
   observées (ex. ligne DK → IE ; ligne PL → IE ; ligne US → FICP ; ligne
   CA → International) : **confirme que "FICP" est bien une valeur
   possible de la colonne mode envoi (M)**, cohérent avec la logique
   Transporteur = FEDEX INTERNATIONAL déjà actée par l'appelant, mais
   montre aussi "IE" et "International" comme valeurs distinctes de
   "FICP" dans cette même colonne pour des lignes classées FEDEX
   INTERNATIONAL — **nuance à vérifier** : la règle Transporteur ∈
   {FICP,IE,RE,International,Europe} → FEDEX INTERNATIONAL semble
   cohérente avec ces observations (IE et International sont bien dans la
   liste), mais aucune ligne "RE" ou "Europe" n'a été repérée dans les
   frames disponibles pour confirmer ces 2 valeurs précises.
9. Colonne TVA = **0,2** pour les lignes FEDEX FRANCE (pays FR) observées,
   et **0** pour les lignes FEDEX INTERNATIONAL (DK, IE, IT, PL, SE, US,
   CA) — confirme la règle déjà actée (TVA 20% si pays UE France
   domestique observé, 0% sinon dans cet échantillon ; **attention** :
   les pays UE hors France (IE, IT, PL, SE = tous UE) montrent ici TVA=0,
   ce qui semble contredire "TVA=20% si pays UE" tel que formulé par
   l'appelant — **à vérifier**, possible que la règle TVA dépende plutôt
   du couple (Zone FICP vs FR) que du seul statut UE du pays, cf. section
   Points ambigus.

## Partie 2 — Onglet "Shipment Detail" : structure et formule E/P (01:10 –
02:20)

10. (01:10 environ) Bascule sur l'onglet **"Shipment Detail"**. En-têtes de
    la ligne 1 (colonnes A à W visibles dans une frame) : **A = "Montant"
    | B = "Clients" | C = "Catégor[ie]" | D = "E/P" | E = "Mode d'[envoi]"
    | F = "Payer Acc[ount]" | G = "Invoice Mo[nth]" | H = "OPCO" |
    I = "Service Ty[pe]" | J = "Service Descriptio[n]" | K = "Pay Type" |
    L = "Shipment D[ate]" | M = "Shipment D[elivery date]" |
    N = "Shipment Tracking Number" | O = "Postal Iden[tification]" |
    P = "Shipper Name" | Q = "Shipper Company Name" | R = "Shipper
    Address" | S = "Shipper Ci[ty]" | T = "Shipper St[ate]" | U = "Shipper
    Co[untry]" | V = "Shipper Po[stal code]" | W = "Shipment..."**.
    **DIFFÉRENCE avec le fichier juin 2026 déjà inspecté par l'appelant** :
    dans cette vidéo (avril 2026), les colonnes de calcul en tête sont
    **A (Montant), B (Clients), C (Catégorie), D (E/P), E (Mode
    d'envoi)** — 5 colonnes, alors que l'appelant avait identifié pour
    juin 2026 seulement A "Montant" (=BK+BN), C "Catégorie" (vide), D
    "E/P" (littéral). **Ici B "Clients" et E "Mode d'envoi" sont des
    colonnes de calcul supplémentaires visibles, absentes de la
    description initiale de l'appelant** — à vérifier si elles existent
    aussi dans le fichier juin 2026 (peut-être simplement pas mentionnées
    car non focalisées par l'appelant, ou réellement absentes/ajoutées
    entre-temps).
11. **Colonne C "Catégor[ie]" NE SEMBLE PAS toujours vide dans ce fichier
    avril 2026** : la frame montre la colonne C remplie avec la valeur
    **"ST"** sur de nombreuses lignes consécutives (ex. lignes 2-43 toutes
    à "ST"), avec occasionnellement **"FICP"** (ligne 2031) — **CONTREDIT
    l'observation de l'appelant sur juin 2026 ("Catégorie" toujours
    vide sur 3607 lignes)**. Deux hypothèses : (a) la colonne "Catégorie"
    a été vidée/dépréciée entre avril et juin 2026 (changement de process
    dans le temps), ou (b) il y a une confusion de colonne entre ce que la
    vidéo montre en colonne C et ce que l'appelant a inspecté — **à
    vérifier en priorité avec le pôle transport**, car cela remet en
    question le statut "vestige jamais utilisé" supposé par l'appelant.
    Notez que les valeurs observées en colonne C ("ST", "FICP") sont très
    proches des valeurs de mode envoi/zone déjà vues dans "Import ERP"
    (T, S, U, H, A) et de la colonne E "Mode d'envoi" de Shipment Detail
    elle-même (valeurs ST/FICP également) — **la colonne C semble en fait
    dupliquer ou précéder la colonne E "Mode d'envoi"**, voir point
    suivant pour la formule exacte de E.
12. **Colonne E "Mode d'envoi" — formule confirmée exactement** :
    `=SI(J2044="FedEx Priority";"ST";"FICP")` (observée en cellule
    `E2044`, barre de nom de cellule lisible) — donc la colonne E dérive
    de la colonne **J "Service Description"** : si "FedEx Priority" →
    "ST" (Standard/Service Type national ?), sinon → "FICP" (International
    Connect Plus / FedEx International Connect Plus, cf. valeurs "FedEx
    International Pr[iority]"/"FedEx Intl Connect Pl[us]" observées en
    colonne J pour les lignes classées FICP, Partie 4). **NUANCE
    IMPORTANTE sur la colonne D "E/P"** : l'appelant pensait cette colonne
    "littérale" (pas une formule) — **CECI EST EN PARTIE CONTREDIT PAR LA
    VIDÉO**, cf. point 13 : la colonne D "E/P" du classeur avril 2026 est
    bien remplie par une formule RECHERCHEX (XLOOKUP), pas une saisie
    manuelle littérale.
13. **Formule de la colonne D "E/P" — confirmée exactement et
    intégralement lisible dans une frame ultérieure (cellule D35)** :
    ```
    =RECHERCHEX(N35;'[2026 01 - Export expéditions_brut.xlsx]exportDemandeExpedition_2026020'!$AP:$AP;'[2026 01 - Export expéditions_brut.xlsx]exportDemandeExpedition_2026020'!$Q:$Q;"")
    ```
    (cette frame précise montre en fait une variante utilisant le fichier
    **`2026 01 - Export expéditions_brut.xlsx`**, onglet
    `exportDemandeExpedition_2026020` — l'opérateur a testé/comparé
    plusieurs fichiers mensuels d'export lors de la construction de cette
    formule, cf. point 15). **CONFIRME DE FAÇON DÉCISIVE le mécanisme "E/P
    via export brut m/m-1/m-2" déjà mentionné dans `registry.js`** : la
    colonne D "E/P" de "Shipment Detail" est calculée par
    **RECHERCHEX(N — tracking FedEx ; colonne AP "PRO_TRACKING" du fichier
    externe ; colonne Q du même fichier externe ; "" si non trouvé)**.
    Résultat observé dans les cellules réelles (colonne D, plusieurs
    lignes) : valeurs **"entreprise"** ou **"particulier"** (texte complet
    en minuscules) — **donc la colonne D n'est PAS une valeur littérale
    figée mais bien le résultat d'une formule XLOOKUP vers un classeur
    externe**, contredisant l'hypothèse de départ de l'appelant. La
    colonne **Q** du fichier externe "Export expéditions_brut.xlsx" est
    donc la colonne qui contient directement le texte "entreprise"/
    "particulier" (à nommer précisément, non lisible dans les frames
    disponibles — l'en-tête de la colonne Q n'a jamais été cadré
    nettement, cf. Points ambigus).
14. Une frame antérieure (cellule **D2**, en cours d'édition) montre le
    DÉBUT de la même formule en train d'être tapé/vérifié :
    `=RECHERCHEX(N2;` avec l'infobulle d'aide de fonction Excel affichée
    (`RECHERCHEX(valeur_cherchée; tableau_recherche; tableau_renvoyé;
    [si_introuvable]; [mode_correspondance]; [mode_recherche])`), puis
    annulée (bouton rouge "X" dans la barre de formule) avant d'être
    validée sous sa forme finale (point 13) — confirme que cette formule
    est reconstruite/vérifiée MANUELLEMENT chaque mois par l'opérateur
    (pas une formule figée copiée automatiquement), avec RÉFÉRENCE
    EXPLICITE au fichier externe du mois concerné.
15. Le fichier externe **`Export expéditions_brut.xlsx`** existe en
    plusieurs versions mensuelles simultanément ouvertes (visible dans la
    barre des tâches Excel avec 5 fenêtres empilées lors d'une frame :
    `0g000e48hq_fedex 04-26...`, `2026_04_Facture Fe...`, **`2026 04 -
    Export expéditio...`**, **`2026 03 - Export expé...`**, **`2026 02 -
    Export expéditio...`**, **`2026 01 - Export expéditio...`**) —
    **confirme explicitement le mécanisme "m/m-1/m-2" (voire m-3)** déjà
    supposé par l'appelant : l'opérateur garde/compare plusieurs mois
    d'export bruts simultanément pour retrouver, par tracking, le
    caractère entreprise/particulier d'une expédition qui pourrait
    remonter à un mois précédent (si le tracking n'existe pas dans
    l'export du mois M, chercher dans M-1, M-2, M-3...). Les noms d'onglet
    internes suivent un schéma `exportDemandeExpedition_20260NN` où NN
    semble incrémenter d'environ +10 à chaque mois (`2026020` pour janvier,
    `2026040` pour mars vu précédemment, `2026050` pour avril) — un
    identifiant d'export séquentiel, pas directement le numéro du mois.
16. En-têtes de colonnes du fichier externe **`Export expéditions_brut.xlsx`**
    (partiel, colonnes AC à AT, lues sur une frame) : **AC = "DES_ETAT" |
    AD = "DES_RELA[TION]" | AE = "INFO_FAC[TURE]" | AF = "INFO_COM[MANDE]"
    | AG = "INFO_DES[CRIPTION]" | AH = "INFO_NB[COLIS]" | AI =
    "INFO_POI[DS]" | AJ = "GEN_ALC[...]" | AK = "GEN_DRO[IT]" | AL =
    "GEN_VAL[EUR]" | AM = "GEN_MOD[E]" | AN = "GEN_ASS[URANCE]" | AO =
    "GEN_COM[...]" | AP = "PRO_TRACKING" | AQ = "ETAT_EXP[EDITION]" | AR =
    "TYPE_EXP[EDITION]" | AS = "TRANSPO[RTEUR]" | AT = "PRECOLISAGE"**.
    Colonne **AS "TRANSPORTEUR"** montre des valeurs variées tous
    transporteurs confondus : CHRONOPOST, KUEHNE, GEODIS,
    MONDIAL-RELAY-FOUTAS, COLISSIMO-APIDURA, DPD-APIDURA — **ce fichier
    "Export expéditions_brut" est donc un export GLOBAL multi-
    transporteurs de gestion des expéditions (probablement l'outil
    métier/TMS interne "La Ruche"), PAS spécifique à FedEx**, utilisé ici
    uniquement pour retrouver, via le tracking, le statut entreprise/
    particulier de l'expédition (colonne D "E/P" de Shipment Detail).
    Colonne "TYPE_EXP" (AR) montre des valeurs **P** et **C** (candidat
    P=Particulier probable, mais la formule confirmée pointe vers la
    colonne **Q** — pas AR — pour la valeur réellement renvoyée, cf. point
    13 ; la colonne AR n'est donc PAS la source de "entreprise"/
    "particulier", ou alors une autre colonne dérivée de AR se trouve en
    Q, hors du cadre visible dans les frames disponibles).
17. **Colonne A "Montant" — formule confirmée EXACTEMENT identique à
    l'observation de l'appelant** : plusieurs frames montrent la formule
    en barre de formule pour différentes lignes, toutes de la forme
    `=BK<n>+BN<n>` (ex. `=BK38+BN38`, `=BK2044+BN2044`, `=BK2045+BN2045`,
    `=BK2666+BN2666`) — **CONFIRME EXACTEMENT** `A = BK + BN` (Fret HT +
    Remise) pour le fichier avril 2026 également, cohérent avec
    l'observation de l'appelant sur juin 2026.
18. Une frame montre une opération de **Rechercher/Remplacer** en cours
    (boîte de dialogue Excel "Rechercher et remplacer", onglet
    "Remplacer") avec un remplacement en cours sur les FORMULES de la
    feuille (option "Regarder dans : Formules") — action non identifiée
    avec certitude (recherche `.` remplacé par `.`? ou remplacement d'une
    référence de plage), probablement liée à une correction de formule en
    masse après collage des nouvelles lignes du mois — **à vérifier
    directement dans le fichier réel, cette action précise n'est pas
    confirmée avec certitude** (cf. Points ambigus).
19. **Le CSV source `0g000e48hq_fedex 04-26_20260512084332.csv` compte
    2676 lignes de données** (dernière ligne de données visible : ligne
    2677 vide, la barre d'état Excel affiche "Nb (non vides) : 2676" avec
    la colonne N "Shipment Tracking Number" sélectionnée en entier) — ceci
    est DIFFÉRENT du nombre de lignes de "Shipment Detail" observé par
    l'appelant pour juin 2026 (3607 lignes), mais il s'agit ici du mois
    d'avril (volumétrie différente d'un mois sur l'autre, donc **PAS
    directement comparable** au nombre de lignes de juin). Confirme que
    ce fichier CSV a bien les colonnes **BK "Shipment Freight Charge
    Billed Currency" | BL "Shipment Miscellaneous Charge Billed Currency"
    | BM "Shipment Duty And Tax Charge Billed Currency" | BN "Shipment
    Discount Billed Currency" | BO "Net Charge Billed Currency"** en
    positions EXACTEMENT identiques aux colonnes du même nom dans l'onglet
    "Shipment Detail" du classeur — les colonnes de montants du CSV brut
    ET du classeur retravaillé partagent donc les MÊMES lettres de colonne
    (décalées uniquement du nombre de colonnes de calcul ajoutées en tête,
    ici 5 : A-E), ce qui explique pourquoi la formule `=BK+BN` fonctionne
    de façon identique après un simple copier-coller des données brutes à
    partir de la colonne F.

## Partie 3 — Onglet "Bilan factures" : structure TCD et réconciliation
PDF (~05:00 – 05:20 vidéo, observé lors d'un aller-retour d'onglets)

20. Onglet **"Bilan factures"** : contrairement à Colissimo/Chronopost où
    cet onglet est un TCD simple à 2 colonnes natives + colonnes de
    contrôle manuel, ici c'est un **vrai TCD avec DEUX DataFields** :
    **Lignes = "Invoice Number"** (numéro de facture, ex. 634243207,
    634243603, 634251047, 634251402, 634259662, 634260023, 634268737,
    634278311, 634278838, 634286408, "(vide)") | **Colonnes B "Frêt HT"
    et C "Total TTC"** (2 champs de valeurs du TCD, probablement Somme de
    Montant et Somme de Net Charge/BO) | **D "PDF"** (saisi manuellement
    depuis le PDF réel de la facture FedEx) | **E "Ecart"** = formule
    (visible dans le panneau des champs TCD : Lignes = "Invoice Number",
    Valeurs = "Frêt HT" et "Total TTC"). **11 lignes de factures pour
    avril 2026** (10 numéros + 1 ligne "(vide)"), alors que seulement **9
    PDF de factures** étaient présents dans le dossier réseau (Partie 0,
    point 1) — **ÉCART CONFIRMÉ ET COHÉRENT avec l'observation de
    l'appelant sur juin 2026** (9 factures dans "Bilan factures" mais
    seulement 8 PDF disponibles) : il semble donc SYSTÉMATIQUE qu'il y ait
    plus de lignes "Invoice Number" dans le TCD que de PDF téléchargés
    dans le dossier réseau — cause exacte non confirmée par la vidéo (PDF
    pas encore téléchargés au moment de la capture ? factures à 0€ non
    utiles ? erreur de récupération PDF sur le portail FedEx ?), **à
    vérifier avec le pôle transport**.
21. Valeurs exactes observées dans le TCD "Bilan factures" (avril 2026) :
    | Invoice Number | Frêt HT | Total TTC | PDF | Ecart |
    |---|---|---|---|---|
    | 634243207 | 102,16 € | 177,92 € | 177,92 | - € |
    | 634243603 | 5 156,31 € | 9 350,82 € | 9350,82 | - € |
    | 634251047 | 88,00 € | 230,29 € | 230,29 | - € |
    | 634251402 | 3 989,28 € | 8 295,23 € | 8295,23 | - € |
    | 634259662 | 103,01 € | 166,39 € | 166,39 | - € |
    | 634260023 | 4 620,77 € | 9 690,80 € | 9690,8 | - € |
    | 634268737 | 6 511,79 € | 13 454,00 € | 13454 | - € |
    | 634278311 | 116,91 € | 182,39 € | 182,39 | - € |
    | 634278838 | 7 605,07 € | 16 072,82 € | (vide) | (vide) |
    | 634286408 | 119,77 € | 193,91 € | (vide) | (vide) |
    | (vide) | | | | |
    | **Total général** | **28 413,07 €** | **57 814,57 €** | | |
    La colonne **D "PDF"** est bien une **valeur TTC** collée à la main
    depuis le PDF réel (confirme la remarque déjà actée dans la mémoire
    projet : "le Total HT affiché est en réalité un Total TTC" — ici la
    colonne s'appelle explicitement "PDF" sans préciser HT/TTC, mais les
    valeurs collées, ex. 177,92 / 9350,82, correspondent numériquement à
    la colonne C "Total TTC" du TCD, PAS à la colonne B "Frêt HT" — donc
    **"PDF" = comparaison contre le Total TTC**, pas contre le Frêt HT).
    La colonne **E "Ecart"** = `Total TTC - PDF` (ou l'inverse), affichée
    à "- €" (zéro) pour les 8 premières factures qui ont un PDF renseigné,
    confirmant l'égalité stricte attendue. Pour les 2 dernières factures
    (634278838 et 634286408), PDF et Ecart sont VIDES — cohérent avec
    l'absence de PDF correspondant constatée dans le dossier réseau
    (Partie 0). **Ceci confirme et NUANCE l'observation de l'appelant** :
    ce n'est pas 9 factures vs 8 PDF mais, pour ce mois d'avril en tout
    cas, **11 lignes de facture (10 numéros + 1 vide) vs 9 PDF
    disponibles**, soit un écart de 2 factures non justifiées par PDF (et
    non 1 comme sur juin) — le mécanisme structurel semble le même
    (toujours au moins 1-2 factures "en retard" par rapport aux PDF
    disponibles au moment de la préparation du fichier).
22. Le panneau des champs TCD (volet droit Excel "Champs de tableau croisé
    dynamique") confirme les champs disponibles pour ce TCD : **Montant,
    Clients, Catégorie, E/P, Mode d'envoi, Payer Account, Invoice Month
    (yyyymm), OPCO, Service Type...** — la source du TCD "Bilan factures"
    est donc bien l'onglet **"Shipment Detail"** (mêmes noms de colonnes
    A-J que Shipment Detail), avec **Lignes = "Invoice Number"** et
    **Valeurs = "Frêt HT"** (renommage d'affichage du champ "Montant"
    probablement, à confirmer) et **"Total TTC"** (renommage d'affichage
    d'un autre champ, probablement "Net Charge Billed Currency"/colonne
    BO, à confirmer directement dans le fichier réel).

## Partie 4 — Onglet "TCD" : structure et formules (07:00 – 08:10)

23. Onglet **"TCD"** : vrai TCD avec **Lignes = "Shipment Tracking Number"**
    et deux niveaux (colonne A regroupée), **B = "Recipient
    Country/Territory"** (pays destinataire, ex. DK, IE, IT, PL, SE, FR,
    GR, US, CA), **C = "Somme de Montant"**, **D = "Somme de Shipment
    Rated Weight (Pounds)"** (poids en LIVRES, natif du TCD) — **CONFIRME
    EXACTEMENT** la structure "RowFields=Tracking+Pays, DataFields=Somme
    de Montant+Somme de Poids" déjà identifiée par l'appelant sur juin
    2026 (à ceci près qu'ici le champ Pays s'appelle explicitement
    "Recipient Country/Territory", et non un simple "Pays").
24. **Colonne E "Poids arrondi sup (kg)" — formule confirmée EXACTEMENT**
    (barre de formule lisible sur plusieurs lignes, ex. cellules E24,
    E2044) :
    ```
    =ARRONDI.SUP(D24*0,453592;1)
    ```
    — **CONFIRME EXACTEMENT** la formule déjà identifiée par l'appelant
    (`=ROUNDUP(D*0.453592,1)`, ici en version française `ARRONDI.SUP`),
    conversion livres → kg avec arrondi au dixième supérieur.
25. **Colonne F "E/P" — formule confirmée EXACTEMENT** (cellule F3,
    barre de formule lisible) :
    ```
    =SI(RECHERCHEX(A3;'Shipment Detail'!N:N;'Shipment Detail'!D:D)="entreprise";"E";"P")
    ```
    — **CONFIRME EXACTEMENT** le mécanisme déjà supposé par l'appelant
    ("F = E/P via XLOOKUP du tracking vers Shipment Detail colonne D") :
    recherche du tracking (colonne A du TCD) dans la colonne **N**
    "Shipment Tracking Number" de "Shipment Detail", renvoie la colonne
    **D** "E/P" de Shipment Detail (déjà elle-même une formule XLOOKUP
    vers le fichier externe "Export expéditions_brut.xlsx", cf. Partie 2) ;
    si le résultat est le texte "entreprise" → **"E"**, sinon → **"P"**.
    **Ceci relie explicitement les 3 couches** : fichier externe TMS
    (colonne "entreprise"/"particulier") → Shipment Detail colonne D
    (XLOOKUP vers le TMS) → TCD colonne F (SI/XLOOKUP vers Shipment Detail,
    condensé en "E"/"P" une seule lettre) → Import ERP colonne H "E/P"
    (probablement XLOOKUP vers TCD colonne F, à confirmer Partie 5).
26. Valeurs observées dans le TCD (échantillon) : tracking `380050840782`
    (DK) → Poids 5,94 lb → 2,7 kg → P (particulier) ; `380060046816` (IE)
    → 21,12 lb → 9,6 kg → P ; `380062847070` (IT) → 7,92 lb → 3,6 kg → P ;
    `380064443798` (PL) → 5,06 lb → 2,3 kg → **E** (entreprise) ;
    `380283487240` (FR) → 22 lb → 10 kg → P ; `870042019264` (FR) → 15,18
    lb → 6,9 kg → **E**. Confirme la cohabitation de lignes "E" et "P"
    pour un même pays (FR), donc le statut E/P est bien indépendant du
    pays — dépend uniquement du TMS externe (type de client final).

## Partie 5 — Détour par le TMS interne "si.laruche-logistique.fr"
(Système d'Information de La Ruche Logistique) (08:10 – 08:40)

27. **Découverte majeure, absente de la liste de points à vérifier
    initiale** : l'opérateur bascule vers un onglet de navigateur déjà
    ouvert sur **`si.laruche-logistique.fr`** (identifié en haut de
    l'écran : "Bienvenue Caroline SCHMITT"), le TMS/ERP interne "Système
    d'Information de La Ruche Logistique", menu **Expéditions >
    Expéditions**, sur la fiche détaillée d'une expédition WMS
    (`EXP20260328-2675251`), onglet **"Devis"** de cette fiche. Le devis
    affiche explicitement :
    - **Transporteur : FEDEX INTERNATIONAL**
    - **Mode de livraison : FICP**
    - **Zone : T**
    - Assurance : 0 €
    - Zone éloignée : 0 €
    - Colis volumineux : 0 €
    - Plus-value Nb. colis : 0 €
    - **Frêt : 17,12 €**
    - Taxe gasoil : 0 € (0 %)
    - **Total HT : 17,12 €**
    - Note en bas de page : *"Hors droits et taxes, hors emballages et
      sous réserve de l'exactitude des informations renseignées : adresse
      de destination, **statut du destinataire (particulier ou
      entreprise)**, poids du/des colis, etc..."* et *"La taxe gazole
      varie chaque semaine. Le montant facturé sera celui correspondant à
      la semaine de départ de vos colis."*
    **CONFIRME DE FAÇON DÉCISIVE et NOUVELLE (absente des points à
    vérifier initiaux)** : le TMS interne "La Ruche Logistique" calcule
    LUI-MÊME, au moment de la demande d'expédition, un **devis prévisionnel
    par transporteur** avec Mode de livraison (FICP/ST/IE/...), Zone, et
    séparément un **statut destinataire "particulier ou entreprise"**
    explicitement mentionné comme donnée d'entrée du devis — ce texte
    "statut du destinataire (particulier ou entreprise)" dans la note du
    TMS **est la source directe et la définition officielle de la
    colonne E/P** que la formule Excel va ensuite chercher via le fichier
    "Export expéditions_brut.xlsx" (qui est très probablement un export
    de CE MÊME TMS, contenant une colonne dérivée de ce "statut du
    destinataire"). Ceci explique pourquoi la colonne D "E/P" de Shipment
    Detail est une recherche externe et non une donnée FedEx native : le
    statut entreprise/particulier est une donnée SAISIE PAR LE CLIENT
    EXPÉDITEUR lors de la demande d'expédition dans le TMS La Ruche, PAS
    une donnée fournie par FedEx dans son propre export de facturation.
28. Ce même écran confirme aussi, de façon INDÉPENDANTE de tout calcul
    Excel, que **"FICP" est bien un "Mode de livraison"** proposé/choisi
    par le TMS pour FedEx International — cohérent avec toutes les
    observations précédentes (Import ERP colonne M, Shipment Detail
    colonne E, TCD colonne F).
29. **Le champ "Zone : T"** du devis TMS est à rapprocher de la colonne
    "Zone" déjà observée dans "Import ERP" (valeurs T, S, U, H, A, France,
    cf. Partie 1, point 8) — **CONFIRME/NUANCE** l'observation de
    l'appelant sur le mécanisme de zonage FedEx (`XLOOKUP(Pays,
    Zoning!B:B, Zoning!C:C)`) : il existe donc, en amont du classeur Excel,
    un système de zonage déjà calculé au niveau du TMS pour chaque
    expédition (zone "T" ici pour un envoi FICP) — la formule Zoning du
    classeur Excel recalcule peut-être cette même zone a posteriori à
    partir du seul pays de destination (moins précise que le TMS qui
    connaît le mode de livraison exact), ou sert de RECOUPEMENT/CONTRÔLE.
    **Point à vérifier avec le pôle transport** : la formule Zoning du
    classeur Excel (XLOOKUP Pays → Zone, colonne C "FICP" uniquement) est-
    elle sensée reproduire fidèlement cette zone "T" calculée par le TMS,
    ou s'agit-il d'un système de zonage différent/simplifié ? Ceci
    éclaire aussi la question déjà posée par l'appelant sur le Zoning
    ("seule la colonne C FICP est utilisée même pour IE/RE") : peut-être
    que TOUTES les expéditions FedEx INTERNATIONAL (FICP, IE, RE...)
    utilisent en réalité la MÊME table de zonage FICP dans le TMS, ce qui
    justifierait que la formule Excel ne référence QUE la colonne C.


## Partie 6 — Onglet "Import ERP" : formules détaillées colonne par colonne
(08:40 – 10:10)

30. **Colonne A "Transporteur" — formule confirmée EXACTEMENT** (cellule
    A3, barre de formule intégralement lisible) :
    ```
    =SI(OU(M3="FICP";M3="IE"; M3="RE"; M3="International"; M3="Europe");"FEDEX INTERNATIONAL";"FEDEX FRANCE")
    ```
    — **CONFIRME EXACTEMENT ET INTÉGRALEMENT** la règle déjà connue de
    l'appelant : Transporteur = "FEDEX INTERNATIONAL" si le "mode envoi"
    (colonne M) fait partie de {FICP, IE, RE, International, Europe},
    sinon "FEDEX FRANCE". Les 5 valeurs candidates sont donc bien FICP/IE/
    RE/International/Europe, telles quelles dans le classeur avril 2026.
31. **Colonne F "N° Tracking" — formule confirmée** : `=TCD!A17` (exemple
    observé en F17) — référence DIRECTE au TCD, colonne A (Shipment
    Tracking Number), décalage de ligne cohérent avec l'en-tête du TCD.
    Confirme le même mécanisme "Import ERP alimenté depuis TCD" déjà
    observé chez Colissimo.
32. **Colonne D "Réf.1" — action de comblement en masse observée** :
    l'opérateur sélectionne toute la colonne D (vide au départ) et
    remplit avec Ctrl+D/glisser vers le bas (icône verte de recopie
    visible dans une frame). Le contenu exact de la formule collée en D3
    n'a pas été lisible net dans les frames (cellule vide au moment du
    screenshot), mais le comportement (recopie en masse après sélection)
    est identique aux étapes déjà vues chez Colissimo pour "Étendre les
    colonnes".
33. **Bug de recopie au-delà des données réelles — cascade #N/A/"inconnu"
    OBSERVÉE, IDENTIQUE au pattern déjà documenté chez Colissimo** :
    plusieurs frames (lignes 2674 à 2707+) montrent la colonne **K "Nbr
    Colis"** avec la formule
    ```
    =RECHERCHEX(F2674;'Shipment Detail'!N:N;'Shipment Detail'!AB:AB)
    ```
    affichant **#N/A** en cascade, la colonne N° Tracking (F) affichant
    **(vide)** ou **#N/A**, la colonne Pays (I) affichant **"inconnu"**
    (avec triangle d'avertissement Excel visible), toutes les colonnes de
    montant (O à W) affichant **0** ou **"- euros"**. Ce bloc de lignes
    "fantômes" s'étend de la ligne 2674 jusqu'à au moins 2707 (limite du
    cadre visible) — cohérent avec le fait que le TCD (donc "Shipment
    Detail") s'arrête à la ligne 2673 (confirmé Partie 4/6, la dernière
    ligne réelle du CSV import final étant 2673, cf. point 36) et que les
    formules colonnes A-W de "Import ERP" ont été étirées PLUS BAS que
    cette limite. **Il s'agit très probablement d'un bug de recopie
    identique à celui déjà documenté chez Colissimo** (Partie 9 de la
    transcription Colissimo) : les formules XLOOKUP sont recopiées trop
    loin, générant une cascade d'erreurs #N/A/"inconnu" sur les lignes
    vides en dessous des données réelles — **ce bug n'a PAS été vu corrigé
    explicitement dans les frames disponibles pour FedEx** (contrairement
    à Colissimo où une suppression de lignes est filmée), mais le fichier
    CSV final exporté (`2026_04_Fedex_Import.csv`, cf. point 36) NE
    contient PAS ces lignes fantômes — donc soit la correction a eu lieu
    hors du cadre capturé par les frames à 6-8s d'intervalle, soit le
    copier-coller final "valeurs seules" vers le CSV exclut naturellement
    ces lignes (à vérifier).
34. **Colonne K "Nbr Colis"** — la formule visible utilise RECHERCHEX vers
    'Shipment Detail'!N:N (tracking) vers 'Shipment Detail'!AB:AB —
    colonne **AB** de Shipment Detail, dont le nom exact n'a pas été
    confirmé dans une frame nette (probablement "Pieces In Shipment", à
    vérifier directement dans le fichier réel).
35. Le fichier `2026_04_Facture Fedex.xlsx` est ensuite **copié en
    valeurs** (probablement Ctrl+C puis collage spécial valeurs, action
    non captée nettement image par image) vers un nouveau classeur/fichier
    nommé **`2026_04_Fedex_Import.csv`**, ouvert et vérifié séparément :
    en-têtes identiques à l'onglet "Import ERP" (Transporteur, Date
    validité tarif, Réf.1, Réf. 2, Id client, N° Tracking, Nom, E/P,
    Pays, Zone, Nbr Colis, Poids, mode envoi, TVA, Droits et taxes,
    Assurance, Zones éloignées, Colis volumineux, Adresses, Frêt,
    plus-value BtoC, Gazole, Nb colis). **Colonne B "Date validité"**
    affiche la date en TEXTE "01/04/2026" pour les premières lignes,
    PUIS un nombre entier **"46113"** (numéro de série Excel de date, non
    reformaté) pour les lignes suivantes — signe d'un collage
    hétérogène/incomplet lors de la construction du CSV final (valeurs
    de date collées sans harmonisation de format sur toutes les lignes).
36. **Nombre de lignes du CSV final confirmé par la barre d'état Excel** :
    en sélectionnant la colonne A entière du fichier
    `2026_04_Fedex_Import.csv`, la barre d'état affiche **"Nb (non vides)
    : 2673"** — donc le fichier CSV d'import final pour avril 2026
    contient **2673 lignes de données** (hors en-tête), très proche des
    2672 lignes réelles de "Shipment Detail" (dernière ligne de données
    observée : ligne 2673, cf. Partie 2 point 19 : 2676 lignes pour le
    CSV source BRUT initial, mais après nettoyage/dédoublonnage,
    "Shipment Detail"/TCD/Import ERP se stabilisent à 2672-2673 lignes
    utiles — écart de quelques lignes probablement des lignes d'en-tête/
    pied de page ou doublons filtrés lors du copier-coller initial).
    **Ce nombre (2672-2673) est à comparer avec le nombre de lignes du
    fichier juin 2026 (3607 lignes selon l'appelant)** : les deux
    volumétries sont plausibles pour des mois différents (avril vs juin),
    donc PAS directement contradictoires — le nombre de lignes SEUL ne
    permet pas d'identifier quel CSV source (`FEDEX 06 2026.csv` avec
    3283/3284 lignes, ou `0g000e48hq_FEDEX 06-2026 v2...csv` avec 4001
    lignes) correspond exactement aux 3607 lignes de "Shipment Detail"
    pour juin — voir Points ambigus, ce point spécifique à juin 2026 n'a
    PAS pu être tranché par cette vidéo qui montre avril 2026.
37. **Confirmation supplémentaire du nom du fichier CSV brut source pour
    avril** : lors d'une recherche Explorateur Windows dans le dossier
    réseau (`$Facturation automatique\1 - Factures transporteurs +
    calculs\2026\2026 04\Fedex`), un SEUL fichier CSV nommé de type
    `0g000e48hq_...` par mois est utilisé comme source pour "Shipment
    Detail" (celui vu en Partie 0, `0g000e48hq_fedex 04-26_2026051208433
    2...csv`, 2676 lignes) — il n'y a pas eu de fusion de 2 fichiers CSV
    distincts dans cette vidéo (contrairement à Colissimo qui fusionne
    "Prestations au colis" + "Frais de douane") : **pour FedEx, un SEUL
    CSV brut mensuel (préfixe `0g000e48hq_`) alimente directement
    "Shipment Detail"**, avec les colonnes BK-BO ("Billed Currency")
    déjà présentes nativement dans ce CSV.
38. Second fichier CSV vu dans le dossier réseau avril
    (`0g000e48hq_Fedex 04-26_202605...csv`, l'autre fichier CSV du
    dossier, non ouvert/utilisé dans les frames disponibles) — cohérent
    avec l'observation de l'appelant sur juin 2026 où 2 fichiers CSV bruts
    sont également présents (`FEDEX 06 2026.csv` et `0g000e48hq_FEDEX
    06-2026 v2_...csv`) : il semble donc que FedEx (ou le processus
    d'extraction du portail FedEx) génère systématiquement 2 exports CSV
    par mois, mais **seul le second (nommé `0g000e48hq_...`, avec les
    colonnes anglaises complètes dont "Billed Currency") est utilisé pour
    construire "Shipment Detail"** — l'autre fichier CSV (nommé
    différemment, sans le préfixe `0g000e48hq_`) n'a jamais été ouvert
    dans cette vidéo, son usage reste à confirmer (peut-être un export
    redondant/de secours, ou un format différent non retenu).


## Partie 7 — QA finale sur le CSV d'import (10:10 – 11:55, fin de vidéo)

39. (10:10 – 11:10) Retour sur le fichier `2026_04_Fedex_Import.csv` :
    l'opérateur applique des **filtres automatiques successifs** pour
    contrôle qualité :
    - Filtre sur colonne **F "N° Tracking"** : confirme le nombre total de
      lignes de données = **2673** (barre d'état "Nb (non vides) : 2673",
      colonne sélectionnée en entier).
    - Filtre sur colonne **M "mode envoi"** ne conservant que **FICP** :
      la liste de valeurs proposée dans le filtre montre exactement
      **3 valeurs distinctes** pour "mode envoi" dans le fichier final :
      **FICP, IE, ST**. Une fois filtré sur FICP seul, **4 lignes**
      trouvées (ligne 20 US/H, ligne 22 CA/A, ligne 25 US/H, et une ligne
      988 avec tracking `870473797522`, pays **NG** (Nigéria), zone **E**,
      poids 4,8 kg, frêt 73,04€) — **DÉCOUVERTE NOUVELLE, absente des
      points à vérifier initiaux** : la colonne "Zone" (J) prend, pour les
      lignes FICP, des valeurs **T, S, U, H, A, E** (au moins 6 lettres de
      zone distinctes), pas seulement les zones observées ailleurs — la
      lettre "E" comme ZONE (à ne pas confondre avec "E" comme valeur de
      la colonne E/P) apparaît ici pour un envoi au Nigéria, laissant
      penser que le zonage FICP est un système de zones lettrées A à H (ou
      plus) représentant des groupes de pays/continents, PAS un simple
      "FICP=International, ST=France" binaire.
    - Filtre sur colonne **I "Pays"** : liste de valeurs proposée montre
      au moins **CA, DK, FR, GR, IE, IT, NG, NL...** (liste tronquée dans
      le cadre de la frame, scroll vertical visible indiquant plus de
      valeurs en dessous) — confirme une bonne variété de pays de
      destination dans le fichier d'avril, cohérent avec un mode
      d'expédition international actif.
    - Filtre sur colonne **U "Frêt"** (ou colonne proche) montrant une
      liste de montants (4,61 / 4,78 / 5,11 / 5,45 / 5,78 / 6,12 / 6,45 /
      6,79...) — action de contrôle visuel des valeurs de frêt présentes,
      sans action de filtrage effective observée (case "Sélectionner
      tout" cochée, fenêtre simplement ouverte puis probablement
      annulée).
40. (11:10 – 11:55, toutes dernières frames) Retour bref sur l'onglet
    **"TCD"** du classeur `2026_04_Facture Fedex.xlsx` : reconfirmation de
    la formule de la colonne F "E/P" en cellule F2 :
    ```
    =SI(RECHERCHEX(A2;'Shipment Detail'!N:N;'Shipment Detail'!D:D)="entreprise";"E";"P")
    ```
    (variante rigoureusement identique à celle déjà confirmée en F3,
    Partie 4 point 25 — juste une ligne au-dessus). La vidéo se termine
    sans étape de clôture/export finale supplémentaire visible après ce
    retour de contrôle (pas de sauvegarde explicite filmée dans les toutes
    dernières secondes, ni de fermeture de classeur).
41. **Onglets "Zoning" et "Bilan clients" jamais ouverts/inspectés dans
    cette vidéo** : bien que visibles en permanence dans la barre
    d'onglets du classeur (`Zoning | Bilan factures | Shipment Detail |
    TCD | Bilan clients | Import ERP`), **aucune frame ne montre le
    contenu de l'onglet "Zoning" ni de l'onglet "Bilan clients"** — ces 2
    onglets ne sont jamais cliqués/activés pendant toute la durée de la
    vidéo. **Ceci est un manque important** par rapport aux points à
    vérifier prioritaires de l'appelant (structure exacte de "Zoning",
    formule XLOOKUP Pays→Zone, utilisation de la seule colonne C "FICP")
    — cette vidéo NE PERMET PAS de confirmer ou contredire ces points
    précis sur l'onglet Zoning lui-même (seules des observations
    indirectes via le TMS "si.laruche-logistique.fr", Partie 5, et le
    filtre sur la colonne Zone du CSV final, point 39 ci-dessus,
    apportent un éclairage partiel).

---

## Synthèse de la réconciliation (point par point, vs. la liste de points
à vérifier fournie par l'appelant)

| # | Point de l'appelant (fichier juin 2026) | Statut | Précision apportée par la vidéo (avril 2026) |
|---|---|---|---|
| 1 | Colonnes BK-BO : BK=Fret HT, BL=TVA+suppléments+gasoil TTC (à ne jamais utiliser), BM=Droits et taxes, BN=Remise, BO=Total TTC | **CONFIRMÉ** | En-têtes CSV source confirmés en toutes lettres : "Shipment Freight Charge Billed Currency" (BK), "Shipment Miscellaneous Charge Billed Currency" (BL), "Shipment Duty And Tax Charge Billed Currency" (BM), "Shipment Discount Billed Currency" (BN), "Net Charge Billed Currency" (BO) — mêmes positions de colonnes que dans "Shipment Detail" (décalage de tête identique). Colonne BL jamais utilisée dans une formule observée, cohérent avec la mémoire projet (à ne pas utiliser). |
| 2 | Colonne A "Montant" = `=BK+BN` | **CONFIRMÉ EXACTEMENT** | Formule `=BK<n>+BN<n>` observée en clair sur plusieurs lignes (ex. `=BK38+BN38`). |
| 3 | Colonne D "E/P" = valeur littérale ('entreprise' ou vide) | **CONTREDIT** | Colonne D "E/P" est une **formule RECHERCHEX/XLOOKUP** vers un fichier externe `Export expéditions_brut.xlsx` (colonne AP tracking → colonne Q valeur), PAS une saisie littérale. Résultat = "entreprise" ou "particulier" (texte complet, pas juste "entreprise"/vide). Le fichier externe est un export du TMS interne "La Ruche" (`si.laruche-logistique.fr`), où le "statut du destinataire (particulier ou entreprise)" est une donnée saisie par l'expéditeur au moment de la demande d'expédition — mécanisme "m/m-1/m-2" confirmé (plusieurs fichiers mensuels ouverts simultanément pour rattraper les trackings non trouvés dans le mois courant). |
| 4 | Colonne C "Catégorie" toujours vide (juin, 3607 lignes) | **CONTREDIT (pour avril)** | Sur le fichier avril 2026, la colonne C "Catégorie" est **REMPLIE** avec des valeurs "ST"/"FICP" sur (quasi) toutes les lignes — pas vide. Hypothèse : la colonne a pu être vidée/dépréciée entre avril et juin, ou alors sa fonction a changé ; à confirmer directement avec le pôle transport, écart significatif avec l'observation juin. |
| 5 | TCD : RowFields=Tracking+Pays, DataFields=Somme Montant+Somme Poids, colonnes calculées E (poids kg, ROUNDUP) et F (E/P via XLOOKUP) | **CONFIRMÉ EXACTEMENT** | Colonne E : `=ARRONDI.SUP(D<n>*0,453592;1)`. Colonne F : `=SI(RECHERCHEX(A<n>;'Shipment Detail'!N:N;'Shipment Detail'!D:D)="entreprise";"E";"P")`. Les 2 formules confirmées caractère par caractère sur plusieurs lignes. |
| 6 | Import ERP : Transporteur = FEDEX INTERNATIONAL si mode envoi ∈ {FICP,IE,RE,International,Europe}, sinon FEDEX FRANCE | **CONFIRMÉ EXACTEMENT ET INTÉGRALEMENT** | Formule complète lue : `=SI(OU(M3="FICP";M3="IE"; M3="RE"; M3="International"; M3="Europe");"FEDEX INTERNATIONAL";"FEDEX FRANCE")`. |
| 7 | Import ERP : Zone = XLOOKUP(Pays, Zoning!B:B, Zoning!C:C), seule colonne C (FICP) utilisée même pour IE/RE — semble suspect | **NI CONFIRMÉ NI INFIRMÉ DIRECTEMENT** | L'onglet "Zoning" n'a **jamais été ouvert** dans cette vidéo (cf. point 41). Éclairage indirect via le TMS "si.laruche-logistique.fr" : un devis d'expédition FICP affiche "Zone : T", cohérent avec les valeurs de zone observées dans "Import ERP"/CSV final (T, S, U, H, A, E). Hypothèse avancée (non confirmée) : toutes les expéditions FedEx INTERNATIONAL (FICP, IE, RE...) utiliseraient la même table de zonage FICP dans le TMS, ce qui justifierait que la formule Excel ne référence que la colonne C — **à confirmer explicitement avec le pôle transport**, cette vidéo ne le prouve pas formellement. |
| 8 | Import ERP : Droits et taxes = XLOOKUP tracking → Shipment Detail!BM | **NI CONFIRMÉ NI INFIRMÉ** | Colonne "Droits et taxes" (O) de Import ERP jamais vue en cours d'édition avec sa formule complète dans les frames disponibles (seule la colonne K "Nbr Colis" et la structure générale ont été confirmées avec formule visible). |
| 9 | Import ERP : Colis volumineux = 10€ si XLOOKUP(tracking, Shipment Detail!AV "Dimmed Length cm") > 60 | **NON OBSERVÉ** | Aucune frame ne montre la formule de la colonne "Colis volumineux" (R) en cours d'édition. Colonne "Dimmed Length (cm)" bien confirmée présente dans l'en-tête du CSV source brut (Partie 0, point 3), donc la matière première existe, mais la formule elle-même n'a pas été lue. |
| 10 | Import ERP : TVA = 20% si pays UE, sinon 0% | **NUANCÉ / POSSIBLEMENT CONTREDIT** | Observations en Partie 1 (aperçu initial "Import ERP", avant les formules détaillées) : TVA = 0,2 pour les lignes FEDEX FRANCE (pays FR), TVA = 0 pour TOUTES les lignes FEDEX INTERNATIONAL observées, y COMPRIS des pays UE (IE, IT, PL, SE, DK — DK n'est pas UE mais IE/IT/PL/SE le sont). Ceci suggère que la règle réelle serait plutôt "TVA=20% si FEDEX FRANCE (national), TVA=0% si FEDEX INTERNATIONAL" plutôt que basée sur le statut UE du pays destinataire — **à vérifier directement dans la formule de la colonne N "TVA" de Import ERP, jamais vue en cours d'édition dans cette vidéo**. |
| 11 | Bilan factures : TCD + Frêt HT/Total TTC/PDF/Écart, 9 factures juin mais 8 PDF disponibles | **CONFIRMÉ ET NUANCÉ** | Structure TCD confirmée à l'identique (Lignes=Invoice Number, Valeurs=Frêt HT + Total TTC, colonnes D "PDF" saisie manuelle et E "Ecart" = calcul, résultat "- €" si égalité). Pour avril 2026 : **11 lignes de facture (10 numéros + 1 "(vide)") vs 9 PDF disponibles** dans le dossier réseau — écart structurel confirmé et même amplifié (2 factures sans PDF, pas juste 1 comme observé par l'appelant sur juin). Cause exacte non montrée dans la vidéo. |
| 12 | 2 CSV bruts juin, lequel est la vraie source de Shipment Detail (3607 lignes) ? | **PARTIELLEMENT ÉCLAIRÉ, NON TRANCHÉ POUR JUIN** | Pour avril : **un seul CSV** (`0g000e48hq_fedex 04-26_...csv`, préfixe `0g000e48hq_`, 2676 lignes brutes) alimente directement "Shipment Detail" (2672-2673 lignes utiles après nettoyage) — colonnes BK-BO "Billed Currency" natives. Le second CSV du dossier (sans préfixe `0g000e48hq_`) n'est JAMAIS ouvert dans la vidéo. **Par analogie, pour juin, le fichier `0g000e48hq_FEDEX 06-2026 v2_20260710125723.csv` (préfixe identique) est très probablement la vraie source**, et non `FEDEX 06 2026.csv` — mais ceci reste une déduction par analogie de nommage, PAS une confirmation directe sur le fichier juin lui-même (le nombre de lignes ne colle pas exactement : 4001 lignes pour le CSV `0g000e48hq_` de juin vs 3607 lignes de Shipment Detail juin, un écart similaire à celui observé en avril entre 2676 lignes brutes et 2672-2673 lignes utiles — cohérent avec un nettoyage/dédoublonnage systématique après collage). |
| 13 | En-têtes français (CSV bruts) vs anglais (xlsx modèle) : mapping/traduction ? | **CONTREDIT (pas de traduction manuelle)** | Le CSV source `0g000e48hq_...` a des **en-têtes NATIVEMENT EN ANGLAIS**, identiques mot pour mot aux en-têtes du xlsx modèle (confirmé en Partie 0, lecture intégrale de la ligne d'en-tête). **Il n'y a donc PAS de traduction manuelle ni d'étape de renommage** pour ce type de fichier. L'hypothèse d'un mapping FR→EN concerne peut-être l'AUTRE CSV (celui sans préfixe `0g000e48hq_`, jamais ouvert dans cette vidéo, qui pourrait être celui avec des en-têtes français comme observé par l'appelant sur `FEDEX 06 2026.csv`) — si tel est le cas, cela renforcerait l'hypothèse du point 12 : le CSV `0g000e48hq_...` (anglais, colonnes Billed Currency) est la vraie source, l'autre CSV (français) étant un export alternatif non utilisé pour "Shipment Detail". |

---

## Découvertes majeures NON présentes dans la liste de points à vérifier
initiale

- **Le TMS interne `si.laruche-logistique.fr`** ("Système d'Information
  de La Ruche Logistique") est la source de vérité amont pour le statut
  "particulier ou entreprise" du destinataire (texte trouvé explicitement
  dans la note de bas de devis TMS) et pour le calcul prévisionnel de
  zone/mode de livraison par transporteur (devis "Transporteur : FEDEX
  INTERNATIONAL / Mode de livraison : FICP / Zone : T"). Ce TMS est
  vraisemblablement la source de l'"export brut" mensuel
  (`Export expéditions_brut.xlsx`) déjà mentionné dans `registry.js`.
- **Le fichier externe `Export expéditions_brut.xlsx`** est un export
  MULTI-TRANSPORTEURS (CHRONOPOST, KUEHNE, GEODIS, MONDIAL-RELAY-FOUTAS,
  COLISSIMO-APIDURA, DPD-APIDURA observés dans sa colonne "TRANSPORTEUR"),
  pas spécifique à FedEx — potentiellement réutilisable pour comprendre le
  mécanisme "E/P" d'autres transporteurs codés dans le projet.
- **La colonne "Zone" (dans Import ERP / CSV final / TCD) prend au moins
  6 valeurs lettrées distinctes** (T, S, U, H, A, E) pour les envois
  internationaux (FICP), suggérant un système de zonage par groupes de
  pays/continents plus fin qu'un simple France/International — la lettre
  "E" comme zone (Nigéria) est à ne pas confondre avec "E" comme valeur
  de la colonne "E/P" (entreprise), ambiguïté de nommage à noter.
- **Bug de recopie au-delà des données réelles** dans "Import ERP" (cascade
  #N/A/"inconnu" sur les lignes 2674+), pattern identique à celui déjà
  documenté sur Colissimo — mais ici la correction n'a pas été vue filmée
  explicitement (le fichier CSV final exporté n'a pourtant pas ces lignes
  fantômes, donc une correction a bien eu lieu hors champ de capture).
- **Le "Total TTC" du Bilan factures ne correspond ni à la colonne A
  "Montant" ni directement à BO** avec certitude formelle dans cette
  vidéo — la formule exacte du champ TCD "Total TTC" n'a jamais été vue
  en cours d'édition (seul le nom du champ dans le panneau latéral des
  champs TCD a été observé) ; sa valeur numérique est cohérente avec
  "Net Charge Billed Currency" (BO) mais ceci reste une déduction, pas une
  lecture directe de formule.

---

## Points ambigus / illisibles à faire confirmer par le pôle transport

1. **Colonne "Catégorie" (C) de "Shipment Detail" remplie en avril
   (valeurs "ST"/"FICP") mais vide en juin selon l'appelant** — écart
   direct entre les 2 mois à faire confirmer/expliquer : dépréciation de
   la colonne entre avril et juin ? Erreur d'inspection sur l'un des deux
   fichiers ? Priorité haute, contredit directement une observation de
   l'appelant.
2. **Onglets "Zoning" et "Bilan clients" jamais ouverts dans cette
   vidéo** — aucune confirmation directe possible sur la structure exacte
   de "Zoning" (colonnes FICP/IE/RE/Zone) ni sur le fait que seule la
   colonne C (FICP) soit utilisée dans la formule Zone de "Import ERP".
   Une autre vidéo FedEx (non fournie ici) montre peut-être cet onglet.
3. **Formules exactes non vues en cours d'édition** pour les colonnes
   "Import ERP" O (Droits et taxes), P (Assurance), Q (Zones éloignées),
   R (Colis volumineux), S (Adresses), T (Frêt), U (plus-value BtoC), V
   (Gazole), N (TVA) — seule la structure globale et le principe RECHERCHEX
   générique (via colonne K "Nbr Colis" observée) ont pu être confirmés
   par analogie, pas chacune de ces formules individuellement.
4. **Règle TVA réellement observée (0,2 si FEDEX FRANCE / 0 si FEDEX
   INTERNATIONAL, y compris pour des pays UE comme IE/IT/PL/SE) semble
   contredire la règle "TVA=20% si pays UE, 0% sinon" formulée par
   l'appelant** — à vérifier directement dans la formule de la colonne N
   "TVA" de "Import ERP" (jamais vue en cours d'édition), c'est un point
   à fort enjeu métier (risque de sous/sur-facturation de TVA).
5. **Colonne Q de l'"Export expéditions_brut.xlsx"** (valeur exacte
   retournée par la formule E/P, contenant le texte "entreprise"/
   "particulier") : son en-tête de colonne n'a jamais été cadré nettement
   dans les frames disponibles — à confirmer directement dans un export
   réel de ce fichier.
6. **Second fichier CSV brut mensuel** (sans préfixe `0g000e48hq_`,
   présent dans le dossier réseau avril ET juin) : jamais ouvert dans
   cette vidéo, son usage exact (redondance, format alternatif, ou
   utilisé pour une autre étape non filmée) reste à clarifier avec le
   pôle transport — cette vidéo ne prouve pas formellement que le fichier
   `0g000e48hq_...` soit la seule/unique source pour tous les mois,
   seulement pour avril 2026.
7. **Écart 11 factures vs 9 PDF (avril) / 9 factures vs 8 PDF (juin)** :
   cause structurelle jamais expliquée dans la vidéo (PDF pas encore
   téléchargés au moment de la capture ? factures à montant nul non
   pertinentes ? délai de mise à disposition des PDF par FedEx après
   émission de la facture ?).
8. **Action de Rechercher/Remplacer observée sur les formules** (Partie 2,
   point 18) : nature exacte de cette opération non identifiée avec
   certitude — à vérifier directement dans le fichier réel ou en
   redemandant au pôle transport lors d'une prochaine session
   d'observation.
9. **Formule exacte du champ TCD "Frêt HT" et "Total TTC" de l'onglet
   "Bilan factures"** : seuls les noms de champs affichés dans le panneau
   latéral Excel ("Frêt HT", "Total TTC") ont été observés, pas les noms
   de champs sources réels de "Shipment Detail" utilisés en interne par le
   TCD (probablement "Montant" pour Frêt HT, et une colonne BO/"Net
   Charge Billed Currency" pour Total TTC, mais non confirmé caractère
   par caractère).
10. **Qualité/résolution de certaines frames** : plusieurs captures de
    formules (notamment autour de 06:00-07:00 et 08:30-09:00, zones de
    texte dense en colonnes lointaines du classeur) ont nécessité un
    zoom mental/une interprétation par cohérence avec les autres
    formules similaires déjà confirmées ailleurs dans la même vidéo —
    signalé ici pour transparence, la plupart des formules citées dans ce
    document ont néanmoins pu être lues caractère par caractère dans au
    moins une frame nette.
