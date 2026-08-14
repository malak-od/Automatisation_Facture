# Transcription vidéo — Colissimo_1_Preparation fichier import.mp4

Durée totale réelle : **12 min 54 s (774,5 s)** (`ffprobe`). Résolution capture
d'écran Windows (1918×1030), Excel (classeur de travail `2026_04_Facture
Colissimo.xlsx`, fichiers sources CSV `CSV_Prestations_au_colis.csv`,
`CSV_Frais_de_douane.csv`, l'onglet Excel `Import CSV` étant à ce stade une
FEUILLE du classeur, pas encore un fichier `.csv` séparé), un navigateur Edge
ouvert brièvement sur un PDF de facture La Poste/Colissimo
(`facture_2026-04-30.pdf`, hébergé sur le partage réseau
`192.168.5.3\Comptabilité La Ruche\$Facturation automatique\1 - Factures
transporteurs + calculs\2026\2026 04\Colissimo\`), et l'Explorateur de
fichiers Windows (dossier réseau `2 - Fichiers csv import\2026\2026 03\`
listant les CSV finaux mensuels de tous les transporteurs).

Aucune bande son exploitable : toute l'info vient de ce qui est visible à
l'écran (barre de formule, onglets, cellules, filtres, menus contextuels).
Timestamps approximatifs en `mm:ss`, dérivés de 129 frames à intervalle fixe
(6 s) + 63 frames de détection de changement de scène, lues intégralement et
dans l'ordre chronologique.

Croisement systématique avec le fichier réel
`Transporteurs/Colissimo/2026_06_Facture Colissimo.xlsx` (ouvert via
openpyxl, formules non recalculées `data_only=False`) et les 3 CSV sources
bruts présents dans le dossier projet (`CSV_Prestations_au_colis.csv`,
`CSV_Frais_de_douane.csv`, `CSV_Indemnisations.csv`) pour confirmer noms
exacts d'onglets/colonnes/formules quand la vidéo ne suffisait pas à elle
seule. **Attention** : le fichier de référence croisé est celui de **juin
2026** (mois M), alors que la vidéo montre le traitement du mois d'**avril
2026** (classeur `2026_04_Facture Colissimo.xlsx`, avec en fond de vues sur
mars `2026_03_...` comme référence du mois précédent) — les formules/
en-têtes de structure sont identiques d'un mois sur l'autre, seules les
valeurs numériques diffèrent.

---

## 0. Onglets du classeur de travail (confirmés à l'écran ET dans le fichier
réel juin 2026, ordre exact identique dans les deux)

Barre d'onglets en bas de `2026_04_Facture Colissimo.xlsx` (vidéo) et
`2026_06_Facture Colissimo.xlsx` (fichier réel), dans cet ordre :

**Table de correspondance | Modes envois | Pays | Poids | Facture Colissimo
| TCD | Import CSV | Bilan Factures | Bilan Client**

Cet ordre ne correspond pas exactement à l'énoncé des notes manuscrites
("Onglet Poids / Onglet TCD / Onglet Bilan Facture / Onglet Import ERP") —
la vidéo montre 4 onglets techniques supplémentaires en tête (Table de
correspondance, Modes envois, Pays, Poids) qui servent de tables de
référence pour les formules RECHERCHEX/XLOOKUP des onglets suivants. L'onglet
appelé "Import ERP" dans les notes correspond à l'onglet **"Import CSV"**
dans le classeur réel (pas "Fichier import").

---

## Partie 1 — Table de correspondance et Modes envois (tables de référence)
(00:00 – 00:30)

1. (00:00 – 00:12) Le classeur `2026_04_Facture Colissimo.xlsx` s'ouvre sur
   l'onglet **Table de correspondance**, 2 colonnes : **A = "Code
   Colissimo" | B = "Corresp."**. Table complète confirmée (26 lignes,
   croisée avec le fichier réel juin 2026) :
   | Code Colissimo | Corresp. |
   |---|---|
   | CH_CAP | **Taxe gazole** |
   | CH_P_6A, CH_P_6C, CH_P_6H, CH_P_6M, CH_P_7R, CH_P_8Q, CH_P_9W, CH_P_CA, CH_P_CB, CH_P_EY, CH_P_8R, CH_P_CG, CH_P_7Q | **Frêt** |
   | CH_O_A_ASSURANCE | **Assurance** |
   | CH_O_FTDDDP_DOUANEMARCHANDISE | **Droits et taxes** |
   | CH_S_D_DESTINATION | **Zones éloignées** |
   | CH_S_NM_NONMECANISABLE | **Colis volumineux** |
   | CH_S_PVR_POIDSVOLUMETRIQUEROUTIER | **Colis volumineux** |
   | CH_S_QA_QUALITEANNONCE | **Adresse** |
   | CH_S_REET_REETIQUETAGE | **Adresse** |
   | CH_S_SDE_DECARBONATION | **Frêt** |
   | CH_S_SI_SURETEINTERNATIONALE | **Frêt** |
   | CH_S_NCD_NONCONFORMITEDOUANIERE | **Droits et taxes** |
   | CH_DOUANE | **Droits et taxes** |
   | CH_S_PT_PARTENAIRETRANSPORT | **Frêt** |
2. (00:12 – 00:30) Onglet **Modes envois** : 6 colonnes — **A = "Concatener"
   (formule `=B2&C2`, préfixe tracking + code pays) | B = "Préfixe
   tracking" | C = "Pays" (code 2 lettres) | D = "Zone" | E = "TVA" (0 ou
   0,2) | F = "Mode transport"**. Valeurs de "Mode transport" confirmées
   dans le fichier réel : **DOM, DOS, A2P, ECO** (liste des 4 valeurs
   distinctes présentes dans la colonne F sur 328 lignes). Exemples de
   lignes : `5R+FR → OM1_DOM/DOM`, `6A+FR → FR_DOM/DOM`, `6C+FR →
   FR_DOS/DOS`, `6H+FR → FR_retrait/A2P`, `6M+FR → FR_retrait/A2P`,
   `7Q+FR → OM1_DOS/DOS`.

## Partie 2 — Onglet "Pays" (table de référence) (00:30 – 00:48)

3. (00:30 – 00:48) Onglet **Pays**, 2 colonnes : **A = "Pays" (nom complet
   en majuscules, ex. GUADELOUPE, GUYANE FRANCAISE, HONG KONG, HONGRIE,
   IRLANDE, ITALIE, LETTONIE, LITUANIE, LUXEMBOURG, MARTINIQUE, MAYOTTE,
   MEXIQUE, MONACO, MONTENEGRO, NORVEGE, PAYS-BAS, POLOGNE, PORTUGAL,
   REPUBLIQUE TCHEQUE, REUNION, ROUMANIE, ROYAUME-UNI, SAINT-BARTHELEMY,
   SAINT-PIERRE-ET-MIQUELON, SECTEUR POSTAL, SERBIE, SLOVAQUIE, SLOVENIE,
   SUEDE, SUISSE, TURQUIE...) | B = "2 lettres" (code ISO, ex. GP, GF, HK,
   HU, IE, IT, LV, LT, LU, MQ, YT, MX, MC, ME, NO, NL, PL, PT, CZ, RE, RO,
   GB, BL, PM, FR, RS, SK, SI, SE, CH, TR)**. 65 lignes au total (fichier
   réel). Cette table sert à convertir le nom de pays en 2 lettres pour la
   correspondance avec l'onglet "Modes envois".

## Partie 3 — Onglet "Poids" et début de "Facture Colissimo" (00:48 – 01:30)

4. Onglet **Poids** : TCD à 2 colonnes — **A = "Étiquettes de lignes"**
   (numéro de tracking colis, ex. `6A06233150369`) | **B = "Moyenne de
   Poids Kg"** (ex. 0,4 / 0,25 / 0,5). Ce TCD est construit à partir de la
   colonne "Poids Kg" de "Facture Colissimo" filtrée sur les lignes ayant un
   poids renseigné (lignes "Charge" avec `Nature du poids facturé = M`), et
   sert de référence pour retrouver le poids d'un colis par son tracking
   dans l'onglet Import CSV.
5. Bascule sur l'onglet **Facture Colissimo** : en-têtes confirmés en ligne
   1, colonnes A à AG (croisé avec le fichier réel) :
   **A = Pays | B = Mode Envoi | C = Zone | D = concatener | E = Préfixe |
   F = Pays | G = Catégorie | H = N de ligne | I = N facture | J = Date |
   K = Compte facturé | L = Compte déposant | M = Description | N = Produit
   | O = N colis | P = Zone | Q = Pays Origine | R = Pays Destination |
   S = Code Postal Destination | T = Pourcentage de remise |
   U = Référence externe colis client | V = Nature du poids facturé |
   W = Poids Kg | X = LEN. cm | Y = HGT. cm | Z = WID. cm | AA = Total HT |
   AB = Taux TVA | AC = Rubrique de la facture | AD = Code charge |
   AE = Balises LIN EDI | AF = Offre | AG = Type de charge**.
   Les colonnes **A à G sont des colonnes de calcul** (remplies par formule,
   fond vert visible dans une frame), les colonnes **H à AG sont les
   données brutes collées** depuis les CSV sources.

## Partie 4 — Formules des colonnes A à G de "Facture Colissimo" (confirmées
via le fichier réel juin 2026, exactement lisibles) (01:10 – 02:00)

6. Formules exactes de la ligne 2 (ligne de données), colonnes A à G :
   - **A (Pays, nom complet)** :
     `=_xlfn.XLOOKUP(R2,Pays!A:A,Pays!B:B,"Pays à créer")`
     — recherche le code Pays Destination (R) dans la colonne B de "Pays"
     (2 lettres) et renvoie... attention, la formule cherche en fait la
     correspondance inverse et renvoie la colonne B ("2 lettres") en
     cherchant R2 dans A ; si R2 (nom en toutes lettres, ex. "FRANCE")
     n'est pas trouvé dans la colonne A de "Pays", renvoie le texte
     **"Pays à créer"** — c'est ce texte qui apparaît en colonne A quand un
     pays est absent de la table de référence (cf. Partie 8, QA finale).
   - **B (Mode Envoi)** :
     `=IF(COUNTIF('Modes envois'!A:A,D2)=0,"zone inconnue",
     _xlfn.XLOOKUP(D2,'Modes envois'!A:A,'Modes envois'!F:F))`
     — si la clé concaténée D2 n'existe pas dans la colonne A "Concatener"
     de "Modes envois", renvoie **"zone inconnue"**, sinon renvoie la
     colonne F "Mode transport" (DOM/DOS/A2P/ECO).
   - **C (Zone)** : même logique,
     `=IF(COUNTIF('Modes envois'!A:A,D2)=0,"zone inconnue",
     _xlfn.XLOOKUP(D2,'Modes envois'!A:A,'Modes envois'!D:D))`
     — renvoie la colonne D "Zone" de "Modes envois" (ex. "FR_DOM").
   - **D (concatener)** : `=E2&F2` (concaténation Préfixe + Pays 2 lettres).
   - **E (Préfixe)** : `=LEFT(O2,2)` — 2 premiers caractères du numéro de
     colis (colonne O "N colis", ex. "6A" extrait de "6A06233150369").
   - **F (Pays, 2 lettres)** :
     `=_xlfn.XLOOKUP(R2,Pays!A:A,Pays!B:B,"")` — recherche le nom complet du
     Pays Destination (R, ex. "FRANCE") dans la colonne A de "Pays" et
     renvoie le code 2 lettres correspondant (colonne B), `""` si non
     trouvé.
   - **G (Catégorie)** :
     `=_xlfn.XLOOKUP(AD2,'Table de correspondance'!A:A,
     'Table de correspondance'!B:B)` — recherche le "Code charge" (AD, ex.
     "CH_CAP", "CH_P_6A") dans la Table de correspondance et renvoie le
     poste ("Taxe gazole", "Frêt", "Assurance", etc.).
7. Ces 7 formules (A à G) sont étirées/recopiées vers le bas sur toute la
   hauteur du tableau — confirme la consigne "**Étendre les colonnes A à
   G**" des notes manuscrites : à chaque import mensuel, il faut réétirer
   ces formules jusqu'à la dernière ligne de données collées.

## Partie 5 — Ouverture des 2 CSV sources ("Prestations au colis" et "Frais
de douane") et copier-coller dans "Facture Colissimo" (01:30 – 04:00)

8. (01:30 – 02:10) Ouverture du fichier **`CSV_Prestations_au_colis.csv`**
   (correspond au "CSV prestation Colis" des notes). En-têtes confirmés
   (colonnes A à Z, délimiteur `;`) : **N de ligne | N facture | Date |
   Compte facturé | Compte déposant | Description | Produit | N colis |
   Zone | Pays Origine | Pays Destination | Code Postal Destination |
   Pourcentage de remise | Référence externe colis client | Nature du poids
   facturé | Poids Kg | LEN. cm | HGT. cm | WID. cm | Total HT | Taux TVA |
   Rubrique de la facture | Code charge | Balises LIN EDI | Offre | Type de
   charge**. Exemple de lignes réelles (juin 2026, croisé fichier) :
   ```
   1;CO01312342;24/06/2026;826035;563067;CAE;6A - Colissimo Domicile Sans Sign. F;6A06233150369;;FRANCE;FRANCE;83250;;;;;;;;0,60;20,00;AJUSTEMENT ENERGIE;CH_CAP;LIN008;;F
   3;CO01312342;24/06/2026;826035;563067;Charge 6A;6A - Colissimo Domicile Sans Sign. F;6A06233150369;NAT;FRANCE;FRANCE;83250;;EXP20260511-2741794;M;0,400;23;8;16;7,71;20,00;FRAIS DE PORT;CH_P_6A;LIN004;;F
   4;CO01312342;24/06/2026;826035;563067;Remise Charge 6A;...;-2,06;20,00;FRAIS DE PORT;CH_P_6A;LIN004;;F
   ```
   Les libellés "Description" observés : "CAE" (ajustement énergie =
   coefficient/taxe gasoil), "Charge 6A/6C/..." (frêt), "Remise Charge 6A"
   (remise, montant négatif), "Article Supplément Participation à la
   décarbonation" (supplément), "Article Suppléments Sûreté
   internationale", etc.
9. (02:10 – 02:40) Copie de la totalité des colonnes B à Y (soit N de ligne
   à Type de charge) du CSV, puis **collage dans l'onglet "Facture
   Colissimo" à partir de la colonne H** (H = "N de ligne") — confirme
   EXACTEMENT la consigne des notes ("Coller dans facture Colissimo à
   partir de la colonne H"). Le nombre de lignes collées correspond aux
   données de "Prestations au colis" du mois (fichier réel juin 2026 :
   plus de 12 000 lignes, jusqu'à la ligne ~12067 pour le dernier compte
   déposant "897361").
10. (02:40 – 03:30) Ouverture du 2ème fichier **`CSV_Frais_de_douane.csv`**
    (correspond au "et douanes" des notes — nom exact du menu déroulant sur
    le site est "CSV prestation Colis et douanes" mais il s'agit en réalité
    de **deux fichiers CSV distincts téléchargés séparément**, "Prestations
    au colis" et "Frais de douane"). En-têtes confirmés : **N de ligne | N
    facture | Date | Compte facturé | Compte déposant | Description |
    Produit | N colis | Droit de douane | TVA à l'importation | Octroi de
    mer | Frais de gestion | Autres taxes | Total HT | Taux TVA | Rubrique
    de la facture | Code charge | Balises LIN EDI | Offre | Type de
    charge**. Exemple de ligne réelle :
    ```
    1;CO01312342;25/05/2026;826035;897361;Frais de douane - Colis CB531574507FR;CB - Colissimo Domicile Sign. INT;CB531574507FR;0,00;88,12;0,00;0,00;5,32;93,44;0,00;PRESTATIONS COMPLEMENTAIRES;CH_DOUANE;LIN005;;D
    ```
    Toutes les lignes de ce fichier ont **Code charge = "CH_DOUANE"**,
    **Rubrique de la facture = "PRESTATIONS COMPLEMENTAIRES"**, et
    **Taux TVA = 0,00** (les frais de douane ne portent pas de TVA
    française, cohérent avec "TVA à l'importation" facturée séparément par
    la douane). Colonne "Total HT" = somme de Droit de douane + TVA à
    l'importation + Octroi de mer + Frais de gestion + Autres taxes (visible
    dans l'exemple : 0+88,12+0+0+5,32 = 93,44).
11. (03:30 – 04:00) Copie des lignes de ce fichier "Frais de douane" et
    **collage À LA SUITE** des lignes déjà collées de "Prestations au
    colis" dans "Facture Colissimo" (mêmes colonnes H à Y, ajoutées en bas
    du tableau existant) — ces lignes ont une structure de colonnes
    légèrement différente (pas de Zone/Pays Origine/Poids/dimensions, mais
    Droit de douane/TVA importation/Octroi de mer/Frais de gestion/Autres
    taxes qui ne mappent PAS sur les mêmes colonnes H-Y — **point à vérifier
    finement avec le pôle transport, cf. section Points ambigus**, la
    correspondance colonne à colonne entre les 2 fichiers sources n'est pas
    strictement identique alors qu'ils sont collés dans le même bloc de
    colonnes H-Y de "Facture Colissimo").

## Partie 6 — Tri par colonne O et étirement des colonnes A-G après collage
(04:00 – 04:40)

12. (04:00 – 04:20) **Tri de tout le tableau "Facture Colissimo" sur la
    colonne O ("N colis" — numéro de tracking colis) de A à Z** — confirme
    exactement l'instruction des notes "**Trier Colonne O de A à Z**".
    Ce tri regroupe les lignes appartenant à un même colis ensemble (Charge,
    Remise Charge, CAE, Article Supplément... pour un même tracking se
    retrouvent côte à côte).
13. (04:20 – 04:40) Sélection des colonnes A à G, poignée de recopie
    (double-clic ou glisser) étirée jusqu'à la dernière ligne du tableau
    (après ajout des nouvelles lignes des 2 CSV) — confirme "**Étendre les
    colonnes A à G et rajouter à la fin de la colonne G**" des notes. Les
    valeurs #N/A ou "Pays à créer"/"zone inconnue" apparaissent
    ponctuellement à ce stade pour les nouvelles lignes dont le Pays/Mode
    Envoi ne matche pas encore une table de référence (à résoudre plus tard
    en fin de vidéo, cf. Partie 8).

## Partie 7 — Onglet "TCD" : structure et formules (04:40 – 06:00)

14. (04:40 – 05:10) Onglet **TCD** (Tableau Croisé Dynamique construit sur
    la source "Facture Colissimo"). En-têtes confirmés dans le fichier réel
    (ligne 1/2) :
    - **Colonnes A/B/C hors-TCD (calcul manuel)** : **A = "ID Client"**,
      **B = "Total hors Gazole"** = `=SUM(E:J,L)` (somme des colonnes
      Adresse+Assurance+Colis volumineux+Droits et taxes+Frêt+plus-value
      BtoC+Zones éloignées, EXCLUANT K "Taxe gazole"), **C = "Total + GO"**
      = `=SUM(E:L)` (même somme mais INCLUANT K "Taxe gazole" — "GO" =
      Gasoil/Gazole).
    - **Zone TCD native (D à M)** : **D = "Étiquettes de lignes"** (numéro
      de tracking colis, ex. `6A06233150369`), **E = Adresse | F =
      Assurance | G = Colis volumineux | H = Droits et taxes | I = Frêt |
      J = plus-value BtoC | K = Taxe gazole | L = Zones éloignées |
      M = (vide)** — ce sont les 7 postes de la "Catégorie" (colonne G de
      "Facture Colissimo") ventilés en colonnes par le TCD (Somme de Total
      HT, une colonne par valeur de "Catégorie").
15. (05:10 – 05:30) Colonne **A "ID Client"** : remplie manuellement/par
    formule à partir d'une jointure externe (les valeurs sont des identifiants
    numériques comme 6739, 7027, 2072, 3752) — correspond à la consigne
    "**Effacer ID client**" des notes : cette colonne est ensuite VIDÉE
    (contenu effacé) une fois son usage de contrôle terminé, probablement
    car elle contient une donnée sensible/temporaire non nécessaire à
    l'export final (confirmé visuellement : une action de sélection de la
    colonne A du TCD suivie d'une touche Suppr/Effacer contenu est visible
    dans une frame vers 05:20, mais le détail exact de la formule d'origine
    de "ID Client" n'a pas pu être lu net avant l'effacement).
16. (05:30 – 06:00) **Tri de l'onglet TCD sur la colonne C ("Total + GO")**
    — les notes mentionnent "Trier Colonne C de A à Z" (a priori un tri
    croissant sur les montants, utile pour repérer visuellement les valeurs
    aberrantes ou nulles en haut/bas de liste). Puis **étirement de la
    colonne B** ("Total hors Gazole") vers le bas pour couvrir toutes les
    nouvelles lignes du TCD après actualisation — confirme "**Étendre
    colonne B**" des notes.
17. Clic droit sur le TCD → **"Actualiser"** visible dans une frame,
    confirmant que le TCD est manuellement rafraîchi après toute
    modification de la source "Facture Colissimo" (nouvelles lignes
    collées + tri).

## Partie 8 — Onglet "Import CSV" : structure et formules (05:40 – 08:30)

18. (05:40 – 06:20) Onglet **Import CSV** : en-têtes confirmés (colonnes A à
    W dans le fichier réel) : **A = " Transporteur " | B = "Date validité
    tarif" | C = "Réf.1" | D = "Réf. 2" | E = "Id client" | F = "N°
    Tracking" | G = "Nom" | H = "E / P" | I = "Pays" | J = " Zone " |
    K = " Nbr Colis " | L = " Poids " | M = "mode envoi" | N = " TVA " |
    O = " Droits et taxes " | P = " Assurance " | Q = " Zones éloignées " |
    R = " Colis volumineux " | S = " Adresses " | T = " Frêt " |
    U = " plus-value BtoC " | V = "Gazole" | W = "Nb Colis"**. (Notez les
    espaces de part et d'autre de nombreux en-têtes, présents dans le
    fichier source réel — probablement un format attendu strict par
    l'import ERP.)
19. Formules exactes de la ligne 2 (croisées avec le fichier réel) :
    - **A** = `"COLISSIMO"` (texte fixe, ligne 1), puis `=A2` pour les
      lignes suivantes (recopie de la valeur fixe).
    - **B** = date fixe du 1er du mois (ex. `01/06/2026`), recopiée avec
      `=B2` sur les lignes suivantes.
    - **C (Réf.1)** =
      `=_xlfn.XLOOKUP(F2,'Facture Colissimo'!O:O,'Facture Colissimo'!U:U)`
      — recherche le N° Tracking (F) dans la colonne O "N colis" de
      "Facture Colissimo", renvoie la colonne U "Référence externe colis
      client".
    - **F (N° Tracking)** = `=TCD!D3` — référence DIRECTE à la colonne D
      du TCD (Étiquettes de lignes), donc l'onglet Import CSV est bien
      **alimenté par formule depuis l'onglet TCD**, ligne par ligne
      (F2=TCD!D3, F3=TCD!D4, F4=TCD!D5, décalage de 1 car le TCD a 2 lignes
      d'en-tête).
    - **I (Pays)** =
      `=_xlfn.XLOOKUP(F2,'Facture Colissimo'!O:O,'Facture Colissimo'!A:A)`
      — recherche le tracking dans "Facture Colissimo" colonne O, renvoie
      la colonne A (Pays nom complet calculé via XLOOKUP, cf. Partie 4).
    - **J (Zone)** =
      `=_xlfn.XLOOKUP(F2,'Facture Colissimo'!O:O,'Facture Colissimo'!C:C,"")`
      — renvoie la colonne C (Zone, ex. "FR_DOM").
    - **K (Nbr Colis)** = `1` en dur (ligne 1), puis `=K2` recopié.
    - **L (Poids)** =
      `=ROUNDUP(_xlfn.XLOOKUP(F2,Poids!A:A,Poids!B:B,""),1)` — recherche le
      tracking dans l'onglet "Poids" (TCD Moyenne de Poids Kg), et
      **ARRONDIT AU DIXIÈME SUPÉRIEUR** (ROUNDUP à 1 décimale) le poids
      trouvé.
    - **M (mode envoi)** =
      `=_xlfn.XLOOKUP(F2,'Facture Colissimo'!O:O,'Facture Colissimo'!B:B)`
      — renvoie la colonne B de "Facture Colissimo" (Mode Envoi = DOM/DOS/
      A2P/ECO, calculé via XLOOKUP sur "Modes envois").
    - **N (TVA)** =
      `=IF(_xlfn.XLOOKUP(F2,'Facture Colissimo'!O:O,'Facture Colissimo'!AB:AB,"")=20,0.2,0)`
      — si le Taux TVA trouvé en colonne AB de "Facture Colissimo" vaut 20
      (%), renvoie 0,2 ; sinon renvoie 0.
    - **O à U (Droits et taxes, Assurance, Zones éloignées, Colis
      volumineux, Adresses, Frêt, plus-value BtoC)** = même motif répété
      pour chaque poste :
      `=IF(_xlfn.XLOOKUP($F2,TCD!D:D,TCD!<col>)=0,"",_xlfn.XLOOKUP($F2,TCD!D:D,TCD!<col>))`
      avec la correspondance colonne TCD → colonne Import CSV suivante,
      confirmée exactement (mapping des lettres de colonnes) :
      | Import CSV | TCD source | Poste |
      |---|---|---|
      | O (Droits et taxes) | TCD!H:H | Droits et taxes |
      | P (Assurance) | TCD!F:F | Assurance |
      | Q (Zones éloignées) | TCD!L:L | Zones éloignées |
      | R (Colis volumineux) | TCD!G:G | Colis volumineux |
      | S (Adresses) | TCD!E:E | Adresse |
      | T (Frêt) | TCD!I:I | Frêt |
      | U (plus-value BtoC) | TCD!J:J | plus-value BtoC |
      (Chaque formule renvoie une chaîne vide `""` si la valeur du TCD vaut
      0, au lieu d'afficher "0" — vraisemblablement pour que l'import ERP
      distingue "pas de charge" de "charge nulle".)
    - Colonne **V (Gazole)** n'a pas été observée avec sa formule exacte
      dans une frame nette (probablement `=XLOOKUP(F2,TCD!D:D,TCD!K:K)`
      par symétrie avec les autres colonnes O-U, TCD colonne K = "Taxe
      gazole", mais non confirmé caractère par caractère dans les frames
      disponibles).
20. Ces formules (colonnes A à W) sont étirées vers le bas jusqu'à la
    dernière ligne du TCD — confirme "**Étendre les colonnes**" pour
    l'onglet Import (ERP).

## Partie 9 — Bug de recopie au-delà des données réelles : cascade de #N/A
(06:20 – 07:10)

21. (06:20 – 06:50) Une frame montre la formule de la colonne I
    `=RECHERCHEX(F2909;'Facture Colissimo'!O:O;'Facture Colissimo'!A:A)`
    dans une cellule **I2909** de l'onglet "Import CSV" (ou son ancêtre —
    à ce stade de la vidéo, "Import CSV" est encore une feuille interne du
    classeur, pas un fichier CSV externe) affichant **#N/A en cascade sur
    toute la ligne 2909 et au-delà** (colonnes I à X), car les formules ont
    été étirées PLUS BAS que la dernière ligne réelle de données du TCD (la
    dernière ligne réelle étant ~2908). Un triangle d'avertissement Excel
    est visible sur ces cellules.
22. (06:50 – 07:10) **Correction en direct** : sélection des lignes 2909
    et suivantes (jusqu'à la fin de la plage étirée par erreur), menu
    contextuel clic droit → **"Supprimer"** → boîte de dialogue "Supprimer"
    avec options "Décaler les cellules vers la gauche / le haut", "Ligne
    entière", "Colonne entière" — l'utilisateur choisit de supprimer les
    lignes en trop (ligne entière). Résultat : le tableau "Import CSV"
    s'arrête proprement à la dernière ligne réelle de données (confirmé par
    une frame ultérieure montrant les lignes s'arrêter à 2908, ligne 2909
    et suivantes vides).

## Partie 10 — Onglet "Bilan Factures" : réconciliation contre le PDF
(01:50 – 02:10 puis retour 08:00 – 08:20)

23. (Vu tôt dans la vidéo, ~01:50, et confirmé à nouveau vers 08:00) Onglet
    **Bilan Factures** : TCD à 2 colonnes natives + 3 colonnes de contrôle
    manuel — **A = "Étiquettes de lignes"** (numéro de facture, ex.
    `CO01192032`) | **B = "Somme de Total HT"** (TCD, somme automatique de
    la colonne AA "Total HT" de "Facture Colissimo") | **D = "PDF HT"**
    (saisi manuellement depuis le PDF de la facture Colissimo réelle) |
    **E = "Indemnisation"** (saisi manuellement) | **F = "Avoir"** (saisi
    manuellement si applicable) | **G = "Somme total"** = `=D4+E4+F4`.
    Exemple de valeurs observées (mois d'avril 2026 dans la vidéo) :
    | Étiquettes de lignes | Somme de Total HT | PDF HT | Indemnisation | Avoir | Somme total |
    |---|---|---|---|---|---|
    | (vide) | — | 24732,08 | 190,51 | (vide) | 24922,59 |
    | CO01192032 | 24922,59 | | | | |
    | **Total général** | **24922,59** | | | | |
    Le TCD "Somme de Total HT" (**24922,59**) est **strictement égal** à la
    "Somme total" calculée en colonne G (**24922,59** = 24732,08 + 190,51 +
    0) — confirme le mécanisme de contrôle : **Somme de Total HT (calculée
    dans le classeur) DOIT être égale à PDF HT + Indemnisation + Avoir
    (valeurs saisies à la main depuis le PDF officiel)**. C'est le
    mécanisme de réconciliation demandé par les notes : "Mettre le prix PDF
    HT et montant indemnisation pour comparer écart" — sauf qu'ici il n'y a
    PAS de colonne "écart" explicite visible (contrairement à Chronopost où
    une colonne "écarts" = Somme HT - pdf existe) ; la comparaison se fait
    par LECTURE VISUELLE de l'égalité entre B6 (Total général TCD) et G4
    (Somme total calculée D+E+F).
24. (08:00 – 08:20, navigateur Edge, onglet PDF `facture_2026-04-30.pdf`
    hébergé sur le partage réseau) : le PDF de facture La Poste-Colissimo
    montre, à la page 1, un tableau récapitulatif "Votre récapitulatif de
    facture HT" avec les postes suivants et leurs montants exacts :
    | Poste | Montant |
    |---|---|
    | Port Brut | 27 385,33 € |
    | Remise | -7 198,85 € |
    | Port Net | 20 186,48 € |
    | Coefficient Ajustement Energie | 1 548,66 € |
    | Suppléments | 349,90 € |
    | Options | 144,00 € |
    | Prestations Complémentaires | 2 693,55 € |
    | Indemnisations | -190,51 € |
    | **TOTAL HT** | **24 732,08 €** |
    Puis un second tableau : TVA 20,00% → Total HT 20 257,73 € / Montant
    TVA 4 051,55 € / Total TTC 24 309,28 € ; TVA 0,00% → Total HT 4 474,35 €
    / Montant TVA 0,00 € / Total TTC 4 474,35 € ; **TOTAL** = 24 732,08 € HT
    / 4 051,55 € TVA / **28 783,63 € TTC**.
    Confirme EXACTEMENT que **la valeur "24732,08" saisie dans "Bilan
    Factures" colonne D "PDF HT" correspond au TOTAL HT du PDF officiel**,
    et que **"190,51" en colonne E "Indemnisation" correspond à la valeur
    ABSOLUE de la ligne "Indemnisations" du PDF (-190,51 €, signe inversé
    car dans "Bilan Factures" l'indemnisation s'AJOUTE positivement — cf.
    formule G4=D4+E4+F4 — alors que dans le PDF elle est soustraite du
    Port Net)**. Le "Coefficient Ajustement Energie" (1 548,66 €) du PDF
    correspond au poste "**Taxe gazole**" identifié par le "Code charge"
    `CH_CAP` = "CAE" (Ajustement Energie) dans "Facture Colissimo" (cf.
    Partie 5, point 8) — c'est la traduction concrète de la "Taxe Gasoil"
    évoquée en tête des notes manuscrites (section suivante).

## Partie 11 — Taxe Gasoil / Coefficient d'Ajustement Énergétique
(observations transversales, pas de navigation web vers le site Colissimo
visible dans cette vidéo)

25. **Aucune frame de cette vidéo ne montre le site
    `colissimo.entreprise.laposte.fr` ni le "coefficient d'ajustement
    énergétique"** mentionné en tête des notes manuscrites — contrairement
    à ce que la consigne suggérait, cette vidéo ne montre PAS la
    consultation du taux officiel sur le site Colissimo. Le mécanisme
    observé dans cette vidéo est différent et plus direct : la taxe gasoil
    (**"CAE" = Coefficient d'Ajustement Énergétique**, Code charge
    `CH_CAP`) est **une ligne de facturation déjà calculée et fournie
    directement par Colissimo** dans le CSV "Prestations au colis" (colonne
    "Total HT" de chaque ligne "CAE", montant déjà en euros, PAS un
    pourcentage à appliquer manuellement). Le classeur ne fait que
    RECLASSER cette ligne dans la catégorie "Taxe gazole" via la Table de
    correspondance (`CH_CAP → Taxe gazole`), puis la reporter telle quelle
    (montant déjà calculé par Colissimo) dans la colonne "Taxe gazole" (K)
    du TCD, puis potentiellement dans la colonne "Gazole" (V, non confirmée
    par formule exacte) de l'onglet Import CSV. **Il n'y a donc, dans cette
    vidéo, aucun taux constaté à l'écran ni de calcul manuel de la taxe
    gasoil** — contrairement à Chronopost (taux AG8/AG9 saisis à la main) —
    le montant de la taxe gasoil Colissimo est **directement fourni ligne
    par ligne par le CSV source Colissimo**, pas calculé par une formule de
    pourcentage dans le classeur. Le lien mentionné dans les notes vers le
    site Colissimo sert donc probablement uniquement à VÉRIFIER a
    posteriori (contrôle de cohérence) que le taux appliqué par Colissimo
    dans ses lignes "CAE" correspond bien au taux publié officiellement,
    plutôt qu'à SAISIR un taux dans le classeur.

## Partie 12 — Nouvelle réconciliation "Bilan Factures" après ajout du CSV
douane (08:20 – 08:40)

26. Retour sur l'onglet "Bilan Factures" : les valeurs restent identiques
    (24732,08 / 190,51 / 24922,59) et confirment l'égalité avec la Somme de
    Total HT du TCD (24922,59) — cette réconciliation a lieu APRÈS l'ajout
    des lignes "Frais de douane" (Partie 5), suggérant qu'à ce stade de la
    vidéo, le fichier était déjà à son état final pour cette étape de
    contrôle.

## Partie 13 — Export et contrôle du CSV final `2026_04_Colissimo_Import.csv`
et comparaison avec le mois précédent (08:40 – 10:30)

27. (08:40 – 09:10) Retour sur l'onglet **Import CSV** du classeur : mise en
    évidence des colonnes A à W avec filtres automatiques activés
    (triangle de filtre sur chaque en-tête). Valeurs observées pour
    plusieurs lignes consécutives : Transporteur = "COLISSIMO", Date
    validité tarif = "01/04/2026" uniforme, Pays = "FR", Zone = "FR_DOM",
    mode envoi = "DOM", Nbr Colis = 1, Poids variable (0,1 à 2,2 kg), TVA =
    0,2, Frêt variable (5,06 à 8,47).
28. (09:10 – 09:40) Ouverture de l'Explorateur de fichiers Windows sur le
    dossier réseau **`$Facturation automatique\1 - Factures transporteurs +
    calculs\2026\2026 04\Colissimo\`** : contient `CSV_Frais_de_douane.csv`,
    `CSV_Indemnisations.csv`, `CSV_Prestations_au_colis.csv`,
    `facture_2026-04-30.pdf` (4 éléments) — confirme la structure exacte du
    dossier mensuel par transporteur, cohérente avec les autres
    transporteurs déjà documentés (Chronopost, Geodis, etc.).
29. (09:40 – 10:10) Le fichier final **`2026_04_Colissimo_Import.csv`** est
    ouvert (ou généré par copie de l'onglet "Import CSV" collée en valeurs
    dans un nouveau classeur/CSV) : en-têtes identiques à l'onglet Import
    CSV (Transporteur, Date validité tarif, Réf.1, Réf. 2, Id client, N°
    Tracking, Nom, E/P, Pays, Zone, Nbr Colis, Poids, mode envoi, TVA,
    Droits et taxes, Assurance, Zones éloignées, Colis volumineux, Adresses,
    Frêt, plus-value BtoC, Gazole, Nb colis).
30. (10:10 – 10:30) Ouverture en parallèle du fichier **`2026_03_Colissimo_
    Import.csv`** (mois précédent, mars 2026) comme référence de structure
    — même en-têtes confirmés, données similaires (Transporteur=COLISSIMO,
    Date=01/03/2026, Pays=FR, Zone=FR_DOM, mode envoi=DOM).

## Partie 14 — QA finale : détection et résolution des lignes "Pays à
créer" / "zone inconnue" par recherche croisée sur le mois précédent
(10:30 – 12:54)

31. (10:30 – 11:00) **Filtre appliqué sur la colonne "Pays" (colonne I de
    Import CSV) ou "Zone" (J) pour isoler les valeurs "Pays à créer" et
    "zone inconnue"** : **26 lignes** trouvées (barre d'état Excel : "26
    enregistrement(s) trouvé(s) sur 2907"). Ces lignes ont pour "N°
    Tracking" des préfixes `CB4...`/`CB5...` (ex. `CB461581203FR`,
    `CB464911513FR`, `CB474258172FR`, `CB489501267FR`, `CB493299166FR`) —
    tous des colis avec préfixe **"CB"** (Colissimo Domicile Signature
    International), cohérent avec des envois internationaux vers des pays
    qui ne sont pas encore répertoriés correctement dans la table "Pays" ou
    "Modes envois" pour ce préfixe/pays précis.
32. (11:00 – 11:40) Pour résoudre ces lignes, l'utilisateur **ouvre le CSV
    du mois précédent `2026_03_Colissimo_Import.csv`** et tape une formule
    de type
    `=RECHERCHEX($F21;TCD!D:D;TCD!E:E)` (variante observée dans une
    frame, avec IF pour gérer les 0) pour tenter de retrouver, PAR NUMÉRO
    DE TRACKING, si ce même colis existe déjà dans les données du mois
    précédent (peu probable pour un nouveau tracking, mais confirme la
    tentative de RÉCUPÉRATION CROISÉE d'informations manquantes -
    Pays/Zone/mode envoi - en cherchant si un tracking similaire ou le même
    modèle de colis a déjà été traité). Une formule plus generale est
    visible :
    `=SI(RECHERCHEX($F21;TCD!D:D;TCD!E:E)=0;"";RECHERCHEX($F21;TCD!D:D;TCD!E:E))`
    (pattern cohérent avec les formules déjà vues en Partie 8, adaptée ici
    pour tester l'existence dans le TCD).
33. (11:40 – 12:20) Vérification/complément manuel : les valeurs "Pays à
    créer"/"zone inconnue" identifiées semblent nécessiter un ajout MANUEL
    dans les tables de référence "Pays" et/ou "Modes envois" (nouveau pays
    ou nouveau préfixe non encore répertorié) — **aucune frame ne montre
    l'ajout effectif d'une nouvelle ligne dans "Pays" ou "Modes envois"
    pendant cette vidéo**, seule la DÉTECTION du problème (filtre à 26
    lignes) est clairement filmée ; la correction elle-même (quelles lignes
    ont été ajoutées aux tables de référence, avec quelles valeurs) n'est
    pas visible dans les frames disponibles.
34. (12:20 – 12:54, fin de la vidéo) Retour sur le fichier `2026_04_
    Colissimo_Import.csv`, cellule B1 sélectionnée ("Date validité tarif"),
    poursuite du défilement/vérification des données. La vidéo se termine
    sans étape de clôture/export finale visible après cette phase de QA
    (pas de sauvegarde explicite filmée dans les toutes dernières
    secondes).

---

## Synthèse de la réconciliation (point par point, vs. notes manuscrites)

| # | Point des notes | Statut | Précision apportée |
|---|---|---|---|
| 1 | Taxe Gasoil, lien vers le site Colissimo, "coefficient d'ajustement énergétique" | **Non observé directement** | Aucune navigation vers le site n'apparaît dans cette vidéo. Le montant de la taxe gasoil ("CAE"/Coefficient Ajustement Energie, Code charge `CH_CAP`) est fourni DIRECTEMENT par Colissimo dans le CSV "Prestations au colis" (déjà calculé en euros), pas recalculé par un taux dans le classeur. Le site sert probablement à un contrôle a posteriori seulement. |
| 2 | Télécharger les Excels, ouvrir le CSV "prestation Colis et douanes" | **Confirmé (précisé)** | Il s'agit en réalité de 2 fichiers CSV distincts : `CSV_Prestations_au_colis.csv` (prestations liées aux colis) et `CSV_Frais_de_douane.csv` (frais de douane), tous deux au format `;`-délimité, avec un 3e fichier `CSV_Indemnisations.csv` non montré collé dans la vidéo mais présent dans le dossier réseau. |
| 3 | Coller dans Facture Colissimo à partir de la colonne H | **Confirmé exactement** | Collage des colonnes "N de ligne" à "Type de charge" du CSV, dans "Facture Colissimo" à partir de la colonne H, pour les 2 fichiers CSV successivement (Prestations au colis puis Frais de douane, ajoutés à la suite). |
| 4 | Trier Colonne O de A à Z | **Confirmé exactement** | Tri du tableau "Facture Colissimo" sur la colonne O "N colis" (numéro de tracking), pour regrouper les lignes d'un même colis. |
| 5 | Étendre les colonnes A à G et rajouter à la fin de la colonne G | **Confirmé exactement** | Colonnes A (Pays), B (Mode Envoi), C (Zone), D (concatener), E (Préfixe), F (Pays 2 lettres), G (Catégorie) sont des formules RECHERCHEX/XLOOKUP réétirées vers le bas après chaque ajout de lignes. |
| 6 | Onglet Poids : MAJ TCD (colonne B), trier Colonne A de A à Z | **Confirmé (structure), tri non observé net** | Onglet Poids = TCD "Étiquettes de lignes" (tracking) / "Moyenne de Poids Kg". Structure confirmée via fichier réel. L'actualisation et le tri n'ont pas été captés comme actions isolées nettes dans les frames. |
| 7 | Onglet TCD : MAJ TCD (colonne L), trier Colonne C de A à Z, étendre colonne B, effacer ID client | **Confirmé (structure et actions), colonne L à vérifier** | TCD confirmé avec colonnes A=ID Client, B=Total hors Gazole (`SUM(E:J,L)`), C=Total + GO (`SUM(E:L)`), D=Étiquettes de lignes, E à L = 7 postes (Adresse/Assurance/Colis volumineux/Droits et taxes/Frêt/plus-value BtoC/Taxe gazole/Zones éloignées). Tri sur C et effacement de la colonne A (ID Client) confirmés visuellement. La consigne "MAJ TCD colonne L" ne correspond pas exactement à une colonne du TCD réel (L = "Zones éloignées", pas une colonne de MAJ identifiable) — possible confusion avec l'ordre des colonnes, à clarifier. |
| 8 | Onglet Bilan Facture : MAJ TCD (colonne B), mettre le prix PDF HT et montant indemnisation pour comparer écart | **Confirmé exactement** | "Bilan Factures" = TCD Somme de Total HT (colonne B) par N facture, comparé à PDF HT (D) + Indemnisation (E) + Avoir (F) = Somme total (G). Égalité stricte observée (24922,59 = 24922,59) pour valider le mois. Valeurs croisées exactement avec le PDF réel `facture_2026-04-30.pdf` (Total HT 24 732,08€, Indemnisations -190,51€). |
| 9 | Onglet Import ERP : étendre les colonnes | **Confirmé, nommé "Import CSV"** | L'onglet s'appelle "Import CSV" (pas "Import ERP") dans le classeur réel. Colonnes A à W étirées avec formules XLOOKUP vers "Facture Colissimo" et "TCD". Bug de recopie au-delà des données réelles observé et corrigé (suppression des lignes en trop, cascade de #N/A). |
| 10 | Vérifier somme OK : somme de toutes les colonnes import ERP + colonne TG de l'onglet TCD | **Partiellement confirmé** | Le mécanisme de vérification observé est plutôt l'égalité "Bilan Factures" (Somme de Total HT = PDF HT + Indemnisation + Avoir), pas une somme explicite des colonnes de "Import CSV" + colonne TG (Taxe Gazole = colonne K du TCD). Aucune formule combinant explicitement "somme Import CSV + TCD colonne K" n'a été repérée dans les frames disponibles — possible contrôle mental/manuel non retranscrit dans une formule visible. |

---

## Points ambigus / illisibles à faire confirmer par le pôle transport

1. **Aucune navigation vers le site
   `colissimo.entreprise.laposte.fr/offres-et-services/tarifs-generaux/
   supplements-tarifaires-colissimo` n'apparaît dans cette vidéo.** Le
   mécanisme réel de la taxe gasoil semble être que Colissimo fournit déjà
   le montant calculé (ligne "CAE"/Coefficient Ajustement Energie) dans son
   CSV "Prestations au colis" — à confirmer si le site sert uniquement de
   contrôle de cohérence a posteriori du taux appliqué par Colissimo, ou
   s'il existe une autre vidéo Colissimo (non fournie ici) montrant cette
   consultation.
2. **Correspondance colonne à colonne entre "Prestations au colis" et
   "Frais de douane"** lors du collage à la suite dans "Facture Colissimo"
   colonnes H-Y : les 2 fichiers CSV ont des colonnes différentes après "N
   colis" (Zone/Pays Origine/Poids/dimensions pour le premier, vs Droit de
   douane/TVA importation/Octroi de mer/Frais de gestion/Autres taxes pour
   le second) — la vidéo ne permet pas de confirmer avec certitude que le
   collage aligne bien "Total HT" du fichier douane avec la colonne AA
   "Total HT" de "Facture Colissimo" (l'alignement des colonnes lors du
   copier-coller mériterait vérification directe dans le fichier réel,
   ligne par ligne, colonne par colonne, en comparant les positions
   relatives des 2 CSV sources).
3. **Formule exacte de la colonne V "Gazole" de l'onglet Import CSV** :
   non confirmée caractère par caractère dans les frames disponibles
   (probablement `=XLOOKUP(F2,TCD!D:D,TCD!K:K)` par symétrie avec les
   colonnes O à U, mais à vérifier directement dans le fichier réel).
4. **Colonne "ID Client" du TCD** : formule d'origine (comment elle est
   remplie AVANT d'être effacée) non capturée nette dans les frames — la
   vidéo montre l'effacement mais pas le remplissage initial.
5. **Consigne "Onglet TCD : MAJ TCD (colonne L)"** des notes manuscrites ne
   correspond à aucune action identifiable clairement sur la colonne L du
   TCD réel (= "Zones éloignées") dans les frames disponibles — à
   clarifier avec le pôle transport (coquille possible dans les notes, ou
   référence à une autre version du classeur).
6. **Mécanisme exact de "vérifier somme OK : somme de toutes les colonnes
   import ERP + colonne TG de l'onglet TCD"** (dernière ligne des notes) :
   aucune formule combinant explicitement cette somme n'a été repérée. Le
   contrôle observé dans la vidéo est l'égalité "Bilan Factures" (TCD vs
   PDF+Indemnisation+Avoir), pas une comparaison somme Import CSV + TG TCD
   vs quelque chose d'autre. Possible confusion ou contrôle mental non
   formalisé en formule Excel — à clarifier directement avec l'utilisateur.
7. **Ajout effectif de nouvelles lignes dans les tables "Pays" / "Modes
   envois"** pour résoudre les 26 lignes "Pays à créer"/"zone inconnue"
   détectées en fin de vidéo (Partie 14) : la DÉTECTION est bien filmée
   (filtre à 26 résultats), mais la CORRECTION elle-même (quel pays/préfixe
   exact a été ajouté, avec quelles valeurs) n'apparaît pas dans les frames
   disponibles — la vidéo se termine sur la phase de diagnostic sans
   montrer la résolution complète ni l'export/sauvegarde final.
8. **Rôle exact du 3ème CSV `CSV_Indemnisations.csv`** (présent dans le
   dossier réseau, en-têtes confirmés via le fichier CSV brut : N de ligne,
   N facture, Date, Compte facturé, Compte déposant, Description, Produit,
   N colis, Référence, Total HT, Taux TVA, Rubrique de la facture = 
   "INDEMNISATIONS", Code charge (ex. CH_I_IAF, CH_I_IN3, CH_I_IN6), Balises
   LIN EDI, Offre, Type de charge = "A") : ce fichier n'a jamais été
   ouvert/collé dans la vidéo — son usage (probablement à coller également
   dans "Facture Colissimo" à la suite des 2 autres, ou à utiliser
   uniquement pour la valeur "Indemnisation" saisie manuellement dans
   "Bilan Factures") reste à confirmer. Le rapprochement avec le montant
   "190,51" du PDF (Indemnisations totales) suggère que ce CSV sert de
   justificatif au montant saisi en E4 de "Bilan Factures", mais la vidéo
   ne montre pas cette étape de calcul/vérification (somme des lignes
   d'indemnisation = 190,51 ?) de façon isolée et nette.
9. **Formule "MAJ colonne L" de l'onglet TCD et interprétation "Effacer ID
   client"** : l'action d'effacement de la colonne A "ID Client" du TCD est
   visible mais son TIMING exact par rapport aux autres étapes (avant ou
   après le tri sur colonne C, avant ou après l'étirement de colonne B)
   n'est pas garanti à 100% dans l'ordre chronologique retranscrit ici
   (plusieurs actions rapides et proches dans le temps entre 05:10 et
   06:00, frames à 6s d'intervalle pouvant manquer l'ordre exact des
   micro-actions).
10. **Résolution de zoom insuffisante sur certaines frames** (notamment la
    zone TCD lignes 06:00-06:20 et certaines cellules de formule de la
    Partie 8/9) : quelques valeurs de colonnes ont été confirmées
    principalement via le fichier réel `2026_06_Facture Colissimo.xlsx`
    plutôt que lues directement dans la vidéo — cohérent avec la démarche
    demandée (croiser vidéo + fichier réel), mais signalé ici pour
    transparence sur la source de chaque affirmation.
