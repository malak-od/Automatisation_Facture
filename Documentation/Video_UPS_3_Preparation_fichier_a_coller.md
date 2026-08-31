# Video UPS 3 - Preparation fichier a coller

Source : `Transporteurs/UPS/UPS_1_Preparation fichier a coller.mp4`
Duree : 413.0 secondes (~6 min 53 s), capture d'ecran silencieuse (Excel/Explorateur Windows).
Methode : extraction 1 frame / 4s (103 frames fixes) + detection de changement de scene
(38 frames), soit ~141 frames lues en ordre chronologique.

## Contexte

Cette video est une **nouvelle capture du meme process** que celui documente dans
`Documentation/Video_UPS_1_Integration_CSV.md` (integration des fichiers CSV UPS bruts dans le
classeur Excel de calcul via Power Query), mais realisee sur une **periode plus recente** :
mois traite = **avril 2026** (dossier source `2026 04`, classeur cible `2026_04_Facture UPS.xlsx`,
copie/renomme a partir de `2026_03_Facture UPS.xlsx` du mois precedent). Compte UPS observe :
`0000A1912W` (transporteur "UPS", pas "UPS_COD").

Aucune etape specifique au probleme "Poids = 0" n'a ete identifiee dans cette video : le
traitement du poids se limite au report brut des colonnes CSV sources ("Poids annonce",
"Poids facture") dans le tableau Power Query, sans recalcul ni repli manuel visible. Voir la
section "Reponse a la question posee" en fin de document.

## Transcription chronologique

### Phase 1 - Localisation des fichiers CSV source (0:00 - 0:20)

1. (0:00-0:08) Classeur Excel vierge ("Classeur1"), onglet **Donnees** actif. L'utilisateur
   engage le meme chemin que la video de reference : **Obtenir des donnees > A partir d'un
   fichier > A partir d'un dossier**.
2. (0:08-0:20) Boite de dialogue **"Parcourir"** : navigation vers le dossier source des CSV
   UPS du mois (chemin reseau `\\192.168.5.3\Comptabilite La Ruche\$Facturation
   automatique\1 - Factures transporteurs + calc...`), sous-dossier du mois **avril 2026**
   (`2026 04`). Liste de fichiers `Invoice_<numero>_<date>.csv`, ex.
   `Invoice_202600369298_040726.csv`, `Invoice_202600378659_041026.csv`,
   `Invoice_202600399330_041726.csv` -- nommage identique au pattern deja documente (numero
   de facture + date JJMMAA).

### Phase 2 - Power Query "Combiner et charger" (0:20 - 1:00)

3. (~0:20-0:28) Fenetre recapitulative du dossier source (colonnes Content/Name/Extension/Date
   accessed/Date modified/Date created/Attributes), clic sur **Combiner > Combiner et charger**.
4. Boite **"Combiner les fichiers"** : "Exemple de fichier" = Premier fichier, Origine =
   "1252 : Europe de l'Ouest (Windows)", Delimiteur = Virgule, Detection du type de donnees =
   "Selon les 200 premieres lignes". Apercu du premier CSV, colonnes generiques Column1 a
   Column11 : Column1=`2.1` (version format), Column2=Column3=`0000A1912W` (compte),
   Column4=`FR`, Column5=`2026-04-07` (date ISO), Column6=`202600369298` (numero facture),
   Column7=`I`, Column8=`6`, Column9=`FR75804949865`, Column10=`EUR`, Column11=`59.50`
   (montant, ligne d'en-tete de facture). Clic **OK**.
5. (~0:40-1:00) Chargement en arriere-plan : cellule A1 = "DonneesExternes_1 : lecture des
   donnees...". Panneau **Requetes et connexions** confirme la meme structure a 5 requetes que
   la video de reference (Transformer le fichier a partir de U... [2], Requetes d'assistance
   [3], Autres requetes [1] = **UPS**). Volume final affiche : **"62 184 lignes chargees"**
   (contre 68 972 lignes pour le mois de mai 2025 dans la video de reference -- volume mensuel
   variable).

### Phase 3 - Tableau "UPS" charge, colonnes generiques (1:00 - 3:20)

6. (1:00-1:20) Onglet **UPS** actif (tableau structure, nom "UPS"), en-tetes generiques
   **Source.Name, Column1 ... Column250+** (Power Query n'a pas renomme les colonnes a ce
   stade, comme dans la video de reference). Colonnes visibles autour de W/X/Y/Z (~1:00) :
   valeurs "K" (kg), "PKG" -- coherent avec les colonnes poids/type de colis deja identifiees
   (positions Poids annonce/Poids facture) mais **non encore renommees**, donc non
   identifiables avec certitude sur cette portion de la video.
7. (~1:04-3:00) Defilement horizontal progressif dans le tableau generique jusqu'a des colonnes
   tres avancees (Column171 a Column250+), confirmant la tres grande largeur du tableau brut
   (au-dela de 250 colonnes generiques avant renommage), avec de nombreuses colonnes vides ou a
   `0`. Colonnes reperees en fin de defilement : `FRAIS DE SERVICE SOLUTION CLIENTELE`,
   `MONTHLY FEE`, `UNDELIVERABLE MANAGE REROUTE`, `APR 2026` -- lignes de frais mensuels
   forfaitaires distinctes des lignes de detail par colis.
8. (~2:52-3:08, cellule BB5) Cellule active affichant `0.00`, colonne **Column53** filtree :
   barre de statut en bas indique **"29299 enregistrement(s) trouve(s) sur 62184"** -- confirme
   qu'un filtre (vraisemblablement "differents de 0" sur une colonne de montant, meme logique
   que la video de reference) a ete applique, isolant une sous-population d'environ 29 300
   lignes sur 62 184.

### Phase 4 - Rechercher/Remplacer point -> virgule (3:20 - 3:40, deduit du contexte)

9. Le meme mecanisme de conversion de separateur decimal (point americain -> virgule
   francaise) que dans la video de reference (`Video_UPS_1_Integration_CSV.md`, etapes 25-27)
   est presume avoir eu lieu dans cette plage -- **non capture avec certitude sur une frame
   dediee dans cet echantillonnage**, mais les valeurs numeriques visibles par la suite
   (colonne "Montant net" au format virgule) sont coherentes avec un traitement identique.

### Phase 5 - Ouverture du classeur modele du mois precedent (~3:20 - 3:36)

10. (~3:24) Fenetre **"Ouvrir"** Excel : navigation dans `1 - Factures transporteurs + calculs
    > 2026`, sous-dossiers **2026 01** a **2026 04** visibles a gauche (**2026 04** en double
    dans la liste des emplacements recents), **2026 03** correspondant au dossier ouvert dans
    la vue de droite (`2026 03`, modifie le 06/05/2026 10:43). Auteur = **Thomas Largeron**.
11. (~3:24-3:28) Deuxieme vue de la meme boite "Ouvrir" : navigation vers le dossier
    **`2026 04`** (`1 - Factures transporteurs + calculs > 2026 > 2026 04`), listant les
    sous-dossiers par transporteur : BLS, CEVA, Chronopost, Colissimo, Delivengo, DHL, DPD,
    Fedex... (meme structure de dossiers que la video de reference, mois different).
12. (~3:28) Le fichier **`2026_03_Facture UPS.xlsx`** est ouvert en **Mode protege** (bandeau
    jaune "Attention aux fichiers provenant d'un emplacement Internet... Activer la
    modification") -- confirme que le classeur du **mois precedent (mars 2026)** est ouvert en
    lecture pour servir de modele/base a dupliquer, exactement comme dans la video de
    reference (avril servant de modele pour mai).
13. (~3:28-3:32) Vue sur l'onglet **Facture UPS** du classeur `2026_03_Facture UPS.xlsx`
    (colonnes V a BX) : cellule active AX27582 = **"Correction d adresse Dom. Standard"**.
    Colonnes confirmees : **W** = "Nombre de coli(s)", **Y** = "Numero de suivi", **AE** =
    "Poids a(nnonce)" (ex. `0`, `38 K`, `16 K`, `15 K`, `19 K`, `5 K`, `10 K`, `20 K`), **AG** =
    "Poids factu(re)" (ex. `0`, `38 K`, `12 K`, `15 K`, `19 K`, `5,5 K`, `10 K`, `20,5 K`),
    **AH** = "Type de co(lis)" = PKG, **AV** = "Code de..." (FRT/TAX/FSC/ACC), **AW** = "Code"
    (RES...), **AX** = "Description des frais" (Correction d adresse Dom. Standard, 20.000 %
    Tax, Dom. Standard, Taxe S/Carburant, Liv.particulier), **BE** = "Montant..." (colonne
    Montant net, valeurs `22,2`, `4,44`, `13,39`, `2,08`...).
    **Point notable pour le poids** : sur les lignes ou "Nombre de colis" = 0 (ex. lignes 24625,
    24626 -- categorie "Correction d adresse" / "20.000 % Tax"), les colonnes "Poids annonce"
    et "Poids facture" affichent **0** -- confirme que les lignes de frais annexes/correctifs
    (pas des lignes de transport physique de colis) ont un poids nul par construction cote UPS
    (le CSV source ne fournit pas de poids pour ces lignes-la), et non un defaut de recuperation
    de donnee cote fichier de calcul.

### Phase 6 - Enregistrer sous le nom du mois traite (~3:36 - 3:44)

14. (~3:36) Boite **"Enregistrer sous"** : dossier cible `1 - Factures transporteurs + calculs
    > 2026`, sous-dossiers **2026 04** (deux fois dans les emplacements recents), **2026 03**,
    **GLS**. Champ **"Nom de fichier"** pre-rempli avec **"2026_03_Facture UPS.xlsx"** (nom du
    modele ouvert), **Type** = "Classeur Excel (*.xlsx)".
15. (~3:40) Champ "Nom de fichier" modifie manuellement en **"2026_04_Facture UPS.xlsx"**
    (remplace "03" par "04" -- meme logique que la video de reference, adaptee au mois avril).
    D'autres emplacements recents visibles dans l'historique : `2026 04\...\2 - Fichiers csv
    i...`, `GLS\...\1 - Factures tran...`, `Fedex\...\1 - Factures tran...`.

### Phase 7 - Classeur cible ouvert, structure d'onglets (3:44 - 4:00)

16. (~3:44) Classeur renomme **"2026_04_Facture UPS.xlsx"** actif, onglet **Facture UPS**
    affiche (colonnes G a Y) : en-tetes confirmes -- **Numero** (facture), **Pays d'or(igine)**,
    **Date de la facture**, **Numero de facture**, **Code de...**, **Code dét...**, **Numero**,
    **Code de...**, **Montant**, **Date de l'operation**, **Reference**, **Numero de l'envoi
    principal**, **Numero**, **Numero de reference**, **Numero**, **Code de**, **Nombre de
    coli(s)**, **Nombre**, **Numero de suivi**. Liste complete des onglets en bas de l'ecran :
    **Charge.CHG_CODE, ST SV, Zone, Clients log, zone colis poids assurance, Bilan factures,
    Facture UPS, Comptes UPS, TCD, Fichier import, Demande avoir, Bilan clients, Adresse**
    (un onglet **"Adresse"** supplementaire par rapport a la liste observee dans la video de
    reference de 2025, qui listait plutot "Gazole", "CODES SVCE LEVEL",
    "CODIFICATION CODE EXCEPTION" en debut de liste -- possible reorganisation/ajout d'onglet
    entre les deux captures).
17. Colonnes **C** ("Mode env(oi)" = ST), **D** ("Categorie" : Adresse, TVA, Fret, Taxe gazole,
    plus-value BtoC, Colis volumineux -- meme jeu de valeurs que la video de reference) sur les
    lignes 24625+ (donnees de mars 2026, avant l'ajout des nouvelles lignes d'avril).

### Phase 8 - Ajout ("append") des nouvelles donnees d'avril a la suite de "Facture UPS" (4:00 - 6:52)

18. (~4:04-6:52) Defilement continu dans l'onglet **Facture UPS** : les lignes ~24625 a ~28570
    montrent encore les donnees de mars 2026 (dates de facture 03/04/2026, montant facture
    globale `197166,06`, tracking `1ZA1912WDK9...`). A partir d'un certain point (repere par le
    changement de date de facture a **05/05/2026** et montant `186092,91`/`185,86`), les
    **nouvelles lignes issues de l'import CSV d'avril sont visibles collees/ajoutees a la suite
    de l'historique existant** -- confirme le mecanisme d'**ajout cumulatif (append)** deja
    suppose dans la video de reference : le nouvel import "UPS" (Power Query) est bien reporte
    a la suite de "Facture UPS" et non recree/vide chaque mois.
19. (~6:44-6:52, dernieres frames) Defilement jusqu'aux lignes **~57479-57510+** puis
    **~43362-43395** (ordre de lecture non strictement lineaire du defilement, l'utilisateur
    navigue/verifie plusieurs zones) : tracking `1ZWV57887...` (compte `0000WV578`, montant
    facture `219,36`), puis tracking `1ZA1912WDK9639...` (compte `0000A1912`, montant facture
    `186092,91`), avec des colonnes Categorie variees (TVA, Fret, Taxe gazole, plus-value BtoC,
    **Colis volumineux**). Cellule active toujours E24625 (curseur reste positionne sur cette
    cellule de reference pendant le defilement, `Ctrl` affiche en bas a droite -- probablement
    en cours de collage Ctrl+V des nouvelles lignes).
    La video se termine sur ce defilement/verification final sans action metier distincte
    supplementaire identifiee.

## Reponse a la question posee (poids = 0 dans le fichier import)

**Aucune methode manuelle specifique de correction/repli du poids n'apparait dans cette
video.** Le traitement du poids se limite strictement au **report brut** des colonnes CSV UPS
sources ("Poids annonce" / "Poids facture", exprimees en kg avec suffixe "K") dans le tableau
Power Query puis dans l'onglet "Facture UPS" -- aucune formule de secours, aucun export
complementaire, aucune saisie manuelle de poids n'est visible. Le seul enseignement pertinent
est que **les lignes de frais annexes sans transport physique de colis (corrections d'adresse,
taxes, frais de service mensuels forfaitaires...) ont un "Nombre de colis" = 0 ET un poids = 0
directement dans le CSV source UPS** -- ce n'est pas un defaut de recuperation de donnee mais
un etat normal du CSV pour ces lignes-la. Si le pipeline actuel (`facturation-app/src/carriers/
ups/index.js`) agrege le poids par tracking via un MAX sur toutes les lignes d'un meme colis
(cf. commentaire code "Poids/Nombre de colis/Montant assurance : agreges par tracking via MAX"),
un poids=0 sur le resultat final signifierait que **toutes les lignes de detail rattachees a ce
tracking ont un poids nul dans le CSV source lui-meme** (pas seulement les lignes de frais
annexes) -- ce cas de figure n'est pas traite explicitement dans cette video. Voir aussi la
video companion `Video_UPS_4_Preparation_fichier_import.md` pour la suite du process
(elaboration du fichier d'import final, ou la colonne Poids est reprise).

## Points ambigus / a confirmer

- **(etape 6-7)** Impossible d'identifier avec certitude, sur les frames capturees, le moment
  exact ou les colonnes generiques (Column1, Column53...) sont renommees en noms metier
  explicites (Poids annonce, Poids facture, etc.) -- contrairement a la video de reference ou
  cette transition est visible autour de 2:00-2:20. Dans cette video, le renommage semble avoir
  deja eu lieu au moment ou l'onglet "Facture UPS" est consulte (etape 13), mais l'operation de
  renommage elle-meme (clic sur chaque en-tete, Power Query "Modifier les colonnes"...) n'a pas
  ete captee sur une frame distincte.
- **(etape 8-9)** L'operation "Rechercher-remplacer point->virgule" sur la colonne Montant net,
  bien que tres probable par analogie avec la video de reference, n'a pas ete observee sur une
  frame dediee dans cet echantillonnage (peut-etre survenue entre deux frames fixes, dans un
  intervalle de moins de 4 secondes).
- **(etape 16)** Presence d'un onglet **"Adresse"** dans la liste des onglets du classeur avril
  2026, absent de la liste observee dans la video de reference (mai 2025). A confirmer avec le
  pole transport : nouvel onglet ajoute entre les deux periodes, ou simplement non visible car
  hors champ dans l'ancienne capture (liste d'onglets tronquee a l'affichage) ?
- **(etape 18)** Le moment exact du copier-coller (Ctrl+V) des nouvelles lignes d'avril dans
  "Facture UPS" n'a pas ete capture sur une frame isolant clairement l'action (barre de formule
  affichant une formule de collage, message de confirmation, etc.) -- seul l'indicateur `Ctrl`
  en bas a droite de l'ecran (etape 19) suggere qu'un collage special est en cours au moment de
  la derniere frame.
- Aucune mention, formule ou export complementaire lie specifiquement au poids (au-dela du
  report brut CSV) n'a ete repere sur l'ensemble de la video -- a confirmer explicitement avec
  le pole transport que cette video ne constitue pas la reponse recherchee au probleme "POIDS =
  0" dans le fichier d'import final (le repli/la formule concernee, si elle existe, serait
  plutot a chercher dans la video companion `Video_UPS_4_Preparation_fichier_import.md`).

(Transcription terminee sur la base des frames a intervalle fixe (0:00 a 6:52) recoupees avec
les frames de detection de changement de scene.)
