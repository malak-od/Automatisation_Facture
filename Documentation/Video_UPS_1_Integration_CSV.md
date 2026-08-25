# Video UPS 1 - Integration des fichiers CSV dans facture Excel UPS

Source : `Transporteurs/UPS/Process Facturation - Facture UPS - 1 - Intégration des fichiers csv dans facture Excel UPS.mp4`
Duree : ~5 min 11 s (311.3 s)

> Transcription complete (frames extraites toutes les 4s de 0:00 a 4:40,
> soit 70 frames couvrant la quasi-totalite de la video ; les timestamps
> sont approximatifs, arrondis a partir du numero de frame x4s).

## Etapes identifiees

### Phase 1 - Localisation des fichiers CSV source (0:00 - 0:16)

1. (0:00) Explorateur Windows ouvert sur
   `Disque 2 (D:) > Drive > Comptabilité La Ruche > $Facturation automatique
   > 1 - Factures transporteurs + calculs > 2025 > 2025 05`. On voit
   l'arborescence par transporteur (Chronopost, DPD, Fedex, GLS, Kuehne, TNT,
   UPS, UPS - 80X7Y5) et des classeurs `2025_05_Facture Chronopost`,
   `2025_05_Facture DPD`, `2025_05_Facture Kuehne` a la racine du mois. Noter
   qu'il existe deux dossiers UPS distincts pour la meme periode : `UPS` et
   `UPS - 80X7Y5` (2 comptes/contrats UPS differents -- confirme plus loin
   dans l'onglet "Comptes UPS" : "80X7Y5" est associe au libelle **UPS_COD**
   alors que les autres comptes sont associes a **UPS**).
2. (0:04-0:08) Clic sur le dossier `UPS` dans l'arborescence de gauche. Le
   volet de droite affiche une longue liste de fichiers `Invoice_<numero>_
   <date>.csv` (l'extension `.csv` est confirmee plus loin dans l'apercu
   Power Query). Exemples de noms : `Invoice_202500468212_050525`,
   `Invoice_202500477321_050625`, `Invoice_202500478200_050725`,
   `Invoice_202500485792_050925`, etc. Tailles variables (3 a 49 Ko), sauf
   deux fichiers nettement plus gros : `Invoice_202500565583_060325`
   (40 013 Ko) et `Invoice_202500569188_060325` (6 704 Ko) -- possibles
   exports globaux/consolides plutot que factures unitaires. 49 elements au
   total dans le dossier.
3. (0:08-0:16) Survol de la souris sur un fichier (`Invoice_202500500978_
   051425`) affichant l'infobulle : Type = "Fichier CSV Microsoft Excel",
   Taille = 6,10 Ko, Modifie le = 03/06/2025 06:15, Statut = "Disponible sur
   cet appareil" (fichier synchronise depuis le cloud/OneDrive ou Synology
   Drive local).

### Phase 2 - Nouveau classeur Excel + Power Query import de dossier (0:16 - 1:04)

4. (0:16-0:20) Bascule vers une fenetre Excel vierge intitulee "Classeur1 -
   Excel", feuille "Feuil1", cellule A1 selectionnee, aucune donnee.
5. (0:24) Ruban "Donnees" active. Clic sur "Obtenir des donnees" (groupe
   "Recuperer et transformer des donnees").
6. (0:24-0:28) Menu deroulant "Obtenir des donnees" -> sous-menu "A partir
   d'un fichier" ouvert, proposant : "A partir d'un classeur Excel", "A
   partir d'un fichier texte/CSV", "A partir d'un fichier XML", "A partir de
   JSON", "A partir d'un fichier PDF", "A partir d'un dossier", "A partir
   d'un dossier SharePoint". Le curseur survole "A partir d'un fichier PDF"
   avant de se diriger vers l'option choisie.
7. (0:28) Selection de **"A partir d'un dossier"** (Power Query import de
   dossier complet, PAS "a partir d'un fichier texte/CSV" unitaire) --
   confirme que l'integration se fait via Power Query sur tout le dossier et
   non par copier-coller manuel fichier par fichier.
8. (0:28-0:36) Boite de dialogue "Parcourir" pour choisir le dossier source.
   Navigation visible dans l'arborescence : `Comptabilité La Ruche` ->
   liste de dossiers `$Facturation automatique`, `$Factures clients`,
   `$Factures fournisseurs`, `$Fichiers pour facturation mensuelle`,
   `$Ouverture de compte`, etc. Puis descente vers
   `1 - Factures transporteurs + calculs > 2025 > 2025 05 > UPS` (le champ
   "Nom de dossier" affiche "UPS"). Dossier `UPS` mis en surbrillance dans
   l'arbre, valide avec "Ouvrir".
9. (0:44) Fenetre recapitulative du dossier source affichant le chemin
   `D:\Drive\Comptabilité La Ruche\$Facturation automatique\1 - Factures
   transporteurs + calculs\2...` et un tableau avec les colonnes **Content,
   Name, Extension, Date accessed, Date modified, Date created, Attributes**.
   Chaque ligne = un fichier `Invoice_<numero>_<date>.csv`, Content =
   "Binary". Message "Les donnees dans l'apercu ont ete tronquees en raison
   de limites de taille." En bas : boutons **Combiner** (menu deroulant),
   **Charger** (menu deroulant), **Transformer les donnees**, **Annuler**.
10. (0:48) Clic sur le menu deroulant **Combiner**, qui propose : "Combiner
    et transformer les donnees", "Combiner et charger", "Combiner et charger
    dans...". Le curseur s'arrete sur **"Combiner et charger"**.
11. (~0:52) Boite de dialogue "Combiner les fichiers" : "Exemple de fichier"
    = "Premier fichier", "Origine du fichier" = "1252: Europe de l'Ouest
    (Windows)", "Delimiteur" = "Virgule", "Detection du type de donnees" =
    "Selon les 200 premieres lignes". Apercu des donnees brutes du premier
    CSV (`Invoice_202500468212...`) sans en-tetes, colonnes generiques
    Column1 a Column11 :
    - Column1 = `2.1` (constante, probablement un numero de version de
      format d'export UPS)
    - Column2 = `0000A1912W` (identifiant compte/expediteur)
    - Column3 = `0000A1912W` (identique a Column2)
    - Column4 = `FR`
    - Column5 = `2025-05-05` (date, format ISO)
    - Column6 = `202500468212` (numero de facture, correspond au nom de
      fichier)
    - Column7 = `I` (code, valeur constante sur les lignes visibles)
    - Column8 = `5`
    - Column9 = `FR75804949865` (numero identifiant, ressemble a un SIRET/
      TVA)
    - Column10 = `EUR`
    - Column11 = `57.88` (montant, uniquement rempli sur la premiere ligne
      visible -- possible ligne d'en-tete de facture avec montant total, les
      lignes suivantes ont Column11 vide).
    Case a cocher "Ignorer les fichiers avec erreurs" (non cochee). Bouton
    **OK** valide.
12. (0:52-0:56) Retour a la feuille Excel : la cellule A1 affiche
    "DonneesExternes_1 : lecture des donnees..." pendant le traitement Power
    Query. Un nouvel onglet **"UPS"** apparait a cote de "Feuil1" en bas.
    Volet lateral droit **"Requetes et connexions"** ouvert, listant
    **5 requetes** organisees en dossiers :
    - "Transformer le fichier a partir de U..." [2]
    - "Requetes d'assistance [3]" : Exemple de fichier, Parametre1 (Exemple
      de fichier), Transformer le fichier, Transformer l'exemple de fichier
      (connexions uniquement, pas de chargement de donnees)
    - "Autres requetes [1]" : **UPS** -- initialement "5,17 Mo depuis
      Invoice_202500565583_0..." pendant le chargement, puis progression
      affichant successivement "40 232 lignes chargees", "68 972 lignes
      chargees" (chiffre final).
13. (~1:00) Barre de progression en bas de l'ecran : "Execution de la
    requete en arriere-plan" puis "Preparation de la feuille de calcul en
    cours..." puis "Ajustement de la largeur des colonnes".

### Phase 3 - Resultat : tableau "UPS" charge (1:04 - 1:24)

14. (1:04) Nouvel onglet feuille **"UPS"** actif, contenant un tableau
    Excel structure (ruban contextuel "Creation de tableau" actif, nom du
    tableau = **"UPS"**, options Ligne d'en-tete cochee, Bouton de filtre
    coche, Lignes a bandes coche). En-tetes de colonnes generiques :
    **Source.Name, Column1, Column2, ... Column17+** (pas de noms de
    colonnes metier -- Power Query n'a pas nomme les colonnes, elles restent
    generiques dans ce classeur temporaire "Classeur1"). Colonne
    supplementaire notee : **Column16** contenant parfois un pays en toutes
    lettres (`ALLEMAGNE`) ou un code numerique (`20250061`, `542`), et
    **Column17** contenant des codes alphanumeriques du type `NNYTJIXLM`,
    `POSTBOTTLING PC24 U...` -- colonnes visiblement heterogenes/optionnelles
    selon les fichiers sources (largeur de ligne CSV variable).
    Exemple de donnees (lignes 2-37, fichier `Invoice_202500468212_
    050525.csv` a `Invoice_202500485795_050925.csv`) :
    - Column1 = `02/01/2025` (date document, differente de la date dans le
      nom de fichier)
    - Column2 = Column3 = `0000A1912W` (identifiant compte)
    - Column4 = `FR`
    - Column5 = date facture format JJ/MM/AAAA (`05/05/2025`, `06/05/2025`,
      `07/05/2025`, `09/05/2025`...)
    - Column6 = numero de facture **affiche en notation scientifique**
      (`2,025E+11`) -- Excel a mal interprete le numero de facture comme un
      nombre -- **point d'attention : colonne a reformater/convertir**.
    - Column7 = code lettre : `I` observe majoritairement, `E` observe sur
      d'autres lignes/fichiers (Invoice_202540005969 notamment) -- semble
      distinguer un type de mouvement (Import/Export ?).
    - Column8 = parfois vide, parfois rempli (valeurs comme `6`)
    - Column9 = `5` ou `6` ou `17` (poids ? quantite ?)
    - Column10 = `FR75804949865` (identifiant fixe, type SIRET/TVA)
    - Column11 = `EUR`
    - Column12 = montant (`57.88`, `17.12`, `21.67`, `56.15`, `7.43`,
      `187.85`, `36.85`...)
    - Column13 = date (`04/04/2025`, `29/04/2025`, `30/04/2025`,
      `23/04/2025`...)
    - Column14 = numero de tracking colis, deux formats observes : court
      type `A1912WMGKYS` / `A1912WSH9NW` et long avec prefixe "1Z" type
      `1ZA1912WDK88614206` / `1ZA1912WDK88724025` (numero de tracking UPS
      complet).
    - Column15 = souvent vide.
    - Column16 = code/texte variable : `542`, `ALLEMAGNE`, `20250061`.
    - Column17 = code alphanumerique variable : `NNYTJIXLM`, `POSTBOTTLING
      PC24 U...`.
15. (~1:16) Defilement rapide vers le bas du tableau jusqu'a la derniere
    ligne (**ligne 68973**, coherent avec les 68 972 lignes chargees + 1
    ligne d'en-tete -- confirme par l'indicateur de statut "Nb (non vides) :
    68973" en bas de l'ecran apres selection de la colonne A entiere).
    Cellule active A68973 = "Invoice_202540005969_051225.csv" (dernier
    fichier source du dossier, nommage different -- prefixe "2254..." au
    lieu de "2025..." pour le numero de facture). Valeurs de la derniere
    ligne : Column1=02/01/2025, Column4=FR, Column5=12/05/2025,
    Column6=2,0254E+11, Column7=E (au lieu de I), Column9=17,
    Column10=FR75804949865, Column11=EUR, Column12=5.60, Column13=30/04/2025,
    Column14=1ZA1912WDK91681100.

### Phase 4 - Enregistrement du classeur sous le nom cible (1:24 - ~1:45)

16. (~1:20-1:24) Retour furtif a l'Explorateur Windows sur le dossier `UPS`
    (meme vue qu'a l'etape 2), puis fenetre Excel devient blanche/en cours
    de chargement, barre de titre affichant **"2025_04_Facture UPS - Excel"**
    -- note importante : le nom affiche est **"2025_04"** (avril) alors que
    les CSV sources sont dates de mai (2025 05) -- possible reouverture d'un
    classeur du mois precedent servant de modele.
17. (~1:28) Ecran "Enregistrer sous" (Fichier > Enregistrer sous) affichant
    la liste des emplacements recents : dossiers `2025 04`, `2025 05`, `UPS`
    (tous sous `D: > Drive > Comptabilité La Ruche > $Facturation
    automatique > 1 - Factures transp...`), `Comptabilité`, `Thomas`. Volet
    "La Ruche Logistique" avec comptes OneDrive/SharePoint
    (thomas@laruche-logistique.fr). Auteur du fichier = **Thomas Largeron**.
18. (~1:32) Clic sur "Ce PC" puis "Parcourir" -> boite de dialogue
    "Enregistrer sous" classique, positionnee sur
    `Disque 2 (D:) > Drive > Comptabilité La Ruche > $Facturation
    automatique > 1 - Factures transporteurs + calculs > 2025 > 2025 05`
    (meme dossier que l'etape 1, contenant les sous-dossiers Chronopost,
    DPD, Fedex, GLS, Kuehne, TNT, UPS, UPS - 80X7Y5 et les classeurs
    `2025_05_Facture Chronopost/DPD/Kuehne`). Champ **"Nom de fichier"**
    pre-rempli avec **"2025_04_Facture UPS"** en surbrillance (edite),
    **"Type"** = "Classeur Excel". L'utilisateur modifie manuellement le nom
    en **"2025_05_Facture UPS"** (remplace "04" par "05" pour correspondre
    au mois traite). Bouton **Enregistrer** actionne.

### Phase 5 - Classeur cible ouvert : structure des onglets et rappel de la table de comptes (1:40 - 2:00)

19. (~1:40) Le classeur est desormais enregistre et renomme
    **"2025_05_Facture UPS - Excel"**. Vue sur l'onglet **"Comptes UPS"**
    (une des feuilles internes du classeur cible, PAS la feuille "UPS"
    Power Query) : colonne A = liste de codes courts (`111R9F`, `11651R`,
    `4F261E`, `4V4555`, `4V47R4`, `535W21`, `6Y06A5`, `764V61`, `764V72`,
    `79A43Y`, `79A7T0`, `80X7Y5`, `9VV644`, `A1912W`, `J40E82` a `J40F00`,
    `W23909`, `W9765E`, `WV5788`) et colonne B = libelle transporteur associe
    : soit **"UPS"** soit **"UPS_COD"**. Table de correspondance
    compte -> transporteur (ex : `80X7Y5` -> `UPS_COD`, `A1912W` -> `UPS`,
    tous les codes `J40Ex...` -> `UPS_COD`). Confirme la logique memorisee
    "UPS vs UPS_COD = 2 transporteurs distincts, table compte->client".
20. (~1:44-1:48) Vue sur l'onglet **"Facture UPS"** du classeur cible avec
    la **liste complete des onglets visibles en bas** : **Gazole, CODES SVCE
    LEVEL, CODIFICATION CODE EXCEPTION, Charge.CHG_CODE, ST SV, zone colis
    poids assurance, Bilan factures, Facture UPS, Demande avoir, TCD,
    Comptes UPS, Fichier import, Bilan clients**. En-tetes de colonnes
    nommees (contrairement au tableau Power Query generique) : **Clients,
    Montant a facturer, Mode envoi, Categorie, Version, Numero de..., Numero
    ..., Pays d'or..., Date de la facture, Numero de facture, Code de...,
    Code de..., Numero..., Code de l..., Montant, Date de l'operation,
    Reference, Numero de l'envoi principal, Numero..., Numero de reference,
    Numero..., Code de l...**. Colonne **Categorie** contient des valeurs
    comme "Surcharge de securite", "TVA", "Fret", "Droits et taxes", "ST"
    (Frêt), "Taxe gazole", "plus-value BtoC", "SV" (Frêt), et -- vues plus
    loin en descendant dans le tableau (lignes ~92 a 128, factures du
    23/05-30/05/2025) -- egalement **"CODE INCONNU"** et **"#N/A"** (memes
    lignes en alternance, l'une servant probablement de calcul intermediaire
    invisible/masque de l'autre), puis a partir de la ligne ~129 (factures
    du 03/06/2025) la categorie devient **"Adresse"** au lieu de "Fret" pour
    les memes types de lignes "ST" -- **correspond exactement au repere
    memoire "colonne AX vide doit ressortir en Fret et non Adresse"** : on
    voit ici le symptome en action, ou l'absence d'une info fait basculer la
    categorie calculee de "Fret"/"ST" vers "Adresse". Egalement observees :
    des lignes avec categorie **"Assurance"** montant **500.00** (ligne
    ~12142, compte 0000A1912V, date facture 03/06/2025) et categorie
    **"Zones eloignees"** puis **"Assurance"** montant **600.00** (ligne
    ~24043) -- lignes speciales avec montant fixe eleve, a rapprocher du
    repere "ligne facture a 4,25€ en cas d'avarie prix colis livre" (valeurs
    differentes ici, mais meme mecanique de lignes a montant fixe insere
    dans le detail).
    Colonne **Montant** (O) = `162431.23` repete sur de nombreuses lignes
    dans les premieres lignes (donnees plus anciennes/avril), puis
    `165204.02` sur les lignes ~12140+ (juin), `15360.56` sur les lignes
    ~30140+ (Column M = "0000J40E82", compte different) -- ce sont des
    valeurs de montant de facture globale repetees sur toutes les lignes
    de detail d'une meme facture, pas des montants ligne a ligne. Colonne
    **Numero de facture** (F) = notation scientifique persistante
    (`2,025E+11`, `2,02501E+11`, etc.) -- **confirme le probleme signale en
    memoire sur toute la hauteur du fichier, pas seulement les premieres
    lignes**.
    Cet onglet "Facture UPS" contient un tres grand nombre de lignes
    (defilement jusqu'a la ligne **86255+** observe plus loin, cf. etape
    34).

### Phase 6 - Retour sur le tableau "UPS" (Power Query) et renommage des colonnes (2:00 - 2:20)

21. (~2:04-2:08) Retour sur l'onglet **"UPS"** (le tableau Power Query issu
    de l'import CSV, dans le meme classeur desormais nomme
    "2025_05_Facture UPS"). Les en-tetes de colonnes ont ete **renommees**
    en noms metier explicites (plus "Column1, Column2..."). Colonnes visibles
    de gauche a droite (defilement horizontal) :
    - **J (13e colonne visible)** : "Code de la..." = EUR
    - **K** : "Montant..." = 57.88 (identique a l'ancien Column12)
    - **L** : "Date de l'o..." = 04/04/2025
    - **M** : "Reference" = A1912WMGKYS (tracking court)
    - **N** : "Numero W..." = (vide, ou "542"/"ALLEMAGNE"/"20250061" selon
      la ligne)
    - **P (colonne Q apres)** : "Numero de reference 1 de l'envoi"
    - **Q** : "Numero de reference 2 de l'envoi" = "NNYTJIXLM",
      "POSTBOTTLING PC24 UD281896"
    - **R** : "Code de l'..." = P/P, F/C, F/D (codes courts repetitifs)
    - **S** : "Nombre d..." = 0 ou 1
    - **T** : "Nombre d..." = 0 (colonne toujours a 0 dans l'echantillon)
    - **U** : **"Numero de suivi"** = tracking long avec prefixe 1Z, ex.
      `1ZA1912W0465646038`, `1ZA1912WDK88614206`, `1ZA1912WDK88724025`,
      `1ZA1912WDK92668454`, `1ZA1912WDK95161790`, `1ZA1912WDK95965958`,
      `1ZA1912WDK98814927` -- **colonne cle differente de "Reference"**
      (tracking court), suggere deux identifiants de colis distincts geres
      en parallele.
22. (~2:12-2:16) Defilement horizontal supplementaire vers les colonnes
    **W a AQ** : en-tetes renommes visibles = **Numero de reference, Numero
    d..., Numero d..., Poids annonce, Unite de m..., Poids facture, Unite de
    m..., Type de co..., Type de po..., Dimensions du..., Zone, Code de
    ty..., Code detai..., Source de..., Code de ty..., Code deta..., Valeur
    det...**. Exemples de valeurs par ligne :
    - **Poids annonce** = 5.0 / 20.0 / 19.0 / 4.0 / 10.0 (unite = K, kg)
    - **Poids facture** = 4.5 / 18.5 / 16.5 / 3.5 / 10.0 / 7.5 (unite = K)
      -- **poids facture different du poids annonce** (ecart facturation
      probable, point business a surveiller).
    - **Type de co...** = "PKG" (package, constant)
    - **Type de po...** = 29, 8, 30 (code zone/service ?)
    - **Zone** = codes numeriques a 3 chiffres : `12`, `355`, `706`, `705`
    - **Code de ty...** (a cote de Zone) = `SHP` (Shipping) ou `RTN`
      (Return)
    - **Code detai...** = `IMP` (import) ou `RTS` (Return To Sender)
    - Colonne finale visible ("G" tronque) = lettre seule (`Fl`, `B`, `G`,
      `E`, `A`, `T`, `In`...) -- code court coupe a l'affichage, valeur
      exacte non lisible sur cette frame.
23. (~2:20) Nouveau defilement horizontal vers les colonnes **AL a BC** :
    en-tetes = **Code de ty..., Code deta..., Valeur det..., Code de ty...,
    Code deta..., Valeur det..., Code de cl..., Code de d..., Description
    des frais, Nombre d'..., Code de d..., Valeur de..., Indicateur...,
    Code de d..., Montant d..., Montant..., Code de d..., Montant...**.
    **Colonne "Description des frais"** (colonne AX, confirme plus loin en
    etape 32) tres riche en information metier, valeurs observees :
    - "WW Express Saver" (FRT/069)
    - "Other Govt. Fees" (BRK/337)
    - "Security Fee" (BRK/348)
    - "Frais sur avances" (BRK/405)
    - "Autres frais de douane" (BRK/495)
    - "Droits de douane" (GOV/201)
    - "Customs GST" (GOV/206)
    - "TB Standard Undeliverable Return" (FRT/011, INF/011)
    - "Taxe S/Carburant" (FSC/FSC) -- **taxe gazole/surcharge carburant**
    - "20.000 % Tax" (TAX/01) -- **taux de TVA a 20% explicitement affiche**
    - "Liv.particulier" (ACC/RES)
    - "WW Standard" (FRT/067)
    - "Frais d'entreposage" (BRK/473)
    - "Droits d'accises" (GOV/204)
    - "TVA douaniere sur marchandises" (GOV/205)
    - "Frais prep. dedouanement" (BRK/404)
    - "Frais de traitement TTB" (BRK/378)
    - "Frais de traitement AMS" (BRK/386)
    - "Frais ligne tarif. suppl." (BRK/410)
    - "FDA Clearance" (BRK/412)
    - "Frais magasinage" (BRK/419)
    - "BTA Prior Notice" (BRK/420)
    - "UPS Hdlg Control" (GOV/202)
    - "Dom. Standard" (FRT/003) -- observe sur des lignes de compte "0000WV578"
    - "Transport" (FRT, ligne sans code de detail visible, compte
      0000A1912V, colis "1ZA1912WDK91708706")
    - "Supplement donnée..." (NPF, meme groupe de lignes)
    - "**Frais de correction d'expédition**" (ACC/SCF) -- observe en tres
      grand nombre sur les lignes ~35800+ (compte 0000A1912V), montant
      quasi-systematiquement **1,5 €** (parfois 1,68 €) -- gros volume de
      petites lignes correctives.
    Colonnes de code associees : "Code de cl..." (colonne AV, ex: FRT, BRK,
    GOV, EXM, FSC, TAX, ACC, INF, COD, NPF) et "Code de d..." /"Code de..."
    (colonne AW, code numerique/alphanumerique, ex: 069, 337, 348, 405, 495,
    201, 206, 011, FSC, 01, RES, 067, 473, 204, 205, 404, 1461, 378, 386,
    410, 412, 419, 420, 202, 003, SCF).
24. (~2:20) Colonne **"Montant..."** (position **BE**, confirme etape 32 :
    en-tete exact = **"Montant net"**) affiche des valeurs en devise avec
    format "Comptabilite" apres une operation de **Rechercher-remplacer**
    (voir etape 25) : `0.00`, `0.06`, `5.83`, `9.01`, `2.39`, `4.41`,
    `36.18`, `12.58`, `1.69`, `2.86`, `11.80`, `2.11`, `4.15`, `3.61`,
    `16.04`, `18.11`, `6.52`, `15.48`, `7.43`, `15.92`, `97.25`, `74.68`,
    `16.05`, `6.27`, `14.53`, `22.98`, `21.39`, `14.63`, `18.01`, `7`,
    `7.31`, `33.31`, `12.95`, `8.33`.

### Phase 7 - Correction du format numerique (separateur decimal) sur la colonne "Montant net" (2:20 - 2:40)

25. (~2:24-2:28) Boite de dialogue **"Rechercher et remplacer"** ouverte
    (onglet "Remplacer" actif) sur la colonne **BE** ("Montant net",
    selectionnee/mise en surbrillance dans le tableau au moment de
    l'ouverture, cellule active BE1 = "Montant net"). Champs :
    - **Rechercher** : `.` (point)
    - **Remplacer par** : `,` (virgule)
    Boutons disponibles : Remplacer tout, Remplacer, Rechercher tout,
    Suivant, Fermer.
26. (~2:32) Clic sur **"Remplacer tout"**. Les valeurs de la colonne BE
    passent du format point-decimal americain (`0.06`, `5.83`, `9.01`...)
    au format virgule-decimale francais (`0,06`, `5,83`, `9,01`...) --
    **conversion texte -> nombre reconnu par Excel FR**. Les valeurs
    initialement `0.00` deviennent simplement `0` (le zero pur, sans
    decimales affichees, car reconnu desormais comme un nombre et non plus
    du texte).
27. (~2:36) Boite de dialogue "Rechercher et remplacer" toujours ouverte
    (fermeture imminente). Selection de la colonne BE entiere confirmee par
    la barre de statut en bas : **"Moyenne : 2,90 € ; Nb (non vides) : 68973
    ; Somme : 200 152,79 €"** -- confirme que la colonne contient
    desormais de vraies valeurs numeriques agregeables (somme totale des
    montants de cette colonne sur les 68 972 lignes = **200 152,79 €**).
28. (~2:36-2:40) Clic sur une cellule (valeur affichee `2,39` avec virgule)
    pour verifier le resultat -- la barre de formule confirme `2,39`
    (nombre, plus du texte).
29. (~2:40) Ouverture du **filtre automatique de la colonne "Montant net"**
    (clic sur le petit triangle d'en-tete), affichant le menu de tri/filtre
    standard Excel : "Trier du plus petit au plus grand", "Trier du plus
    grand au plus petit", "Trier par couleur", "Effacer le filtre de
    'Montant net'", "Filtres numeriques", champ de recherche, puis la
    **liste des valeurs uniques cochees** pour filtrage : `0`, `0,06`,
    `0,3`, `0,31`, `0,32`, `0,34`, `0,42`, `0,46`... (toutes cochees,
    "Selectionner tout" actif).

### Phase 8 - Verification apres filtre/tri sur la colonne "Montant net" (2:40 - 2:52)

30. (~2:44) Apres fermeture du menu de filtre, la colonne "Montant net" est
    reaffichee mais la **barre de statut indique desormais "Moyenne : 5,52
    € ; Nb (non vides) : 36255 ; Somme : 200 152,79 €"** -- la somme totale
    reste identique (200 152,79 €) mais le nombre de valeurs non vides
    comptees chute de 68973 a 36255, et la moyenne passe de 2,90 € a
    5,52 € -- suggere qu'un filtre a ete applique excluant les lignes a 0.
31. (~2:48) Selection de la cellule active correspondant a la ligne "BRK/405
    - Frais sur avances", valeur **16,05** confirmee en barre de formule.

### Phase 9 - Onglet "Facture UPS" : la formule de calcul de la colonne "Categorie" (colonne D) (2:52 - 4:00)

**Decouverte majeure** : en cliquant sur des cellules de la colonne
**D ("Categorie")** de l'onglet **"Facture UPS"**, la barre de formule
revele la logique complete de calcul de cette colonne, lue integralement
sur la cellule **D36254** :

```
=SI(NB.SI('ST SV'!Q:Q;AX36254)<>0;"Adresse";
   SI(NB.SI('ST SV'!D:D;AX36254)<>0;"plus-value BtoC";
      SI(AV36254="FRT";"Frêt";
         SI(AV36254="TAX";"TVA";
            SI(NB.SI(Charge.CHG_CODE!A:A;AW36254)=0;"code inconnu";
               RECHERCHE(AW36254;Charge.CHG_CODE!A:A; ... )
            )
         )
      )
   )
)
```

(formule tronquee a l'affichage au-dela de "RECHERCHE(AW36254;
Charge.CHG_CODE!A:A;" -- la suite exacte, tres probablement
`Charge.CHG_CODE!B:B` en 3e argument du RECHERCHE, n'est pas lisible sur la
frame capturee.)

Cette formule **confirme et precise exactement** le repere memoire projet :
- **Colonne AW** = "**Code de...**" (code de detail des frais, ex. `337`,
  `405`, `011`, `067`, `SCF`, `EVS`...) -- utilise en cle de recherche dans
  la table `Charge.CHG_CODE` (onglet dedie visible dans la liste des
  onglets). Si le code n'existe pas dans `Charge.CHG_CODE!A:A`, la formule
  renvoie le texte **"code inconnu"** -- confirme par la frame ~3:16
  (filtre applique sur la colonne Categorie de "Facture UPS", cf. etape
  32bis) montrant "91 enregistrement(s) trouve(s) sur 36254" pour la valeur
  "code inconnu" seule cochee. **Convertir la colonne AW en nombre retire
  ces #N/A/"code inconnu"** probablement car le RECHERCHE() en mode
  approximatif echoue si le type (texte vs nombre) du code AW ne correspond
  pas exactement au type des codes references dans `Charge.CHG_CODE!A:A`.
- **Colonne AX** = "**Description des frais**" texte libre (ex. "WW
  Express Saver", "Autres frais de douane", "Frais de correction
  d'expédition"...) -- utilisee en cle de recherche dans l'onglet
  **'ST SV'** (colonnes Q et D). Si `AX` correspond a une valeur presente
  dans `'ST SV'!Q:Q`, la categorie devient **"Adresse"**. Si `AX` est
  **vide** (ou ne correspond a aucune ligne de `'ST SV'!Q:Q` ni
  `'ST SV'!D:D`), la formule redescend dans les conditions suivantes : si
  `AV` (colonne "Code de classe" FRT/TAX/etc.) vaut "FRT", la categorie
  devient **"Frêt"**. **Ceci confirme precisement le repere memoire
  "colonne AX vide doit ressortir en Fret et non Adresse"** : si `AX` est
  vide, `NB.SI('ST SV'!Q:Q;"")` peut a tort correspondre a des cellules
  vides presentes dans la plage `'ST SV'!Q:Q` (si cette colonne contient
  elle-meme des cellules vides), ce qui ferait ressortir "Adresse" au lieu
  de "Frêt" -- **bug de la formule NB.SI avec critere vide**, exactement le
  symptome observe (categorie "Adresse" apparaissant a tort a partir des
  lignes de juin dans "Facture UPS", etape 20).
- **Colonne AV** = "Code de classe" (FRT, TAX, BRK, GOV, ACC, EXM, FSC,
  COD, NPF...), teste directement par egalite ("FRT" -> "Frêt", "TAX" ->
  "TVA").

32. (~2:56-3:08) Lecture des en-tetes de colonnes du meme onglet
    "Facture UPS" confirmant les positions exactes : colonne **AW** =
    "**Code de...**" (valeurs `337`, `348`, `405`, `386`, `SCF`...) ;
    colonne **AX** = "**Description des fra(is)...**" (valeurs "Other Govt.
    Fees", "Security Fee", "Frais de traitement", "Frais de correction
    d'expédition"...) ; colonne **BE** = "**Montant net**" (confirme
    l'en-tete complet observe en etape 25, valeurs `33,31`, `1,5`, `1,68`).
    Une icone d'alerte jaune (triangle "!") apparait sur une cellule de la
    colonne AV -- probablement une alerte de coherence Excel (nombre
    stocke en texte) sur cette ligne precise.
    Egalement confirme : colonnes **BT a CB** = bloc adresse destinataire
    -- "**Nom de l'...**", "**Adresse...**", "**Ville de l'...**", "
    **Code...**", "**Pay(s)...**", "**Nom du destinataire**" (en-tete
    complet lu sur CB1 : "Nom de l'entreprise du destinataire"). Exemples
    de valeurs : "JEAN-PIERRI[...] 114", "DOUZENS", "11700", "FR", "CITY
    M[...] 127" -- ce bloc de colonnes Nom/Adresse/Ville/Pays est
    vraisemblablement lie a la source de l'onglet **'ST SV'** utilisee
    dans la formule Categorie (colonnes Q/D) pour distinguer "Adresse" du
    reste des categories.
32bis. (~3:16-3:20) **Filtre applique sur la colonne D "Categorie"** de
    l'onglet "Facture UPS" : menu de filtre ouvert listant les valeurs
    uniques cochables **"code inconnu", "Colis volumineux", "Droits et
    taxes", "Frêt", "plus-value BtoC", "Surcharge de sécurité", "Taxe
    gazole", "TVA", "Zones éloignées"** (liste alphabetique complete des
    categories possibles dans cette colonne). Seule la valeur **"code
    inconnu"** reste cochee -> resultat filtre : **"91 enregistrement(s)
    trouve(s) sur 36254"** affiche en bas a gauche -- confirme un nombre
    limite mais non nul de lignes en erreur de mapping (91 lignes sur
    36254 lignes visibles a cet instant, probablement apres qu'un premier
    filtre soit deja actif sur une autre colonne, cf. 36255 note en etape
    30). Les lignes filtrees "code inconnu" partagent le motif : colonne
    AV = "ACC", colonne AW = "SCF", colonne AX = "Frais de correction
    d'expédition", montant net = **1,5** (ou 1,68 sur une ligne) --
    **confirme que le code "SCF" (Frais de correction d'expédition) n'est
    pas reference dans la table `Charge.CHG_CODE`**, ce qui declenche la
    categorie "code inconnu" pour toutes ces lignes.

### Phase 10 - Retour final sur "Facture UPS", confirmation du grand volume cumulatif (4:00 - 4:40)

33. (~4:00-4:04) Nouveaux defilements dans l'onglet "Facture UPS" : lignes
    ~36211-36255, compte "0000WV578", montant facture globale **2236.24 €**
    repete, categories toutes "Droits et taxes"/"Frêt"/"Taxe gazole"/"TVA"/
    "plus-value BtoC" alternees, mode de paiement "CHEQUE, ESPECE ET CB A..."
    visible en colonne T, dates d'operation mi-mai 2025.
34. (~4:08-4:12) Ligne **35327** : cellule active B35327 (Montant a
    facturer) avec formule visible **`=SI(AW35327="EVS";BA35327;0)`** --
    **autre formule cle utilisant la colonne AW** : si le "Code de..." (AW)
    vaut "EVS", alors le montant a facturer (colonne B) reprend la valeur
    de BA, sinon 0. Confirme que la colonne B "Montant a facturer" est
    **filtree/conditionnee par un code specifique ("EVS")** de la colonne
    AW, tandis que la colonne D "Categorie" (etape 31/32) utilise cette
    meme colonne AW pour le mapping de libelle via `Charge.CHG_CODE`.
    "EVS" n'a pas ete rencontre par ailleurs parmi les valeurs de AW listees
    (337, 348, 405, 495, 201, 206, 011, FSC, 01, RES, 067, 473, 204, 205,
    404, 1461, 378, 386, 410, 412, 419, 420, 202, 003, SCF) -- **a
    rechercher/confirmer le sens du code "EVS"** (peut-etre "Excess Value
    Service"/assurance valeur declaree chez UPS, ou "Enlevement").
35. (~4:16-4:36) Defilement continu dans l'onglet "Facture UPS", lignes
    ~86058-86255+, compte "0000WV578" toujours, montant facture **2236.24
    €**, tracking `1ZWV57887...`, mode paiement identique "CHEQUE, ESPECE
    ET CB A...", dates de mai/juin 2025. Confirme la tres grande taille du
    tableau "Facture UPS" (au moins 86 000+ lignes visibles en defilement,
    bien au-dela des 68 972 lignes du seul import "UPS" du mois de mai) --
    **fort indice que "Facture UPS" est un historique cumulatif multi-
    mois, alimente au fil des imports mensuels successifs plutot que
    recree/vide chaque mois**. La video se termine sur ce defilement final
    (fin de capture ~4:40-5:11, derniere partie potentiellement une
    lecture passive/fermeture sans nouvelle action distincte -- a verifier
    avec les frames de detection de changement de scene si des
    manipulations supplementaires ont lieu entre 4:40 et 5:11).

## Points ambigus / a confirmer

- **(etape 16)** Le classeur rouvert juste avant l'enregistrement affiche
  "2025_04_Facture UPS - Excel" dans la barre de titre alors que le mois
  traite dans les CSV est mai 2025 ("2025 05"). Confirme par la suite
  (etape 35) : l'onglet "Facture UPS" contient un historique tres
  volumineux couvrant plusieurs mois (avril/mai/juin releves dans les
  dates de facture), ce qui suggere que le classeur est **copie depuis le
  mois precedent puis renomme**, et que "Facture UPS" est un cumul jamais
  purge. A confirmer explicitement avec le pole transport : le nouvel
  import "UPS" (Power Query, feuille "UPS") est-il ensuite **ajoute a la
  suite** de l'historique "Facture UPS" (append, etape non visible dans
  cette video -- peut-etre traitee dans la video n°2 "Elaboration du
  fichier d'import"), ou l'historique visible ici est-il simplement un
  reliquat du mois precedent non encore purge au moment de la capture ?
- **(etape 31)** Fin exacte de la formule de la colonne "Categorie"
  (`RECHERCHE(AW36254;Charge.CHG_CODE!A:A; ...)`) tronquee a l'affichage --
  le 3e argument du RECHERCHE (vraisemblablement `Charge.CHG_CODE!B:B`)
  n'a pas pu etre lu avec certitude. A confirmer en ouvrant directement le
  classeur "Facture UPS" pour lire la formule complete dans la barre de
  formule elargie.
- **(etape 34)** Le sens exact du code **"EVS"** teste dans la formule
  `=SI(AW35327="EVS";BA35327;0)` de la colonne "Montant a facturer" n'a pas
  ete identifie parmi les valeurs de AW observees dans la video. A
  confirmer avec le pole transport.
- **(etape 27-30)** Ecart entre "Nb (non vides) : 68973" (etape 27, avant
  filtre) et "Nb (non vides) : 36255" (etape 30, apres ouverture/fermeture
  du menu de filtre colonne "Montant net") alors que la Somme reste
  identique (200 152,79 €) -- le mecanisme exact du filtre applique
  (differents de 0 ? ou juste une selection partielle de cellules ?) n'est
  pas explicitement visible a l'ecran.
- **Ligne facture a 4,25€ (avarie)** : pas observee explicitement (valeur
  exacte 4,25€ non rencontree dans les frames lues). En revanche, des
  lignes a montant fixe eleve et categorie speciale ont ete reperees :
  **"Assurance" = 500,00 €** (ligne ~12142) et **"Zones eloignees"/
  "Assurance" = 600,00 €** (ligne ~24043) dans l'onglet "Facture UPS" --
  meme famille de lignes "hors normes" a montant fixe, mais valeurs
  differentes de 4,25€. Possible que la ligne a 4,25€ soit traitee dans
  la video n°2 ou n°2-1 (avaries), pas dans cette premiere video dediee a
  l'integration des CSV bruts.
- Deux gros fichiers CSV (`Invoice_202500565583_060325` 40 013 Ko et
  `Invoice_202500569188_060325` 6 704 Ko) sont nettement plus lourds que les
  autres factures individuelles -- probablement des exports consolides
  (toutes commandes du mois) plutot que des factures unitaires. A confirmer
  s'ils sont bien traites de la meme maniere par le Power Query "Combiner et
  charger" (structure identique en colonnes) ou s'ils necessitent un
  traitement a part.
- Fin de la video (~4:40 a 5:11, soit les 30 dernieres secondes) : couverte
  uniquement par du defilement continu dans "Facture UPS" (etape 35) sans
  action metier distincte identifiee -- aucune scene-change significative
  detectee dans cette plage lors de l'extraction complementaire (voir
  clusters scene-change concentres entre 80s et 301s, deja couverts par les
  etapes ci-dessus). Si une action existe entre 4:40 et 5:11, elle n'a pas
  ete captee distinctement par l'echantillonnage 4s -- a signaler comme
  zone de moindre certitude.

(Transcription terminee sur la base des frames a intervalle fixe (0:00 a
4:40) recoupees avec les frames de detection de changement de scene
concentrees entre 1:20 et 5:00.)
