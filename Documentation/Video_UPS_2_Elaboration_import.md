# Video UPS 2 - Elaboration du fichier d'import

Source : `Transporteurs/UPS/Process Facturation - Facture UPS - 2 - Elaboration du fichier d'import.mp4`
Duree : 393.5 secondes (~6 min 33 s), capture d'ecran silencieuse (Excel/CSV/navigateur).
Methode : extraction 1 frame / 4s (98 frames fixes) + detection de changement de scene (29 frames), soit
127 frames lues en ordre chronologique.

Statut : TRANSCRIPTION COMPLETE (video verifiee de 0:00 a 6:33/fin) — verification finale du
segment 4:05-6:33 effectuee a densite 1 frame/3s pour confirmer qu'aucune etape n'a ete omise.

## Contexte

Cette video documente la 2e etape du process de facturation UPS : l'elaboration du fichier d'import ERP
final (23 colonnes standard) a partir des donnees deja integrees dans le classeur UPS (etape 1, integration
des CSV).

Le classeur ouvert au debut de la video s'appelle **"2025_05_Facture UPS"** (mois de mai 2025, capture
anterieure au fichier `2026_06_Facture UPS.xlsx` present dans le dossier). Ses onglets, dans l'ordre visible
en bas d'ecran : `Gazole`, `CODES SVCE LEVEL`, `CODIFICATION CODE EXCEPTION`, `Charge.CHG_CODE`, `ST SV`,
`zone colis poids assurance`, `Bilan factures`, `Facture UPS`, `Demande avoir`, `TCD`, `Comptes UPS`,
`Fichier import`, `Bilan clients`.

**En-tetes EXACTES des 23 colonnes du fichier d'import**, confirmees via le fichier CSV final
`2025_04_UPS_Import` (colonnes A a W, cf. etape 12) :
`Transporteur` (A) / `Date validite` (B) / `Ref.1` (C) / `Ref.2` (D) / `Id client` (E) / `N° Tracking` (F) /
`Nom` (G) / `E / P` (H) / `Pays` (I) / `Zone` (J) / `Nbr Colis` (K) / `Poids` (L) / `mode envoi` (M) /
`TVA` (N) / `Droits et taxes` (O) / `Assurance` (P) / `Zones eloignees` (Q) / `Colis volumineux` (R) /
`Adresses` (S) / `Fret` (T) / `plus-value BtoC` (U) / `Gazole` (V) / `Nb Colis` (W).
NB : cet ordre differe legerement de celui de l'onglet `Fichier import` du classeur Excel source (qui
commence par `E/P ERP`, `Zone`, `Transport`... — voir etape 5) : le CSV final semble reordonne/renomme par
rapport a la mise en page interne de calcul de l'onglet Excel.

## Transcription chronologique

### Phase 1 — Onglet "Facture UPS" : etat des donnees sources (0:00 - 0:20)

1. (0:04-0:08) Le classeur `2025_05_Facture UPS.xlsx` est ouvert sur l'onglet **Facture UPS**. On voit un
   grand tableau avec les colonnes H a AAAE environ : `Pays d'origine`, `Date de la facture`, `Numero de
   facture`, `Code de ...`, `Code de ...`, `Numero`, `Code de ...`, `Montant`, `Date de l'operation`,
   `Reference`, `Numero de l'envoi principal`, `Numero`, `Numero de reference`, `Numero`, `Code de ...`,
   `Nombre de colis`, `Nombre ...`, `Numero de suivi`, puis colonnes `Type`, `Zone...`. Les lignes montrent
   des trackings UPS `1ZA1912...`, avec montants ~165204,02 EUR, dates de facture 03/06/2025 ou 05/05/2025.
   Cette feuille semble etre l'extraction brute post-integration CSV (etape 1 de la video precedente).

2. (0:10) Changement de scene : la barre de formule montre la cellule active contenant un numero de suivi
   `1ZA1912W6856349024` — pas de formule, valeur brute collee.

### Phase 2 — Colonne Zone/Assurance : calcul sur l'onglet "Facture UPS" (0:12 - 0:20)

3. (0:12-0:20, cellules R15/R10303) Formule tapee/visible dans la barre de formule :
   ```
   =SI(RECHERCHE(H10303;'zone colis poids assurance'!D:D;'zone colis poids assurance'!G:G)=0;"";
       MAX(10;ARRONDI.SUP(0,02*RECHERCHE(H10303;'zone colis poids assurance'!D:D;'zone colis poids assurance'!G:G);2)))
   ```
   Interpretation : recherche du n° de suivi (H) dans la colonne D de l'onglet `zone colis poids assurance`,
   renvoie la valeur trouvee en colonne G de ce meme onglet, si le resultat est 0 on laisse une cellule
   vide, sinon on prend le MAX entre 10 (plancher) et l'arrondi superieur a 2 decimales de 2% de la valeur
   trouvee. Le calcul en pourcentage (2%) avec plancher (10 centimes ?) est typique d'une **prime
   d'assurance transport** proportionnelle a la valeur declaree — cette formule alimente tres probablement
   la colonne **Assurance** du fichier import (colonne P), pas la colonne Zone (voir Points ambigus).
4. On observe un dialogue Excel "Attention — Microsoft Excel a trouve des donnees pres de votre
   selection. Comme elles ne sont pas selectionnees, elles ne seront pas triees." avec options "Etendre la
   selection" / "Continuer avec la selection en cours" (visible deux fois de suite, frames ~0:16-0:20) — la
   personne trie une colonne sur l'onglet Facture UPS.

### Phase 3 — Onglet "Fichier import" (interne au classeur Excel) : structure de calcul (0:32 - 0:52)

5. (0:32-0:36) Bascule sur l'onglet **Fichier import** (feuille de calcul interne, differente du CSV final
   du meme nom). En-tetes de colonnes visibles en ligne 1 (a partir de la colonne A) : `E/P ERP`, `Zone`,
   `Transport`, `Date validite`, `Ref.1`, `Ref.2`, `Id cli...`, `N° Tracking`, `N...`, `E`, `Pays`, `Zone`,
   `Nbr Co...`, `Poi...`, `mode en...`, `TVA`, `Droits et ta...`, `Assura...`, `Zones eloign...`,
   `Colis volumin...`, `Adress...`, `Fret`, `plus-value B...`. Une note en rouge tout a droite indique :
   **"Si nb colis = 0 -> voir montant plus-value BtoC"**.
   Les lignes visibles (0276 a 0314 environ) montrent des donnees deja remplies : colonne A = `particulier`
   / `entreprise` (E/P), B = `Zone` (valeurs 1 a 4, un `33` isole en ligne 0289 — POINT AMBIGU / A
   VERIFIER), C = `Transport` = `UPS_COD` pour toutes les lignes visibles, D = `Date validite` = `01/05/2025`
   uniforme, E = `Ref.1` = `CHEQUE, ESPECE` pour tout le lot visible, F = `Ref.2` = un numero
   (ex. `266990752`), H = `N° Tracking` (`1ZWV5788...`), colonne `E`/`P` secondaire (I) = `E` ou `P`,
   `Pays` = `FR`, `Zone` (colonne L, texte) = `France`, `Nbr Colis` = `1`, `Poids` = `1 ST`/`1,5 ST`/`2 ST`/
   `0,5 ST` (unite "ST" affichee a cote du chiffre = mode de transport Standard, par opposition a "SV" vu
   plus loin = Saver ?), `TVA` = `0,2` pour toutes les lignes visibles (= 20%), colonne `Fret` = montants en
   euros (`10,82 €`, `6,05 €`, `6,06 €`, `4,75 €`, `4,76 €`, `8,64 €`...).
   Les colonnes `Droits et taxes`, `Assurance`, `Zones eloignees`, `Colis volumineux`, `Adresses`,
   `plus-value BtoC` apparaissent vides pour ce lot de lignes.

6. (0:36) Un menu de couleur de remplissage est ouvert (palette "Couleurs du theme" / "Couleurs standard" /
   "Couleurs utilisees recemment") sur la cellule d'en-tete "Numero de suivi" — mise en forme, pas une
   formule.

7. (0:40-0:44) Cellule active `H10314` sur l'onglet **Fichier import**, barre de formule affiche
   `=TCD!E10315` — la colonne **N° Tracking** du fichier import est alimentee directement par une reference
   croisee vers l'onglet **TCD** (tableau croise dynamique), decalee d'une ligne (E10315 pour la ligne
   10314 du fichier import — decalage du a la ligne d'en-tete du TCD).

8. (0:44) Bascule sur l'onglet **TCD**. Tableau croise dynamique avec en colonnes : `Logistique` / `Cout` /
   `Code Client` / `Poids` / `Numero de suivi` / `Categorie` / `Adresse` / `Assurance` / `Colis volumineux` /
   `Droits et taxes` / `Fret` / `plus-value BtoC` / `Surcharge de securite` / `Taxe gazole` / `TVA` /
   `Zones eloignees` / `(vide)` / puis une zone de controle a droite avec titre en rouge **"CONTROLER
   ASSIETTE TAXE GAZOLE"** et les colonnes `% gazole`, `date d'enlevement`, `mode transport`,
   `gazole theo...`, `ecarts`, `Gazole vendu`, `retours`.
   Valeurs observees : `% gazole` ~7,76% a ~13,19% selon les lignes, `mode transport` = `ST` pour toutes les
   lignes visibles, `Gazole vendu` = `19,75%` ou `20,00%` ou `13,23%` selon la ligne, `ecarts` = valeurs
   negatives (`-5,48%`, `-3,34%`, `-4,54%`, `-0,04%`...). Controle de coherence entre le taux de gazole
   applique par UPS sur la facture et un taux de reference/theorique attendu, avec calcul de l'ecart.
   En bas du TCD, lignes speciales : `A1912WSTGYP` et `A1912WTW9TC` avec poids = `0 kg` et des montants
   isoles en colonne "Surcharge de securite" (`32,76 €`, `92,40 €`) sans les autres colonnes remplies, et
   une ligne `AB` avec `#N/A` partout — probablement des frais annexes (surcharge carburant/redevance) sans
   tracking colis standard associe (a confirmer).

9. (0:44) Retour sur **Fichier import**, cellule `H10314` toujours `=TCD!E10315`.

10. (0:48) Zoom out sur l'onglet **Fichier import** : dernieres lignes du tableau (jusqu'a ~0327) avec un
    bloc de lignes en surbrillance grise (selection) couvrant les lignes 0315 a 0327 environ. Ces lignes
    montrent des valeurs `#N/A` en cascade dans plusieurs colonnes (Ref.2, N° Tracking = vide, Zone =
    `inconnu`, Poids = `0`/`inconnu`) — lignes en fin de tableau correspondant a des enregistrements
    incomplets/non apparies issus du TCD (jointure ratee ou lignes de totaux residuelles). Compteur Excel en
    bas : **"Nb (non vides) : 315"**.

11. (0:48-0:52) Cellule `Q10313` (colonne "Droits et taxes" du Fichier import), barre de formule :
    ```
    =SI(TCD!I10314=0;"";TCD!I10314)
    ```
    Confirme le pattern general : chaque colonne numerique du Fichier import (onglet Excel) est alimentee
    par une formule de renvoi vers la colonne correspondante du TCD, avec un test "si 0 alors vide" pour ne
    pas afficher de zeros parasites.

### Phase 4 — Exploration de l'arborescence reseau et du fichier CSV d'import du mois precedent (0:56 - 1:20)

12. (0:56-1:04) Bascule vers l'**explorateur de fichiers Windows**. Navigation dans l'arborescence reseau :
    `Ce PC > Disque 2 (D:) > Drive > Comptabilite La Ruche > $Facturation automatique > 2 - Fichiers csv
    import > 2025 > 2025 05`. Contenu du dossier `2025 05` : `2025_05_DPD_Import` (217 Ko, statut OK vert),
    `2025_05_Kuehne_Import` (29 Ko, OK vert), `A IMPORTER 2025_05_UPS_Import_2` (1 Ko, statut "cloud" bleu =
    pas encore synchronise/traite). Le prefixe **"A IMPORTER"** dans le nom de fichier signale un fichier
    d'import en attente de traitement/depot dans l'ERP. C'est ce fichier `A IMPORTER 2025_05_UPS_Import_2`
    (mois courant, en cours de construction) qui sera complete plus loin dans la video (Phase 5).
    Le panneau lateral gauche liste aussi les autres dossiers du meme niveau : `1 - Factures transporteurs +
    calculs`, `3 - Fichiers csv export + calculs`, `4 - Import facturation`, `5 - Facturation ecommerce`,
    `10 - Grilles tarifaires`, `11 - Grilles zoning`, `12 - Grilles zones eloignees`, `Contrats`, `Process`
    (avec sous-dossiers `1 - Factures transporteurs Excel`, `2 - ERP - [TRANSPORT] - Imports`,
    `3 - ERP - [TRANSPORT] - Prix de vente`).

13. (1:04-1:08) Navigation vers le dossier `2025 04` (mois precedent) du meme chemin `2 - Fichiers csv
    import`. Contenu : `2025_04_GLS_Import`, `2025_04_Mondial Relay_import_2`, `2025_04_TNT_Import`,
    `2025_04_Fedex_Import`, `2025_04_UPS_Import_abeille` (2 Ko, statut cloud), `2025_04_Kuehne_Import_2`,
    `2025_04_Lettre_Timbre_Allemagne_SLAA...`, `2025_04_Lettre suivie prepa_Import fichier`,
    `2025_04_Lettre suivie_Import fichier`, `2025_04_Delivengo_LPPAQ_Import`, `2025_04_Colissimo_Import`,
    `2025_04_Mondial Relay_import`, `2025_04_DHL_Import`, `2025_04_DPD_Import`,
    `2025_04_Chronopost_Import_2`, `2025_04_Chronopost_Import`, `2025_04_Kuehne_Import`, et surtout
    **`2025_04_UPS_Import`** (884 Ko, statut OK vert) — c'est ce fichier (mois precedent complet) qui est
    ouvert ensuite, manifestement pour servir de reference de comparaison (regle memoire projet : "comparer
    au fichier import du mois precedent" pour les zones ambigues).

14. (1:08-1:20) Ouverture du fichier **`2025_04_UPS_Import.csv`** dans Excel (titre de fenetre
    "2025_04_UPS_Import - Excel"). **En-tetes de colonnes en ligne 1, lisibles integralement** :
    `Transporteur` (A) / `Date validite` (B) / `Ref.1` (C) / `Ref.2` (D) / `Id client` (E) /
    `N° Tracking` (F) / `Nom` (G) / `E / P` (H) / `Pays` (I) / `Zone` (J) / `Nbr Colis` (K) / `Poids` (L) /
    `mode envoi` (M) / `TVA` (N) / `Droits et taxes` (O) / `Assurance` (P) / `Zones eloignees` (Q) /
    `Colis volumin[eux]` (R) / `Adresses` (S) / `Fret` (T) / `plus-value B[toC]` (U) / `Gazole` (V) /
    `Nb Colis` (W).
    Donnees de la premiere ligne : Transporteur = `UPS`, Date validite = `01/04/2025`, Ref.1 = vide,
    Ref.2 = `32432`, Id client = `EXP20250325-2102627`, N° Tracking = `1ZA1912W0468980706`, E/P = `E`,
    Pays = `CA` (Canada), Zone = `10`, Nbr Colis = `1`, Poids = `1 ST`, TVA = `0`, Droits et taxes =
    `67,50 €`. Autres lignes notables : Pays = `LU`/`DE`/`FR`/`PT`, Zone = `5`/`10`/`356`/`6`/`11`/`52`/
    `703`/`9` selon la destination, mode envoi = `ST` ou `SV`, TVA = `0` (hors France/UE) ou `0,2` (France).
    Colonne `Gazole` (U, `plus-value B...` avant elle) = valeurs `1,41` ou `4,15` selon les lignes.
    Colonne `Assurance` (P) : quelques valeurs isolees, ex. `15` (ligne Zone GB=703).
    On remarque **Zone = `703`** a plusieurs reprises pour Pays = `GB` (Royaume-Uni) — code de zone distinct
    special post-Brexit ou DOM/TOM (a confirmer), et **Zone = `356`** pour Pays = `PT` (Portugal) sur une
    ligne "Pickup Portugal" avec Poids = `15 ST` (colis volumineux).

15. (~1:20) Scroll dans le fichier `2025_04_UPS_Import` jusque vers la ligne ~919-963 : lignes toutes
    `UPS_COD` / `CHEQUE, ES[PECE]` / Pays `FR` / Zone `France` / mode envoi `ST` / TVA `0,2` — un menu
    contextuel clic-droit est ouvert (Couper/Copier/Options de collage/Coller special/Inserer/Supprimer/
    Effacer le contenu/Format de cellule/Hauteur de ligne/Masquer/Afficher), suggerant une manipulation de
    lignes (insertion ou suppression) sur ce fichier de reference — action precise non determinable sur
    cette frame seule.

### Phase 5 — Onglet "Fichier import" (Excel source) : notes de regles metier en rouge (~1:16 - 1:24)

16. Retour sur le classeur **`2025_05_Facture UPS`**, onglet **Fichier import**, vue de la partie superieure
    du tableau (lignes 2 a 44). Colonnes visibles : `E/P ERP`(A), `Zone`(B), `Transport`(C),
    `Date validite`(D), `Ref.1`(E), `Ref.2`(F), `Id cli[ent]`(G), `N° Tracking`(H), `N[om]`(I), `E`(J,
    Pays ?), `Pays`, `Zone`(J), `Nbr Co[lis]`(K), `Poi[ds]`(L), `mode en[voi]`(M), `TVA`(N),
    `Droits et ta[xes]`(O), `Assur[ance]`(P), `Zones eloign[ees]`(Q), `Colis volumin[eux]`(R), `Adress[es]`
    (S), `Fret`(T), `plus-value B[toC]`(U).
    **Trois notes en rouge/jaune, tres importantes, lisibles integralement dans le coin superieur droit du
    tableau (colonnes AAAE et suivantes)** :
    - *"Si nb colis = 0 -> voir montant plus-value BtoC p[our determiner Nb Colis reel]"* (ligne 1, deja
      notee en Phase 3).
    - *"Si 1,38 € -> 1 colis, sinon arrondi 20 kg / colis"* — regle de determination du nombre de colis a
      partir du poids/montant quand l'info n'est pas fiable directement : si un montant/poids specifique de
      1,38 € correspond a exactement 1 colis, sinon on arrondit le poids total par tranches de 20 kg pour
      en deduire le nombre de colis.
    - *"Comparer avec exports expedies depuis ERP"* — confirme la logique de recoupement avec un export ERP
      externe pour fiabiliser certaines valeurs (zone, E/P, nb colis).
    - *"ATTENTION modes ST / SV"* — mise en garde explicite sur la distinction entre les deux modes d'envoi
      `ST` (Standard) et `SV` (Saver/Express ?) rencontres dans les donnees, a ne pas confondre lors du
      calcul tarifaire/zoning.
    - *"ATTENTION zones 505, 506, etc"* — mise en garde sur des codes de zone specifiques (505, 506...)
      necessitant un traitement particulier (zones eloignees / DOM-TOM probablement, a rapprocher des zones
      703/356 vues en Phase 4).
    - **Encadre jaune en gras, tres visible : "E / P : facturer en P si Plus-value Paperless -..."** (texte
      coupe a droite, fin de phrase non visible sur cette frame) — regle centrale du calcul du mode E/P
      (Entreprise/Particulier) : si une "plus-value Paperless" est presente/appliquee sur l'envoi, alors le
      mode de facturation doit etre force a **Particulier (P)**, quelle que soit la valeur E/P d'origine
      (confirme et precise la note memoire projet "E/P reconstruit... Point relais -> entreprise").
    Donnees de lignes observees dans ce meme tableau (echantillon) : ligne 2 = E/P `inconnu`, Zone `12`,
    Transport `UPS`, Pays `NZ` (Nouvelle-Zelande), Zone(J) `12`, Poids `3`, mode envoi `#N/A`, TVA `0`,
    Fret `61,65 €`. Ligne 9 = E/P `entreprise`, Zone `4`, Pays vide, N° Tracking `1ZA1912W6895597648`,
    Poids `1`/`10 ST`, TVA `0,2`, Fret `13,38 €`, plus-value BtoC `1,41`. Ligne 36 = E/P `inconnu`, Zone
    `703`, Pays `GB`, Poids `0`/`0 ST`, mode envoi `#N/A`, Fret `48,69 €` — confirme que Zone `703`
    correspond bien au Royaume-Uni meme quand Pays n'est pas renseigne dans cette colonne locale.

### Phase 6 — Construction effective du fichier CSV d'import du mois (copier-coller Excel -> CSV) (~1:24 - 1:36)

17. (~1:24-1:28) Sur l'onglet **Fichier import** du classeur `2025_05_Facture UPS`, selection integrale du
    bloc de donnees (lignes 0276 a 0314, colonnes A a W) — bordure de selection en pointilles verts visible
    (copie), message en bas d'ecran Excel : **"Nb (non vides) : 195968"** (nombre eleve = somme sur toute la
    selection, pas un nombre de lignes).

18. (~1:28-1:32) Bascule vers le fichier CSV **`2025_05_UPS_Import`** (le fichier "A IMPORTER" vu en Phase
    4, etape 12) et **collage** du bloc copie en `A1`. Un menu contextuel de collage special est ouvert
    (icones "Options de collage"). **Juste apres le collage, la feuille affiche temporairement des erreurs
    en cascade `#REF!`** dans plusieurs colonnes (E/P, Nom, Zone...) avec des valeurs constantes visibles
    par ailleurs (`10`, `38`, `9,5`, `inconnu`, `0`) — signe que le premier collage a ete fait en formules
    (liees aux onglets source TCD/zone colis poids assurance du classeur d'origine, invalides hors
    contexte) et non en valeurs. Bordeaux/orange toujours present sur les colonnes N° Tracking, Zone,
    TVA (mise en forme conditionnelle heritee).

19. (~1:32-1:36) Correction : nouveau collage, cette fois en **valeurs seules** — la feuille `2025_05_UPS_
    Import` affiche desormais des donnees propres et coherentes (colonnes A a W remplies, Transport =
    `UPS_COD`, Pays = `FR`, Zone = `France`, N° Tracking `1ZWV5788...`, TVA `0,2`, Fret en euros). Derniere
    ligne visible : `0314 inconnu 01/05/2025 FACTURE 16 ... 9100021938...` (une ligne "FACTURE" isolee sans
    tracking standard, E/P = `inconnu`). Barre d'etat Excel en bas : **"CSV Enregistrement en cours
    2025_05_UPS_Import. Appuyez sur Echap pour annuler."** — confirmation explicite que le fichier est
    sauvegarde au format CSV a ce moment precis (Ctrl+S sur un fichier .csv declenche cette barre de
    progression dans Excel).

### Phase 7 — Corrections manuelles et controle qualite sur le fichier "2025_05_UPS_Import" (~1:40 - 1:56)

20. (~1:40-1:44) Sur le fichier CSV `2025_05_UPS_Import` (mois en cours), la colonne **mode envoi** (M)
    contient des `#N/A` sur plusieurs lignes en tete de tableau (lignes 2 a 4 : Pays `NZ`/`NZ`/`MQ`, Zone
    `12`). La personne **saisit manuellement `SV`** dans ces cellules (`M3` = `SV` tape au clavier, collage
    visible via l'icone flottante "(Ctrl)" caracteristique d'un Ctrl+V) pour corriger ces valeurs manquantes
    ligne par ligne.

21. (~1:44-1:48) **Filtre applique sur la colonne "Droits et taxes" (colonne O)** : liste de valeurs
    cochables affichee = `26,66 €`, `48,69 €`, `56,15 €`, `57,88 €`, `60,41 €`, `86,62 €`, `187,85 €`,
    `1 062,79 €`, `(Vides)` (seule "Vides" est cochee). Objectif : isoler les lignes qui ont un montant de
    "Droits et taxes" renseigne (frais de douane/dedouanement) pour verification ou traitement particulier
    distinct des lignes standard. Apres filtre, barre d'etat Excel : **"99 enregistrement(s) trouve(s) sur
    10313"** puis **"110 enregistrement(s) trouve(s) sur 10313"** — total du fichier `2025_05_UPS_Import`
    = **10 313 lignes**.

22. (~1:48-1:56) Sur les lignes filtrees (99-110 lignes avec Droits et taxes non vides), on observe des
    N° Tracking `1ZA1912WDK9...` (prefixe distinct de `1ZA1912WD99...` et `1ZWV5788...` vus precedemment),
    Poids = `0` presque systematiquement, Zone = `0` ou vide, mode envoi = `ST`, TVA = `0,2`. Une colonne
    "Nb (non vides)" / "Somme" en bas d'ecran varie selon la selection (`Moyenne : 9155,64 / Nb (non vides) :
    76 / Somme : 183112,8` puis `Moyenne : 7632,133333 / Nb (non vides) : 19 / Somme : 45792,8`) — verification
    manuelle de sous-totaux sur la selection filtree, probablement pour controler la coherence du total
    "Droits et taxes" par rapport a un montant de reference (facture PDF UPS ?). Une derniere capture montre
    un **filtre reinitialise a 0 resultat** ("Nb (non vides) : 0 / Somme : 0") suite a une manipulation sur
    la colonne A (nom "UPS" tape dans la zone de nom, pas une formule) — probablement fin de la verification
    et retour a la vue complete.

### Phase 8 — Filtres complementaires et confirmation des zones speciales GB (~1:56 - 2:16)

23. (~1:56-2:00) Nouveau filtre sur la colonne **Droits et taxes** : liste de valeurs affichee = `7,00 €`,
    `7,06 €`, `7,31 €`, `7,43 €`, `22,25 €`, `26,66 €`, `30,50 €`, `44,39 €`... — barre d'etat : **"10184
    enregistrement(s) trouve(s) sur 10225"**, confirmant un total proche de 10225-10313 lignes selon l'etat
    du filtre a cet instant (les totaux varient legerement au fil de la video, signe de filtres empiles/
    retires successivement plutot que d'un total fige).

24. (~2:00-2:04) Filtre sur la colonne **Fret** (T), cette fois par **valeurs numeriques croissantes**
    (menu "Trier du plus petit au plus grand" / "Trier du plus grand au plus petit" / "Filtres numeriques"),
    liste de montants `56,56 €` a `60,29 €` proposee — tri/filtre sur le montant du fret pour reperer des
    valeurs extremes ou verifier une plage.

25. (~2:04-2:08) Resultat du filtre combine (Droits et taxes non vide) : **22 enregistrements trouves sur
    10225**. Lignes observees avec **Zone = `705`** et **Zone = `706`** pour Pays = `GB`, en plus du `703`
    deja vu — confirme l'existence d'au moins **trois codes de zone distincts pour le Royaume-Uni (703, 705,
    706)**, tous avec Droits et taxes renseignes (`187,85 €`, `26,66 €`, `86,62 €`, `56,15 €`, `7,43 €`,
    `60,41 €`...) — coherent avec des frais de douane/dedouanement post-Brexit variables selon la nature de
    l'envoi. Une ligne Zone `9` / Pays `US` montre Droits et taxes = `1 062,79 €` (le plus eleve observe).

26. (~2:08-2:16) Retour a la vue non filtree (99 lignes -> vue complete), verification de la cellule `J6`
    (colonne Zone) qui contient la valeur texte **"France"** (barre de formule confirme : pas de formule,
    texte brut "France" tape/colle directement dans la colonne Zone pour les envois nationaux — coherent
    avec la regle memoire projet "zone=0 + pas de pays -> mettre France").

### Phase 9 — Verification de la colonne "plus-value BtoC" (regle Nb Colis = 0) (~2:16 - 2:32)

27. (~2:16-2:20) Cellule `U1` confirmee par la barre de formule : en-tete exact = **"plus-value BtoC"**
    (texte, pas de formule) — colonne U du fichier Excel/CSV. Confirme le nom exact de cette colonne parmi
    les 23 colonnes standard.

28. (~2:20-2:24) Filtre applique sur la colonne **Zone** (J) : liste de valeurs cochables `0`, `3`, `4`, `5`,
    `6`, `7`, `8`, `9`... (toutes cochees par defaut, "Selectionner tout"). Puis filtre sur la colonne
    **plus-value BtoC** (U) : valeurs `1,33`, `1,6`, `4,15`, `4,97`, `4,98`, `(Vides)`. Barre d'etat : "1012
    enregistrement(s) trouve(s) sur 10225" (etape intermediaire).

29. (~2:28-2:32) Filtre final combine (Zone = 0 ou vide, plus-value BtoC = `1,33` ou `1,6`) : **4
    enregistrements trouves sur 10225**. Lignes obtenues : N° Tracking `1ZA1912WDK90537161` (Zone 3, Poids
    10, plus-value BtoC 1,6), `1ZA1912WDK91445633` (Zone 4, Poids 1, plus-value BtoC 1,6),
    `1ZA1912WDK91708706` (Zone 0, Poids 0, Fret 18,93 €, plus-value BtoC 1,33), `1ZA1912WDK98316288` (Zone 3,
    Poids 25, plus-value BtoC 1,6). Ce filtrage croise confirme concretement la regle notee en Phase 5 :
    **quand Nbr Colis (K) = 0, on utilise le montant de la colonne "plus-value BtoC" pour determiner/valider
    le nombre de colis reel** — les valeurs `1,33`/`1,6`/`4,15`/`4,97`/`4,98` semblent etre des montants
    forfaitaires par colis associes a certains types d'envoi (probablement lies au mode envoi ST/SV ou a une
    option Paperless), utilises comme indice indirect quand Nbr Colis est absent/nul.

### Phase 10 — Verification exhaustive zone par zone + liste complete des zones (~2:32 - 3:00)

30. (~2:32-2:40) Sur le fichier `2025_04_UPS_Import` (mois precedent), on observe des lignes avec des
    valeurs riches en colonne Ref.1/Ref.2/Nom, confirmant le contenu de certaines colonnes : ligne 2
    Id client `EXP20250325-2102627`, Pays `CA`, Zone `10`, Droits et taxes `67,50 €` (deja vu Phase 4).
    D'autres lignes : Ref.1 `2123433`, Pays `FR`, Zone `France`, Fret `7,24 €`, plus-value BtoC `1,41` ;
    Nom `CROSS DOCKING`, Pays `LU`, Zone `5`, Nbr Colis `10`, Poids `210`, Fret `98,51 €` ; Nom `DEMANDE
    ANOTHERWAY`, Pays `FR`, Zone `France` ; Ref.1 `EXP2025040`, Nom **`Pickup Portugal`**, Pays `PT`,
    **Zone `356`**, Poids `15`, mode envoi `ST`, Fret `45,14 €`, plus-value BtoC `4,15` — confirme que la
    zone 356 est bien liee a un "Pickup" (enlevement) special au Portugal, pas une erreur de saisie.

31. (~2:40-2:48) **Liste complete des valeurs de zone disponibles dans le filtre de la colonne Zone**,
    lue integralement dans le volet deroulant : `71`, `355`, `357`, `703`, `705`, `706`, `707`, `757`,
    **`France`** (en plus des valeurs numeriques simples 1-12, 52 etc. vues precedemment, non toutes
    visibles sur cette capture precise car defilement partiel de la liste). Confirme que les zones a 3
    chiffres (355, 356, 357, 703, 705, 706, 707, 757) forment une famille de codes distincts des zones
    "standard" 1-12/52, tous associes a des pays hors France (GB notamment pour 703/705/706, PT pour 356).

32. (~2:48-3:00) Sur le fichier `2025_05_UPS_Import` (mois courant), tri/filtre affichant les lignes avec
    Poids = `0` : **20 enregistrement(s) trouve(s) sur 10225**. Echantillon : Zone `12`/Pays `NZ` (Poids 1),
    Zone `703`/Pays `GB` (Droits et taxes `48,69 €`), Zone `6`/Pays `NO` (Droits et taxes `7,00 €` et
    `7,31 €`), Zone `9`/Pays `US` (Droits et taxes `1 062,79 €`), plusieurs lignes Zone `705`/`706`/Pays
    `GB` avec Droits et taxes variables (`56,15 €`, `7,43 €`, `187,85 €`, `26,66 €`, `86,62 €`, `60,41 €`),
    et lignes `UPS_COD` en bas de liste (Zone `France`, Adresse `9,5`) — cette vue confirme que les
    quelques dizaines de lignes a Poids=0 se concentrent presque exclusivement sur des envois internationaux
    speciaux avec Droits et taxes/douane (GB, NO, US, NZ), coherent avec des envois documentaires ou factures
    commerciales sans poids physique associe.

### Phase 11 — Correction manuelle du Poids sur les lignes a Poids = 0 (~3:00 - 3:20)

33. (~3:00-3:12) Sur le fichier `2025_05_UPS_Import`, la colonne **Poids** (L) est filtree/triee. Le
    volet de filtre liste les valeurs disponibles : `0`, `0,5`, `1`, `1,5`, `2`, `2,5`, `3`, `3,5`, `4`...
    — confirme que **les poids sont exprimes par increments de 0,5** (coherent avec la regle memoire projet
    "poids interdit d'avoir des virgules hors 0,1 et 0,5" et l'usage d'ARRONDI.SUP). Vue triee sur les
    20 lignes a Poids = 0 identifiees en Phase 10 (etape 32).

34. (~3:12-3:20) Sur la derniere ligne du tableau filtre (ligne `9768`, N° Tracking `1ZJ40E856831381684`,
    Transport `UPS_COD`, Ref.1 `SJ40E381684`, Pays vide, Zone `France`, Fret `5,42 €`), la personne
    **modifie manuellement la valeur de la colonne Poids (L) de `0` a `1`** (saisie clavier directe, cellule
    `L9768`). Resultat visible : barre d'etat passe de "20 enregistrement(s) trouve(s)" a **"17
    enregistrement(s) trouve(s) sur 10225"** (le filtre "Poids = 0" perd des lignes au fur et a mesure que
    chaque poids nul est corrige manuellement a la main). Cette correction manuelle ligne-par-ligne confirme
    qu'un poids de `0` est traite comme une anomalie a corriger avant l'import final (probablement en se
    referant a la facture PDF UPS source ou a l'export ERP pour la valeur reelle).

35. (~3:20-3:28) Meme logique de correction appliquee cette fois a la colonne **Nbr Colis** (K) sur les
    lignes precedemment reperees a Poids/Nbr Colis = 0 : plusieurs cellules K passent de `0` a `1`
    (saisie manuelle), avec mise a jour du filtre en cascade ("17 enregistrement(s) trouve(s) sur 10225").
    Confirme que la correction porte a la fois sur **Nbr Colis et Poids** pour les lignes anormales (les
    deux colonnes sont remises a 1 quand elles sont a 0, sauf cas particulier deja traite via la regle
    plus-value BtoC de la Phase 9).

### Phase 12 — Verification finale du fichier CSV et acces a l'ERP interne "Taxe gasoil" (~3:28 - 3:56)

36. (~3:28-3:36) Retour sur les dernieres lignes du fichier `2025_05_UPS_Import` (lignes 10206 a 10226,
    toutes `UPS_COD` / `CHEQUE, ES...` / Pays `FR` / Zone `France` / mode envoi `ST` / TVA `0,2`, Poids en
    increments de 0,5). **Derniere ligne du fichier = ligne 10226** : Transport `UPS` (pas UPS_COD),
    Ref.1 `FACTURE 16`, N° Tracking `91000219382` (format numerique pur, pas un vrai tracking UPS `1Z...`),
    Pays `FR`, Zone `France`, Poids `10`, Fret `8,64 €` — ligne de type "facture"/frais divers sans tracking
    colis standard, en toute fin de fichier. Menu contextuel clic-droit ouvert sur une ligne vide (Couper /
    Copier / Options de collage / Inserer / Supprimer / Effacer le contenu...) — verification/nettoyage des
    lignes vides en fin de tableau.

37. (~3:36-3:40) Retour en haut du fichier (cellule A2, valeur "UPS"), vue d'ensemble complete du tableau
    (lignes 2 a 46+) confirmant la structure finale du fichier `2025_05_UPS_Import` pret a etre depose dans
    l'ERP.

38. (~3:40-3:56) **Bascule vers le navigateur web**, application interne **"Systeme d'Information de La
    Ruche"** (URL : `si.laruche-logistique.fr/facturation/transporteur/home`), utilisateur connecte
    "Thomas LARGERON". Ecran **"Facturation - Gestion des transporteurs"** (24 transporteurs au total).
    Recherche "ups" -> 2 resultats : **`UPS`** (libelle "UPS") et **`UPS_COD`** (libelle "UPS_COD") —
    confirme que UPS et UPS_COD sont bien geres comme **deux transporteurs/codes distincts** dans l'ERP,
    coherent avec la colonne Transport du fichier import qui alterne entre les deux valeurs.
    Fiche transporteur UPS ouverte, avec 4 onglets : **General**, **Transporteurs ERP** (10), **Plus-values
    colis** (7), **Taxe gasoil** (306, onglet actif). Dans l'onglet Taxe gasoil : filtre "Mode de livraison"
    = **`[65] - express saver`**, tableau paginable (116/116 resultats pour ce mode, 12 pages) avec colonnes
    **Mode livraison / Date debut / Date fin / Pourcentage**. Valeurs lues (ordre decroissant de date) :
    09/06/2025-15/06/2025 = **29.25%**, 26/05/2025-08/06/2025 = **30.25%**, 19/05/2025-25/05/2025 =
    **29.25%**, 12/05/2025-18/05/2025 = **29.75%**, 05/05/2025-11/05/2025 = **30.50%**, 28/04/2025-04/05/2025
    = **29.75%**, 21/04/2025-27/04/2025 = **29.50%**, 14/04/2025-20/04/2025 = **30.75%**, 07/04/2025-
    13/04/2025 = **31.25%**, 31/03/2025-06/04/2025 = **30.75%**. Cet ecran ERP est **la source de reference
    du "taux theorique" gazole** utilise dans le controle "CONTROLER ASSIETTE TAXE GAZOLE" du TCD (Phase 3,
    etape 8) : le taux varie chaque semaine et depend du mode de livraison (ici `express saver`, present
    aussi dans la liste des codes de service UPS).

### Phase 13 — Import final du fichier CSV dans l'ERP (~3:56 - 6:33, fin de la video)

39. (~3:56-4:00) Retour bref sur le classeur `2025_05_Facture UPS`, onglet **Fichier import**, vue complete
    (barre en haut confirme "2025_05_Facture UPS - Excel", onglet actif surligne en bas "Fichier import").
    **Les 3 notes rouges + l'encadre jaune sont maintenant entierement lisibles en colonnes Y-AA** :
    - *"Si nb colis = 0 -> voir montant plus-value BtoC p..."* (suite coupee mais sens confirme par Phase 9).
    - *"Si 1,38 € -> 1 colis, sinon arrondi 20 kg / colis"*.
    - *"Comparer avec exports expedies depuis ERP"*.
    - *"ATTENTION modes ST / SV"*.
    - *"ATTENTION zones 505, 506, etc"*.
    - Encadre jaune, texte lu en quasi-totalite : **"E / P : facturer en P si Plus-value Paperless -"**
      (la phrase semble se couper juste apres le tiret, fin de ligne hors du cadre visible sur toutes les
      frames disponibles — voir Points ambigus).
    Colonnes A-B de l'onglet Fichier import (a gauche) confirment aussi le mapping **E/P ERP** (colonne A :
    `inconnu`/`entreprise`/`particulier`) associe a une **Zone** (colonne B, valeurs numeriques 0-703) —
    exemple ligne 2 : E/P `inconnu`, Zone `12`, ligne 9 : E/P `entreprise`, Zone `4`, ligne 36 : E/P
    `inconnu`, Zone `703`.

40. (~4:00-4:04) Bascule vers l'**explorateur de fichiers Windows**, boite de dialogue **"Envoi du fichier"**
    ouverte a partir du navigateur (upload), navigation deja positionnee sur `Drive > Comptabilite La Ruche >
    $Facturation automatique > 2 - Fichiers csv import > 2025 > 2025 05`. Contenu du dossier : `2025_05_DPD_
    Import` (217 Ko, OK), `2025_05_Kuehne_Import` (29 Ko, OK), **`2025_05_UPS_Import`** (905 Ko, statut OK
    vert — la taille est passee de vide/en cours a **905 Ko**, signe que le fichier vu construit et
    sauvegarde en Phase 6-11 est maintenant complet), et `A IMPORTER 2025_05_UPS_Import_2` (1 Ko, toujours en
    statut cloud non synchronise, laisse de cote). **Selection du fichier `2025_05_UPS_Import.csv`** dans la
    boite de dialogue.

41. (~4:04-4:08) Retour dans l'application web ERP, sur la page **"[TRANSPORT] - Imports"** (menu lateral
    gauche, section "Gestion Facturation" ; menu complet visible sur cette page : "[TRANSPORT] - Imports",
    "[TRANSPORT] - Prix de vente (OLD)", "[TRANSPORT] - Prix de vente", "[TRANSPORT] - Reclamations",
    "[TRANSPORT] - Grilles tarifaires", "[TRANSPORT] - Transporteurs", **"[TRANSPORT] - Grilles zoning"**,
    **"[TRANSPORT] - Grilles zones eloignees"**, puis section [APPRO] et [EXPEDITION] — confirme l'existence
    d'un ecran ERP dedie "Grilles zoning" distinct du fichier Excel de calcul, potentiellement la source de
    reference du zoning cote ERP, a explorer si besoin pour la logique de zone). Fenetre modale
    **"Import fichier"** avec deux champs obligatoires : **"Fichier"** (deja renseigne :
    `2025_05_UPS_Import.csv` via le bouton "Parcourir...") et **"Periode"** (champ date avec picker
    calendrier). Sur une frame legerement posterieure (juste avant le clic sur Valider), le champ **Periode
    est rempli avec la valeur `05/2025`** — confirme que la periode d'import est saisie manuellement (mois/
    annee) en plus du fichier, avant validation. Tableau en arriere-plan (vide, "Aucune donnee disponible
    dans le tableau") avec les colonnes **Id / Transporteur / Date validite tarifs / Client / Tracking / Nom
    / E/P / Pays / Zone / Nb. Colis / Poids / Mode envoi / Tva / Droits & taxes / Assurance / Zone eloignee /
    Colis volumineux / Adresse / Fret / Plus-value BtoC / Taxe gasoil** — confirme une nouvelle fois
    l'intitule exact des colonnes cote ERP (tres proche du CSV, "Taxe gasoil" au lieu de "Gazole", "Zone
    eloignee" singulier au lieu de "Zones eloignees").

42. (~4:08-6:33) Clic sur **"Valider"**. Une notification cloche apparait : **"L'import FACTURATION_IMPORT
    est en cours"** avec barre de progression (0% 0/0 puis progression), et une bannitere verte de
    confirmation en haut a droite : **"Fichier bien pris en compte. L'import est en cours."** — ceci marque
    la fin du process de la video : le fichier `2025_05_UPS_Import.csv` (10 226 lignes de donnees, hors
    en-tete) construit et verifie manuellement tout au long de la video est **deverse dans l'ERP interne La
    Ruche** via cet ecran d'import dedie.

## Points ambigus / a confirmer par le pole transport (mise a jour incrementale)




- Etape 3 : la formule `=SI(RECHERCHE(H10303;'zone colis poids assurance'!D:D;'zone colis poids
  assurance'!G:G)=0;"";MAX(10;ARRONDI.SUP(0,02*RECHERCHE(...);2)))` est situee dans une colonne dont le nom
  exact n'est pas lisible directement sur cette frame de l'onglet Facture UPS. Sur la base du calcul (2% +
  plancher 10), il s'agit tres probablement de la colonne **Assurance**, pas Zone — a confirmer par
  relecture cellule-par-cellule ou aupres du pole transport.
- Etape 5 : la valeur `Zone = 33` en ligne 0289 de l'onglet Excel "Fichier import" parait aberrante comparee
  aux valeurs 1-4 observees partout ailleurs sur ce meme lot de lignes (mais coherente avec des valeurs de
  zone a 2-3 chiffres vues plus tard dans le CSV final, ex. 356, 703 — donc peut-etre pas une erreur,
  simplement une zone eloignee/DOM-TOM) — a confirmer.
- Etape 8 : les lignes `A1912WSTGYP`, `A1912WTW9TC`, `AB` en bas du TCD (poids 0, colonnes eparses, `#N/A`)
  ne sont pas clairement identifiees — possible frais annexes (surcharge carburant/redevance) sans tracking
  colis associe, a confirmer.
- Etape 14-16 : signification exacte des codes de zone a 3 chiffres (`703` pour GB, `356` pour PT, et `505`/
  `506` mentionnes dans la note rouge "ATTENTION zones 505, 506, etc") par rapport aux zones a 1-2 chiffres
  (`5`, `6`, `9`, `10`, `11`, `52`) — sont-ce des zones internationales speciales (DOM-TOM, zones eloignees
  Brexit) distinctes de l'echelle "1 a 12" habituelle ? A confirmer avec le Guide des Services Viticolis /
  onglet `zones` mentionne dans les reperes projet.
- Etape 16 : le texte complet de la regle **"E / P : facturer en P si Plus-value Paperless -..."** est
  coupe a droite sur la frame disponible (hors cadre visible) — la condition exacte et la suite de la phrase
  ne sont pas lisibles. A relire sur une frame plus large ou zoomee, ou a confirmer directement aupres du
  pole transport : est-ce "si Plus-value Paperless > 0" ou une autre condition ? Cette regle semble
  s'ajouter (ou preciser) la regle memoire projet existante ("SI(X2="";"entreprise";"particulier")").
  De meme, la regle "Si 1,38 € -> 1 colis, sinon arrondi 20 kg / colis" n'est pas totalement limpide : le
  contexte exact (quelle colonne source, quel montant a 1,38 €) n'est pas visible sur les frames disponibles.
- Etape 16 : la distinction entre modes d'envoi `ST` et `SV` (note "ATTENTION modes ST / SV") n'est pas
  expliquee davantage dans la video — a confirmer si SV = Saver/Standard Ground vs ST = Standard/domestique,
  et si cela influe sur le calcul tarifaire zoning.
- Etape 17-19 : le detail exact du deuxieme collage "en valeurs" (quel raccourci/option de collage special
  utilise — Ctrl+Alt+V puis "Valeurs" probablement) n'est pas visible frame par frame ; seul le resultat
  avant/apres (erreurs #REF! puis donnees propres) est observable.
- Etape 41 : la page laterale de l'ERP "Gestion Facturation" comporte un menu **"[TRANSPORT] - Grilles
  zoning"** et **"[TRANSPORT] - Grilles zones eloignees"** distincts, non ouverts/explores dans cette video —
  il pourrait s'agir de la source de reference "officielle" du zoning cote ERP (par opposition/en complement
  a l'onglet `zones` du fichier Excel de calcul mentionne dans les reperes memoire projet). A verifier lors
  d'une prochaine video ou aupres du pole transport si cette grille ERP doit primer sur le fichier Excel pour
  le codage du carrier.
- Non couvert par cette video (a chercher dans une autre video ou a confirmer par le pole transport) : la
  gestion explicite du "1Z79" (colis a exclure + demande d'avoir) mentionnee dans les reperes memoire
  generaux UPS n'apparait pas distinctement dans cette video precise (elle documente uniquement l'elaboration
  du fichier d'import a partir des donnees deja integrees, pas le traitement des demandes d'avoir/litiges).

## Verification finale (complement de transcription)

La video a ete revisionnee integralement sur le segment final 4:05-6:33 (49 frames supplementaires extraites
a 1 frame/3s, densite superieure a la premiere passe) pour s'assurer qu'aucune etape n'avait ete omise apres
l'interruption de la tentative precedente. Resultat : les phases 9 a 13 deja redigees dans ce fichier
couvrent fidelement tout le contenu observe jusqu'a la toute fin de la video (bandeau de confirmation ERP
"Fichier bien pris en compte. L'import est en cours."). Deux precisions mineures ont ete ajoutees a l'etape
41 (valeur exacte du champ "Periode" = `05/2025`, liste complete du menu lateral "Gestion Facturation" de
l'ERP). Aucune phase manquante n'a ete identifiee : la transcription est consideree complete du debut
(0:00) a la fin (6:33) de la video.
