# Video UPS - Preparation fichier import (UPS_Preparation fichier import.mp4)

Source : `Transporteurs/UPS/UPS_Preparation fichier import.mp4`
Duree : 2168,4 secondes (~36 min 08 s), capture d'ecran silencieuse (Excel/CSV/navigateur).
Methode : extraction 1 frame / 8s (271 frames fixes, video complete) + detection de changement de
scene (138 frames, video complete), plus un passage dense complementaire (`focus23`/`focus23b`,
~330 frames) deja exploite lors d'une tentative precedente sur le segment 20:00-27:30. Les deux
jeux de frames ont ete relus en ordre chronologique par lots de 10.

Mois traite : **mai 2026** (classeur `2026_05_Facture UPS.xlsx`, fichier `2026_05_UPS_Import.csv`,
compte `0000A1912W`, transporteur `UPS`). A ne pas confondre avec `UPS_2_Preparation fichier
import.mp4` (plus ancien/different mois, hors perimetre de cette transcription) ni avec
`UPS_1_Preparation fichier a coller.mp4` (integration CSV du mois d'avril 2026, cf.
`Video_UPS_3_Preparation_fichier_a_coller.md`).

Statut : TRANSCRIPTION COMPLETE (0:00 a 36:08). Le segment 20:00-27:30 reprend, en le
reformulant, le resultat d'une analyse anterieure deja validee (repli manuel du Poids via
billing.ups.com) ; le reste de la video (0:00-20:00 et 27:30-36:08) a ete transcrit integralement
pour cette tache. **Mise a jour** : le segment 19:00-28:00 a ete re-examine en detail (echantillonnage
dense 1 frame/2s puis 1 frame/1s) pour verifier un repli du Poids jusqu'a mars 2026 (mois-2) signale
par le pole transport ; CONFIRME a 27:04-27:06, cf. etape 32bis de la Phase 9.

## Contexte general

Cette video documente la suite du process de facturation UPS pour mai 2026, apres l'integration
des CSV bruts (etape documentee separement, cf. `Video_UPS_3_Preparation_fichier_a_coller.md` /
`Video_UPS_1_Integration_CSV.md`) : verification du classeur de calcul `2026_05_Facture
UPS.xlsx`, lecture/confirmation des formules des onglets `Fichier import`, `TCD`, `Comptes UPS`,
`zone colis poids assurance`, `Bilan factures`, `Charge.CHG_CODE`, `ST SV`, puis construction et
controle qualite du fichier CSV d'import final `2026_05_UPS_Import.csv` (7249 lignes de donnees,
hors en-tete), avec en particulier le traitement approfondi des lignes a **Poids = 0**.

Onglets du classeur `2026_05_Facture UPS.xlsx` observes dans la video (barre d'onglets en bas) :
`ST SV`, `Zone`, `Clients log`, `zone colis poids assurance`, `Bilan factures`, `Facture UPS`,
`Comptes UPS`, `TCD`, `Fichier import`, `Demande avoir`... (structure coherente avec celle deja
documentee dans `Video_UPS_2_Elaboration_import.md`, `Charge.CHG_CODE` en plus).

**En-tetes des 23 colonnes du fichier d'import CSV**, confirmees a l'ecran (`2026_05_UPS_
Import.csv`, colonnes A a W) : `Transporteur`(A) / `Date validite`(B) / `Ref.1`(C) / `Ref.2`(D) /
`Id client`(E) / `N° Tracking`(F) / `Nom`(G) / `E / P`(H) / `Pays`(I) / `Zone`(J) / `Nbr Colis`(K)
/ `Poids`(L) / `mode envoi`(M) / `TVA`(N) / `Droits et taxes`(O) / `Assurance`(P) /
`Zones eloignees`(Q) / `Colis volumineux`(R) / `Adresses`(S) / `Fret`(T) / `plus-value BtoC`(U) /
`Gazole`(V) / `Nb Colis`(W). Identique aux videos precedentes.

## Transcription chronologique

### Phase 1 - Chargement Power Query "UPS" et controle du volume (0:00 - 1:00)

1. (0:00-0:20) Classeur `Classeur1` (Excel), panneau **Requetes et connexions** ouvert : 5
   requetes visibles (`Transformer le fichier a partir de U...` [2], `Requetes d'assistance` [3]
   avec `Exemple de fichier`, `Parametre1 (Exemple de fichier)`, `Transformer le fichier`, et
   `Autres requetes` [1] = **UPS**, message **"52 104 lignes chargees"**). Meme mecanique Power
   Query que dans `Video_UPS_1`/`Video_UPS_3` (combiner les CSV Invoice_* du dossier du mois).
   Onglet actif **Creation de tableau** (ruban contextuel du tableau structure "UPS").
2. (0:20-0:40) Le tableau `UPS` affiche des colonnes generiques renommees en partie (`Source.Name`
   /`Version`/`Numero du...`/`Numero de...`/`Pays d'or...`/`Date de la...`/`Numero de...`/
   `Code de t...`/`Code det...`/`Numero de t...`/`Code de la...`/`Montant d...`/`Date de
   l'o...`/`Reference`/`Numero de l'envoi pr...`/`Numero`/`Numero de reference`...) sur des lignes
   de fin de tableau (~26777-26804) : donnees issues des fichiers `Invoice_202600556979_
   060326.csv` (compte `0000WV5788`, montant net `1223,54`), `Invoice_202600498843_051526.csv`
   (compte `0000J40E95`), `Invoice_202600552608_060326.csv` et `Invoice_202600564118_060326.csv`
   (comptes `0000A1912W` / `000079A7T0`). Deux dernieres lignes du tableau brut : trackings
   `A1912WTYQCK` (montant `14,53`) et `A1912WVVFCB` (montant `7,01`), issus de la facture
   `Invoice_202600477649/477650_050826.csv` — **ce sont precisement les deux trackings qui seront
   corriges manuellement plus loin dans la video (Phase 20:00-27:30) pour un Poids = 0**,
   confirmant qu'ils sont deja presents des la source CSV brute avec un montant facture mais un
   poids absent des colonnes generiques a ce stade.
3. (0:40-1:00) Deuxieme fenetre Excel ouverte en parallele : `2026_05_Facture UPS.xlsx`. Deux
   fenetres cote a cote dans la barre des taches (`2026_05_Facture UPS.xlsx` et `Classeur1`).

### Phase 2 - Onglet "Facture UPS" : donnees post-integration (1:00 - 2:20)

4. (1:00-1:40) Onglet **Facture UPS** du classeur `2026_05_Facture UPS.xlsx` actif. Colonnes
   visibles : `Clients`(A), `Montant...`(B), `Mode env...`(C), `Categorie`(D, valeurs `Surcharge
   de securite`/`TVA`/`Fret`/`Taxe gazole`/`plus-value BtoC`/`#N/A`), `Version`(E), `Numero de...`
   (F/G), `Pays d'or...`(H) = `FR`, `Date de la facture`(I) = `03/06/2026`, `Numero de facture`
   (J), `Code de t...`(K), `Code det...`(L) = `1`, `Numero de t...`(M) = `FR75804949865`, `Code de
   la...`(N) = `EUR`, `Montant...`(O) = `163726,16` / `162,67`, `Date de l'operation`(P), et plus
   loin `Numero de l'envoi principal`, `Numero de reference` (`1Z79A7T06802040157` etc. pour le
   compte `079A7T0`, `1992187050` pour le compte `A1912W`). Confirme la structure deja documentee
   (colonnes generiques issues du Power Query, alimentant ensuite le TCD).
5. (1:40-2:20) Barre d'onglets en bas du classeur `2026_05_Facture UPS.xlsx`, en ordre : `ST SV`,
   `Zone`, `Clients log`, `zone colis poids assurance`, `Bilan factures`, `Facture UPS`, `Comptes
   UPS`, `TCD`, `Fichier import`, `Demande avoir`... Onglet **Bilan factures** : tableau croise
   dynamique (ou tableau recapitulatif) listant les numeros de facture avec dates et montants,
   destine a etre recoupe avec le "Centre de facturation" billing.ups.com (cf. etape suivante).

### Phase 3 - Controle croise avec billing.ups.com "Centre de facturation" (2:20 - 3:20)

6. (2:20-3:00) Bascule navigateur, page **billing.ups.com/ups/billing/invoice** — **"Centre de
   facturation"** UPS, utilisateur "Camille". Tableau des factures avec colonnes **Numero de
   facture / Numero de compte / Date de facture / Statut de facture / Statut de paiement / Montant
   facture / Montant exigible / Date d'echeance / Type / PDF / Actions**. Lignes observees :
   `000202600568775` et `000202600568776` (compte `A1912W`, 04/06/2026, `63,53 €` et `9,95 €`,
   type **Import**), `000202600561549`/`000202600561569` (comptes `4V4555`/`4V47R4`, statut
   "Cloturee - Non payable en ligne", `0,00 €`), `000202600552608` (compte `A1912W`, `163 726,16
   €`, type **National/Export**, 03/06/2026), `000202600556979` (compte `WV5788`, `1223,54 €`),
   `000202600555499/555492/555490/555491` (comptes `J40F00`/`J40E85`/`J40E82`/`J40E83`),
   `000202600564118` (compte `79A7T0`, `162,67 €`, statut **"En retard"**), `000202600543689`
   (compte `A1912W`, `27,26 €`, type Import, 02/06/2026), `000202600533268` (compte `A1912W`,
   `82,29 €`, 29/05/2026). Cette page confirme que l'onglet `Bilan factures` du classeur Excel sert
   a comparer ligne a ligne les numeros/dates/montants de facture generes par Power Query avec la
   liste officielle du Centre de facturation UPS (controle de non-omission d'une facture).
7. (3:00-3:20) Retour classeur Excel, poursuite de la lecture de l'onglet `Facture UPS` /
   `Bilan factures`.

### Phase 4 - Onglet "zone colis poids assurance" : TCD de calcul du poids (3:20 - 5:00)

8. (3:20-4:00) Onglet **zone colis poids assurance** : tableau croise dynamique avec colonnes
   observees **Numero de suivi** (ligne), **Max de Poids facture**, **Poids UPS**, **Poids
   UPS_COD** (ou equivalent) — confirme l'existence d'un TCD dedie qui agrege le poids par
   tracking (coherent avec le commentaire du code `facturation-app/src/carriers/ups/index.js`
   "Poids/Nombre de colis/Montant assurance : agreges par tracking via MAX").
9. (4:00-4:40) Meme onglet, defilement : colonnes **Code Client**, **Poids**, valeurs de zone,
   d'assurance. Ce TCD alimente directement (par renvoi de formule) les colonnes correspondantes
   de l'onglet **Fichier import** et du **TCD** principal (cf. pattern deja documente en Phase 3
   de `Video_UPS_2_Elaboration_import.md` : `=SI(TCD!I10314=0;"";TCD!I10314)`).
10. (4:40-5:00) Formule observee dans une cellule du TCD principal (colonne D) :
    `=RECHERCHEX(E7266;'zone colis poids assurance'!D:D;'zone colis poids assurance'!I:I)` —
    recherche du n° de suivi dans l'onglet `zone colis poids assurance` (colonne D), renvoie la
    colonne I de ce meme onglet (poids ou zone selon la position, non confirmee avec certitude sur
    cette frame — cf. Points ambigus).

### Phase 5 - Onglet "Charge.CHG_CODE" et TCD principal (5:00 - 6:30)

11. (5:00-5:40) Onglet **Charge.CHG_CODE** : table de correspondance entre codes de frais UPS
    (colonne "Code" ou equivalent, valeurs `FRT`/`TAX`/`FSC`/`ACC`...) et la **Categorie** deja
    vue en Phase 2 (`Fret`/`TVA`/`Taxe gazole`/`Surcharge de securite`/`Adresse`/`Assurance`/
    `Colis volumineux`/`plus-value BtoC`) : c'est cette table qui pilote la formule de la colonne
    Categorie de l'onglet `Facture UPS` (recherche du code de charge UPS -> categorie metier).
12. (5:40-6:30) Retour onglet **TCD** (tableau croise dynamique principal) : colonnes confirmees
    identiques a celles deja documentees dans `Video_UPS_2_Elaboration_import.md` — `Logistique` /
    `Cout` / `Code Client` / `Poids` / `Numero de suivi` / `Categorie` / `Adresse` / `Assurance` /
    `Colis volumineux` / `Droits et taxes` / `Fret` / `plus-value BtoC` / `Surcharge de securite` /
    `Taxe gazole` / `TVA` / `Zones eloignees`, avec zone de controle a droite **"CONTROLER
    ASSIETTE TAXE GAZOLE"** (% gazole, date d'enlevement, mode transport, gazole theorique, ecarts,
    Gazole vendu, retours).

### Phase 6 - Onglet "Fichier import" : formules detaillees colonne par colonne (6:30 - 12:00)

13. (6:30-7:30) Onglet **Fichier import**, cellule d'en-tete **E/P ERP** (colonne A ou B selon la
    frame). **Formule confirmee dans la barre de formule** :
    ```
    =SI(X2="";"entreprise";"particulier")
    ```
    Regle exacte de determination du statut Entreprise/Particulier (E/P ERP) a partir d'une
    colonne X (probablement liee a l'option "Plus-value Paperless" mentionnee dans la note rouge
    documentee en Phase 5 de `Video_UPS_2_Elaboration_import.md`) : confirme et precise la regle
    memoire projet `SI(X2="";"entreprise";"particulier")`.
14. (7:30-8:10) Formule observee sur une autre ligne de la meme colonne : `=SI(B7="particulier";
    "P";SI(X7="";"E";"P"))` — cette deuxieme formule (colonne finale **E/P**, code court E/P
    utilise dans le fichier import) transforme le resultat texte `entreprise`/`particulier` de
    l'etape precedente en code court `E`/`P`, avec le meme test sur la colonne X en cascade.
15. (8:10-9:00) Formule de la colonne **Zone** (F ou J selon la frame), tres importante :
    ```
    =SI(NB.SI('Comptes UPS'!A:A;DROITE(GAUCHE(I2;8);6))=0;"inconnu";
        RECHERCHEX(DROITE(GAUCHE(I2;8);6);'Comptes UPS'!A:A;'Comptes UPS'!B:B))
    ```
    Extrait les 6 caracteres en positions 3-8 du numero de suivi (colonne I, `DROITE(GAUCHE(I2;8)
    ;6)` = sous-chaine du tracking), verifie si ce code est present dans l'onglet **Comptes UPS**
    (colonne A), et si oui renvoie la zone associee (colonne B de ce meme onglet) ; sinon renvoie
    le texte `"inconnu"`. C'est le mecanisme central de determination de la Zone a partir du
    numero de compte UPS encode dans le tracking (cf. memoire projet "Numeros comptes UPS").
16. (9:00-9:40) Formule de la colonne **Nbr Colis** :
    ```
    =MAX(RECHERCHEX(I2;'zone colis poids assurance'!D:D;'zone colis poids assurance'!E:E);A2)
    ```
    Prend le MAX entre la valeur trouvee par recherche du tracking (colonne I) dans l'onglet
    `zone colis poids assurance` (colonne E, "Nombre de colis") et une valeur de secours en
    colonne A du fichier import (probablement "Nb Colis ERP", cf. la note rouge documentee
    precedemment "Si nb colis = 0 -> voir montant plus-value BtoC").
17. (9:40-10:30) Formule de la colonne **TVA** : `=SI(TCD!N5="";0;0,2)` — si la cellule
    correspondante du TCD (colonne N, "TVA") est vide, la TVA du fichier import vaut 0, sinon
    0,2 (20%). Logique binaire simple (TVA francaise standard ou 0 pour l'international).
18. (10:30-11:10) Formule de la colonne **Pays** :
    ```
    =RECHERCHEX(I2;'[2026 05 - Export expéditions_brut.xlsx]exportDemandeExpedition_2026050'!$AP:$AP;
                 '[2026 05 - Export expéditions_brut.xlsx]exportDemandeExpedition_2026050'!$Q:$Q;"")
    ```
    Meme mecanisme de recherche croisee que celui deja documente pour le Poids (voir Phase
    "20:00-27:30" ci-dessous) : recherche du n° de suivi dans le fichier reseau externe `2026 05 -
    Export expéditions_brut.xlsx` (onglet `exportDemandeExpedition_2026050`), colonne AP
    (PRO_TRACKING), renvoie la colonne Q (Pays) de ce meme classeur. Confirme que **le fichier
    d'export ERP mensuel sert de source de repli pour plusieurs colonnes (Poids, Pays, et
    vraisemblablement Zone), pas seulement pour le Poids**.
19. (11:10-12:00) Manipulation **Rechercher et remplacer** (Ctrl+H) sur l'onglet Fichier import,
    remplacement d'une virgule par du vide dans une plage de formules — vraisemblablement un
    correctif technique lie au format des references de plage dans les formules RECHERCHEX
    (nettoyage de separateurs). Detail exact de l'operation non totalement lisible (cf. Points
    ambigus).

### Phase 7 - Enregistrement du classeur et export CSV (12:00 - 13:30)

20. (12:00-12:40) Boite de dialogue **"Enregistrer sous"** : dossier cible reseau (`1 - Factures
    transporteurs + calculs`), champ "Nom de fichier" = `2026_05_Facture UPS.xlsx`.
21. (12:40-13:30) Sur le fichier CSV **`2026_05_UPS_Import`** (feuille active `2026_05_UPS_
    Import.csv` visible dans les frames suivantes), collage du bloc de donnees calcule dans
    l'onglet Fichier import (meme mecanisme de collage en deux temps — formules puis valeurs —
    deja documente dans `Video_UPS_2_Elaboration_import.md`, non re-capture en detail ici).

### Phase 8 - Controle qualite du fichier CSV : filtres colonne par colonne (13:30 - 20:00)

22. (13:30-14:20) Sur `2026_05_UPS_Import.csv`, filtre sur la colonne **N° Tracking**, recherche
    textuelle `"1zw"` dans la liste de valeurs du filtre (trackings `1ZWV5788900115871`,
    `1ZWV5788900877352`, `1ZWV5788900964687`, `1ZWV5788901097836`, `1ZWV5788901141368`,
    `1ZWV5788901695985`...) — verification cible sur le prefixe de compte `WV5788` (compte
    identifie en Phase 3, facture `000202600556979`).
23. (14:20-15:00) Filtre sur la colonne **Zone**, recherche du code pays `AT` (Autriche) : liste
    de lignes avec Zone = `6` pour toutes les lignes AT observees (trackings `1ZA1912WD9...`
    divers, Poids 2 a 97, mode envoi `ST`, TVA `0,2`) — confirme Zone=6 pour l'Autriche dans le
    zonage UPS.
24. (15:00-15:40) Filtre sur la colonne **Zone** = `6`, Pays = `NO` (Norvege) et `BE` (Belgique) :
    lignes Zone `6`/Pays `BE`... et Zone `5`/Pays `BE` observees sur un large echantillon (~40
    lignes), confirmant Zone 5 pour la Belgique (deja vu dans `Video_UPS_2`) mais aussi Zone 6
    pour certaines lignes BE — **incoherence apparente a confirmer** (cf. Points ambigus).
25. (15:40-16:20) Menu contextuel (clic droit) sur l'en-tete de colonne **Zone** : options
    standard (Couper/Copier/Options de collage/Collage special/Inserer/Supprimer/Effacer le
    contenu/Format de cellule/Largeur de colonne/Masquer/Afficher) — ajustement de mise en forme,
    pas une modification de donnees.
26. (16:20-17:00) Indicateur de collage actif (icone flottante `(Ctrl)`) sur la colonne **Nbr
    Colis** (K), valeur `3` — collage en cours dans cette colonne, pas de detail supplementaire
    visible.
27. (17:00-18:00) Bascule navigateur : onglet **PDF "Guide des services Viticolis - 2026"** (page
    15/19), tableau de zones destination/origine par departement francais (colonnes 49 a 98) et
    tableau international ("A destination de" / "En provenance de") avec valeurs de zone par pays
    : Allemagne (zones 5/6/7 selon code postal), **Angleterre/Ecosse/Pays de Galles** (zones
    **704/705/706/707** selon code postal de destination — plus fin que le simple "703" deja
    documente : la video precedente notait 703/705/706, ici le guide montre 704 egalement selon
    la plage de codes postaux), Belgique (zone 4 ou 5 selon origine/destination), Danemark (zone
    6-7), Espagne (zone 5-6 selon code postal), Irlande du Nord/Republique (zone 6-7), Italie
    (zone 4-7 selon code postal), Luxembourg (zone 4-5), Pays-Bas (zone 5-6), Portugal (zone
    6-7). **Ce PDF `Guide des services Viticolis - 2026` (chemin local
    `C:/Users/AT-HYD-010/Desktop/Outils/Utiles/Guide des services Viticolis - 2026.pdf`) est donc
    la reference officielle du zoning UPS**, deja mentionnee dans la memoire projet ("zoning page
    14 du Guide Viticolis" — ici c'est la page 15 qui est montree, tableau international).
28. (18:00-19:00) Retour Excel, `2026_05_UPS_Import.csv`, ligne 5076 : tracking
    `1ZA1912WDK97468698`, E/P = `P`, Nbr Colis = `1`, Poids = `38`, mode envoi `ST`, TVA `0,2`,
    Frêt `27,58`, plus-value `4,6`. Lignes 7131, 7249, 7250 deja visibles en bas de tableau :
    `1ZJ40E856800012818` (Poids `40`), `A1912WTYQCK` (Poids encore vide/0 a ce stade, Frêt
    `14,53`), `A1912WVVFCB` (Poids encore vide/0, Frêt `7,01`) — **ce sont exactement les deux
    dernieres lignes du fichier CSV, correspondant aux deux trackings reperes des la Phase 1
    (etape 2)**, qui seront traitees dans la section suivante (20:00-27:30).
29. (19:00-20:00) Filtre applique sur la colonne **Poids** : **144 puis 142 enregistrement(s)
    trouve(s) sur 7249** — confirme le volume total de 7249 lignes de donnees deja connu, et
    montre le compteur de lignes a Poids=0 en cours d'affinage juste avant le debut de la section
    documentee precedemment (146/7249 -> 144 -> 142, le nombre baisse au fur et a mesure que des
    lignes sont deja retraitees).

### Phase 9 - Traitement du Poids = 0 via export ERP et facture PDF UPS (20:00 - 27:30)

*(Reprise, reformulee, du resultat d'une analyse anterieure deja validee sur ce segment precis.)*

30. Dans le fichier `2026_05_UPS_Import.csv`, colonne **Poids** (L), filtre applique sur
    **Poids = 0** : **146 lignes sur 7249** au total identifiees.
31. Formule utilisee en colonne temporaire pour tenter de resoudre chaque poids manquant :
    ```
    =RECHERCHEX(F2;'[2026 05 - Export expéditions_brut.xlsx]exportDemandeExpedition_2026050'!$AP:$AP;
                 '[2026 05 - Export expéditions_brut.xlsx]exportDemandeExpedition_2026050'!$AI:$AI;"")
    ```
    Recherche du **N° Tracking** (colonne F) dans le fichier reseau
    `\\192.168.5.3\Laruche\$Informatique\Exports 2026\2026 05\2026 05 - Export
    expéditions_brut.xlsx` (onglet `exportDemandeExpedition_2026050`), colonne **AP**
    (PRO_TRACKING), renvoie la colonne **AI** (poids ERP).
32. Comme le mois courant (05) ne matche presque rien (~18 resultats sur 146), la reference du
    classeur externe est **changee vers le mois precedent** (`2026 04 - Export
    expéditions_brut.xlsx`, onglet `exportDemandeExpedition_2026040`) -> **140/146 resultats**
    obtenus. Un troisieme repli vers le fichier CSV `2026_04_UPS_Import.csv` local
    (`=RECHERCHEX(F2;'2026_04_UPS_Import.csv'!$F:$F;'2026_04_UPS_Import.csv'!$J:$J;"")`, observe
    egalement pour la colonne **Zone**) est utilise en complement pour certaines lignes.
32bis. **CONFIRME par reexamen dense (1 frame/2s puis 1 frame/1s) du segment 19:00-28:00** : il
    existe bien un **4e niveau de repli, vers le mois -2 (mars 2026)**, en plus du repli vers avril
    deja documente ci-dessus. A l'ecran vers **27:04-27:06**, sur la ligne 2 du fichier
    `2026_05_UPS_Import.csv` (tracking `1ZA1912W0466714131`, colonne L temporaire, filtre "Poids
    = 0" actif, 146 enregistrements sur 7249), la barre de formule affiche explicitement :
    ```
    =RECHERCHEX(F2;'[2026 03 - Export expéditions_brut.xlsx]exportDemandeExpedition_2026040'!$AP:$AP;
                 '[2026 03 - Export expéditions_brut.xlsx]exportDemandeExpedition_2026040'!$AI:$AI;"")
    ```
    Reference explicite au classeur **`2026 03 - Export expéditions_brut.xlsx`** (mars 2026, soit
    mois-2 par rapport au mois traite mai). Resultat renvoye pour cette ligne : **Poids = 38**
    (colonne L), puis normalise en colonne M via `=ARRONDI.SUP(...)` (visible juste apres, meme
    mecanisme qu'a l'etape 33). **Point notable** : le nom d'onglet interne reste
    `exportDemandeExpedition_2026040` (suffixe "040", identique a celui utilise pour le classeur
    d'avril) alors que le nom du CLASSEUR (fichier .xlsx) est bien celui de mars (`2026 03 -
    Export expéditions_brut.xlsx`) — vraisemblablement un nom d'onglet herite/non renomme dans le
    fichier source ERP, sans consequence sur le resultat puisque RECHERCHEX cible explicitement le
    bon fichier externe via le prefixe `[2026 03 - ...]`. La barre des taches Excel (visible sur
    une frame voisine, ~26:53) confirme egalement l'ouverture simultanee de 3 classeurs d'export
    ERP : `2026 05 - Export expéditio...`, `2026 04 - Export expéditio...` et `2026 03 - Export
    expéditio...`, plus le classeur de calcul et `2026_05_UPS_Import.csv` — **preuve visuelle
    supplementaire que le repli va bien jusqu'a 3 mois en arriere (mai -> avril -> mars)**. Ceci
    confirme et complete la sequence de repli deja documentee : **mois courant (05) -> mois-1
    (04) -> mois-2 (03)** avant le repli manuel final sur facture PDF UPS (etape 34).
33. Le resultat est recopie en colonne M temporaire, puis normalise via
    `=ARRONDI.SUP(L;1)` (arrondi superieur a 1 decimale, coherent avec la regle poids par pas de
    0,5/0,1 deja documentee).
34. Pour les **6 dernieres lignes non resolues** (trackings hors format standard `1ZA1912...`, ex.
    `1ZY09K266718233483`, `A1912WTYQCK`, `A1912WVVFCB`) : repli manuel sur la **facture PDF
    officielle UPS** via `billing.ups.com` (recherche par tracking, page **"Detail d'envoi
    import"**, colonne **"Poids/Conteneur"**, valeurs lues ex. `13,6/17,2 B PKG` — la lettre **"B"
    signifiant "Poids rectifie suite audit UPS"**), report manuel de la valeur, avec **controle
    croise sur ups.com/track** (page de suivi publique) pour confirmer la coherence de l'envoi.
35. En toute fin de cette phase, les deux dernieres lignes du fichier (`A1912WTYQCK`,
    `A1912WVVFCB`, deja reperees en Phase 1/Phase 8) sont corrigees manuellement : **Poids = 8**
    pour `A1912WTYQCK` et **Poids = 13,5** pour `A1912WVVFCB` (valeurs confirmees a l'ecran juste
    apres 27:00) — le filtre "Poids = 0" passe alors de 4 a 0 enregistrement trouve sur 7249.

### Phase 10 - Poursuite du controle qualite : Zone et autres colonnes (27:30 - 31:00)

36. (27:30-28:10) Filtre sur la colonne **Zone** = `0` (vide/inconnu), 128 puis 21 enregistrements
    trouves sur 7249 : lignes avec Nbr Colis eleve (`13`, `59`), Poids variable (`0,1` a `64,5`),
    mode envoi `ST`, majoritairement TVA `0` ou `0,2` — verification des lignes ou la Zone n'a pas
    ete resolue par les formules de recherche croisee (memes lignes potentiellement liees a des
    trackings non-standard type `1957899306`, `1957901183` = comptes/references internes hors
    format tracking UPS classique).
37. (28:10-29:00) Cellule active `J2`, valeur `0`, formule affichee :
    `=RECHERCHEX(F2;'2026_04_UPS_Import.csv'!$F:$F;'2026_04_UPS_Import.csv'!$J:$J;"")` — **confirme
    que la colonne Zone beneficie du meme mecanisme de repli vers le fichier CSV d'import du mois
    precedent que la colonne Poids** (recherche du tracking dans le fichier CSV `2026_04_UPS_
    Import.csv`, colonne F, renvoie la colonne J = Zone de ce fichier). C'est une decouverte
    complementaire importante par rapport a la Phase 20:00-27:30 (qui ne documentait le repli que
    pour le Poids) : **le repli "mois precedent" s'applique donc a la fois au Poids et a la
    Zone**, potentiellement via la meme logique generale pour d'autres colonnes.
38. (29:00-29:40) Filtre sur la colonne **Zone**, recherche/verification sur les valeurs `1957901xxx`
    (references de type "commande" plutot que tracking UPS standard) : ~69-72 lignes affichees
    avec Zone = `0`, Nbr Colis variable, Poids `0,1` a `64,5` — lignes correspondant probablement a
    des envois particuliers/references internes non couvertes par le fichier d'export ERP ni par
    le CSV du mois precedent.
39. (29:40-30:20) Filtre Zone, recherche cible sur pays **`BE`** avec valeurs `France` observees
    en colonne Zone (texte) pour certaines lignes marquees Pays vide — confirme regle deja connue
    "zone=0 + pas de pays -> France" pour les envois nationaux.
40. (30:20-31:00) Bascule navigateur : application interne **"Systeme d'Information de La Ruche"**
    (`si.laruche-logistique.fr/demande/expedition/home`), page **"Visualisation de l'expedition :
    EXP20251209-2521771 (Colis - Livre)"**, utilisatrice connectee "Caroline SCHMITT". Onglets
    **Expediteur / Destinataire / Marchandises a expedier / Mode d'expedition / Programmation
    enlevement / Suivi / Documents / Devis / Informations / Avarie**. Coordonnees visibles :
    societe "SAS VINBIOME", adresse d'expedition "26 RUE DES ECOLES", code postal partiellement
    masque/selectionne (`67...`), ville "Epfig", pays `FR`. Menu lateral complet de l'ERP
    "Expeditions" : `Tableau de bord`, `Expeditions` (sous-menu `Expeditions`, `Donnees UPS`,
    `Expeditions sans mouvement`, `Demande enlevement en erreur`, `Carnet d'adresses
    destinataires`, `Alias correction adresse`, `Codes api transporteurs inconnus`), `Devis`,
    `Gestion Facturation`, `Gestion contre-remboursement`, `Tickets`, `Gestion des avaries`,
    `Notifications Extranet`, `Ordres Expedition`, `Attendu Reception`, `Crm`, `Transporteurs`,
    `Produits`, `Emballages`, `Referentiel Geographique`, `Exports`... **Cette page ERP "Donnees
    UPS"/"Expeditions"** (menu lateral gauche, distincte de la page "[TRANSPORT] - Imports" deja
    documentee dans `Video_UPS_2`) semble etre consultee pour verifier manuellement l'adresse
    et donc la zone/le code postal d'un envoi specifique cite en exemple (verification croisee
    d'une adresse ambigue, cf. Points ambigus).

### Phase 11 - Derniers controles et cloture (31:00 - 36:08)

41. (31:00-33:00) Retour sur `2026_05_UPS_Import.csv`, filtres successifs sur diverses colonnes
    (Zone, Nbr Colis, Droits et taxes) pour verification finale, dans la continuite de la logique
    de controle qualite deja documentee dans `Video_UPS_2_Elaboration_import.md` (Phases 7 a 11 de
    cette derniere : filtres sur Droits et taxes, plus-value BtoC, verification des zones 703/705/
    706/707 pour le Royaume-Uni).
42. (33:00-35:00) Poursuite du defilement/verification du fichier CSV, sans nouvelle formule ou
    ecran significatif capture sur les frames disponibles a cette densite d'echantillonnage (8s) —
    probablement une repetition/generalisation des controles deja vus (filtre par colonne, tri,
    verification de sous-totaux).
43. (35:00-36:08, fin de la video) Derniere frame disponible : fichier `2026_05_UPS_Import.csv`
    toujours actif, aucune action de sauvegarde finale ou d'upload ERP explicitement capturee sur
    cette derniere portion (contrairement a la video `Video_UPS_2_Elaboration_import.md` qui se
    terminait sur l'ecran de confirmation ERP "Fichier bien pris en compte") — la video semble se
    terminer sur le controle qualite du fichier CSV, sans montrer l'etape finale d'upload dans
    l'ERP (cf. Points ambigus).

## Points ambigus / a confirmer par le pole transport

- **Etape 10** : la formule `=RECHERCHEX(E7266;'zone colis poids assurance'!D:D;'zone colis poids
  assurance'!I:I)` du TCD principal n'a pas ete lue avec certitude quant a la colonne resultat
  exacte qu'elle alimente (Poids ? Zone ? Assurance ?) — a confirmer par relecture cellule par
  cellule.
- **Etape 15** : la formule Zone `=SI(NB.SI('Comptes UPS'!A:A;DROITE(GAUCHE(I2;8);6))=0;"inconnu";
  RECHERCHEX(...))` utilise une extraction de 6 caracteres en position 3-8 du tracking — la
  logique exacte de cette extraction (pourquoi ces positions precises) et la structure de
  l'onglet `Comptes UPS` (colonnes A/B) n'ont pas ete entierement verifiees visuellement au-dela
  de ce qui est deja documente dans la memoire projet ("Numeros comptes UPS").
- **Etape 19** : le detail exact de l'operation "Rechercher et remplacer" (Ctrl+H, remplacement
  d'une virgule) sur l'onglet Fichier import n'est pas totalement clair — a confirmer s'il s'agit
  d'un correctif technique recurrent (format des references de plage dans les formules) ou d'une
  correction ponctuelle liee a une erreur de saisie.
- **Etape 24** : des lignes Pays = `BE` (Belgique) ont ete observees avec **Zone = 5 ET Zone = 6**
  selon les lignes dans le meme fichier — incoherence apparente par rapport au Guide des services
  Viticolis (qui indique une zone unique par pays/code postal selon la Phase 27) : a confirmer si
  cela reflete une difference legitime selon le code postal beige exact (le guide montre des
  variations regionales), ou une anomalie de donnee a corriger.
- **Etape 27** : le Guide des services Viticolis - 2026 (PDF) indique des zones **704/705/706/707**
  pour le Royaume-Uni selon le code postal exact de destination, alors que les videos precedentes
  (`Video_UPS_2_Elaboration_import.md`) ne documentaient que 703/705/706 — a confirmer si la zone
  704 existe egalement dans les donnees reelles ou si elle est simplement presente dans le guide
  sans jamais apparaitre en pratique.
- **Etape 37** : la decouverte que la colonne **Zone** beneficie du meme mecanisme de repli vers
  le fichier CSV du mois precedent que la colonne Poids (formule RECHERCHEX vers `2026_04_UPS_
  Import.csv`) n'avait pas ete documentee dans l'analyse precedente du segment 20:00-27:30 (qui ne
  mentionnait ce mecanisme que pour le Poids) — a confirmer aupres du pole transport si ce repli
  "mois precedent" s'applique de facon generale a toute colonne non resolue par les formules
  primaires (Poids, Zone, et potentiellement d'autres), ou seulement a ces deux colonnes
  specifiquement.
- **Etape 40** : l'usage exact de la page ERP "Visualisation de l'expedition" (module
  Expeditions/Donnees UPS de si.laruche-logistique.fr, distinct du module "[TRANSPORT] - Imports"
  deja documente) dans le contexte de cette video n'est pas totalement clair — semble servir a
  verifier une adresse/code postal pour resoudre un cas de zone ambigu, mais le tracking ou la
  ligne CSV precise concernee par cette verification n'a pas pu etre identifiee avec certitude sur
  les frames disponibles.
- **Etape 43 / fin de video** : contrairement a `Video_UPS_2_Elaboration_import.md` (qui se
  terminait sur la confirmation d'upload ERP "Fichier bien pris en compte. L'import est en
  cours."), cette video ne montre pas d'ecran de confirmation d'upload final sur les dernieres
  frames disponibles (jusqu'a 36:08) — soit l'upload a eu lieu apres la fin de l'enregistrement,
  soit une action rapide entre deux frames fixes a ete manquee malgre le passage scene-change. A
  confirmer si une suite existe ou si l'upload est traite dans une video separate non fournie ici.
- Plusieurs segments a densite d'echantillonnage 8s (notamment 33:00-35:00) n'ont pas revele
  d'action metier nouvelle distincte : il est possible que des filtres/verifications
  supplementaires rapides (moins de 8s) aient ete manques sur cette portion malgre le recours aux
  frames de changement de scene — a confirmer si necessaire par une relecture plus dense de ce
  seul segment.
