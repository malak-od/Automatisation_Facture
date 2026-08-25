# Video UPS 2-1 — Expeditions declarees en particulier mais facturees en entreprise par UPS

Source : `Transporteurs/UPS/Process Facturation - Facture UPS - 2-1 - Expés déclarées en particulier mais facturées en entreprise par UPS.mp4`
Duree : 90.7 secondes (video courte, capture d'ecran silencieuse, Excel).
Methode : extraction 1 frame / 2s (45 frames fixes) + detection de changement de scene (9 frames) sur ffmpeg.
Toutes les frames ont ete lues (transcription complete, 0s -> 90,7s).

## Resume de la regle de resolution (a retenir en priorite)

Formule Excel trouvee et confirmee sur plusieurs lignes (J14, J20, J23, J24) de l'onglet
**"Fichier import"** du classeur `2025_05_Facture UPS.xlsx` (`2025 05` = mai 2025) :

```
=SI(A14="particulier";"P";SI(W14="";"E";"P"))
```
(la reference de ligne s'adapte : A20/W20 pour J20, A23/W23 pour J23, etc. -- meme formule copiee vers le
bas sur toute la colonne J)

- Colonne **A** = "E/P ERP" = declaration interne (valeurs vues : `particulier` / `entreprise` / `inconnu`).
- Colonne **W** = "plus-value E..." (plus-value Paperless -- alimentee elle-meme par une formule
  `=SI(TCD!K24="";"";TCD!K24)`, donc issue du TCD/retour UPS).
- Colonne **J** ("E" en entete visible, mais contient des valeurs "P"/"E") = statut E/P retenu, resultat
  de la formule.

Logique de la formule : si l'ERP dit "particulier" -> **P**, sans condition. Sinon (ERP = "entreprise" ou
"inconnu") -> on regarde si une **plus-value Paperless** a ete relevee sur la ligne UPS (colonne W issue
du TCD, non vide) : si oui -> force **P** ; si non -> **E**.

**Attention -- observation factuelle qui contredit cette lecture simple sur les dernieres frames de la
video (lignes 20 et 24) :** sur ces deux lignes, la colonne A affiche `entreprise` et la colonne **W
apparait visuellement vide** dans la vue Excel, et pourtant le resultat affiche en colonne J est **"P"**
(pas "E" comme la lecture litterale de la formule le laisserait attendre si W est reellement vide). Deux
explications possibles, non tranchees par la video : (1) W contient en realite une valeur non nulle mais
non visible sur la capture (largeur de colonne, decimales, format masque) ; (2) c'est precisement le cas
particulier illustre par la video : malgre une plus-value Paperless non detectee/vide en apparence, UPS a
neanmoins facture en entreprise, et une resolution manuelle (hors formule automatique) s'applique alors.
Voir section "Points ambigus" -- **ce point doit etre leve avec le pole transport avant de coder cette
regle dans le carrier**, car il pourrait s'agir du coeur meme du cas documente par cette video.

## Transcription chronologique

### Phase 1 -- Navigation et ouverture des fichiers (0s -> 15s)

1. (0-12s) Navigation dans l'explorateur Windows, arborescence reseau
   `D:\Drive\Comptabilité La Ruche\$Facturation automatique\1 - Factures transporteurs + calculs\2025\2025 05`.
   Liste des dossiers par transporteur (Chronopost, Colissimo, DPD, Fedex, GLS, Kuehne, Mondial Relay,
   TNT, UPS, UPS - 80X7Y5) et des classeurs `2025_05_Facture <Transporteur>.xlsx`, dont
   `2025_05_Facture UPS.xlsx` (53 485 Ko) et `2025_05_Facture UPS_Admosphere.xlsx` (43 611 Ko -- probable
   second compte/filiale UPS "Admosphere").

2. (~13-14s) Deux fenetres Excel apparaissent en apercu dans la barre des taches : `2025_05_Facture
   UPS.xlsx` et `2025 05 - Export expédition...` (fichier d'export ERP).

### Phase 2 -- Onglet "Fichier import" du classeur de calcul UPS (15s -> 28s)

3. (~15s) Bascule vers `2025_05_Facture UPS.xlsx`, onglet actif **"Fichier import"** (onglets visibles en
   bas : Gazole, CODES SVCE LEVEL, CODIFICATION CODE EXCEPTION, Charge.CHG_CODE, ST SV, zone colis poids
   assurance, Bilan factures, Facture UPS, Demande avoir, TCD, Comptes UPS, Fichier import, Bilan clients).
   Colonnes visibles en fin de tableau (lignes ~276-315) :
   - **A = "E/P ERP"** : declaration interne (`particulier` / `entreprise` / `inconnu`).
   - **B = "Zone"**, **C = "Transport"** (`UPS_COD` ou `UPS`).
   - **D = "Date validite"**, E = "Ref.1", F = "Ref.2", G = "Id cli...".
   - **H = "N° Tracking"** (ex. `1ZWV57887933811137`) -- alimentee par formule `=TCD!E<n>` (voir pt.5).
   - **I** : colonne fine, nom coupe (non lisible).
   - **J = "E"** (entete tronque) : contient les valeurs **P / E** -- colonne du **statut E/P retenu**
     (voir formule detaillee section suivante).
   - K = "Pays" (FR), L = "Zone" (France), M = "Nbr Co...", N = "Poi...", O = "mode en...",
     P = "TVA" (valeurs 0,2 = taux 20%), Q = "Droits et ta...", R = "Assura...", S = "Zones eloign...",
     T = "Colis volumin...", U = "Adres...", V = "Frêt" (montant en €), **W = "plus-value E..."**
     (= plus-value Paperless).
   - Notes rouges en colonnes X-AB, en marge droite du tableau.

4. (~16s, cellule A10283 selectionnee) Barre de formule affiche la valeur texte brute `entreprise` (pas
   une formule) -- confirme que la colonne A "E/P ERP" est une valeur figee/collee ligne par ligne, pas
   calculee automatiquement.

5. (~17-24s) Retour en haut du tableau (lignes 1-44). Notes rouges lisibles en colonnes X-AB :
   - Ligne 1 : "Si nb colis = 0 -> voir montant plus-value BtoC p..." (coupe)
   - Ligne 2 : "Si 1,38 € -> 1 colis, sinon arrondi 20 kg / colis"
   - Ligne 3 : "Comparer avec exports expedies depuis ERP"
   - Ligne 5 : "ATTENTION modes ST / SV"
   - Ligne 7 : "ATTENTION zones 505, 506, etc"
   - **Ligne 9 (grande police, encadre jaune) : "E / P : facturer en P si Plus-value Paperless -..."**
     (fin de phrase non entierement visible sur cette frame, mais confirmee/completee par la formule
     lue en phase 3 : la regle = plus-value Paperless presente -> facturer en P).

   Sur les lignes 2-44, colonne A contient de nombreuses valeurs **"inconnu"** (lignes 2, 3, 5, 8, 11,
   28, 36...) en plus de particulier/entreprise. Colonne J est toujours renseignee (P ou E), y compris
   quand A = "inconnu".

   Formules colonne H (N° Tracking) confirmees :
   - **H7 = `=TCD!E8`**
   - **H15 = `=TCD!E16`**
   -> la colonne H de "Fichier import" est entierement alimentee par reference a l'onglet **TCD**
   (Tableau Croise Dynamique), avec un decalage constant de 1 ligne (ligne Excel N -> ligne TCD N+1).

### Phase 3 -- Export ERP source (28s -> 44s)

6. (~28-29s) Bascule dans l'explorateur vers `D:\Drive\Laruche\$Informatique\Exports 2025\2025 05\` --
   liste des exports mensuels : "2025 05 - Export expéditions.xlsx", **"2025 05 - Export
   expéditions_brut.xlsx"** (13 544 Ko, selectionne), "2025 05 - Export préparations.xlsx", "2025 05 -
   Export stock ERP.xlsx", etc.

7. (~30-44s) Ouverture de **`2025 05 - Export expéditions_brut.xlsx`**, feuille
   "exportDemandeExpedition_202 (2)". Export source ERP brut (avant tout traitement transporteur),
   colonnes cles :
   - A = "CODE_EXPE" (ex. `EXP20250501-2155017`)
   - B = "DATE_EX...", C = "DATE_EN...", D = "CLI..." (ID client numerique)
   - E = "CLIENT_DENOM" (nom/raison sociale du destinataire)
   - **F = "EXP_SOC_"** : colonne cle, valeurs **"particulier"** / **"entreprise"** pour (quasi) toutes
     les lignes visibles -- **declaration interne du statut client** faite dans l'ERP La Ruche au moment
     de l'expedition.
   - G a Z : adresse (ville, code postal, pays), reference, poids, mode d'envoi ("standard", "enlevement
     client", "B2B - Livraison", "DOM - Livraison", "palette"...), etc.
   - **AP = "PRO_TRACKING"** (colonne surlignee en bleu, valeur 0 sur les lignes visibles, format Nombre)
   - AQ = "ETAT_EX..." (ex. "En attente d...", "Commande P...")
   - **AR = "TRANSPORTEUR"** : UPS, TNT, DPD-HYDRATIS, KUEHNE, COLISSIMO-HYDRATIS...
   - AS = "PRECOLISAGE" (ex. "EMB_EXT_A12", "EMB_EXT_DC_OMA_780x580x500"...)
   - Une colonne "DES_SOC..." (~colonne Q, visible en Q2/Q3) contient aussi "particulier" -- semble
     dupliquer/reprendre l'info a un autre niveau (destinataire vs expediteur ?), non confirme.

   -> Cet export "brut" est la **source de verite ERP** pour la declaration E/P (colonne EXP_SOC_).

### Phase 4 -- Formule de resolution E/P (45s -> 90,7s, fin de la video)

8. (~45s -> fin) Retour sur `2025_05_Facture UPS.xlsx`, onglet "Fichier import". L'utilisateur clique et
   affiche successivement la formule de la colonne J sur plusieurs lignes (J14, J20, J23, J24), et de la
   colonne W sur une ligne (W23), en les laissant affichees longuement a l'ecran (probablement pour bien
   les montrer/documenter) :

   - **J14 (A14="entreprise") : `=SI(A14="particulier";"P";SI(W14="";"E";"P"))` -> resultat "E"**
   - **J20 (A20="entreprise") : meme formule -> resultat "P"**
   - **J23 (A23="particulier") : `=SI(A23="particulier";"P";SI(W23="";"E";"P"))` -> resultat "P"**
   - **J24 (A24="entreprise") : `=SI(A24="particulier";"P";SI(W24="";"E";"P"))` -> resultat "P"**
   - **W23 : `=SI(TCD!K24="";"";TCD!K24)`** -- confirme que la colonne W ("plus-value E/P" / Paperless)
     est elle-meme une formule qui va chercher dans le TCD (colonne K) : vide si le TCD ne renvoie rien,
     sinon la valeur du TCD.

   **Lecture caractere par caractere de la formule cle (colonne J)** :
   `=SI(A14="particulier";"P";SI(W14="";"E";"P"))`
   - `SI(` ouverture de la fonction SI
   - `A14="particulier"` : test -- la cellule A14 (colonne "E/P ERP") est-elle exactement le texte
     "particulier" ?
   - `;"P"` : si vrai, la formule renvoie le texte "P"
   - `;SI(W14="";"E";"P"))` : sinon, on evalue une seconde condition -- la cellule W14 (colonne
     "plus-value E/P" / Paperless) est-elle une chaine vide ? Si oui -> "E". Si non (W14 contient une
     valeur) -> "P". Fermeture des deux parentheses de SI imbriquees.

   **Consequence metier / regle de resolution E/P pour UPS :**
   1. La declaration ERP interne ("particulier") est **prioritaire et absolue** : des qu'un client est
      marque "particulier" dans l'ERP, le statut retenu (colonne J) est **P**, quoi qu'il arrive sur la
      colonne W.
   2. Pour tout le reste (ERP = "entreprise" ou "inconnu"), c'est la **presence d'une plus-value
      Paperless renvoyee par UPS via le TCD** (colonne W non vide) qui **decide** : plus-value presente
      -> P (l'expedition a ete traitee/facturee comme "entreprise" par UPS, ce qui declenche cette
      plus-value -- indice indirect que UPS a vu un compte "entreprise" plutot qu'un colis
      particulier standard) ; plus-value absente -> E.

   -> **Observation cle non expliquee par le simple texte de la formule** : sur les lignes 20 et 24
   observees en fin de video, A="entreprise" et W apparait vide a l'ecran, pourtant J="P". Cela peut
   indiquer soit une valeur W non visible (arrondi/formatage), soit -- plus probable vu le titre de la
   video -- que **ces lignes 20 et 24 sont precisement les exemples du cas "declare particulier... /
   entreprise, mais facture differemment par UPS"** que la video veut illustrer, et qu'il existe une
   correction/situation supplementaire non entierement capturee par les frames disponibles (par exemple
   un forcage manuel de la colonne A ou W en amont, hors champ de la capture). A confirmer imperativement
   avec le pole transport avant d'implementer cette regle dans le carrier Node.js.

## Points ambigus / a confirmer par le pole transport

1. **Point le plus important** : sur les lignes 20 et 24 (A="entreprise"), la colonne W (plus-value
   Paperless) apparait vide dans la capture mais le resultat J affiche est "P" et non "E" comme la
   formule le suggererait litteralement. A verifier en ouvrant le fichier reel : soit W20/W24 contiennent
   une valeur non visible sur la capture (a confirmer par une capture d'ecran zoomee ou l'ouverture du
   fichier), soit il y a une autre couche de logique (correction manuelle, colonne masquee, calcul du TCD
   different de ce qui est affiche) qui explique cet ecart. **Ce point pourrait etre exactement le "cas
   particulier" que la video cherche a documenter** (expedition declaree en particulier -- ou en tout cas
   pas explicitement "entreprise" avec plus-value -- mais facturee en entreprise par UPS), et merite une
   clarification directe du pole transport avant tout codage.
2. **Sens exact des codes "P" et "E" en sortie de la formule J.** Colonnes A (E/P ERP) et F/EXP_SOC_
   (export brut) utilisent les mots complets "particulier"/"entreprise". La colonne J n'affiche que des
   lettres seules "P"/"E". A confirmer que P = particulier et E = entreprise dans le meme sens que la
   colonne A (tres probable au vu du contexte, mais pas litteralement demontre par du texte complet dans
   la video).
3. Le texte complet de la note rouge en cellule ~AB9 ("E / P : facturer en P si Plus-value Paperless
   -...") n'a jamais ete lu integralement/non coupe sur une frame -- la fin de phrase reste a confirmer en
   ouvrant le fichier reel (probablement quelque chose comme "... est renseignee/positive/non nulle").
4. Colonne "DES_SOC..." (~colonne Q) de l'export brut, vue avec valeur "particulier" : semble distincte
   de F/EXP_SOC_ (meme type de valeur mais colonne differente) -- role exact (destinataire vs expediteur,
   ou doublon) non confirme.
5. Le contenu exact de la colonne I (entre H "N° Tracking" et J) n'a jamais ete lisible (colonne tres
   fine, nom tronque) -- pourrait contenir un numero de colis ou un indicateur complementaire lie au
   calcul E/P.
6. Aucune formule ni valeur de l'onglet "TCD" lui-meme (colonne E pour le tracking, colonne K pour la
   plus-value Paperless) n'a ete visible dans cette video -- uniquement les references depuis "Fichier
   import". Le detail du calcul dans l'onglet TCD (d'ou vient la plus-value Paperless : montant, %, seuil
   de declenchement chez UPS) reste a documenter separement, idealement avec une autre video montrant
   l'onglet TCD lui-meme.
