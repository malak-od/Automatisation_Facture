# Transcription vidéo -- Lettres (Suivie / Suivie Prépa / Timbre Allemagne SLAACE)

Vidéo source : `Transporteurs/Lettres/Process Facturation - Facture Lettre suivie -
Lettre timbre_1.mp4`
Durée : 4 min 17 s (257,4 s)
Résolution capturée : 1906x1026 (bureau Windows, Excel + navigateur "La Ruche
Logistique")

Méthode : extraction 1 frame/4s + détection de changement de scène (seuil
0.15), lecture chronologique des deux jeux d'images. Pas d'audio dans la
vidéo (process silencieux) -- toute l'info vient de ce qui est visible à
l'écran (barre de formule, cellules, filtres, onglets navigateur).

Cette vidéo montre exclusivement la correction manuelle du fichier WMS brut
"Export expéditions_brut.xlsx" (onglet `exportDemandeExpedition_202 (2)`) :
identification et remplacement des `PRO_TRACKING` invalides par le
`CODE_EXPE`, avec vérification croisée dans le WMS "La Ruche" (interface web
`si.laruche-logistique.fr`, à priori l'ERP/WMS métier, pas un export
distinct). Elle ne montre PAS encore la construction des fichiers d'import
finaux (voir vidéo 2).

## Colonnes du fichier brut identifiées (onglet `exportDemandeExpedition_202 (2)`)

Ligne d'en-tête observée, de gauche à droite (lettres de colonne Excel
entre parenthèses) :
- `CODE_EXPE` (A)
- `DATE_EX...` (B) -- date d'expédition
- `DATE_EN...` (C)
- `CLI...` (D) -- code client
- `CLIENT_DENOM...` (E) -- dénomination client
- `EXP_SOC...` (F)
- `EX...` (G)
- colonnes intermédiaires H à Z : adresse destinataire (DES_SOC, DES_CP =
  code postal, DES_VILL = ville, DI..., etc.)
- `INFO_CO...`, `INFO_NB...`, `INFO_...` (AH-AI) -- dont une colonne poids
  (`INFO...`, ex. valeurs 0,02 / 0,05 / 1 kg)
- `GEN_MO...` (AM/AN zone) -- mode d'expédition ("Lettre suivi...")
- **`PRO_TRACKING`** (colonne AP) -- numéro de tracking transporteur, **en
  fond bleu clair/surligné** dans les captures (colonne mise en évidence
  visuellement par une sélection/surbrillance bleu clair pendant toute la
  vidéo, cohérent avec la consigne utilisateur "colonne bleue")
- **`ETAT_EX...`** (AQ) -- état de l'expédition (valeurs vues : "Livré",
  "En cours de P...", "Commande P...", "En attente ...")
- **`TRANSPORTEUR`** (AS) -- colonne filtrée. 4 valeurs "lettre" distinctes
  vues dans le filtre déroulant : **"Lettre Suivie"**,
  **"LETTRE-SUIVIE-HYDRATIS"**, **"LETTRE-SUIVIE-PREPA"**,
  **"LETTRE-TIMBRE-SLAACE"**
- `PRECOLIS...` (AT) -- ex. valeurs "EMB_EXT_E..."
- `Qté` (AU)

Nombre total de lignes du brut affiché en bas de fenêtre Excel : "56124"
enregistrements (barre de statut), filtré à plusieurs reprises sur
TRANSPORTEUR pour isoler les sous-groupes.

## Chronologie détaillée

### 0:00 -- 0:12 (frames 1-3)
Excel ouvert sur `2025 05 - Export expéditions_brut.xlsx`, onglet
`exportDemandeExpedition_202 (2)`. Vue sur les colonnes du brut avec filtre
actif sur `TRANSPORTEUR`.

### 0:24 -- 0:44 (frames 6-11, scene 1-2)
Le filtre `TRANSPORTEUR` est ouvert : liste déroulante des valeurs
disponibles, incluant les 4 valeurs "lettre" citées ci-dessus. Zoom sur la
colonne `PRO_TRACKING` avec une plage de cellules sélectionnée (fond
bleu clair) -- confirme que la colonne PRO_TRACKING est visuellement mise en
évidence en bleu pendant l'inspection, mais il ne s'agit pas d'une couleur
de remplissage permanente enregistrée dans le fichier : c'est la
surbrillance de sélection Excel (colonne cliquée/sélectionnée). **Aucune
couleur de police/fond "bleu foncé" distincte des "bleu clair" n'a été
observée à ce stade -- à confirmer plus loin.**

### 0:48 -- 1:24 (frames 12-21, scene 3-4)
Bascule vers le navigateur : interface web **"La Ruche Logistique"**
(`si.laruche-logistique.fr` ou domaine similaire), écran de recherche
d'expédition. Un `CODE_EXPE` est saisi/recherché (ex. `EXP20250519-2184563`)
directement dans cette interface -- **pas un export distinct, mais une
consultation live du WMS/ERP**. Le détail de l'expédition s'affiche avec :
- Numéro d'expédition
- **Code statut transporteur : "101"**, avec message explicite du type
  "numéro de tracking non valide" (texte partiellement lisible)
- Un bouton/action pour supprimer ou vider le tracking associé

### 1:24 -- 1:44 (frames 19-26, scene 5)
Retour dans Excel. Cellule `AP7690` sélectionnée, barre de formule affiche :
**`=A7690`** -- c'est-à-dire une formule qui recopie directement la valeur de
la colonne `CODE_EXPE` (colonne A) de la même ligne dans la colonne
`PRO_TRACKING`. Résultat affiché dans la cellule : `EXP20250519-2184563`.
Ceci confirme le mécanisme "tracking invalide -> remplacé par CODE_EXPE".

Le tracking d'origine, avant remplacement, était **`3760304101977`** --
un tracking partagé/dupliqué entre au moins 2 lignes (voir étape suivante),
d'où son invalidité signalée par La Ruche (code 101).

### 1:44 -- 2:08 (frames 22-27, scene 6)
Répétition du même mécanisme pour la ligne suivante : `CODE_EXPE` =
`EXP20250520-2186965`, même tracking d'origine `3760304101977` détecté comme
dupliqué. Formule `=A<ligne>` collée dans `PRO_TRACKING`. Donc **le
tracking "invalide" traité ici est un cas de duplication d'un même numéro
sur 2 expéditions différentes** (pas un tracking vide au sens "cellule
vide", ni un "0").

### 2:08 -- 2:32 (frames 28-30, scene 7)
Excel : la colonne `TRANSPORTEUR` d'une ligne isolée est retapée
manuellement de "Colissimo" à **"Lettre Suivie"** (ligne ~52701) -- indique
qu'en plus des corrections de tracking, certaines lignes ont aussi leur
`TRANSPORTEUR` corrigé manuellement quand il est mal renseigné par le WMS.
Point isolé, non retrouvé ailleurs dans les 2 vidéos -- **à confirmer si
récurrent ou anecdotique**.

### 2:32 -- 3:04 (frames 31-38, scene 8-9)
Filtre `TRANSPORTEUR` réappliqué, confirmation des 4 valeurs "lettre"
(Lettre Suivie / LETTRE-SUIVIE-HYDRATIS / LETTRE-SUIVIE-PREPA /
LETTRE-TIMBRE-SLAACE) dans le menu déroulant du filtre. Défilement dans la
colonne `PRO_TRACKING` : plusieurs lignes affichent des valeurs au format
**`BAC****`** (préfixe "BAC" suivi de chiffres, PAS "BAC25" précisément) :
`BAC1018`, `BAC1037`, `BAC1099`, `BAC1171`, `BAC316`, `BAC418`, `BAC677`,
`BAC711`, `BAC753`, `BAC870`, `BAC884`, `BAC886`, `BAC903`, `BAC989` (liste
non exhaustive, valeurs lues sur les captures disponibles). Toutes ces
lignes sont sur le groupe `LETTRE-SUIVIE-HYDRATIS`.

**Le préfixe exact observé est "BAC" (3 lettres), pas "BAC25". Le nombre de
chiffres qui suit varie (3 à 4 chiffres). Il n'y a pas eu, dans cette
vidéo, de confirmation visuelle explicite d'un préfixe "BAC25" précis --
seulement "BAC" + nombre. A faire confirmer par le pôle transport si le
préfixe attendu dans le code doit être "BAC" (générique) ou strictement
"BAC25".**

### 3:04 -- 3:28 (frames 39-45, scene 10-11)
Sélection de toute la plage de cellules `PRO_TRACKING` correspondant aux
lignes "BAC***" repérées, et collage en masse de la formule `=CODE_EXPE`
(équivalent `=A<ligne>`) sur toute la plage. Résultat : chaque tracking
"BAC***" est remplacé par le `CODE_EXPE` correspondant à sa ligne (ex.
`AP52789` -> `EXP20250501-2154905`, `AP52790` -> `EXP20250521-2189059`,
etc. -- correspondance ligne à ligne, pas une valeur unique collée
partout).

Bascule navigateur "La Ruche" : consultation du détail d'une expédition avec
mention **"Livraison complète sans anomalie de poids"**, état "Livré" pour
`EXP20250520-2186236` (rattachée au tracking dupliqué `7743814314641836`).
Une autre expédition affiche **"Etat indéterminé"**.

### 3:28 -- 3:48 (frames 46-53, scene 12-14)
Nouveau cas de duplication de tracking détaillé dans La Ruche : deux lignes
`AP20008` et `AP20009` partagent exactement le même tracking dupliqué
(`7743814314641830` / variantes `...36`, `...38` selon les captures -- un
même préfixe de 13 chiffres avec les 2 derniers chiffres qui varient d'une
ligne à l'autre, donc bien 2 numéros distincts mais tous deux détectés comme
non valides / code statut "101"). Les deux lignes reçoivent la même logique
`=CODE_EXPE`, donnant `EXP20250520-2186236` et une valeur similaire.

Dans La Ruche, actions vues : suppression du tracking invalide dans la
fiche expédition (bouton "supprimer"/croix sur le champ tracking),
consultation du code statut transporteur associé.

### 3:48 -- 4:04 (frames 54-59, scene 15)
Poursuite de la même boucle : recherche `CODE_EXPE` dans La Ruche,
suppression tracking invalide (ex. `7743814314641838`, code statut "101"),
retour Excel, formule `=A<ligne>` collée en `PRO_TRACKING`.

### 4:04 -- 4:17 (frames 60-64, scene 16-22, fin de vidéo)
Suite et fin du traitement en boucle des lignes `PRO_TRACKING` = "BAC***"
(`BAC1018`, `BAC1037`...) : recherche dans La Ruche, **code statut
transporteur "400"** affiché avec message du type "le numéro que vous avez
saisi n'est pas valide (nombre de caractères incorrect)" -- confirme que le
format "BAC****" est un tracking **structurellement invalide** (trop court/
mauvais format pour être un vrai tracking), différent du cas "dupliqué"
(code 101) vu plus tôt. La vidéo se termine sur cette boucle de vérification
ligne par ligne dans La Ruche, sans montrer la suite (construction des
fichiers d'import), qui est couverte par la vidéo 2.

## Points clés relevés (transporteur "Lettres") -- vidéo 1

1. **Colonne "bleue" citée par l'utilisateur** : correspond très
   vraisemblablement à la colonne `PRO_TRACKING`, qui est systématiquement
   sélectionnée/surlignée en bleu clair (bleu de sélection Excel) pendant
   toute la vidéo lors de son inspection -- mais aucune preuve d'un
   remplissage de cellule permanent en bleu clair n'a été observée dans
   cette vidéo 1. Voir vidéo 2 pour un éventuel bleu foncé de mise à jour.
2. **Deux mécanismes distincts de tracking invalide identifiés dans cette
   vidéo** :
   - **Tracking dupliqué** (le même numéro de tracking apparaît sur au
     moins 2 expéditions différentes) -> code statut transporteur **"101"**
     dans La Ruche, message "numéro non valide".
   - **Tracking mal formé / préfixe "BAC" + chiffres** (ex. BAC1018,
     BAC316...) -> code statut transporteur **"400"** dans La Ruche,
     message "nombre de caractères non valide".
   Dans les deux cas, le traitement Excel appliqué est identique : la
   formule `=CODE_EXPE` (`=A<ligne>`) est collée dans `PRO_TRACKING` pour
   remplacer la valeur invalide par le numéro d'expédition.
   **Aucun cas de tracking "vide" (cellule réellement vide) ni de tracking
   "= 0" (valeur zéro) au sens strict n'a été observé dans cette vidéo 1** --
   voir precisions utilisateur n°1/2, à investiguer en vidéo 2 ou à faire
   confirmer, cette vidéo ne montrant que les cas "dupliqué" et "BAC***".
3. **Pas de mention ni de préfixe exact "BAC25"** trouvé dans les captures
   lisibles -- uniquement "BAC" + 3-4 chiffres. A faire confirmer le
   préfixe exact attendu par le pôle transport (le code actuel n'implémente
   pas ce cas du tout, pour rappel).
4. **Aucune mention ni action visible concernant des expéditions "en cours
   de préparation"** dans cette vidéo 1 -- la colonne `ETAT_EX...` affiche
   parfois "En cours de P..." dans les captures de fond (hors filtre actif
   sur les cas problématiques), mais aucune action de filtrage/exclusion
   explicite n'est visible dans cette vidéo. Voir vidéo 2.
5. **Pas d'export ERP distinct identifié** : l'outil "La Ruche" utilisé
   pour vérifier/corriger les trackings est une interface web consultée en
   direct (recherche + fiche détail expédition), pas un fichier d'export
   séparé. Aucune mention ni capture d'un second export "ERP" distinct
   du fichier "Export expéditions_brut.xlsx" n'apparaît dans cette vidéo.
6. **Correction de la colonne TRANSPORTEUR observée une fois** (Colissimo
   -> Lettre Suivie sur une ligne isolée) -- mécanisme distinct du
   tracking, à confirmer si récurrent.

## Points ambigus / illisibles à faire confirmer par le pôle transport

- Le préfixe exact du cas "erreur atelier" : "BAC" générique ou "BAC25"
  strictement (précision utilisateur n°4 non confirmée textuellement dans
  cette vidéo -- seuls "BAC1018", "BAC1037", "BAC316", etc. ont été lus).
- Aucune trace visible dans cette vidéo 1 des cas "tracking vide" (cellule
  vide) et "tracking à zéro" (valeur "0") distincts l'un de l'autre -- la
  vidéo ne montre que "tracking dupliqué" (code 101) et "tracking mal
  formé BAC***" (code 400). Il est possible que ces cas apparaissent dans
  la vidéo 2, ou que "vide"/"zéro" fassent référence aux mêmes mécanismes
  observés ici sous une terminologie différente (à faire confirmer).
- Aucune mention ni capture claire d'une "couleur bleu foncé" appliquée
  après correction, ni d'une mise à jour explicite "dans l'ERP" du
  tracking corrigé (l'action vue dans La Ruche est une **suppression** du
  tracking invalide, pas une saisie du nouveau numéro de tracking dans
  cette interface) -- à confirmer si une saisie manuelle du nouveau
  tracking a lieu hors champ de la capture, ou si "mise à jour ERP" désigne
  uniquement cette suppression.
- La correction ponctuelle de la colonne TRANSPORTEUR (Colissimo -> Lettre
  Suivie) : mécanisme, fréquence et règle de détection non documentés
  dans cette vidéo -- semble être une correction manuelle au cas par cas,
  pas une règle automatisable identifiée.
- Certaines captures (notamment les frames de type "scene" autour de
  1min-1min30 et 3min30) sont partiellement floues sur les codes statut
  exacts de La Ruche (texte petit, capture en transition) -- les codes
  "101" et "400" ont pu être lus avec confiance sur plusieurs occurrences
  mais le texte descriptif complet du message d'erreur n'est pas
  intégralement lisible.
