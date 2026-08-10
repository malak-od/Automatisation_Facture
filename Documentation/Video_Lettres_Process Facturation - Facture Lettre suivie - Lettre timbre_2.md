# Transcription vidéo -- Lettres (Suivie / Suivie Prépa / Timbre Allemagne SLAACE)

Vidéo source : `Transporteurs/Lettres/Process Facturation - Facture Lettre suivie -
Lettre timbre_2.mp4`
Durée : 4 min 45 s (284,97 s)
Résolution capturée : 1906x1026 puis 1920x1080 selon les fenêtres (bureau
Windows, Excel + Explorateur de fichiers réseau).

Méthode : extraction 1 frame/4s + détection de changement de scène (seuil
0.15), lecture chronologique. Pas d'audio (process silencieux).

Cette vidéo montre la suite logique de la vidéo 1 : construction des **3**
fichiers d'import CSV finaux à partir du fichier brut WMS corrigé
(`2025 05 - Export expéditions_brut.xlsx`), pour le mois de mai 2025,
groupe par groupe (Lettre Suivie -> Lettre Suivie Prépa -> Lettre Timbre
Allemagne SLAACE), avec enregistrement de chaque fichier CSV dans
l'arborescence réseau partagée.

## Arborescence réseau observée

Chemin complet vu dans l'Explorateur de fichiers / boîtes "Enregistrer
sous" : `Réseau > 192.168.5.3 > Comptabilité La Ruche > $Facturation
automatique > 2 - Fichiers csv import > 2025 > 2025 05` (et `2025 04` pour
consultation/comparaison).

Autres dossiers frères visibles dans `2025 05` (sous `2 - Fichiers csv
import`) : Chronopost, Colissimo, DPD, Fedex, GLS, Kuehne, Mondial Relay,
TNT, UPS, UPS - 80X7Y5. **Aucun dossier "Lettres" séparé** : les fichiers
CSV Lettres sont enregistrés directement à plat dans le dossier du mois
(`2025 05`), comme les autres transporteurs à fichier unique.

## Fichiers de sortie -- confirmation du nombre exact

Le dossier `2 - Fichiers csv import\2025\2025 04` (consulté en référence,
avant recopie/adaptation pour mai) contient explicitement, parmi les
fichiers listés à l'écran :
- `2025_04_Lettre_Timbre Allemagne_SLAA...csv` (nom tronqué à l'affichage ;
  taille **1 Ko**, modifié 08/05/2025 14:40)
- `2025_04_Lettre suivie prépa_Import fichie...csv` (taille **35 Ko**,
  modifié 08/05/2025 14:37)
- `2025_04_Lettre suivie_Import fichier.csv` (taille **207 Ko**, modifié
  08/05/2025 14:28)

Soit **3 fichiers de sortie distincts confirmés**, pas 2 : "Lettre suivie"
(qui regroupe en réalité 2 valeurs `TRANSPORTEUR` du brut : "Lettre Suivie"
+ "LETTRE-SUIVIE-HYDRATIS", fusionnées dans un seul fichier de sortie),
"Lettre suivie prépa", et "Lettre Timbre Allemagne SLAACE". Le dossier
`2025 05` en cours de constitution montre la même liste de 3 fichiers une
fois complétée (`2025_05_Lettre suivie prépa_Import fichie...csv` 48 Ko,
`2025_05_Lettre suivie_Import fichier.csv` 252 Ko, plus SLAACE traité en
fin de vidéo).

**Confirmation du point 5 de la demande** : le filtre en sortie produit
bien 2 "familles" au sens onglet/fichier visibles pour Suivie+Hydratis (un
seul fichier, deux valeurs `TRANSPORTEUR` sources) et Prépa d'un côté,
mais avec SLAACE cela fait bien **3 fichiers CSV distincts au total**, pas
2.

## Chronologie détaillée

### 0:00 -- 0:16 (frames 1-4)
Excel ouvert sur le brut `2025 05 - Export expéditions_brut.xlsx`, filtré
sur `TRANSPORTEUR` = `LETTRE-TIMBRE-SLAACE` en fin de liste (queue du
fichier, lignes ~6097-6103 sur 56124 au total). Colonnes visibles :
`PRO_TRACKING`, `ETAT_EX...` (valeurs "En attente d..."), `TRANSPORTEUR`
= LETTRE-TIMBRE-SLAACE, `PRECOLIS` = EMB_EXT_E..., `Qté` = 1.

### 0:16 -- 0:32 (frames 5-9, scene 1-2)
Réapplication du filtre `TRANSPORTEUR`, saisie "let" dans la zone de
recherche du filtre pour isoler les 4 valeurs "lettre" (Lettre Suivie,
LETTRE-SUIVIE-HYDRATIS, LETTRE-SUIVIE-PREPA, LETTRE-TIMBRE-SLAACE) --
coché/décoché successivement pour isoler chaque groupe.

### 0:32 -- 0:52 (frames 10-13, scene 3)
Ouverture de l'Explorateur de fichiers réseau, navigation jusqu'à
`\\192.168.5.3\Comptabilité La Ruche\$Facturation automatique\2 - Fichiers
csv import\2025\2025 04`, listant les fichiers CSV import d'avril 2025 pour
tous les transporteurs (Chronopost, Colissimo, DPD, Fedex, GLS, Kuehne,
Mondial Relay, TNT, UPS...). Les 3 fichiers Lettre y sont visibles (voir
section ci-dessus).

### 0:52 -- 1:28 (frames 14-22, scene 4-5)
Boîte de dialogue **"Enregistrer sous"** : le fichier
`2025_04_Lettre suivie_Import fichier.csv` est rouvert/dupliqué pour être
réenregistré sous le nom `2025_05_Lettre suivie_Import fichier.csv` dans
le dossier `2025 05` nouvellement créé. Type de fichier : "CSV
(séparateur : point-virgule) (*.csv)".

### 1:28 -- 1:44 (frames 23-25, scene 6)
Le fichier `2025_05_Lettre suivie_Import fichier.csv` est ouvert dans
Excel. En-têtes de colonnes confirmées dans l'ordre exact :
`Transporteur | Date validité | Réf.1 | Réf.2 | Id client | N° Tracking |
Nom | E/P | Pays | Zone | Nbr Colis | Poids | mode envoi | TVA | Droits et
taxes | Assurance | Zones éloignées | Colis volumineux | Adresses | Frêt |
plus-value B2C | gazole | Nb colis`

Valeurs observées sur les premières lignes copiées depuis le mois d'avril
(pour référence de structure, avant remplacement par les données de mai) :
Transporteur = "LETTRE-SUIV..." (LETTRE-SUIVIE), mode envoi = "LS", TVA =
0,2, Frêt = **3,83** (collé/répété sur toute la colonne).

### 1:44 -- 2:16 (frames 26-30, scene 7-8)
Retour sur le fichier brut. Collage des colonnes utiles (Tracking, Pays,
Zone, Nbr Colis, Poids...) depuis le brut filtré `LETTRE-SUIVIE` +
`LETTRE-SUIVIE-HYDRATIS` vers le fichier d'import. Une ligne isolée
affiche en `N° Tracking` la valeur **`94598546`** -- un tracking numérique
de 8 chiffres, sans le format habituel "2L0..." ni le préfixe "EXP..." des
corrections CODE_EXPE vues en vidéo 1. **Ce pourrait être le cas "tracking
à zéro" évoqué par l'utilisateur (numéro trop court / non conforme mais
non vide), mais aucune action de correction n'est visible sur cette ligne
précise dans les frames disponibles -- valeur simplement recopiée telle
quelle.** D'autres lignes affichent les valeurs `EXP20250519-2184563` et
`EXP20250520-2186965` -- correspondant exactement aux corrections
CODE_EXPE observées en vidéo 1 (cas trackings dupliqués).

### 2:16 -- 2:44 (frames 31-34, scene 9)
Poursuite du collage des colonnes (Poids en valeurs, Frêt = 3,83 recopié
sur toute la colonne). Le fichier `2025_05_Lettre suivie_Import fichier.csv`
atteint sa taille finale de **3067 lignes** de données (observée via la
barre de nom `AI3557` / défilement, et confirmée par la référence de
1re vidéo évoquant ~3683 lignes pour juin -- ordre de grandeur cohérent
pour un fichier mensuel).

### 2:44 -- 3:16 (frames 35-39, scene 10-12)
Retour au fichier brut, nouveau filtre `TRANSPORTEUR` (recherche "let"),
isolement du groupe **`LETTRE-SUIVIE-PREPA`** cette fois (549
enregistrements trouvés sur 56124 au total, affiché en barre de statut
Excel). Colonne `PRO_TRACKING` sélectionnée pour copie.

**Colonne `ETAT_EX...` bien visible sur ce groupe avec plusieurs valeurs
distinctes** : "Livré", **"En cours de P..."** (très probablement "En
cours de Préparation"), "Commande P..." (probablement "Commande Préparée"
ou "Commande en Préparation"). Aucune action de filtre supplémentaire sur
cette colonne `ETAT_EX...` n'est visible dans les captures disponibles --
**les lignes semblent toutes copiées indifféremment de leur `ETAT_EX...`
dans cette portion de la vidéo**, ce qui contredirait la consigne
utilisateur n°3 ("expés en cours de préparation à ne pas facturer") sauf
si un filtrage a lieu hors-champ (scroll/action rapide non capturée) ou en
amont (dans La Ruche, vidéo 1) plutôt que visuellement dans ce fichier
Excel. **Point à faire confirmer explicitement par le pôle transport.**

### 3:16 -- 3:44 (frames 40-45, scene 13-15)
Boîte "Enregistrer sous" pour `2025_04_Lettre suivie prépa_Import
fichier.csv` -> `2025_05_Lettre suivie prépa_Import fichier.csv`. Fichier
ouvert, en-têtes identiques au fichier Suivie, avec **Frêt = 0,01€** collé
sur toute la colonne "Frêt". `mode envoi` = "LS", TVA = 0,2, Zone = France
(implicite, pas de colonne Zone visible distincte -- "Zone" reste vide ou
héritée). Les valeurs `N° Tracking` collées sont au format `2L0...`
standard (pas de préfixe EXP ni BAC visible sur ce groupe dans les
captures disponibles).

### 3:44 -- 4:04 (frames 46-51, scene 16-17)
Poursuite du remplissage du fichier "Lettre suivie prépa". Le fichier
final `2025_05_Lettre suivie prépa_Import fichier.csv` atteint **549
lignes de données** (confirmé par la barre de statut Excel : "549
enregistrement(s) trouvé(s) sur 56124" côté brut, et le nombre de lignes
non vides "Nb (non vides) : 549" sur la colonne Poids côté fichier
import). Frêt total = 549 x 0,01 = 5,38€ (somme affichée : "Somme : 5,38"
en bas de la sélection de la colonne Frêt).

### 4:04 -- 4:20 (frames 52-56, scene 18-20)
Retour au fichier brut, dernier filtre du groupe : `LETTRE-TIMBRE-SLAACE`
(6-7 lignes en toute fin de fichier, lignes ~6097-6103). Colonne
`PRO_TRACKING` copiée : valeurs longues (13 chiffres, préfixe "870009...").
Colonne `ETAT_EX...` = **"En attente d(e préparation)"** pour toutes ces
lignes SLAACE (texte tronqué mais lisible "En attente ...P" cohérent avec
"En attente de Préparation" ou "En attente d'expédition").

### 4:20 -- 4:36 (frames 57-63, scene 21-24)
Boîte "Enregistrer sous" : `2025_04_Lettre_Timbre Allemagne_SLAACE_
Import.csv` (nom de fichier complet visible cette fois dans le champ
"Nom de fichier :" de la boîte de dialogue). Fichier ouvert :
`2025_05_Lettre_Timbre Allemagne_SLAACE_Import.csv`. En-têtes identiques
(mêmes 23 colonnes). Valeurs collées :
- `N° Tracking` : 7 valeurs, format `87000936255908` / `87000860428850`
  etc. (13-14 chiffres, préfixe "870...")
- `Pays` : **variable selon la ligne** -- valeurs observées : `DE`, `PL`,
  `CH`, `CH`, `BE`, `CH`, `IT` (collées depuis la colonne ville/pays du
  brut, PAS une valeur fixe)
- `Zone` : **`DE` fixe** sur toutes les lignes (contrairement à Pays qui
  varie)
- `Nbr Colis` = 1, `Poids` variable (0,02 à 0,03 kg par ligne)
- `mode envoi` = **`DOMDE`** sur toutes les lignes
- `TVA` = **0** sur toutes les lignes
- `Frêt` = **1,65** collé sur toutes les lignes

### 4:36 -- 4:45 (frames 64-71, scene 25-29, fin de vidéo)
Finalisation du fichier SLAACE : les 2 dernières lignes (F7, F8) reçoivent
leurs valeurs Pays (`CH`, `IT`), Zone (`DE`), Nbr Colis (1), mode envoi
(`DOMDE`), TVA (0), Frêt (1,65) -- une frappe manuelle de la colonne
`Transporteur` est visible transitoirement affichant "LETTRE-TIMBRE-
ALLEMAGNE" pour ces 2 lignes (avant harmonisation avec le reste, probablement
"LETTRE-TIMBRE-SLAACE" au format final -- **valeur exacte du champ
Transporteur pour ce groupe non confirmée avec certitude sur la toute
dernière frame, la colonne A affiche "LETTRE-TIM..." tronqué partout
ailleurs**). Colonne `Date validité` = 01/05/2025 harmonisée sur toutes
les 7 lignes en toute fin de vidéo. Fichier final SLAACE = **7 lignes de
données**. La vidéo se termine sur cet état (fichier complet, non encore
vu enregistré/fermé).

## Points clés relevés (transporteur "Lettres") -- vidéo 2

1. **3 fichiers de sortie confirmés** (pas 2) : `Lettre suivie_Import
   fichier.csv` (regroupe TRANSPORTEUR="Lettre Suivie" +
   "LETTRE-SUIVIE-HYDRATIS"), `Lettre suivie prépa_Import fichier.csv`,
   `Lettre_Timbre Allemagne_SLAACE_Import.csv`. Nomenclature de fichier
   confirmée : `AAAA_MM_Lettre suivie_Import fichier.csv`,
   `AAAA_MM_Lettre suivie prépa_Import fichier.csv`,
   `AAAA_MM_Lettre_Timbre Allemagne_SLAACE_Import.csv`.
2. **Tarifs Frêt confirmés à l'identique du code actuel** :
   - Lettre Suivie (+ Hydratis) : **3,83 €** par ligne
   - Lettre Suivie Prépa : **0,01 €** par ligne
   - Lettre Timbre Allemagne SLAACE : **1,65 €** par ligne
   Aucun écart constaté avec `config.json` (`fret: 3.83 / 0.01 / 1.65`).
3. **SLAACE : le champ `Pays` varie par ligne** (DE, PL, CH, BE, IT
   observés) alors que `Zone` reste fixe à "DE" et `mode envoi` = "DOMDE".
   Ceci est cohérent avec le code actuel (`Pays: String(r[c.pays])` lit la
   valeur réelle de la ligne brute, `Zone: g.zone` = "DE" fixe) -- **pas
   d'écart identifié sur ce point**, juste confirmation visuelle.
4. **SLAACE = 7 lignes en mai 2025** dans cette vidéo (contre 0 ligne en
   juin 2026 selon le commentaire du config.json actuel, et 5 lignes selon
   la même note pour mai 2025 -- **léger écart de comptage, 7 vs 5,
   probablement du fait que la vidéo capture un état encore en cours de
   finalisation en toute fin ou que le commentaire du code se basait sur
   une lecture partielle -- à revérifier sur le fichier brut réel de mai
   2025 si disponible**).
5. **Cas "tracking à 8 chiffres non standard" repéré** (`94598546` sur le
   groupe Lettre Suivie) sans action de correction visible dans cette
   portion de vidéo -- possible candidat pour le cas "tracking à zéro"
   évoqué par l'utilisateur, mais non confirmé formellement (pas de
   zéro dans la valeur elle-même, juste un format court/non standard).
6. **Colonne `ETAT_EX...` avec valeurs "En cours de P..." et "Commande
   P..." visibles sur le groupe Prépa**, sans qu'un filtre d'exclusion
   explicite ne soit visible à l'écran dans les captures disponibles --
   point le plus important à faire confirmer/vérifier par le pôle
   transport (voir section ambiguïtés).
7. **Aucune "colonne bleue" ni "bleu foncé" identifiée avec certitude**
   dans cette vidéo 2 -- les seules couleurs de fond observées sont les
   surbrillances de sélection standard d'Excel (bleu clair de sélection de
   plage), pas des remplissages de cellule permanents.
8. **Aucune mention ni capture d'un export ERP distinct** dans cette
   vidéo 2 non plus -- uniquement le fichier brut WMS et les fichiers CSV
   d'import, plus l'Explorateur de fichiers réseau pour l'enregistrement.

## Points ambigus / illisibles à faire confirmer par le pôle transport

- **Exclusion des expéditions "en cours de préparation" (précision
  utilisateur n°3)** : la colonne `ETAT_EX...` du brut affiche bien des
  valeurs "En cours de P..." et "Commande P..." pour une partie des lignes
  du groupe Prépa, mais aucune action de filtre/suppression n'est
  visible dans les frames disponibles de cette vidéo avant le collage
  dans le fichier final. Il est possible que ce filtrage soit fait ailleurs
  (dans La Ruche en amont, cf vidéo 1) ou lors d'une étape non capturée par
  l'échantillonnage de frames (toutes les 4s). **A vérifier précisément
  avec le pôle transport : les 549 lignes du fichier Prépa final
  incluent-elles bien des lignes "en cours de préparation", ou ont-elles
  été exclues par un mécanisme invisible dans cette vidéo ?**
- **Cas "tracking = 0" strictement** : non observé distinctement d'un
  tracking "vide" ou "dupliqué" dans les 2 vidéos. Le seul candidat
  possible est la valeur `94598546` (8 chiffres, non standard) vue sans
  action de correction. A confirmer si un vrai "0" ou une cellule
  totalement vide apparaît ailleurs dans le fichier (hors échantillon de
  frames capturé).
- **Couleur "bleu foncé" de mise en évidence après correction** (précision
  utilisateur n°2) : non observée avec certitude dans les 2 vidéos --
  seules des sélections de plage standard Excel (bleu clair, non
  permanentes) ont été vues. Possible que cette mise en couleur soit un
  usage informel non systématiquement démontré dans les captures
  disponibles, ou qu'elle intervienne à un moment non échantillonné.
- **Mise à jour du tracking "dans l'ERP" (au-delà du fichier d'import)** :
  dans La Ruche (vidéo 1), seule une action de **suppression** du tracking
  invalide a été observée, jamais une saisie explicite du nouveau numéro
  de tracking (= CODE_EXPE) dans cette interface. A confirmer si cette
  saisie a lieu hors-champ ou si "mise à jour ERP" désigne uniquement
  l'action de suppression vue.
- **Valeur exacte de la colonne `Transporteur` pour le groupe SLAACE dans
  le fichier import final** : la colonne A affiche "LETTRE-TIM..." tronqué
  sur la quasi-totalité des lignes, avec une frappe transitoire
  "LETTRE-TIMBRE-ALLEMAGNE" visible sur 2 lignes en toute fin de vidéo (non
  confirmé si valeur finale ou saisie provisoire avant correction/
  harmonisation). Le code actuel utilise `"LETTRE-TIMBRE-SLAACE"` comme
  `nomTransporteurErp` -- **à confirmer que c'est bien la valeur finale
  retenue et non "LETTRE-TIMBRE-ALLEMAGNE"**.
- **Écart de comptage SLAACE** (7 lignes vues dans cette vidéo pour mai
  2025, vs 5 lignes mentionnées dans le commentaire du `config.json`
  actuel pour la même période) -- à revérifier sur le fichier brut réel.
- Plusieurs captures de la boîte "Enregistrer sous" et du menu filtre sont
  partiellement tronquées à l'affichage (noms de fichiers coupés par la
  largeur de colonne/fenêtre) -- les noms exacts complets de fichiers ont
  été reconstitués par recoupement entre plusieurs frames mais n'ont pas
  été lus intégralement en un seul coup d'œil sur une unique capture.
