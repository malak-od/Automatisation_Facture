# Script vidéo — GLS — "Process Facturation - Facture GLS_1.mp4"

Vidéo silencieuse, durée 478 s (~8 min). Fichier source :
`Transporteurs/GLS/Process Facturation - Facture GLS_1.mp4`.
Classeur manipulé : `2025_05_Facture GLS.xlsx`.

## Avertissement préalable (important pour la question posée)

**La vidéo ne montre jamais l'onglet "Bilan factures" activé/rempli.** Un
clic est bien effectué sur la zone des onglets vers 144 s, mais il atterrit
sur l'onglet **"TCD"** (adjacent dans la liste des onglets), pas sur "Bilan
factures" — cf. onglets visibles en bas de fenêtre tout au long de la
vidéo : `PA GLS | Zoning | Poids | Comptes GLS | Catégories | Facture GLS |
TCD | Import csv | Bilan factures | Bilan clients`.

En revanche, la vidéo montre un **mécanisme de contrôle HT/TTC très proche
de celui décrit dans la question, mais situé dans l'onglet "TCD"** (pas
"Bilan factures"), sous forme d'un petit tableau de cellules figées à droite
du TCD principal (colonnes O à S, lignes 1 à 9). C'est la meilleure source
disponible dans cette vidéo pour répondre aux 4 points — à confirmer avec le
pôle transport si "Bilan factures" reproduit exactement la même logique ou
une logique différente (formule figée vs dynamique, colonnes A-F).

## Repères chronologiques

1. (0:00–0:20) Ouverture du classeur `2025_05_Facture GLS.xlsx`, onglet
   "PA GLS" actif.
2. (~0:07–1:30) Navigation dans les onglets préparatoires : "Zoning",
   "Poids", "Comptes GLS" — consultation/vérification de colonnes de
   référentiel (zones, poids, comptes clients GLS).
3. (~0:50–1:00) Défilement rapide dans l'onglet "Facture GLS" (table
   source de ~3000+ lignes, colonnes SENEA / Transporteur / Date valid. /
   Réf.1 / Réf.2 / Id client / N° Tracking / Nom / E-P / Pays / Zone /
   Nbr Colis / Poids / mode envoi / TVA / Droits et taxes / Assurance /
   Zones éloignées / Colis volumineux / Adresse / Frêt / plus-value /
   Gazole).
4. (~1:26) Onglet "Catégories" : table de correspondance à 2 colonnes
   "Catégorie GLS" -> "Catégorie E..." (ex. EXPREDOMESTANDARD -> Frêt,
   PARCEDOMESPRTN -> Frêt - retour, SURCHALL ADM -> Frais de gestion,
   SURCHALL GEO -> Zone éloignée, SURCHALL NC -> Colis volumineux,
   SURCHALL SNGLPRCL -> ZZ_Ramasse unicolis, etc.). Cette table sert de
   base à un `RECHERCHE()` utilisé dans "Facture GLS" (colonne A,
   catégorie) : formule vue en barre de formule à 144 s —
   `=SI(NB.SI(Catégories!A:A;X1561)=0;"catégorie inconnue";
   RECHERCHE(X1561;Catégories!A:A;Catégories!B:B))`.
5. (~1:33 / 86 s) Onglet **"TCD"** actif pour la première fois avec le
   panneau de contrôle en O1:S9 déjà rempli — voir section dédiée
   ci-dessous.
6. (~1:53–2:26 / 113–143 s) Ouverture d'un PDF (visualiseur externe,
   probablement Adobe/navigateur) : "Bordereau de Contrôle Facturation"
   GLS n° facture **2501254050**, daté 31 mai (25), CE 9000037814,
   Class FDS — liste des colis (N° Colis, Date, Produit, Réf.
   destinataire, Destinataire, Points, Poids, Nb Colis). Puis bascule
   vers la page "**Facture**" du même PDF (Payeur 2507012387, Client SAP
   2500034092, Numéro de document 2501254050, Date 31 mai 25, Devise EUR)
   — le cadrage ne montre pas le total TTC de ce document précis.
7. (~2:24 / 144 s) Retour dans Excel, onglet "Facture GLS" (formule
   catégorie visible en barre de formule, cf. point 4).
8. (~2:25–2:28) Clic sur les onglets en bas -> atterrit sur **"TCD"**
   (pas "Bilan factures").
9. (~2:28–4:00) Manipulations dans le panneau "Champs de tableau croisé
   dynamique" du TCD : champs disponibles listés — Catégorie, Total HT,
   Gazole, Total HT (hors gazole), Poids, Facture, Document de vente,
   Date de la facture, n° client, raison sociale, n° chargeur, chargeur
   Alpha, Nom chargeur, etc. Champ "Colonnes" = Catégorie ; champ
   "Lignes" = Numéro de colis ; champ "Valeurs" = Somme de Total HT (ou
   Total HT hors gazole selon le test). Cellule `Q8`/`R8` "Ecart" change
   dynamiquement selon le champ choisi (1850,79€ avec "Total HT" vs
   1342,45€ avec "Total HT (hors gazole)") — démonstration/test manuel
   de la personne, pas forcément la configuration finale.
10. (~4:05–7:00) Onglet "Import csv" : construction/vérification du
    fichier `2025_05_GLS_Import.csv` (colonnes Transporteur, Date
    validité, Réf.1, Réf.2, Id client, N° Tracking, Nom, E/P, Pays,
    Zone, Nbr Colis, Poids, mode envoi, TVA, Droits et taxes, Assurance,
    Zones éloignées, Colis volumineux, Adresse, Frêt, plus-value, Gazole,
    Nb colis) — présence de lignes `#N/A` / "pays inconnu" / "zone
    inconnu taux inconnu" en fin de fichier (colis non trouvés dans le
    référentiel Comptes GLS/Zoning), à filtrer/traiter en amont.
11. (~7:00–7:58) Poursuite du scroll/vérification du CSV import jusqu'à
    la fin de la vidéo (~478 s). L'onglet "Bilan factures" n'est jamais
    cliqué/affiché pendant cette phase.

## Réponses précises aux 4 points demandés

### 1. Comment la colonne "TTC" est-elle calculée/remplie ?

Non observable directement dans l'onglet "Bilan factures" (jamais
affiché). Mais dans le TCD de contrôle observé (onglet "TCD", cellules
O1:S9, visible dès 86 s et jusqu'à ~239 s), le TTC affiché en **S1**
("5 614,18 € TTC") apparaît comme une valeur affichée à côté du HT
("Total / 4 678,48 € / HT"), sans qu'on voie la formule exacte être
tapée à l'écran (cellule jamais sélectionnée pour lire la barre de
formule). Le ratio observé 5614,18 / 4678,48 ≈ **1,2000...** — cohérent
avec un calcul `HT * 1,20` (TVA 20%, confirmée par ailleurs par le PDF
facture GLS réel qui affiche "Tax : 20,00 %"). **Ceci reste une
déduction par calcul de ratio, pas une lecture directe de formule** :
aucune frame ne montre la cellule S1 sélectionnée ni sa formule.

À confirmer par le pôle transport : si la colonne C "TTC" du "Bilan
factures" utilise bien `=B*1,2` (formule qui se recalcule), ou si elle a
été figée/copiée-collée en valeur pour un mois donné.

### 2. Réconciliation avec le PDF — quelle valeur, quel libellé, comment saisie ?

Le PDF facture GLS lu dans le classeur exemple du dossier
(`Facture_2507012387_2501382993_20260703_135525.pdf`, même structure que
celui manipulé dans la vidéo) affiche le TTC à deux endroits identiques :
- en bas de la page 1, encadré "**A joindre à votre règlement**" :
  `Client SAP / Date / Montant T.T.C. : 3 683,04` ;
- en page 4 (dernière page), tableau récapitulatif final :
  `Montant H.T. : 3 069,20 / Tax : 20,00 % 613,84 / Montant T.T.C. : 3 683,04`.

Le libellé exact à rechercher sur le PDF pour la réconciliation est donc
**"Montant T.T.C."**.

Dans le TCD de contrôle observé dans la vidéo (onglet "TCD", pas "Bilan
factures"), la valeur de comparaison n'est pas ce TTC du PDF mais un
**"Fichier importé"** : cellule **Q7/R7** = "Fichier importé — 2 016,11 €"
(à 86-143 s) puis "2 524,45 €" (à ~239 s, après changement de config du
TCD) — c'est-à-dire un total recalculé à partir du CSV d'import généré en
amont (onglet "Import csv"), **pas une saisie manuelle du TTC lu sur le
PDF**. L'écart en **Q8/R8** ("Ecart" 1850,79 € ou 1342,45 € selon le
champ de valeur choisi dans le TCD) compare donc `Total HT (du TCD sur
Facture GLS) - Fichier importé`, et non `TTC théorique - TTC PDF`.

**Conclusion pour le point 2** : le mécanisme de contrôle réellement
filmé dans "GLS_1" compare le TCD interne (Facture GLS) au CSV d'import,
pas au TTC du PDF fournisseur. La logique décrite dans la question
(colonnes D/E "TTC (PDF)" + F "Ecart" dans "Bilan factures") n'est **pas
visible dans cette vidéo** — probablement montrée dans une autre vidéo
GLS (`_2`, `_3`...) si elle existe, ou dans un onglet non filmé ici. Il
n'est pas possible de confirmer si le TTC PDF est saisi à la main ou lié
par formule/RECHERCHEV dans "Bilan factures" à partir de cette vidéo
seule.

### 3. Colonne "ok" ou équivalent — rôle exact ?

Aucune colonne "ok" n'est visible dans le TCD observé. On voit en
revanche, à droite de la cellule "Ecart", un **texte libre non structuré
en colonne dédiée** : `"Frais de gestion ok"` suivi de `"août 2024"`
(cellules R8:S8 environ, visible à 86-239 s). Cela ressemble à une **note
manuelle laissée par la personne** (commentaire de suivi mensuel/annuel
plutôt qu'une colonne "ok" structurée avec logique conditionnelle) —
probablement pour se rappeler qu'à cette date le poste "Frais de
gestion" avait été vérifié/validé. Ce n'est donc pas un indicateur
formule (type `=SI(ABS(Ecart)<seuil;"ok";"KO")`) mais une annotation
texte libre. À confirmer avec le pôle transport : cette case est-elle
réellement une simple note manuelle, ou existe-t-il ailleurs (onglet
"Bilan factures" non filmé) une vraie colonne "ok" avec logique de seuil ?

### 4. Le TTC théorique est-il figé (copié-collé) ou une formule vivante ?

Non déterminable avec certitude depuis cette vidéo — la cellule
contenant le TTC (S1 du TCD, "5 614,18 € TTC") n'est jamais sélectionnée
à l'écran, donc sa barre de formule n'est jamais visible. Le seul indice
disponible est le ratio TTC/HT = 1,20 exactement (5614,18/4678,48),
compatible avec une formule `=HT*1,2` aussi bien qu'avec une valeur
figée qui aurait été calculée une fois avec ce même taux. **Aucune preuve
visuelle de copier-coller de valeur, ni de formule tapée** n'apparaît
dans les frames disponibles pour cette cellule précise.

## Points ambigus / illisibles à faire confirmer par le pôle transport

- **Onglet "Bilan factures" jamais montré dans cette vidéo** : impossible
  de répondre avec certitude aux 4 questions posées à partir de "GLS_1"
  seule. Il faut vérifier s'il existe une vidéo `GLS_2` (ou suivante) qui
  couvre spécifiquement cet onglet, comme c'est le cas pour Geodis
  (`Geodis_2_Preparation Fichier Import.mp4`).
- Formule exacte de la colonne "TTC" (S1 du TCD observé, ou colonne C de
  "Bilan factures") : jamais lue en barre de formule dans la vidéo ;
  déduite uniquement par calcul de ratio (1,20).
- Le TCD de contrôle vu dans l'onglet "TCD" compare "Total HT" au
  "Fichier importé" (CSV), pas au TTC du PDF — à confirmer si c'est un
  contrôle différent et complémentaire de celui du "Bilan factures", ou
  si la question du pôle transport concerne en réalité ce même mécanisme
  mais mal localisé (onglet "TCD" au lieu de "Bilan factures").
  L'"Ecart" affiché change de valeur selon la configuration du TCD
  (1850,79 € avec "Total HT", 1342,45 € avec "Total HT hors gazole") —
  cela ressemble à une manipulation de test/démonstration plutôt qu'à la
  configuration finale figée du process.
- Texte "Frais de gestion ok / août 2024" : nature exacte (commentaire
  manuel libre vs. contenu de cellule avec logique) non confirmée à
  100 % — la police/mise en forme suggère un texte simple, mais la
  cellule n'a jamais été sélectionnée pour vérifier sa formule/valeur
  source.
- Le PDF ouvert pendant la vidéo (facture n° 2501254050, 31 mai 25) n'est
  pas le même exercice que celui utilisé en exemple pour illustrer le
  libellé "Montant T.T.C." (j'ai dû utiliser à la place le PDF
  `Facture_2507012387_2501382993_20260703_135525.pdf` déjà présent dans
  `Transporteurs/GLS/`, facture du 30 juin 26, pour confirmer le libellé
  exact affiché sur une facture GLS réelle). Le principe (libellé
  "Montant T.T.C." en bas de la 1ère page et en dernière page) est
  vraisemblablement identique d'un mois sur l'autre vu qu'il s'agit du
  même template GLS, mais ce n'est pas une capture directe de la vidéo.
