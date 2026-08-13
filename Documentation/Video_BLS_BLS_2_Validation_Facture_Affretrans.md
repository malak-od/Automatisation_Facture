# Transcription vidéo — BLS_2_Validation_Facture_Affretrans.mp4

Durée totale : ~1 min 08 s (67,9 s). Résolution capture d'écran Windows,
navigateur (Chrome, deux onglets : "planeo.laruche-logistique-france.fr" et
"Tickets - GLPI") + fenêtres Excel superposées/réduites au premier plan par
moments. Timestamps approximatifs en `mm:ss`.

Aucune bande son exploitable : toute l'info vient de ce qui est visible à
l'écran.

Cette vidéo montre le mécanisme exact de **validation/rapprochement de la
facture BLS directement dans le portail web "AffreTrans"**
(`planeo.laruche-logistique-france.fr/affrtrans/...`), complémentaire à la
préparation du fichier Excel vue dans la vidéo 1.

---

## Étape 1 — Filtrage de l'historique des commandes AffreTrans (00:00 – 00:12)

1. (00:00) Portail **AffreTrans** ouvert sur la page **"Historique"**
   (menu de gauche : Tableau de bord / Affrètement / Historique /
   Statistiques / Export / Stats / Annuaire / Import / Modèles Emails /
   Administration). Titre de page : **"Historique des commandes"**.
   Filtres visibles en haut : "Tous les clients" / "Tous les transporteurs"
   / nombre de résultats par page ("20...") / barre de recherche
   "Rechercher parmi tous les affrètements". Onglets de statut sous les
   filtres : **Toutes / Brouillon / Nouvelle demande / En attente
   transporteur / En attente retour client / Validé / Refusé / Annulé /
   Livraison conforme / Livraison avec anomalie / Facturé / Urgentes**.
   Onglet actif : **"Livraison conforme"**. Compteur : **9 commandes**.
   Cartes de commandes visibles : "Nez Commerce" (N° NEZ03260001,
   Transporteur BLS), "Hydratis" (N° HYD04260003, Transporteur BLS),
   "Hydratis" (N° HYD04260006, Transporteur BLS), "V. Marchand..." (N°
   VMA03260012, Transporteur BLS), "Hydratis" (N° HYD03260012,
   Transporteur BLS), "Atelier Marta" (N° ATE04260001, Transporteur BLS).
2. (00:02) Filtre déroulant **"Tous les transporteurs"** cliqué (curseur
   dessus).
3. (00:04 – 00:06) Sélection du transporteur **"BLS"** dans le filtre
   déroulant → la liste se réduit à **8 commandes** (toutes Transporteur =
   BLS), statut "Livraison conforme" toujours actif comme filtre.

## Étape 2 — Ouverture d'une commande individuelle et confirmation du "N° Tracking" (00:08 – 00:14)

4. (00:08) Clic sur la première carte de commande (probablement "Nez
   Commerce" ou équivalent) → ouverture d'une page détail "Enlèvement"
   horodatée **04/05/2026 12:03:16**.
5. (00:08 – 00:10) Page détail "Enlèvement" : formulaire avec Type de lieu
   = "La Ruche Logistique", Nom de l'entreprise = "La Ruche Logistique",
   Adresse = "4 Voie communale", Code postal = "21320", Ville = "Créancey".
   À droite : bloc **"Transporteurs"** avec 4 cases de tarif : **Géodis
   278,92€ / BLS 195€ (encadré en surbrillance) / Kuehne 81,96€ / CEVA
   0,00€**, et une case **FedEx 0,00€** séparée. En dessous, bloc orange
   **"TRANSPORTEUR — Demande de cotation / Confirmation"** avec boutons
   Géodis/BLS/Kuehne/CEVA/FedEx (pas d'email pour FedEx), un sélecteur de
   modèle d'email, un bouton **"Envoyer aux 0 transporteur(s)
   sélectionné(s)"**, et en bas boutons **"Demander cotation"** /
   **"Confirmer commande"**. Statut de la commande (barre du bas) :
   **"Livraison conforme"**.
6. (00:12) Défilement vers le bas de la même page : section **"Prestation
   — Détails concernant la marchandise et le service"** : Marchandise =
   "Livres", Type de colis = Palette, Nombre de palettes/colis = **1**
   (case "Palette Europe (consignée)" cochée), **Poids total (kg) = 400**,
   dimensions Longueur 80 / Largeur 120 / Hauteur 160 cm.
   À droite : bloc **"Commentaires / Notes internes"** avec deux champs
   clés : **"N° Tracking"** (vide, placeholder "Numéro de suivi") et
   **"N° Récépissé" = 11111**. Case à cocher "Problèmes signalés" (non
   cochée).

## Étape 3 — Ouverture des fichiers Excel de référence en parallèle (00:13 – 00:16)

7. (00:13-00:14) Une barre de tâches Windows apparaît en bas de l'écran
   avec **4 fichiers ouverts en aperçu miniature** :
   - **`2026_04_Facture BLS.x...`** (classeur de travail xlsx)
   - **`2026_04_BLS_Import.csv - ...`**
   - **`2600772 (1).xlsx - Excel`** (fichier source brut reçu du
     transporteur, même référence que celle vue dans la vidéo 1)
   - **`export_affretement_2026-...`** (export du portail AffreTrans, même
     fichier que celui généré/utilisé dans la vidéo 1, cohérent avec le nom
     `export_affretement_2026-05-04_1108.csv`)
   Cela confirme que le pôle transport garde ces 4 fichiers ouverts
   simultanément pendant tout le processus de rapprochement/validation.
8. (00:14 – 00:16) Bascule vers le classeur **`2026_04_Facture BLS.xlsx`**,
   onglet **"Factures BLS"** actif. Cellule D5 sélectionnée, contenu
   "11157" (colonne "Dossier"). Le tableau affiche les mêmes 16 lignes déjà
   vues dans la vidéo 1 (Dossier 11111 à 11609, n° facture 2600772 pour
   toutes les lignes, colonne "ID Client" = colonne A, actuellement VIDE
   sur toutes les lignes visibles à l'écran à ce stade).

## Étape 4 — Sélection de la colonne "ID Client" et confirmation qu'elle est vide (00:16 – 00:20)

9. (00:16) Colonne A "ID Client" sélectionnée entièrement (surlignage vert
   pâle sur toute la colonne, lignes 2 à 16) — aucune valeur visible dans
   les cellules à ce stade (colonne vide confirmée pour ce fichier avril
   2026, cohérent avec la vidéo 1).
10. (00:16 – 00:18) Cellule D2 sélectionnée, contenu "11111" — barre de
    formule affiche "11111" en valeur brute (pas de formule).

## Étape 5 — Retour au portail AffreTrans, saisie du statut / facturation (00:19 – 00:24)

11. (00:19) Retour à l'onglet navigateur AffreTrans, page détail de la
    commande "Hydratis" N° HYD04260003 (ou équivalent) — champ **"N°
    Tracking"** toujours vide, **"N° Récépissé" = 11111** inchangé.
12. (00:20) Clic sur le menu déroulant **"Statut"** en bas de page
    (actuellement "Livraison conforme") : liste déroulante affichant
    toutes les valeurs possibles de statut : **Brouillon / Nouvelle
    demande / En attente transporteur / En attente retour client / Validé
    / Refusé / Annulé / Livraison conforme (coché ✓) / Livraison avec
    anomalie / Facturé**.
13. (00:22 – 00:24) Sélection du statut **"Facturé"** dans la liste
    déroulante (changement d'état de la commande de "Livraison conforme"
    à "Facturé").

## Étape 6 — Renseignement du montant de facturation transporteur (00:24 – 00:30)

14. (00:24) Après changement de statut, une nouvelle section apparaît/est
    scrollée à droite : **"Facturation transporteur"** (horodatage
    04/05/2026 12:03:30), avec deux champs : **"N° Facture"** (vide,
    placeholder "Numéro de facture") et **"Montant facture (€)"**
    (contient déjà "0.00", curseur cliqué dedans).
    En dessous, deux lignes informatives non modifiables :
    - **"Tarif transporteur (BLS)" = 195,00 €**
    - **"Coût attendu transporteur" = 195,00 €**
    (Ces deux valeurs sont égales à ce stade, correspondant au tarif "BLS"
    déjà vu dans le bloc "Transporteurs" de l'étape 5, ligne "BLS 195€".)
15. (00:26 – 00:28) L'utilisateur tape **"195"** dans le champ "Montant
    facture (€)".
16. (00:28 – 00:30) Après saisie, une nouvelle ligne apparaît sous "Coût
    attendu transporteur" : **"Écart"** avec un badge **vert "Conforme"**
    (195,00 tarif = 195,00 facturé = écart nul → validation automatique
    affichée par un badge coloré).
    **Mécanisme confirmé : AffreTrans compare automatiquement le "Tarif
    transporteur (BLS)" (tarif système/attendu) au "Montant facture (€)"
    saisi manuellement par le pôle transport (lu depuis la vraie facture
    BLS), et affiche un badge "Conforme" (vert) ou un écart en euros
    (orange/rouge) selon le résultat.**

## Étape 7 — Sauvegarde et retour à l'historique filtré (00:30 – 00:34)

17. (00:30) Clic sur **"Enregistrer l'évolution"** (bouton en bas à
    droite). Une notification toast apparaît : **"Progrès enregistré —
    Commande mise à jour."**
18. (00:32 – 00:34) Le champ "N° Tracking" (bloc Commentaires/Notes
    internes) est cliqué et l'utilisateur y saisit/confirme la valeur
    **"2600772"** (le numéro de facture BLS, PAS un numéro de tracking
    classique — probablement pour lier cette commande AffreTrans au numéro
    de facture officiel BLS pour traçabilité). Notification "Progrès
    enregistré" réapparaît après cette saisie.

## Étape 8 — Retour à la liste "Historique", vue élargie (00:34 – 00:44)

19. (00:36) Retour à la page "Historique des commandes", filtre "Tous les
    clients" / "Tous les transporteurs" (réinitialisé), statut
    **"Toutes"** actif → **114 commandes** au total affichées. Cartes
    visibles : "Nez Commerce" (Facturé), "Renight Store" (Facturé, avec
    badge d'alerte orange ⚠), "Atelier Marta" (Facturé), "Atelier Marta"
    (Facturé), "Weloca" (Facturé), "SAS LE MOU..." (Nouvelle demande,
    Transporteur "Non défini").
20. (00:38 – 00:40) Nouveau filtrage : retour du filtre transporteur sur
    **"BLS"**, statut **"Livraison conforme"** → **7 commandes**
    restantes (la commande traitée à l'étape 6-7 étant passée en
    "Facturé", elle disparaît de cette vue filtrée "Livraison conforme").
    Cartes visibles : "Hydratis" (HYD04260003), "Hydratis" (HYD04260006),
    "V. Marchand..." (VMA03260012), "Hydratis" (HYD03260012), "Atelier
    Marta" (ATE04260001), "La Ruche" (LAR04260001).

## Étape 9 — Traitement d'un second dossier ("Hydratis" HYD04260003, Dossier 11301) (00:42 – 00:60)

21. (00:42) Clic sur la carte "Hydratis" N° HYD04260003 → page détail
    **"Détail de l'Affrètement pour 'Hydratis' n°: HYD04260003"**. Barre
    de progression du statut visible en haut : Brouillon → Nouvelle
    demande → En attente transporteur → En attente retour client → Validé
    → Refusé → **Livraison conforme (point actif)** → Livraison avec
    anomalie → Facturé.
    Section **"Informations Client"** : Client = "Sélectionner un client"
    (vide, liste déroulante non renseignée), **"Code Client" / "ID
    Client"** = champ vide (placeholder "ID Client"), Numéro d'expédition
    vide, Date d'expédition vide.
    Section **"Cotations Reçues"** à droite : Prestation = Affrètement/
    Messagerie/Course/Aérien (aucun coché à l'écran), bloc Transporteurs
    Géodis/BLS/Kuehne/CEVA tous à 0,00€ à ce stade (cotations non encore
    saisies pour cette commande précise), bloc FedEx 0,00€.
    **Confirme que le champ "Code Client"/"ID Client" existe bel et bien
    dans AffreTrans au niveau "Informations Client" de chaque commande,
    mais n'est PAS automatiquement renseigné — reste à "Sélectionner un
    client" / vide dans cet exemple.**
22. (00:44 – 00:46) Section adresse de livraison scrollée : "DB SCHENKER
    LE HAVRE C/O NOUMEA TRANSIT", "VOIE DES VANNEAUX", France, 76430,
    Saint-Vigor-d'Ymonville. Blocs email "Envoyer le devis au client" et
    "Envoyer en interne" visibles à droite, ainsi qu'un bloc "Documents
    joints" en bas (zone de dépôt de fichiers, vide).
23. (00:48) Défilement vers la section **"Prestation"** de ce dossier :
    Marchandise = "Equipements", Type de colis = Palette, **Nombre de
    palettes/colis = 5**, case "Palette Europe (consignée)" cochée,
    **Poids total (kg) = 1750**. 5 lignes de dimensions détaillées (Qté 1,
    Longueur 80, Largeur 120, Hauteur variable : 165/165/165/135/95 cm
    pour les dimensions 1 à 5).
    Bloc "Commentaires / Notes internes" à droite : **"N° Tracking"**
    vide, **"N° Récépissé" = 11301** (au lieu de 11111 du dossier
    précédent).
    Bloc "Service créateur" : "Demandé par : Marie-Pierre CARLES
    (atelier@laruche-logistique.fr)".
24. (00:50 – 00:52) Bascule vers le classeur Excel `2026_04_Facture
    BLS.xlsx`, onglet "Factures BLS", cellule B2 sélectionnée, contenu
    "2600772" (n° facture). Défilement pour localiser la ligne Dossier
    "11301" — visible ligne 10 : Dossier=11301, Date=13/04/2026,
    Libellé="De 21 Créancey A 76 Saint-Vigor-d'Ymonville" (correspond bien
    à l'adresse de livraison "Saint-Vigor-d'Ymonville" vue à l'étape 22),
    Montant H.T.=365,00.
25. (00:54 – 00:56) Sélection de la cellule D10 (Dossier=11301) et de la
    plage autour (multi-sélection "6L x 1C" visible dans la zone de nom) —
    confirmation visuelle du rapprochement entre le "N° Récépissé"=11301
    d'AffreTrans et la ligne "Dossier"=11301 du classeur Factures BLS.
    **Confirme la clé de rapprochement : "N° Récépissé" (côté AffreTrans)
    = "Dossier" (côté classeur BLS) — et non "N° Tracking" qui lui reste
    vide côté AffreTrans dans ces deux exemples.**
26. (00:58) Info-bulle "ID Client" survolée sur l'en-tête de colonne A du
    classeur Excel (confirmation du nom exact de la colonne).

## Étape 10 — Saisie de la facturation pour le second dossier, avec ERREUR puis correction (00:58 – 01:08)

27. (00:58) Retour à AffreTrans, changement du statut de la commande
    "Hydratis" HYD04260003 vers **"Facturé"** (via le menu déroulant,
    liste des statuts affichée à nouveau).
28. (01:00) Section **"Facturation transporteur"** (horodatage 04/05/2026
    12:04:12) : champ **"N° Facture" = 2600772** déjà saisi, champ
    **"Montant facture (€)"** vide (placeholder "0.00").
    Lignes informatives : **"Tarif transporteur (BLS)" = 365,00 €**,
    **"Coût attendu transporteur" = 365,00 €** (cohérent avec le Montant
    H.T. de 365,00 lu dans le classeur Excel à l'étape 24).
29. (01:02) L'utilisateur tape **"36"** (erreur de frappe, montant
    incomplet — il manque le dernier chiffre) dans le champ "Montant
    facture (€)". Résultat immédiat : ligne **"Écart" = badge orange/rouge
    "-329,00 €"** (365 - 36 = 329, écart significatif signalé visuellement
    par une couleur d'alerte, PAS "Conforme").
30. (01:04 – 01:06) Correction : le champ est complété pour afficher
    **"365"** (montant correct cette fois). Résultat : ligne "Écart"
    repasse en **badge vert "Conforme"** (365 = 365, écart nul).
31. (01:06 – 01:08) Notification toast **"Progrès enregistré — Commande
    mise à jour."** réapparaît (sauvegarde automatique ou après clic sur
    "Enregistrer"). Fin de la vidéo sur cet état.

---

## Synthèse du mécanisme observé (pour rappel, sans interprétation ajoutée au-delà de ce qui est vu)

- **"AffreTrans" est un portail web interne** (URL
  `planeo.laruche-logistique-france.fr/affrtrans/...`), PAS un logiciel
  tiers du transporteur BLS ni un simple fichier. Il gère les commandes
  d'affrètement pour plusieurs transporteurs (Géodis, BLS, Kuehne, CEVA,
  FedEx).
- Le **rapprochement "trackings dans l'ERP"** mentionné dans le registre se
  fait via la page **Historique** d'AffreTrans, filtrée par transporteur
  ("BLS") et par statut ("Livraison conforme" → à traiter, puis
  "Facturé" → traité), en ouvrant individuellement chaque commande.
- La **clé de rapprochement** entre une commande AffreTrans et une ligne du
  classeur Excel "Factures BLS" est le champ **"N° Récépissé"** côté
  AffreTrans, qui correspond à la colonne **"Dossier"** côté classeur BLS
  (PAS le champ "N° Tracking" d'AffreTrans, qui reste vide dans les deux
  exemples observés).
- La **validation de facturation** se fait DIRECTEMENT dans AffreTrans :
  changement du statut de la commande vers "Facturé", puis saisie du
  "N° Facture" et du "Montant facture (€)" dans la section "Facturation
  transporteur". AffreTrans compare alors automatiquement ce montant saisi
  au "Tarif transporteur" pré-calculé par le système, et affiche un badge
  "Conforme" (vert) ou un écart en euros signé (orange/rouge) en cas de
  différence.
- Le champ **"Code Client"/"ID Client"** existe dans la section
  "Informations Client" d'AffreTrans mais n'était PAS renseigné dans les
  deux commandes consultées lors de cette vidéo (reste à "Sélectionner un
  client..." / vide) — cette vidéo ne montre donc PAS le mécanisme exact
  de détermination/saisie du code client observé dans le TCD "Bilan
  client" du classeur Excel (cf. point ambigu correspondant ci-dessous).

---

## Points ambigus / illisibles à faire confirmer par le pôle transport

1. **Mécanisme de détermination du code client ("ID Client")** : NI la
   vidéo 1 NI la vidéo 2 ne montrent le pôle transport en train de
   consulter un outil/ERP pour DÉTERMINER un code client à partir d'un
   Dossier ou d'un trajet, ni le saisir concrètement dans le classeur
   Excel "Factures BLS". Dans les deux vidéos, la colonne "ID Client" du
   classeur de travail (avril 2026) reste vide de bout en bout, et le
   champ "Code Client"/"ID Client" d'AffreTrans reste également vide pour
   les deux commandes consultées. **Ce point spécifiquement demandé par
   l'appelant n'a PAS pu être élucidé avec ces deux vidéos — à
   redemander/reconfirmer directement auprès du pôle transport** (peut-être
   un mécanisme réalisé sur un fichier "Bilan client" pour un mois
   antérieur, non capturé ici, ou une étape effectuée en dehors de la
   plage filmée).
2. **Premier dossier traité (étapes 4-7)** : le nom exact de la commande
   ("Nez Commerce" ou une autre carte de la liste) n'est pas identifié
   avec 100% de certitude au moment précis du clic (00:08) — la carte
   cliquée n'est pas visible sur la frame immédiatement suivante montrant
   déjà la page détail. Le "N° Récépissé"=11111 et le tarif BLS=195€
   permettent cependant de la relier avec confiance à la ligne
   Dossier=11111 du classeur Excel (Montant H.T.=195,00, cohérent).
3. **Signification exacte du champ "N° Tracking" (toujours vide dans
   AffreTrans, distinct du "N° Récépissé")** : à clarifier — semble être
   un champ prévu pour un numéro de tracking transporteur externe (non
   utilisé/non rempli par BLS ou par le pôle transport pour ce
   transporteur), tandis que "N° Récépissé" sert de clé métier interne
   correspondant au "Dossier" BLS.
2bis. **Pourquoi taper "2600772" (n° facture) dans le champ "N°
   Tracking"** (étape 18) plutôt que dans un autre champ dédié : ce choix
   semble délibéré (traçabilité du numéro de facture dans la fiche
   commande) mais le raisonnement exact n'est pas explicité à l'écran — à
   confirmer si c'est bien la pratique standard ou une correction ad hoc
   de ce cas précis.
4. **Colonne "Réf.1"/"Réf.2" et champs "Numéro d'expédition" côté
   AffreTrans** : non renseignés dans les exemples vus, rôle exact non
   déterminé par cette vidéo.
5. **Aucune mention de taxe gasoil / Gazole** n'apparaît dans cette vidéo
   2 (le champ "Tarif transporteur (BLS)" semble être un montant unique
   tout compris, sans détail visible d'une éventuelle surcharge gasoil
   séparée) — cohérent avec le statut "à définir" du registre, mais ne
   permet pas de le préciser davantage.
6. **Aucun "MAJ TCD" (rafraîchissement de tableau croisé dynamique
   Excel) n'est visible dans cette vidéo 2** — cette action a été
   observée uniquement dans la vidéo 1 (onglet "Bilan PDF"). Le mécanisme
   "MAJ TCD" pour "Bilan client" spécifiquement n'a été vu dans aucune des
   deux vidéos.
7. **Numéro exact des commandes / N° Récépissé pour les cartes non
   cliquées** (Nez Commerce, V. Marchand, Atelier Marta, La Ruche, etc.)
   : non vérifiés individuellement, seules les deux commandes explicitement
   ouvertes (N° Récépissé 11111 et 11301) ont été documentées en détail.
