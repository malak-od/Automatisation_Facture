# Notes de facturation — observations terrain

Notes prises en observant la facturation se faire (transcrites telles quelles,
2026-07). Complète le manuel `FACTURATION EXCEL.pdf` et les analyses vidéo par
transporteur. Certains points contredisent ou précisent ce qui est déjà codé
dans `facturation-app/` — signalé en **⚠️** avec le contexte.

## Concepts généraux

- **Taxe gasoil** : à faire entrer dans l'ERP *avant* la facturation. Elle est
  **par transporteur** — certains l'ont, d'autres non (ex. Delivengo n'en a pas).
- Les transporteurs facturent **par tracking**.
- **Avarie d'import** = tracking non trouvé dans l'ERP. Causes : appel API
  transporteur, expédition pas encore partie de chez nous.
- **Avarie de vente** = grille tarifaire erronée/obsolète.
- **Viticolis** = principalement **UPS**, sinon Geodis / Kuehne / BLS.
- **Navette** = tarifs négociés avec le transporteur (factures internes) :
  Kuehne (inclus dans son Excel), Geodis (PDF à part), BLS (pas encore démarré
  à la date de la note).

## Kuehne

- EDI = ce qu'on envoie **tous les jours** au transporteur.
- Recherche en ligne sur le site du transporteur pour certaines infos.
- TVA : France/Europe = 0,2 ; ailleurs = 0,0 *(✓ déjà dans `kuehne/config.json`)*.
- **Remarque post-V1** : garder les formules dans toutes les feuilles du
  classeur fonctionne pour juin (06), mais testé avec les fichiers de mai (05)
  ça renvoie un fichier différent de celui de juin → le classeur modèle n'est
  pas robuste d'un mois à l'autre en l'état (`finaliser_kuehne.py`).

## Delivengo

- Fret = 1,0, on ne facture rien dessus (à confirmer). Fini, attend l'export
  expéditions brut.
- **Remarque post-V1** : colonne Poids du fichier brut Delivengo → **poids/100**
  → à mettre dans les colonnes **Poids et Assurance** du fichier d'import.
  ⚠️ **Contredit un choix déjà fait** : lors de la correction du carrier
  Delivengo (2026-07), le même arbitrage (/1000 selon le PDF vs /100 selon
  cette note) avait été posé explicitement et tranché en faveur de **/1000**
  (actuellement en prod dans `finaliser_delivengo.py`). À reconfirmer avec une
  vraie facture avant de changer quoi que ce soit — l'écart x10 change le
  montant d'Assurance.
- Colonne Droits et taxes du fichier d'import : chercher le n° de suivi dans
  l'export des expéditions brutes **du mois en cours** (colonne
  `PRO_TRACKING`) pour ressortir la colonne `INFO_POIDS`. *(✓ implémenté)*
  Le logiciel doit indiquer clairement où déposer le fichier transporteur reçu
  **et** le fichier des exports bruts du mois en cours — aujourd'hui ce
  dernier est retrouvé automatiquement par nom de fichier, jamais uploadé
  explicitement.
- Si après cette recherche des Droits et taxes restent vides : refaire la
  même recherche dans l'export brut **du mois précédent**. *(✓ implémenté :
  priorité mois courant puis repli mois-1)*.

## Geodis

- Colonnes décalées vers la fin — **toujours vérifier**. Bien vérifier que la
  colonne décalée est incluse dans la formule du fret.
- Numéro récépissé : **9 = messagerie**, **4 = affrètement**.

## BLS

- Note : « BLS est fait ». État à clarifier — dans le registre du logiciel
  BLS est encore `planned` (non codé). À vérifier si ça veut dire que le
  process manuel Excel est stabilisé (prêt à automatiser) ou autre chose.

## GLS

- Même procédure que les autres : recopier le fichier du mois précédent,
  coller le CSV en valeur, le tableau du bas sert de base au TCD, actualiser
  la feuille Poids, effacer les ID clients.
- Bilan facture = comparaison avec le PDF.
- Vérifier les frêts à zéro.
- Gazole à rajouter sur l'ERP.

## DPD

- **⚠️ Écart constaté sur le fichier fait à la main de juillet (2026-07)** :
  comparaison ligne à ligne entre `2026_07_Facture DPD.xlsx` (fait à la main)
  et le classeur généré par `finaliser_dpd.py` à partir des mêmes fichiers
  sources reçus (`.../2026 07/DPD/excel/*.xlsx`). Écart total de **15,50 €**
  sur la colonne D (« Total hors GO ») de l'onglet « Facture DPD » :
  30 440,80 € (généré) vs 30 425,30 € (fait main).
  - Réparti sur 3 colonnes : **C — Zones éloignées** (+14,32 €, 70 lignes),
    **I — Frêt** (+2,66 €, 32 lignes), **K — Retour** (−1,48 €, 26 lignes).
  - Pour chaque ligne vérifiée (par N° Colis), la valeur retenue par le
    finaliseur correspond **exactement** à la valeur brute du fichier source
    reçu de DPD (ex. tracking `10214001115721` : « Fact. Retour expédition »
    = 5,70 dans le fichier source, contre 6 dans le fichier fait à la main ;
    tracking `10214001114728` : « Supplément île et montagne » = 5,28 dans le
    fichier source, contre 5 dans le fichier fait à la main).
  - **Confirmé avec le pôle transport (2026-08-07) : notre méthode est la
    bonne.** Le fichier fait à la main arrondit ces montants (à l'entier le
    plus proche sur 122 des 128 lignes en écart), le classeur généré garde
    les décimales exactes des fichiers source DPD.

- **Écart persistant entre « Bilan factures » (colonne Ecart) et les factures
  PDF DPD — cause identifiée (2026-08-07), pas une erreur** : le TCD compare
  « Somme de Total GO » (colonne E de « Facture DPD ») au montant HT de
  chaque PDF, et un écart subsiste presque toujours (ex. juillet : de
  −19,93 € à +116,15 € selon le compte). Cause : la formule Total GO de
  juillet (`E = SUM(G:L)+SUM(A:C)+N`) utilise la colonne **N « Frais de
  dossier modifié »** (total mensuel des « Frais de tenue de compte » réparti
  également sur toutes les lignes du mois), alors que **chaque PDF facture
  le vrai montant réel M « Frais dossier réel »** propre à ce compte (souvent
  20 € fixe, indépendant du nombre de lignes de ce compte). L'écart mesuré
  correspond exactement à `N − M` pour chaque compte (vérifié à l'exactitude
  du centime sur les 15 comptes de juillet). C'est la formule officielle du
  fichier fait à la main (confirmée par le pôle transport) — l'écart est donc
  structurel et attendu, pas un signe d'erreur de calcul.


## Lettres

- Filtrer sur les expéditions brutes, vider tous les trackings vides
  (colonne bleue), les rechercher dans un export depuis l'ERP.
- Pour ceux à zéro : tracking = numéro d'expédition, mis en **bleu foncé**,
  et mise à jour du tracking dans l'ERP — **sauf** les expés en cours de
  préparation (ne pas facturer, bien vérifier).
- `PRO_TRACKING` commençant par **BAC25\*** = erreur atelier → mettre le
  numéro d'expédition en tracking.
  

## UPS

- Exclure Multitasker et Vitawave (comptes séparés).
- Télécharger les factures par paquets de 10 sur le site du transporteur.
- Remplacer les points par des virgules dans toutes les colonnes **sauf AX**
  *(✓ cohérent avec le PDF p.5)*.
- Colonne 21 = tracking, à trier de A à Z.
- Colonne D : vérifier vide/N-A → si pas de tracking, retirer la ligne.
- Fichier import — nombre de colis : pour les expés multi-colis, une seule
  référence tracking « maître ».
- Vérifier les références de destination des expés selon le PDF Viticolis
  (à demander/obtenir).
- Import ERP — vocabulaire des erreurs :
  - **Mappage entité** = numéro de tracking qui n'existe pas dans l'ERP.
  - **Doublon consolidé** = déjà facturé le mois précédent.
  - **Doublon fichier** = cette expédition apparaît deux fois dans le
    fichier d'import.

## Mondial Relay

- Plusieurs fichiers ZIP par facturation ; transformer tous les points en
  virgules depuis le CSV de base envoyé par le transporteur.
- Une facture PDF **par pays** (avant : par pays ET par client — le format a
  changé).
- Les écarts sont souvent sur la facture France/La Ruche : la **collecte**
  n'est pas facturée aux clients, il faut l'ajouter dans la feuille contrôle.
- Mondial Relay = Europe uniquement → TVA fixe 0,2.
- Fichier import, colonne mode envoi : remplacer **24RC par 24R** pour avoir
  des zones connues (sinon avaries d'import) *(✓ cohérent avec le PDF p.9)*.
- Les frêts à zéro se suppriment, **en vérifiant** que les autres colonnes à
  facturer sont aussi à zéro (sinon ne pas supprimer).

## Colissimo

- Le fichier du mois précédent renvoie une erreur — à refaire.

## FedEx

- Fichier CSV téléchargé par la compta.
- Les modes d'envoi FICP sont en mauvais état (« c'est la merde » — beaucoup
  de cas à traiter manuellement).
- Voir le wiki / dossier transports pour valider certaines notations —
  emplacement à préciser (ni l'un ni l'autre identifiés dans le dépôt pour
  l'instant).

## Chronopost

- Deux factures Excel envoyées par le transporteur.
- Nouveau classeur : copier les infos des factures, trier les trackings de
  A à Z.
- **CAP** = taxe gazole, **ECO** = taxe éco, **SUR** = taxe sûreté.
- Colonne AF : un calcul par label, trié par la colonne L.
- Actualiser le TCD poids et le TCD général, en supprimant les ID clients ;
  tout éteindre en formule (pas de valeurs figées).
- **Deux contrats distincts** : Standard et 2Shop — non différenciés dans le
  PDF `FACTURATION EXCEL.pdf`, à prendre en compte si/quand Chronopost est
  automatisé.

## Autres notes (contexte à préciser — pas clairement liées aux imports transporteurs)

- Le SMIC qui augmente doit être mappé dans la base de données.
- Normer les factures par code d'expédition (impacte uniquement la main
  d'œuvre) ; dans l'onglet TCD prépa des tarifs, catégoriser chaque prix,
  arrondir au centime.
