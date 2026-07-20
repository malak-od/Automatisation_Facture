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

- Fait.

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
