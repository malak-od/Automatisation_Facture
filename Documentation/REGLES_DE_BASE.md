# Règles de base — facturation transporteurs

Document de référence transversal (valable pour **tous les transporteurs**),
à consulter avant de livrer un import ERP ou de coder/réviser un carrier.
Complète `FACTURATION EXCEL.pdf`, `Notes_Facturation_Observations.md` et les
analyses vidéo par transporteur (spécificités propres à chacun).

## Glossaire

- **Taxe gasoil** : à faire entrer dans l'ERP *avant* la facturation. Elle est
  **par transporteur** — certains l'ont, d'autres non (ex. Delivengo n'en a pas).
- Les transporteurs facturent **par tracking**.
- **Avarie d'import** = tracking non trouvé dans l'ERP. Causes : appel API
  transporteur, expédition pas encore partie de chez nous.
- **Avarie de vente** = grille tarifaire erronée/obsolète.
- **Viticolis** = principalement **UPS**, sinon Geodis / Kuehne / BLS.
- **Navette** = tarifs négociés avec le transporteur (factures internes) :
  Kuehne (inclus dans son Excel), Geodis (PDF à part), BLS (pas encore
  démarré).

## Checklist qualité — à appliquer pour TOUS les transporteurs

### Dans le fichier de calcul

- Supprimer les ID clients.
- Checker le tarif des PDF avec celui du TCD.
- Dans l'import ERP, étendre toutes les colonnes jusqu'à voir le même nombre
  de lignes que dans le TCD + changer la date de la 1ère ligne.

### Dans le fichier d'import

- Checker que le nom du transporteur soit bien mis partout.
- Checker que la date soit au format date courte (et le bon mois).
- Checker que les numéros de tracking soient corrects.
- Checker qu'il y ait E/P sur toutes les lignes.
- Checker qu'il n'y ait pas de pays inconnu ou de zone inconnue.
- Checker qu'il n'y ait pas de colis = 0.
- Checker qu'il n'y ait pas de poids = 0 kg → comparer à l'export des
  expéditions brutes du mois M ou M-1 (voir « Export expéditions brut » ci-dessous).
- Checker que les poids possèdent 1 décimale après la virgule (si besoin,
  faire un `=ARRONDI.SUP` dans le fichier de calcul) : évite les avaries de
  prix dans l'ERP.
- Checker de ne pas avoir de fret à 0 € ; si oui, enlever le zéro et ne rien
  mettre.
- Si fret vide **et** colonnes « droits et taxes » jusqu'à « adresse » toutes
  à 0/vide → supprimer la ligne entière. Si fret vide mais au moins une de
  ces colonnes a une info → laisser la ligne.
- Checker que chaque somme des colonnes (droits et taxes, assurance, zones
  éloignées, colis volumineux, adresses, fret) soit égale à celle du fichier
  de calcul.
- Checker qu'il y ait des zones partout.

## Export expéditions brut (`Automatisation/AAAA MM - Export expéditions_brut.xlsx`)

Fichier WMS mensuel, source d'arrière-plan partagée par plusieurs
transporteurs — **jamais** à uploader manuellement dans l'app, il est
retrouvé automatiquement par convention de nom (mois de la facture, avec
repli mois-1 si besoin) via `facturation-app/src/core/exportBrut.js`.

Usages actuels :
- **Delivengo** : recherche `PRO_TRACKING` → colonne `INFO_POIDS` /
  `INFO_POIDSRETENU` pour Droits et taxes / poids.
- **GLS** : 3ᵉ repli poids (après poids brut CSV BCF et grille tarifaire
  prix→poids) via `INFO_POIDSRETENU`, pour les colis dont aucune ligne du
  CSV BCF ne porte de poids (ex. catégorie `SURCHALL RETURN`).
- **Lettres** : source unique de calcul (filtré par colonne `TRANSPORTEUR`).

À généraliser aux autres transporteurs si la même règle "poids=0kg" s'y
applique.

## Navette (factures internes, tarifs négociés transporteur)

- **Kuehne** : incluse directement dans son classeur Excel.
- **Geodis** : PDF à part (pas dans le fichier transporteur principal).
- **BLS** : pas encore démarrée (à la date des observations terrain, 2026-07).

## Notes par transporteur

### Kuehne
- EDI = ce qu'on envoie **tous les jours** au transporteur.
- Recherche en ligne sur le site du transporteur pour certaines infos.
- TVA : France/Europe = 0,2 ; ailleurs = 0,0 (✓ dans `kuehne/config.json`).
- **Point ouvert** : garder les formules dans toutes les feuilles du classeur
  fonctionne pour juin (06), mais testé avec les fichiers de mai (05) ça
  renvoie un fichier différent → le classeur modèle n'est pas robuste d'un
  mois à l'autre en l'état (`finaliser_kuehne.py`).

### Delivengo
- Fret = 1,0, rien facturé dessus (à confirmer).
- **Point ouvert / arbitrage à reconfirmer** : colonne Poids du fichier brut
  Delivengo → deux hypothèses vues dans les notes terrain, poids/100 vs
  poids/1000. Le choix **/1000** a été tranché explicitement et est
  actuellement en prod (`finaliser_delivengo.py`) — /100 (vu dans une note
  ultérieure) contredit ce choix. Vérifier sur une vraie facture avant de
  changer quoi que ce soit : l'écart x10 change le montant d'Assurance.
- Droits et taxes : recherche du n° de suivi (`PRO_TRACKING`) dans l'export
  des expéditions brutes du mois en cours, colonne `INFO_POIDS` (✓
  implémenté). Si vide, repli sur l'export du mois précédent (✓ implémenté).

### Geodis
- Colonnes décalées vers la fin — **toujours vérifier**, notamment que la
  colonne décalée est incluse dans la formule du fret.
- Numéro récépissé : **9 = messagerie**, **4 = affrètement**.
- Navette = PDF à part (voir section Navette ci-dessus).

### BLS
- État à clarifier : les notes terrain disent « BLS est fait », mais dans le
  registre du logiciel BLS est encore `planned` (non codé). À vérifier si ça
  signifie que le process manuel Excel est stabilisé (prêt à automatiser) ou
  autre chose.
- Navette pas encore démarrée.

### GLS
- Même procédure que les autres : recopier le fichier du mois précédent,
  coller le CSV en valeur, le tableau du bas sert de base au TCD, actualiser
  la feuille Poids, effacer les ID clients.
- Bilan facture = comparaison avec le PDF.
- Vérifier les frêts à zéro.
- Gazole à rajouter sur l'ERP.
- Poids=0 : voir « Export expéditions brut » ci-dessus (repli implémenté).

### DPD
- Fait (pas de remarque particulière dans les notes terrain).

### Lettres
- Filtrer sur les expéditions brutes, vider tous les trackings vides
  (colonne bleue), les rechercher dans un export depuis l'ERP.
- Pour ceux à zéro : tracking = numéro d'expédition, mis en **bleu foncé**,
  et mise à jour du tracking dans l'ERP — **sauf** les expés en cours de
  préparation (ne pas facturer, bien vérifier).
- `PRO_TRACKING` commençant par **BAC25\*** = erreur atelier → mettre le
  numéro d'expédition en tracking.

### UPS
- Exclure Multitasker et Vitawave (comptes séparés).
- Télécharger les factures par paquets de 10 sur le site du transporteur.
- Remplacer les points par des virgules dans toutes les colonnes **sauf AX**.
- Colonne 21 = tracking, à trier de A à Z.
- Colonne D : vérifier vide/N-A → si pas de tracking, retirer la ligne.
- Nombre de colis : pour les expés multi-colis, une seule référence tracking
  « maître ».
- Vérifier les références de destination des expés selon le PDF Viticolis.
- Vocabulaire des erreurs d'import ERP :
  - **Mappage entité** = numéro de tracking qui n'existe pas dans l'ERP.
  - **Doublon consolidé** = déjà facturé le mois précédent.
  - **Doublon fichier** = cette expédition apparaît deux fois dans le
    fichier d'import.

### Mondial Relay
- Plusieurs fichiers ZIP par facturation ; transformer tous les points en
  virgules depuis le CSV de base envoyé par le transporteur.
- Une facture PDF **par pays** (avant : par pays ET par client — le format a
  changé).
- Les écarts sont souvent sur la facture France/La Ruche : la **collecte**
  n'est pas facturée aux clients, il faut l'ajouter dans la feuille contrôle.
- Europe uniquement → TVA fixe 0,2.
- Colonne mode envoi : remplacer **24RC par 24R** pour avoir des zones
  connues (sinon avaries d'import).
- Les frêts à zéro se suppriment, **en vérifiant** que les autres colonnes à
  facturer sont aussi à zéro (sinon ne pas supprimer).

### Colissimo
- Le fichier du mois précédent renvoie une erreur — à refaire.

### FedEx
- Fichier CSV téléchargé par la compta.
- Les modes d'envoi FICP sont en mauvais état (beaucoup de cas à traiter
  manuellement).
- Voir le wiki / dossier transports pour valider certaines notations —
  emplacement pas encore identifié dans le dépôt.

### Chronopost
- Deux factures Excel envoyées par le transporteur.
- Nouveau classeur : copier les infos des factures, trier les trackings de
  A à Z.
- **CAP** = taxe gazole, **ECO** = taxe éco, **SUR** = taxe sûreté.
- Colonne AF : un calcul par label, trié par la colonne L.
- Actualiser le TCD poids et le TCD général, en supprimant les ID clients ;
  tout en formule (pas de valeurs figées).
- **Deux contrats distincts** : Standard et 2Shop — non différenciés dans le
  PDF `FACTURATION EXCEL.pdf`, à prendre en compte si/quand Chronopost est
  automatisé.

## Autres notes (contexte à préciser)

- Le SMIC qui augmente doit être mappé dans la base de données.
- Normer les factures par code d'expédition (impacte uniquement la main
  d'œuvre) ; dans l'onglet TCD prépa des tarifs, catégoriser chaque prix,
  arrondir au centime.

---
*Source des notes par transporteur : `Notes_Facturation_Observations.md`
(observations terrain, 2026-07). Se référer à ce fichier pour le contexte
brut d'origine et les mises à jour futures.*
