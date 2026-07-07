# Analyse de la vidéo de process — Chronopost

> **Chronopost** = colis **express** (Groupe La Poste). Client « LA RUCHE LOGISTIQUE », plusieurs **sous-comptes**.
> Dépouillement de la vidéo `Process Facturation - Facture Chronopost.mp4` (9 min 19), croisée avec les fichiers réels (`facture_chronopost_<compte>_202606.xlsx`, classeur `2026_06_Facture Chronopost.xlsx`, PDF `CHRONOPOST_*.pdf`).
> C'est le process le **plus riche** (14 feuilles, multi-comptes, redistribution du gazole). Les `f_xxx` renvoient aux images.

---

## 1. Sources d'entrée

- **N fichiers Excel par sous-compte** : `facture_chronopost_<compte>_AAAAMM.xlsx` — feuille « Données », titre « Détail de facture Chronopost (Période MM/AAAA) / Compte <n°> - LA RUCHE LOGISTIQUE », **en-têtes en ligne 4**, données à partir de la ligne 5.
  - ⚠️ **Nombre de comptes variable selon le mois** : juin 2026 = 2 (`51291303`, `65481903`) ; mai 2025 = 3 (ajout `59272403`). → à paramétrer dynamiquement.
- **Colonnes source** : No Facture, Sous-compte, Date LT, Code postal départ/arrivée, Pays départ/arrivée, Ref Destinataire, Ref Expéditeur, No Groupage tarifaire, **Numero LT** (= n° de suivi), Groupage, **Type prestation**, TVA, Observations, **Zone Tarifaire**, **Poids**, **Produit**, **Montant HT**, Raison sociale.
- **Factures PDF** `CHRONOPOST_*.pdf` : servent à la réconciliation.
- Stockage : `…\$Facturation automatique\1 - Factures transporteurs + calculs\AAAA\AAAA MM\Chronopost\`.

---

## 2. Le classeur de travail `AAAA_MM_Facture Chronopost.xlsx` (14 feuilles)

| Feuille | Rôle |
|---|---|
| **Facture Chronopost** | Table centrale = brute consolidée + colonnes de calcul (split gazole / hors gazole, %, catégorie, totaux ; sous-totaux par poste en col. AF) |
| **Catégories** | Correspondance **Type prestation Chronopost → poste** (liste de validation : Adresse, Assurance, Colis volumineux, Corse, Droits et taxes, Frais facturation, Frêt, Gazole, Zones éloignées) |
| **Sous-comptes** | Mapping n° de sous-compte (0, 1, 2, 5…) → **client / Id client** (ex. 0=La Ruche, 1=Ker Sun, 2=Foutas) |
| **Bibliothèque transporteurs** | Lookup produit/zone Chrono → **mode envoi ERP / zone ERP / Transporteur** |
| **Pays TVA** | Pays → régime TVA (0,2 / 0) |
| **Zoning 2shop** | Code postal 2shop → zone |
| **TARIFS** / **cap à 5%** | Références de tarifs + **plafonnement d'un surcharge** (règle à confirmer) |
| **TCD poids** / **TCD** | Agrégations (par poids ; par Numéro LT avec ventilation par poste + colonnes ERP + redistribution gazole) |
| **Contrôle pdf** | Réconciliation TCD ↔ factures PDF |
| **Fichier import** | Sortie normalisée pour l'ERP |
| **Bilan clients** / **Avoir** | Synthèse par client ; demandes d'avoir |

Le classeur du mois précédent sert de **modèle** (formules et feuilles pré-construites) ; on fait « Enregistrer sous » pour le nouveau mois.

---

## 3. Déroulé

1. **Consolider** les N fichiers de sous-compte (copier-coller) dans la feuille `Facture Chronopost`. Lignes « spéciales » (forfaits, facturés à part par Chronopost) : `CAPI/CAPN/CAPO` = Surcharge Carburant, `ECORx` = Participation Éco-Responsable, `SURTx` = Sûreté colis, + `Transport`, `Supplement`.
2. **Reclasser** chaque ligne via la feuille `Catégories` (la colonne « Catégories » a une liste de validation).
3. **Mettre à jour le taux gazole du mois** : chercher le taux officiel (Google « taxe carburant chronopost »), le saisir dans **le SI `si.laruche-logistique.fr` ▸ [TRANSPORT] Transporteurs ▸ CHRONOPOST/CHRONO-BAL/CHRONO-INTER ▸ onglet « Taxe gasoil »** (une entrée **par mode de livraison**, avec période et %), puis le reporter dans le classeur (ex. **16,85 % routier / 27,55 % aérien**, mai 2025).
4. **Réconcilier** (feuille `Contrôle pdf`) : TCD par No Facture × Σ Montant HT vs total de chaque PDF ; **écart doit = 0,00 €**.
5. **Agréger** (TCD / TCD poids) + ajouter les colonnes ERP par RECHERCHE dans `Bibliothèque transporteurs`, puis **redistribuer le gazole** au prorata.
6. **Construire `Fichier import`** (formules depuis le TCD) → export **CSV** dans `2 - Fichiers csv import` → import via **[TRANSPORT] - Imports**.

---

## 4. Règles de calcul (le distinctif Chronopost)

**Mapping prestation → poste** (feuille `Catégories`) :
| Prestation Chronopost | Poste |
|---|---|
| Surcharge Carburant (CAPI/CAPN/CAPO) | **Gazole** |
| Transport | **Frêt** |
| Participation Éco-Responsable (ECORx) | **Frêt** |
| Sûreté colis (SURTx) | **Sûreté** (regroupé) |
| Zones difficiles d'accès / Correction d'adresse | **Adresse** |
| Supplément Corse | **Corse** |
| Supp. Zone Internationale / Retour Expéditeur / Zones éloignées | **Zones éloignées** |
| Assurance | **Assurance** |

**Formules repérées :**
- Sous-totaux par poste (col. AF) : `=SOMME(T2:T10)`, `=SOMME(T11:T17)`…
- Lookup mode envoi ERP : `=SI(NB.SI('Bibliothèque transporteurs'!C:C;F1141)=0;"inconnu";RECHERCHE(F1141;'Bibliothèque transporteurs'!C:C;'Bibliothèque transporteurs'!A:A))`
- **Pool gazole à redistribuer** (en-tête « Gazole + facturation + divers ») : `=SOMME(N:N)+SOMME(P:P)+SOMME(O3:O11)`
- 🔑 **Redistribution du gazole au prorata du frêt** : `=C1141/SOMME(C:C)*(SOMME(C:C)+$W$2)` — chaque ligne reçoit sa **quote-part du pool gazole** selon son poids dans le total *hors gazole* (col. C).

> ⚠️ **Différence majeure avec Kuehne/Delivengo** : chez Chronopost, le gazole est facturé par le transporteur en **lignes forfaitaires séparées** (CAP…), qui sont **mises en pool puis redistribuées** sur chaque expédition au prorata du frêt — ce n'est ni un résiduel (Kuehne) ni un taux par pays (Delivengo).

---

## 5. Contrôles + export

- **Contrôle PDF** : Σ HT par No Facture (TCD) = total imprimé sur chaque facture PDF → **écart 0,00 €** obligatoire avant de continuer (total ≈ 12 970,60 €).
- **Cohérence gazole/ERP** : le % gazole du classeur doit correspondre aux taux saisis dans l'ERP par mode de livraison.
- **Export** : feuille `Fichier import` (Transporteur, Date validité, Réf.1/2, Id client, N° Tracking, Nom, E/P, Pays, Zone, Nbr Colis, Poids, mode envoi, TVA, Droits et taxes, Assurance, Zones éloignées, Colis volumineux, Adresse, plus-value, Gazole, Frêt…) → **CSV** dans `2 - Fichiers csv import` → import via **[TRANSPORT] - Imports**.
- **Garde-fous annotés** : « **2SHOP : ne pas mettre de gazole dans les coûts de revient** », « **Renseigner "Pays"** ».

---

## 6. Règles métier à coder (automatisation Chronopost)

1. **Ingestion multi-comptes** : charger un nombre **variable** de fichiers `facture_chronopost_<compte>_AAAAMM.xlsx` (en-têtes en ligne 4), les consolider.
2. **Reclassement** Type prestation → poste via la table `Catégories`.
3. **Client** = lookup Sous-compte → Id client (`Sous-comptes`).
4. **Colonnes ERP** (mode envoi, zone) = lookup `Bibliothèque transporteurs` (produit/zone Chrono).
5. **Gazole** : pooler les lignes de surcharge carburant + **redistribuer au prorata du frêt** par expédition ; taux par mode de livraison (maintenu par mois). Gérer la **remise gazole** (Gazole réel vs avant remise).
6. **Réconciliation** : Σ HT par facture = total PDF (garde-fou, écart 0).
7. **Cas particuliers** : `cap à 5%` (plafonnement — à préciser), Corse, 2SHOP (pas de gazole en coût), lignes « inconnu » (forfaits sans produit/zone).

---

## 7. Points à confirmer

1. **`cap à 5%`** et **`TARIFS`** : logique exacte (plafonnement de quoi à 5 % ?) — feuilles non ouvertes dans la vidéo.
2. **Remise gazole** : règle exacte (Gazole / Gazole réel / Gazole avant remise).
3. **Suffixes I/N/O** des forfaits (CAP**I**/CAP**N**/CAP**O**…) = familles produit (aérien / route / outre-mer ?).
4. **Étape CSV finale + mapping ERP** : non filmés (nom exact du CSV, colonnes).
5. **`Avoir` / `Bilan clients`** : rôle décrit par contexte, feuilles non affichées.
