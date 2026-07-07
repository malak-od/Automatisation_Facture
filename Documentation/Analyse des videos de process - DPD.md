# Analyse des vidéos de process — DPD

> **DPD** = transporteur de **colis** (France/Europe), agence « DPD 21 » (`dpd021`).
> Dépouillement des **5 vidéos** `Process Facturation - Facture DPD_1..5.mp4` du dossier `Transporteurs/DPD/`.
> ⚠️ Le dossier DPD ne contient **que les vidéos** (pas de fichiers de données) : ce document repose sur l'analyse des images. Les `f_xxx` renvoient aux images.

---

## Inventaire des vidéos (5 parties séquentielles d'un même process)

| Partie | Durée | Contenu |
|---|---|---|
| DPD_1 | 5 min 57 | Récupération des fichiers source (mails) + import Power Query |
| DPD_2 | 2 min 44 | Consolidation Power Query + nettoyage des sous-totaux |
| DPD_3 | 4 min 32 | Empilage dans `Facture DPD` + calcul des postes + réconciliation |
| DPD_4 | 0 min 16 | (court) navigation `Import ERP` ↔ `Facture DPD`, lecture d'un total |
| DPD_5 | 4 min 44 | Extraction du fichier import + nettoyage anomalies + enregistrement |

Outil : **Excel** (+ Power Query + Outlook + Notepad++). Cible finale : **`si.laruche-logistique.fr`** (import non filmé).

---

## 1. Récupération de la source (partie 1)

- DPD envoie les données **par email**, expéditeur **`general@dpd021.dpd.fr`** (ou « Agnès »), **un mail par compte/client**. Objet : « complément de facture DPD 21 (N) ».
- Pièce jointe : **`complément_facture02100_2505_0XXXX-31.05.2025.xlsx`** (`02100` = préfixe compte, `2505` = 2025-05, `0XXXX` = n° de facture). ~15 fichiers/mois.
- ⚠️ Le mail précise « **le document n'est pas une facture** » (la facture PDF est séparée, utilisée pour la réconciliation).
- ⚠️ Certaines pièces jointes nommées `.xlsx` sont en réalité des **archives (WinRAR/`.gz`)** → **décompression obligatoire** (« Extraire tout »).
- En réalité, ces fichiers sont des **CSV délimités `;`** (visible dans Notepad++). En-tête : `Type (Slave Export);No facture;Date de facture;Valable du;Valable au;No de compte;Sous-compte;…`
- Stockage : `…\$Facturation automatique\1 - Factures transporteurs + calculs\AAAA\AAAA MM\DPD\`.

---

## 2. Consolidation (parties 2-3)

- **Power Query** : *Données ▸ Obtenir des données ▸ À partir d'un dossier* → **Combiner** (Sheet1 de chaque fichier) → une table unique **`DPD`** (~97 colonnes, ~2304 lignes), avec une colonne `Source.Name`.
- Les fichiers `complément` contiennent : des **lignes de détail** (`Type (Slave Export)` = 1 par expédition) **+ des lignes de sous-total** (« Total expéditions », « Total poids », « Dont total frais de gazoil »…).
- **Nettoyage** : filtre sur `Type (Slave Export)` pour **exclure les lignes de sous-total (texte) et les vides** (42 lignes parasites sur 2304).
- Collage dans le classeur de travail **`AAAA_MM_Facture DPD.xlsx`**, onglet **`Facture DPD`**.

**Colonnes brutes DPD** — identité : No facture, No de compte, Sous-compte, expéditeur/destinataire (nom+adresse+pays), **N° Colis** (tracking à 14 chiffres), **DPD ID** (très grand nombre, affiché en notation scientifique), Votre référence 1/2/3, Nombre de colis, **Poids / Poids initial / dimensions / Poids volumétrique / Type poids (« Poids repesé »)**, Colis refacturé.
**Colonnes de surcharges DPD** : Prix transport, Supplément, **Coût de la VD** (valeur déclarée), Taxe fixe, Taxe Consolidat, **Indexation gasoil**, Indexation kérosène, Supplément prédict, **Supplément DPD Sécure**, Contribution Logistique Responsable, **Frais de tenue de compte**, Prix cumulé, TVA + compteurs de statut (Ramasse programmée, Liv. domicile privé, Étiquette manuelle, Consigne, Taxe triangulaire).

---

## 3. Reclassement en postes (onglet `Facture DPD`)

- Classeur `AAAA_MM_Facture DPD.xlsx`, onglets : **Zoning, Bilan factures, Facture DPD, Import ERP, Bilan clients, Tarifs achat DPD** (même **famille de modèle que Chronopost**).
- À **gauche** de `Facture DPD` : les **postes calculés** (surlignés vert) — `Total, Total hors GO, Total GO, Client, Colis, Adresses, Frêt, BtoC, Retour, Gazole, Frais dossier, Zones éloignées`. (**« GO » = gasoil** ; Total GO = avec taxe gasoil, hors GO = sans.)
- À **droite** : les **données brutes** recollées. Les postes de gauche sont des **formules pointant les colonnes brutes** (`=AV1755`, `=BD2247` — décalage de colonne).
- **Rechercher/Remplacer global** (≈ 3397 remplacements) pour normaliser les valeurs de surcharge.
- **Mapping surcharge → poste** : *pré-câblé dans le modèle, formules non lisibles à l'écran*. Correspondances **déduites** (à confirmer) : Indexation gasoil → Gazole ; Supplément DPD Sécure → Assurance ; Coût de la VD → Droits et taxes / Assurance ; Frais de tenue de compte → Client ; Supplément prédict → Adresses ; Colis refacturé / Poids repesé → Colis volumineux.

---

## 4. Réconciliation (onglet `Bilan factures`)

- **TCD** : lignes = `No de compte`, valeurs = **Σ Total GO** et **Σ Total hors GO**.
- Table manuelle **« Factures pdf »** à côté : saisie des **montants HT des factures DPD PDF** par compte.
- **Contrôle** : Σ Total GO (calculé) doit égaler le montant PDF par compte → surlignage **vert** si ça matche. Total ≈ **13 795,17 € HT / 15 462,24 € TTC**.
- **1 compte DPD (No de compte) = 1 client** : 9748 = LA RUCHE LOGISTIQUE, 9779 = NALKAA, 9808 = HYDRATIS, 9836 = NUBIANCE, 9902 = FLUID MARKET, 9854 = LABORATOIRES OMI, etc.

---

## 5. Fichier import + nettoyage (onglets `Import ERP` → export, parties 4-5)

- L'onglet **`Import ERP`** reprend le **gabarit ERP normalisé** (Transporteur=DPD, Date validité tarif=01/MM, Réf.1/2, Id client, N° Tracking, Nom, E/P=P, Pays=FR, Zone, Nbr Colis=1, Poids, mode envoi, TVA=0,2, Droits et taxes, Assurance, Zones éloignées, Colis volumineux, Adresses, Frêt, plus-value, Gazole, nb colis).
- Alimenté par **formules** depuis `Facture DPD` (ex. `=SI('Facture DPD'!G2721=0;"";'Facture DPD'!G2721)`).
- **Contrôles qualité (annotés en rouge)** avant export :
  - **Format tracking** : le **DPD ID** s'affiche en notation scientifique (`2500215894255270000`) ; le **N° Colis** est le bon (14 chiffres). Reformater (Standard → Nombre/Texte) et nettoyer préfixe/espaces : `021-105124544 2` → `1051245442`.
  - **Tracking = 0** : filtre → suppression des lignes parasites (9 → 3) puis **complétion manuelle** du tracking depuis DPD ID / N° Colis.
  - **Poids = 0** (« ATTENTION AUX POIDS A 0 ») : revue et re-remplissage.
  - Suppression des lignes « **inconnu** » / `#REF!` (lookups zone/tarif en échec).
- **Extraction** de l'onglet vers un **classeur autonome `AAAA_MM_DPD_Import`** (mono-feuille) → enregistré dans `…\2 - Fichiers csv import\AAAA MM\` (vraisemblablement en **CSV `;`**, l'étape de choix du type n'est pas filmée).
- Import ensuite dans **`si.laruche-logistique.fr`** (non filmé).

---

## 6. Règles métier à coder (pour l'automatisation DPD)

1. **Ingestion** : récupérer les mails DPD (`general@dpd021.dpd.fr`), **décompresser** les pièces jointes archivées, parser les **CSV `;`** `complément_facture…`, **fusionner tous les comptes** (équivaut au Power Query « combiner »).
2. **Filtrage** : exclure les lignes de **sous-total** (`Type (Slave Export)` non numérique) et les vides.
3. **Reclassement** surcharges → postes (Gazole, Assurance, Adresses, Frêt, Client…) — *mapping exact à récupérer dans les formules du classeur `Facture DPD`*.
4. **Normalisation tracking** : utiliser **N° Colis** (14 chiffres), retirer préfixe/espaces, éviter la notation scientifique. *(Même problème que Delivengo → règle générique.)*
5. **Gestion poids = 0** et lignes sans données (à exclure/corriger).
6. **Client** = lookup `No de compte` → client (table `Sous-comptes` / `Bilan clients`).
7. **Réconciliation** : Σ Total GO par compte = montant HT du PDF DPD correspondant (garde-fou).
8. **Règle spéciale** : client **2SHOP** → **ne pas mettre le gazole dans les coûts de revient**.

---

## 7. Points à confirmer

1. **Mapping exact surcharge DPD → poste** : à extraire des formules réelles de l'onglet `Facture DPD` (non lisibles dans les vidéos).
2. **Format d'enregistrement final** : xlsx affiché vs dossier « csv import » — confirmer CSV `;` + encodage.
3. **Zoning / Tarifs achat DPD / zones éloignées** : onglets de lookup cités mais formules non montrées (zone, tarif d'achat).
4. **Import SI + avaries DPD** : non filmés (s'arrête à l'enregistrement du fichier).
5. **`Type (Slave Export)`** : signification exacte du champ (indicateur de type de ligne).
