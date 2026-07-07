# Facturation transporteur Kuehne+Nagel — Documentation A→Z

*Document de référence pour l'automatisation. Rédigé à partir du reverse-engineering du classeur Excel de calcul, des 2 CSV bruts K&N, des 2 factures PDF, des docs de process (.docx) — puis vérifié indépendamment (audit multi-agents + prototype Python validé 187/187 lignes contre l'export ERP réel de juin 2026).*

---

## 0. Résumé en une phrase

Chaque mois, Kuehne+Nagel (K&N) facture La Ruche Logistique via **2 fichiers CSV** (une facture de transport + une facture d'« événements ») et **2 PDF** ; le process consiste à **traduire ces ~183 colonnes techniques K&N en un fichier d'import ERP** de 23 colonnes (une ligne par expédition, ventilée en 8 postes de charge), **réconcilié au centime avec les PDF**, pour que l'ERP puisse ensuite **refacturer les clients de La Ruche**.

---

## 1. Le contexte métier (le « pourquoi »)

- **La Ruche Logistique** expédie des colis (vin — univers « Viticolis ») pour ses clients, via plusieurs transporteurs : UPS (principal), Geodis, **Kuehne+Nagel**, BLS…
- **Deux niveaux de facturation :**
  1. **Achat** : le transporteur (K&N) facture La Ruche pour le transport réalisé.
  2. **Vente / refacturation** : La Ruche refacture ce transport à ses propres clients, à un **tarif de vente** (indexé par zone), avec une marge.
- **Le fichier d'import ERP est le pivot** : on y injecte, ligne par ligne, le coût/refacturation de chaque expédition. C'est ce fichier que le présent process produit.
- **Rôle spécifique de K&N :**
  - Flux **EDI** : chaque jour La Ruche envoie ses expéditions à K&N ; les données reviennent avec une **Ref EDI** (= le n° d'expédition interne de La Ruche).
  - K&N fait partie du dispositif **Navette** (tarifs négociés, factures internes) : sa facture Navette est **incluse dans le classeur Excel** (contrairement à Geodis, PDF à part).
  - **Taxe Gasoil** : lue **directement dans l'Excel envoyé par K&N** (ni site web, ni ERP) — spécificité K&N.
  - Cycle **mensuel** ; la facture K&N est récupérée/préparée par Pauline et déposée sur le serveur.

---

## 2. Glossaire (à connaître avant de lire la suite)

| Terme | Définition |
|---|---|
| **EDI** | Échange de Données Informatisé : flux automatisé d'expéditions envoyé quotidiennement au transporteur. |
| **Messagerie** | Colis groupés, tarifés au poids (taxable). Sous-types K&N : `STANDARD`, `FIRST`. |
| **Affrètement (AFFR)** | Camion / lot complet, tarifé à la tonne — logique et taux gazole différents de la messagerie. |
| **Ref EDI** | Identifiant d'expédition **de La Ruche** (ex. `5001333` en messagerie, `EXP20260603-2777654` sur les événements, **vide** en affrètement). |
| **ComRefExpedition** | Référence **K&N** de l'expédition (ex. `26A050838106`). Devient **Réf.1** dans l'import. |
| **N° Tracking** (import) | = **Ref EDI si présente, sinon ComRefExpedition**. |
| **TCD** | Tableau Croisé Dynamique (pivot). Ici : **outil de contrôle uniquement**, PAS de production de l'import. |
| **ERP** | Progiciel de gestion dans lequel on injecte le fichier d'import pour facturer. |
| **Zone** | Clé de tarification `DeptExp + France/étranger + pad2(DeptDest)` (ex. `21France78`). |
| **Poids taxable** | Poids facturable (peut différer du poids réel). ⚠️ L'import utilise le **Poids réel**, pas le taxable. |
| **E / P** | Entreprise / Particulier. Conditionne la **plus-value BtoC** (surcoût livraison particulier). |
| **Souffrance** | Colis bloqué / non livrable en attente de traitement. |
| **Présentation** | Nouvelle tentative de livraison, facturée. |
| **Correction de poids** | Réajustement après pesée réelle. |
| **Taxe Gasoil / surcharge gazole** | Surcoût carburant en %, indexé sur un prix de référence (juin 2026 : 1,7834 €/l). |
| **Incoterm DAP** | *Delivered At Place* : répartition des frais/responsabilités transport & douane. |
| **Demande d'avoir** | Correction/remboursement en faveur du client. |
| **Postes de charge** | Les 8 rubriques cibles : Fret, Zones éloignées, Colis volumineux, Adresses, Assurance, Droits & taxes, Plus-value BtoC, Gazole. |
| **Colonnes `T_*`** | ~120 colonnes de détail tarifaire du CSV K&N, reclassées vers les 8 postes. |

---

## 3. Les données d'entrée

| Fichier | Contenu | Format |
|---|---|---|
| `FcCSV#...F2606017122....csv` | **Facture principale** : 181 lignes (169 messagerie + 9 affrètement + **3 lignes de synthèse**). | latin-1, séparateur `;`, décimale `,`, 183 colonnes. |
| `FcCSV#...F2606017123....csv` | **Facture « événements »** : 9 lignes (souffrance, présentation, correction poids, portuaire, aéroport). | idem |
| `..._F2606017122_ORIGINAL.pdf` | Facture PDF officielle (totaux, TVA, TTC, ventilation gazole). | PDF |
| `..._F2606017123_ORIGINAL.pdf` | Facture PDF des événements. | PDF |
| Excel envoyé par K&N | Source de la **Taxe Gasoil**. | xlsx |
| `2026_06_Facture Kuehne.xlsx` | **Classeur de calcul** (6 onglets : Barème gazole, Fichier Kuehne, TCD, Kuehne_Import, Bilan clients, Demande avoir). | xlsx |
| `2026_06_Kuehne_Import.csv` | **Sortie** : fichier d'import ERP (187 lignes, 23 colonnes). | latin-1, `;`, `,` |
| 3 fichiers `.mp4` | **Vidéos de process** (préparation import ; facture + fichier import ; **prix de vente + doublon prestation**). Sources primaires. | vidéo |

---

## 4. Le processus manuel actuel (ce que fait l'opérateur, A→Z)

Le classeur a 6 onglets ; le flux est **CSV bruts → onglet *Fichier Kuehne* (reclassement) → *TCD* (contrôle) → *Kuehne_Import* → fichier CSV d'import**.

### Phase A — Préparation
1. Ouvrir le classeur du **mois précédent**, l'enregistrer sous le nouveau mois.
2. `CTRL+MAJ+L` sur la 1ʳᵉ ligne pour vérifier qu'aucune donnée résiduelle ne subsiste.
3. Récupérer la **Taxe Gasoil K&N** dans l'Excel fourni par K&N.

### Phase B — Collage & calcul (onglet *Fichier Kuehne*)
4. **Coller** les données des CSV K&N.
5. **Trier les Ref EDI de A→Z**.
6. Vérifier le **nombre de lignes** (aucune perte).
7. **Étendre les colonnes B→L** (formules de reclassement).
8. Supprimer les ID clients.

### Phase C — Consolidation & contrôle (onglet *TCD*)
9. Mettre à jour le TCD (colonne S).
10. Effacer la colonne A ; étendre T/U/V/W.
11. **Contrôler les 2 PDF** : y saisir le **montant taxable** et comparer au calcul.
12. Vérifier tarif PDF = tarif TCD.

### Phase D — Import (onglet *Kuehne_Import* → fichier CSV)
13. Vérifier le nb de lignes + étendre les colonnes (aligné sur le TCD).
14. Changer la **date** de la 1ʳᵉ ligne.
15. Ouvrir le fichier d'import du mois précédent, l'enregistrer sous le nouveau mois, **supprimer les anciennes données** (garder la 1ʳᵉ ligne comme trame).

### Phase E — Validations (check-list)
16. Transporteur = Kuehne partout · date au bon format/mois · trackings corrects.
17. **E/P présent partout** · **zones : pas d'étranger, pas de zone 0/inconnue** · **colis ≠ 0** · **poids ≠ 0 (1 décimale, ARRONDI.SUP sinon)** · **fret ≠ 0 €**.
18. **Égalité des sommes par poste** entre fichier de calcul et import.
19. **TVA** = 0,2 France/UE, 0 hors-UE.

---

## 5. La logique de transformation (le cœur — règles vérifiées)

### 5.1 Identifiants
- **N° Tracking** = `Ref EDI` si non vide, **sinon** `ComRefExpedition`. (174/187 via Ref EDI, 13/187 fallback.)
- **Réf.1** = `ComRefExpedition` (toujours).
- ⚠️ La colonne « Ref EDI » du CSV **change de format** selon le type : `5001333` (messagerie), `EXP...` (événements), **vide** (affrètement).

### 5.2 Reclassement des charges — LA règle centrale
Chaque ligne CSV porte `Mt HT (hors frais)`, `Mt HT (avec frais)` et ~120 colonnes `T_*`. On les regroupe en **8 postes** (formules réelles des colonnes C→L de *Fichier Kuehne*) :

| Poste ERP | = somme des colonnes CSV |
|---|---|
| **Droits et taxes** | `Is_Douane` + `Is_Douane_Export_ST` + `Is_Douane_Import_ST` + `T_DOUANE_ANDORRE` + `T_DOUANE_CANARIES` + `T_DOUANE_EXPORT` + `T_DOUANE_IMPORT` + `T_DOUANE_INSPECT` + `T_ATTENTE_DOUANE` + `T_IMMO_DOUANE` |
| **Assurance** | `T_ASSUR` |
| **Zones éloignées** | `T_ILE` + `T_LIV_SUPERMARCHE` + `T_LIV_URBAINE` + `T_LIV_AEROPORT` + `T_FERRY` + `T_LIV_PORTUAIRE` |
| **Colis volumineux** | `T_CORR_POIDS` |
| **Adresses** | `T_GESTION_PARTICULIER` + `T_LIV_ETAGE` + `T_SOUFFRANCE` |
| **Plus-value B2C** | `T_EMBALLAGE` + `T_PALETTE_EUR` |
| **Frêt** | `Mt HT (hors frais)` + **tous les autres `T_*`** (approche, admin, manutention, préavis, `T_ENERGIE`…) — c'est un **fourre-tout** |
| **Gazole** | `Mt HT (avec frais)` − (somme des 7 postes ci-dessus) → **résidu** |

**Propriétés vérifiées :** les 120 colonnes `T_*` forment une **partition parfaite** (20 vers les 6 postes spécifiques + 99 vers Frêt + 1 = `T_GAZOLE` isolé). Aucune colonne oubliée, aucune comptée deux fois.
`Total hors GO` = somme des 7 postes ; `Total + GO` = les 8 = `Mt HT (avec frais)`.

⚠️ **Piège** : des codes sémantiquement « taxe » comme `T_EKAER` (taxe routière) ou `T_MAUT` (péage allemand) tombent dans **Frêt**, pas dans Droits & taxes. Tout **nouveau** code tarifaire K&N tombera silencieusement dans Frêt.

### 5.3 Le gazole
- La colonne `T_GAZOLE` est **vide** sur toutes les lignes : le gazole n'est **pas** ventilé par expédition.
- Il apparaît sous forme de **2 lignes de synthèse dans le CSV brut** (facture principale), identifiées par le libellé en colonne 17 « Agence Saisie/Chgt » :
  - `TAXE SURCOUT GAZOLE MESS 1,7834 EUR/l : 11,70 %` = **1 176,27 €**
  - `TAXE SURCOUT GAZOLE AFFR 1,7834 EUR/l : 8,43 %` = **249,11 €**
  - (+ 1 ligne `TAXE FIXE DE FACTURATION` = 0 €)
- Ces lignes ont un **tracking vide** → **exclues de l'import** (voir 5.4). Le gazole (**1 425,38 €**) est donc **refacturé globalement, à part** ; la colonne « Taxe Gasoil » de l'import reste **vide**.
- Le calcul « résidu » du poste Gazole capte l'intégralité de ces montants (sur les lignes normales il vaut 0).
- Les taux (MESS 11,70 %) proviennent du **barème gazole** (onglet dédié : prix €/l → %) ; l'affrètement (8,43 %) suit un indice distinct.

### 5.4 Granularité : 1 ligne CSV = 1 ligne import
- L'export ERP est **1:1** avec les lignes CSV brutes — **AUCUN regroupement par tracking**. La TCD (pivot par tracking) ne sert **qu'au contrôle**.
- **Seul filtrage** : on **exclut les lignes dont le tracking est vide** (Ref EDI ET ComRef vides) = les lignes de synthèse.
- Identité : **190 lignes brutes − 3 sans tracking = 187 = lignes ERP** (exact).
- ⚠️ **7 trackings apparaissent dans les 2 CSV** (principale + événements) et donnent **2 lignes ERP distinctes** : un regroupement naïf par tracking les fusionnerait et **casserait les totaux**.

### 5.5 Zone
`Zone = DeptExp & (Pays_dest=="FR" ? "France" : "étranger") & pad2(DeptDest)` — ex. `21France78`, `21étranger10` (BE).
- Dept récupéré par **XLOOKUP(tracking, colonne Ref EDI, colonne Dept)**. Si le tracking vient de ComRefExpedition (Ref EDI vide, cas affrètement), le lookup **échoue → 0 → `0France00`**.
- **Seul le dept destinataire est paddé** sur 2 chiffres ; le dept expéditeur reste brut.
- Le **pays** provient d'une source distincte (reste correct même quand le dept tombe à 0).
- ⚠️ **Bug latent** : les 13 lignes `0France00` partent en facturation avec une **zone tarifaire factice** (dept perdu). À corriger dans le logiciel (fallback lookup sur ComRef, ou alerte « zone inconnue »).

### 5.6 Champs fixes (K&N)
`Transporteur=KUEHNE` · `E/P=E` · `mode envoi=ST` · `Date validité tarif = 1er du mois` (187/187, aucune exception).
`TVA = 0,2` France/UE (Belgique incluse) / `0,0` hors-UE (US, GB, CH…). La branche hors-UE est **documentée mais non exercée** en juin 2026.
`Poids` = colonne **Poids** (jamais Poids taxable). `Nbr Colis` = colonne **Nb colis**.

### 5.7 Réconciliation — l'équation maîtresse
```
Total hors gazole (13 335,88)  +  Gazole (1 425,38)  =  14 761,26 €
        ▲ (= import ERP)              ▲ (= 1176,27 + 249,11)
   = Facture F2606017122 (14 433,99, taxable)  +  Facture F2606017123 (327,27)
```
TVA 20 % et TTC : 17 320,79 € (F...122) + 392,72 € (F...123).

### 5.8 Les « événements » (2e facture) — illustration du reclassement
| Événement PDF | Colonne `T_*` | Poste ERP |
|---|---|---|
| SOUFFRANCE (77,10) | `T_SOUFFRANCE` | Adresses |
| CORRECTION_POIDS (30,00) | `T_CORR_POIDS` | Colis volumineux |
| LIVRAISON_PORTUAIRE (21,00) | `T_LIV_PORTUAIRE` | Zones éloignées |
| LIVRAISON_AEROPORT (21,00) | `T_LIV_AEROPORT` | Zones éloignées |
| PRESENTATION (178,17) | (T_* de type manutention) | **Frêt** (fourre-tout) |

---

## 6. Méthodologie — les manipulations d'investigation (reverse-engineering)

1. **Exploration** du dossier (`ls`) : repérage des CSV, PDF, xlsx, docx, mp4.
2. **Extraction des .docx** : un `.docx` est un zip → `unzip -p fichier.docx word/document.xml` puis suppression des balises XML → texte du process.
3. **Lecture du .xlsx** avec `openpyxl`, en **deux passes** :
   - `data_only=True` → les **valeurs** (résultats).
   - `data_only=False` → les **formules** (la logique métier réelle).
4. **Décodage des formules** : mapping colonne Excel → colonne CSV via l'offset (colonne M d'Excel = colonne 1 du CSV) pour reconstituer les 8 postes.
5. **Analyse des pivots** : lecture des `pivotCacheDefinition` (source du TCD = *Fichier Kuehne* ; source de *Bilan clients* = TCD).
6. **Parsing des CSV** : attention **encodage latin-1**, séparateur `;`, décimale `,`.
7. **Parsing des PDF** avec `pypdf` → totaux officiels, TVA, ventilation gazole.
8. **Prototype Python** rejouant tout le pipeline depuis les CSV bruts → comparaison **multiset** avec l'export ERP réel : **187/187 lignes identiques, 0 écart**.
9. **Vérification adversariale multi-agents** (workflow) : 4 agents ré-dérivent chacun une facette pour tenter de casser les règles, 1 extrait le process manuel, 1 critique la complétude. Double preuve : lecture des formules **+** validation empirique contre la sortie réelle.

> Leçon de méthode : on prouve une règle métier de **deux façons indépendantes** (lire la formule *et* reproduire la sortie réelle). Quand les deux concordent, la confiance est maximale.

---

## 7. Chiffres de contrôle (juin 2026)

| Poste | Montant (€) |
|---|---|
| Frêt | 13 137,78 |
| Adresses | 77,10 |
| Zones éloignées | 91,00 |
| Colis volumineux | 30,00 |
| Assurance / Droits & taxes / Plus-value B2C | 0,00 |
| **Sous-total import (hors gazole)** | **13 335,88** |
| Gazole (hors import) | 1 425,38 |
| **TOTAL HT** | **14 761,26** |

187 lignes d'import · 180 trackings uniques + 7 en doublon · 46 zones distinctes (top `21France34`).

---

## 8. Pièges à éviter pour l'automatisation

- **Encodage** : lire en latin-1 (pas UTF-8) ; **mapper les colonnes par INDEX**, pas par libellé (accents cassés + espaces de garde ` Zone `).
- **Décimale** : `,` → `.` au parsing, et `.` → `,` à l'écriture de l'import.
- **Exclure les lignes à tracking vide** (sinon +3 lignes et **faux gazole de 1 425,38 €**). Critère = « tracking vide », PAS « ligne gazole ».
- **Ne jamais regrouper/dédupliquer par tracking** (fusionnerait les 7 paires inter-fichiers).
- **Zone `0France00`** = lookup échoué → prévoir fallback/alerte.
- **Padding** : seul le dept destinataire est paddé sur 2 chiffres.
- **Nouveau code `T_*`** → tombe dans Frêt : prévoir une alerte sur codes inconnus.
- **Poids** (pas Poids taxable) ; interdits : poids 0, colis 0, fret 0 €.
- **Colonne « Taxe Gasoil » de l'import doit rester VIDE** ; colonne « Nb colis » (22) vide, « Nbr Colis » (10) remplie.
- **Arrondi TVA** au centime.
- **TVA hors-UE (0)** jamais testée en juin 2026 → à valider sur un mois avec US/GB/CH.

---

## 9. Zones d'ombre à éclaircir (prochaines investigations)

1. **Refacturation vente & marge** : grille de vente par zone, lien achat→vente (le « pourquoi » final). → **vidéo « Prix de vente - Doublon prestation »**.
2. **Mécanisme concret de refacturation du gazole** (ligne globale ? prorata du fret ?).
3. **Origine de l'ID client** (supprimé de l'import) et comment la refacturation retrouve le client (via tracking / Bilan clients).
4. **Barème gazole** : indice de référence, pourquoi MESS ≠ AFFR.
5. **Dérivation E/P** (forcé à E pour K&N ? depuis le nom ? un flag ?).
6. **Règle 1Z79** (colis à sortir de l'import → demande d'avoir) — spécifique UPS, non testée ici.
7. **Spec normative des 23 colonnes** de l'import attendue par l'ERP.
8. **Contradiction « pas d'étranger » vs ligne Belgique (UE)** : étranger UE toléré, hors-UE interdit ?

---

## 10. Prochaines étapes logicielles

1. Durcir le module K&N (moteur de validation + file d'exceptions).
2. Externaliser la table de reclassement en **config par transporteur** (YAML) et valider sur un 2ᵉ transporteur (UPS/Geodis).
3. Traiter les 2 maillons manuels : parsing PDF (totaux/gazole) et attribution client.
