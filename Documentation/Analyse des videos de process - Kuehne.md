# Analyse des vidéos de process — Kuehne

> Complément à `Documentation_Facturation_Kuehne.md` (doc maîtresse : contexte, règles de reclassement, gazole, zone, réconciliation).
> Ce document **dépouille les 3 vidéos `.mp4`** du dossier `Transporteurs/Kuehne/`, image par image, pour en extraire les étapes manuelles et les règles métier — en particulier le **volet VENTE / avaries** qui restait une zone d'ombre.
>
> Méthode : extraction d'images (ffmpeg, 1 img/12 s ou détection de changement d'écran) puis lecture par agents. Les n° `f_xxx` renvoient aux images.

---

## Inventaire des vidéos

| Vidéo | Durée | Outil montré | Ce qu'elle apporte |
|---|---|---|---|
| `KUEHNE_1_Preparation fichier import.mp4` | 9 min 22 | **Excel** | Construction du fichier import (achat) |
| `Process Facturation - Kuehne - Facture + Fichier import.mp4` | 15 min 44 | **Excel** | Vue d'ensemble + contrôles qualité |
| `Process Facturation - Kuehne - Prix de vente - Doublon prestation.mp4` | 3 min 20 | **Logiciel interne** `si.laruche-logistique.fr` | **VENTE + avaries** (nouveau) |

**Constat clé** : les 2 premières vidéos confirment le process Excel *achat* déjà documenté et déjà porté en code (pipeline validé 187/187). La 3ᵉ montre un **autre système** — le logiciel interne où se fait la **refacturation client (vente)** et la gestion des **avaries** : c'est du contenu neuf.

---

## Vidéo 1 & 2 — Process Excel (ce qu'elles CONFIRMENT)

Rien de contradictoire avec la doc maîtresse. Points confirmés « au pixel » :

- **Onglets du classeur** : `Bareme gazole | Fichier Kuehne | TCD | Kuehne_Import | Bilan clients | Demande avoir correction poids`.
- **Collage** du CSV brut dans `Fichier Kuehne` à partir de la **colonne M** ; colonnes **A→L = postes calculés** (`A=ID Client, B=Tracking, C=Total hors, D=Total+GO, E=Droits et taxes, F=Assurance, G=Zones éloignées, H=Colis volumineux, I=Adresse, J=Frêt, K=Plus value, L=Gazole`).
- **Tracking** : `B = SI(RefEDI<>"";RefEDI;ComRefExpedition)`.
- **Gazole** : `L = (Mt HT avec frais) − (Total hors)` par ligne (vérifié : `215 − 196,23 = 18,77`).
- **Réconciliation** (bloc de contrôle du TCD) : saisie manuelle des **totaux HT des 2 PDF**, égalité `Total avec frais (TCD) = Σ HT PDF` → bascule **« Import OK »** ; colonne **« Autres frais » = `Mt avec frais − Σ postes ventilés`**, rouge si ≠ 0 (surcharge non reclassée à corriger).
- **Export** : figer en valeurs → *Enregistrer sous → CSV (séparateur « ; »)* dans `2 - Fichiers csv import\AAAA\AAAA MM\`.

### Détails manuels utiles (à intégrer/automatiser côté contrôles)
- **Nom du destinataire** rempli par `RECHERCHEX(Tracking; 'Fichier Kuehne'!E:E; 'Fichier Kuehne'!Z:Z)`, avec **fallback** : les lignes « ramasse/palette » (n° interne `25-06-xxxx`) renvoient `#N/A` → on **change la colonne de recherche en `G:G`**.
- **Corrections récurrentes** repérées à l'écran : `Date validité tarif = 01/MM/AAAA` ; **TVA = 0,2 constante** (attention à l'artefact de recopie `0,2 ; 1,2 ; 2,2…`) ; **Zone `0France… → 21France…`** ; anomalies de Frêt à **0 €** (lignes B2C) ou **1 €** (à investiguer) ; tracking des lignes ramasse renseigné avec le n° d'expédition interne.
- **Barème gazole** : taux effectif rapproché du barème officiel (MESS 12,45 % / AFFR 8,73 % base 1,8658 €/l ; 8,85 % constaté juin 2025) — **contrôle**, pas driver de calcul.

---

## Vidéo 3 — « Prix de vente / Doublon prestation » (NOUVEAU : volet VENTE + avaries)

Cette vidéo **ne montre pas Excel** mais le logiciel interne :
`si.laruche-logistique.fr` → **Gestion Facturation ▸ [TRANSPORT] - Prix de vente**.
C'est là que se fait la **refacturation au client** (prix de vente, marge/écart) et la **gestion des avaries**.

### Les types d'avaries (anomalies) gérés
- **`DOUBLON_PRESTATION`** — même expédition facturée sur 2 lignes.
- **`GT_CLIENT_DOUBLON`**, **`GT_LIGNE_MANQUANTE`**, **`GT_MANQUANTE`** — anomalies de **grille tarifaire** (GT) : tarif/ligne manquant, doublon client.

### Le cas « DOUBLON_PRESTATION » (exemple `EXP20250502-2157393`, client Renight Medical)
**Problème** : Kuehne facture une même expédition (même tracking, même poids, même zone) sur **2 lignes** — l'une porte le **Frêt**, l'autre **une surcharge seule** (ici *Zone éloignée*). Sans correction, le client serait facturé deux fois.

**Correction dans le SI** :
1. Filtrer *Gestion des prix de vente* sur **Avarie(s) = DOUBLON_PRESTATION** (un compteur d'avaries est affiché).
2. Repérer les 2 lignes de même tracking (badge orange « avarie »).
3. Ouvrir la ligne (crayon) → modal *Modification de la ligne* : section **« Liste des avaries »** + section **« Ligne à valider »** (ligne consolidée à conserver).
4. **Reporter la charge du doublon** (ex. *Zone éloignée = 25,7*) sur la ligne à valider qui portait déjà le Frêt → le **PV Frêt** et l'**Écart** (marge = PV − coût) se recalculent, la **Taxe gasoil** est recalculée sur le frêt consolidé.
5. Cocher **« Ne plus considérer comme avarie »** → *Enregistrer*. Le compteur d'avaries décrémente.
6. Vérifier : relancer le filtre Avarie → **« 0/0 – aucune donnée »**.
7. **Consolidation batch finale** (bouton *Consolider*) : « Vous êtes sur le point de consolider N ligne(s) » → *Confirmer*. Le récap client montre **Qté tracking** qui **ne compte pas** les doublons (ex. 30) et le **Prix** total (ex. 2 502,79 €).

### Règle métier à coder (détection + correction du doublon)
- **Détection** : au sein d'une même période / transporteur / client, **grouper par n° de tracking**. Si un tracking apparaît sur **≥ 2 lignes** ⇒ `DOUBLON_PRESTATION`. Signature typique : le **frêt sur une ligne** et **une surcharge seule sur l'autre** (poids/zone identiques).
- **Correction** : conserver **1 ligne par tracking**, **sommer les postes de charge** des doublons (Frêt, Zone éloignée, Assurance, Droits & taxes, Colis volumineux…), **recalculer** Taxe gasoil (sur frêt consolidé), **PV Frêt** (grille tarifaire) et **Écart**. Facturer le client **une seule fois**.

> ⚠️ À ne pas confondre avec le reclassement *achat* (doc maîtresse) : la fusion « 1 ligne par tracking » est une opération **côté VENTE/SI**, alors que l'export d'**achat** reste **1 ligne CSV = 1 ligne import** (aucun regroupement).

---

## Ce que ça change pour le projet

- **Volet ACHAT (import)** : les vidéos 1 & 2 **valident** le pipeline existant (`automatisation/facturation_kuehne.py`, appli Node). Rien à refaire ; éventuellement ajouter les **contrôles** vus à l'écran (Autres frais ≠ 0, #N/A noms, Frêt 0/1 €, zones 0France) comme garde-fous automatiques.
- **Volet VENTE (nouveau)** : la vidéo 3 documente enfin la **refacturation client** et les **avaries** — jusqu'ici zone d'ombre. C'est le prochain gros chantier fonctionnel :
  - moteur de **détection d'avaries** (doublon prestation, GT manquante/doublon),
  - **grille de prix de vente** par zone → PV Frêt + marge/écart,
  - **consolidation par tracking** côté vente.

---

## Questions ouvertes précisées (issues des vidéos)

1. **Recalcul du gasoil après consolidation d'un doublon** : automatique (assis sur le frêt total) ou saisi ? (valeur transitoire `0,07` vs finale `7,06` observée à l'écran).
2. **Devenir de la 2ᵉ ligne du doublon** : supprimée, ou neutralisée (charges reportées puis mises à 0) ? À l'écran elle reste visible jusqu'à la consolidation batch.
3. **Écart consolidé** : somme des écarts des 2 lignes, ou recalcul sur la marge de la ligne unique ?
4. **Grille tarifaire de VENTE** : structure exacte (par zone / poids / transporteur) — source des avaries `GT_*`.
5. **Lien résolution unitaire d'avarie ↔ consolidation batch** (« N lignes ») à cadrer.
