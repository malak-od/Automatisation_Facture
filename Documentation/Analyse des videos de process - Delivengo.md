# Analyse des vidéos de process — Delivengo

> **Delivengo** = offre courrier international **avec suivi** de La Poste. Compte : « LA RUCHE LOGISTIQUE - **LPPAQ** ».
> Ce document dépouille les **3 vidéos `.mp4`** du dossier `Transporteurs/Delivengo/`, croisées avec les fichiers réels (export de suivi, classeur `2026_06_Delivengo_LPPAQ.xlsx`).
> Méthode : extraction d'images (ffmpeg) + lecture par agents + vérification des **formules réelles** du classeur. Les `f_xxx` renvoient aux images.

---

## Inventaire des vidéos

| Vidéo | Durée | Outil | Apport |
|---|---|---|---|
| `Delivengo_1_Recuperation facture excel.mp4` | 1 min 19 | Portail **MyDelivengo** | D'où vient le fichier source |
| `Delivengo_2_Preparation fichier import.mp4` | 7 min 06 | **Excel** | Construction du fichier import (cœur) |
| `Process Facturation - Delivengo - Avarie Import.mp4` | 1 min 54 | Logiciel interne **si.laruche-logistique.fr** | Cas d'avarie à l'import |

---

## 1. Récupération de la source (vidéo 1)

La source **n'est pas une facture PDF** mais un **export de suivi**.

- Portail **MyDelivengo** : `https://mydelivengo.laposte.fr` → bascule `profil.mydelivengo.laposte.fr`.
- Menu **Suivi ▸ Suivi des envois** (`/suivi/index/lister`).
- Filtre **Période = « Le mois dernier »** (≈ 557 plis/mois).
- Bouton **Outils ▸ Exporter au format Excel**.
- ⚠️ **Export asynchrone** : le portail envoie un **email** (via le service tiers `www2.easyreco.com`) à `transport@laruche-logistique.fr` avec un **lien de téléchargement** (`/downloads/download/key/<clé>`).
- Fichier obtenu : **`AAAA-MM-JJ HH_MM_SS-Export_du_suivi_des_envois.xls`** (~134 Ko ; en réalité un `.xlsx` déguisé).

**Colonnes utiles de l'export de suivi** : Numéro de suivi (`LD…FR`), Destinataire (nom, pays en clair), **Statut Général** (Distribué / Déposé), **Remise en poste** (date), **Poids** (col AE, en grammes — ⚠️ *valeur nominale peu fiable, souvent 150 g*).

---

## 2. Construction du fichier import (vidéo 2) — le cœur

### Les 3 sources combinées
1. **Export de suivi** (ci-dessus) → identité de l'envoi (tracking, nom, pays, statut, date).
2. **« Export expéditions brut » (.xlsx)** = export du WMS interne → **le POIDS RÉEL** (colonne `INFO_POI` = AI), rapproché par **N° de suivi** (`PRO_TRACKING` = AP). ➜ **source découverte via la vidéo, invisible dans les seuls fichiers.**
   - Il faut aussi l'**export brut du mois précédent** : ~25 plis sur 557 sont expédiés le mois M-1 mais distribués/facturés le mois M (2ᵉ `RECHERCHEX` de repli).
3. **Feuille « Pays »** (dans le classeur) = table de correspondance.

### Feuille « Pays » (clé = nom de pays en clair)
| Col | Contenu |
|---|---|
| A | Pays (nom complet) — **clé de recherche** |
| B | Code ISO2 (DE, CH, BE, US…) |
| C | Zone (`1` / `2`, variantes `1_CH`, `2_US`…) |
| D | TVA (`0,2` UE / `0` hors-UE) |
| E | Gazole = **1,50 €/kg uniquement pour US, Canada, Australie, Nouvelle-Zélande, Japon** |

### Feuille « Fichier import » — colonne par colonne
| Col | En-tête | Origine / règle |
|---|---|---|
| A | Date remise | export suivi (copie) — *colonne de travail, absente du CSV* |
| B | Transporteur | **constante** `DELIVENGO-LPPAQ` |
| C | Date validité tarif | **constante** = 1er du mois (recopie `=C7`) |
| D, E | Réf.1 / Réf.2 | vides |
| F | Statut | export suivi (copie) |
| G | **N° de suivi** | export suivi (copie) — **clé de rapprochement** |
| H | Nom | export suivi (copie) |
| I | E/P | épars `1`/`2` — **signification à confirmer** |
| J | Pays (code) | `RECHERCHE(X → Pays!B)` |
| K | Zone | `RECHERCHE(X → Pays!C)` |
| L | Nbr Colis | **constante** `1` |
| M | **Poids facturé (kg)** | `=ARRONDI(MAX(P;Q);2)` |
| N | mode envoi | **constante** `avecsuivi` |
| O | TVA | `RECHERCHE(X → Pays!D)` |
| P / Q | *(en-têtes « Droits et taxes »/« Assurance » réutilisés)* | **colonnes de travail du poids** : l'une = **poids réel** via `RECHERCHEX(G → brut!AP → brut!AI)`, l'autre = **constante `0,15`** (plancher). Rôles P/Q inversés selon les mois ; peu importe car `M = MAX`. |
| R, S | Zones éloignées / Colis volumineux | vides |
| T | Adresses | colonne helper surlignée, **vidée dans le CSV** |
| U | Frêt | **constante** `1` |
| V | plus-value BtoC | vide |
| W | Taxe Gasoil | `=SI(NB.SI(Pays!A;X)=0;"";RECHERCHE(X → Pays!E))` |
| X | Pays (nom complet) | export suivi (copie) — **clé** des RECHERCHE J/K/O/W ; *absente du CSV* |

**Règle métier centrale — poids facturé** :
> `Poids = ARRONDI( MAX( poids_réel_WMS , 0,15 ) , 2 )` en kg.
> Le poids réel vient de l'**export WMS brut** (pas du poids nominal du suivi), rapproché par tracking, avec **repli sur le brut du mois précédent** si non trouvé. `0,15 kg` est un **plancher de facturation**.

### Export final
- **Enregistrer sous → CSV (séparateur « ; »)** : `AAAA_MM_Delivengo_LPPAQ_Import.csv` dans `…\2 - Fichiers csv import\AAAA\AAAA MM\`.
- Le **CSV ≠ la feuille** : il démarre à **Transporteur** (colonnes de travail A « Date remise » et X « Pays nom » retirées), ajoute **« Nb colis »**, et renomme des en-têtes (« N° Tracking », « Nom »). C'est le **gabarit d'import** du système aval, pas un simple export à l'identique.
- Puis import dans **`si.laruche-logistique.fr`** (Gestion Facturation ▸ [TRANSPORT] Imports).

---

## 3. Avarie à l'import (vidéo 3)

**Cas** : à l'import dans le SI, la ligne d'un pli tombe en avarie avec **3 libellés** : `CHAMPS_MANQUANT`, `MAPPAGE_EXPEDITION_ERP`, `MAPPAGE_ENTITE_ERP`.

**Cause racine** : le **numéro de suivi** avait été saisi **avec des espaces** côté expédition WMS (`LD 21 539 131 7 FR`) alors que la facture Delivengo porte `LD215391317FR` → le rapprochement automatique (tracking ↔ expédition ↔ client) échoue, d'où « champ manquant » (Client/Transporteur non déductibles).

**Correction (dans le SI)** : ouvrir l'expédition → onglet Suivi → **supprimer le code mal formaté**, garder `LD215391317FR` → revenir sur la ligne d'import, **renseigner le Client** (+ transporteur) → Enregistrer. L'avarie se lève (compteur décrémenté).

**Règle à coder (simple et à fort impact)** :
> Avant tout rapprochement, **normaliser les N° de suivi des deux côtés** : `upper()` + suppression de tout caractère non alphanumérique (espaces, tirets). `"LD 21 539 131 7 FR"` → `"LD215391317FR"`.
> Valider le format à la saisie : `^[A-Z]{2}\d{9}[A-Z]{2}$`. En cas d'échec de matching exact, **retenter avec le tracking normalisé** avant de lever une avarie.

---

## 4. Ce que ça change pour le projet

- Delivengo est déjà partiellement codé (`automatisation/finaliser_delivengo.py`), mais les vidéos apportent **2 précisions importantes à vérifier dans le code** :
  1. **Poids** = `MAX(poids_réel_WMS, 0,15)` avec le poids réel **issu de l'export WMS brut** (+ repli mois précédent), **pas** `poids_suivi/1000`. À confirmer que le finaliseur fait bien ce rapprochement (le poids nominal du suivi = 150 g est trompeur).
  2. **Gazole** limité à 5 pays (US, CA, AU, NZ, JP) via la table Pays.
- **Normalisation des trackings** (règle de la vidéo 3) : garde-fou à ajouter au moteur générique — utile pour **tous** les transporteurs.
- Le reste (constantes, lookups Pays/Zone/TVA) est simple et déjà cerné.

---

## 5. Points à confirmer

1. **Colonne I « E/P »** : valeurs éparses `1`/`2` — signification (Enveloppe/Paquet ? Entreprise/Particulier ?).
2. **Colonne K « Zone »** : format exact (`1_CH` vs `1`) et formule précise (non captée à l'écran).
3. **Colonne U « Frêt » = 1** constant : est-ce un montant, un flag, ou un placeholder (le prix réel étant calculé en aval par le SI) ?
4. **Récupération automatique de l'export** : dépend d'un email (service `easyreco.com`) → parsing IMAP ou API MyDelivengo à investiguer.
5. **Nommage** : léger décalage `Delivengo-LPPQ` (classeur juin) vs `DELIVENGO-LPPAQ` (vidéo) — à trancher.
