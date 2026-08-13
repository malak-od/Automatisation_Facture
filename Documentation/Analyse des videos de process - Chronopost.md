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


Résultats de l'inspection — Chronopost (classeur modèle + factures brutes + PDF)
a) Feuille "Facture Chronopost" (A1:AJ2351, 2350 lignes de données)
En-têtes ligne 1 (A à AJ, 36 colonnes) :
A ID Clients, B No Facture, C Sous-compte, D Date LT, E Code postal depart, F Code postal arrivee, G Pays depart, H Pays arrivee, I Ref Destinataire, J Ref Expediteur, K No Groupage tarifaire, L Numero LT, M Groupage, N Type prestation, O TVA, P Observations, Q Zone Tarifaire, R Poids, S Produit, T Montant HT, U Raison sociale, V Raison sociale 2, W Code produit modifié, X Zoning 2shop, Y Gazole %, Z frêt, AA sureté + eco, AB gazole, AC hors gazole, AD Catégories, AE Total avec GO, AF Total (label), AG1 = =SUM(AG2:AG6), AH/AI = zone annexe (voir plus bas), AJ vide.

Colonnes A-U = import brut (venant de la facture Chronopost source, collé tel quel). Colonnes W à AE = formules ajoutées. Formules exactes ligne 2 :

W2 (Code produit modifié) : =IF(C2<>2,"",IF(S2="5X","5XK",IF(S2="5Y","5YK",IF(S2="6B","6BK",IF(S2="6C","6CK",""))))) — ne s'applique que si Sous-compte (C) = 2 (Foutas), sinon vide.
X2 (Zoning 2shop) : =IF(S2="6B",_xlfn.XLOOKUP(H2,'Zoning 2shop'!A:A,'Zoning 2shop'!B:B),IF(S2="6C",_xlfn.XLOOKUP(G2,'Zoning 2shop'!C:C,'Zoning 2shop'!D:D),"")) — H2 = pays arrivée pour code 6B, G2 = pays départ pour 6C.
Y2 (Gazole %) : =AB2/Z2
Z2 (frêt) : =T2 (simple copie du Montant HT de la ligne)
AA2 (sureté + eco) : =IF(OR(S2="6C",S2="6B",S2="5X",S2="5Y"),0.08,0.5) — taux forfaitaire selon le code produit (0.08 pour 2Shop, 0.5 sinon). Attention : c'est un taux, pas un montant en euros — nom de colonne trompeur.
AB2 (gazole) : =$AG$5/$AG$2*Z2 — répartition du pool gazole au prorata du fret de la ligne (AG5 = total pool "gazole avant remise" T2:T11, AG2 = total fret T28:T3000).
AC2 (hors gazole) : =Z2+AA2
AD2 (Catégories) : =IF(COUNTIF(Catégories!A:A,N2)=0,"catégorie inconnue",LOOKUP(N2,Catégories!A:A,Catégories!B:B))
AE2 (Total avec GO) : =SUM(Z2:AB2) — répétée sur les 2350 lignes.
Zone récapitulative AF/AG/AH/AI (lignes 1-9), grand total et sous-totaux pour juin 2026 :


AF1='Total'    AG1='=SUM(AG2:AG6)'              → 21141.82
AF2='Frêt'     AG2='=SUM(T28:T3000)'            → 17949.52
AF3='eco'      AG3='=SUM(T12:T19)'              → 122.5
AF4='sureté'   AG4='=SUM(T20:T27)'              → 393.4
AF5='Gazole'   AG5='=SUM(T2:T11)'               → 2676.4
AF6='Gestion'  AG6= (vide)
AF8='Gazole réel'      AH8=0.2315   AI8='routier'
AF9='(avant remise)'   AH9=0.4375   AI9='aérien'
AH5='=AG5/AG2' (=0.14910705…)   ← taux gazole effectif réel calculé, utilisé indirectement via AB
Il n'y a pas de AG8/AG9 — les taux de gasoil routier (0.2315 = 23.15%) et aérien (0.4375 = 43.75%) sont saisis en dur en AH8/AH9, avec le libellé en AI8/AI9. Ce sont les mêmes plages T2:T11 (pool aérien+routier+eco+sureté saisi manuellement en tête de feuille avant le report des lignes détaillées) — donc T2:T27 = zone "pool spécial" (aérien/routier/eco/sûreté agrégés), T28:T3000 = lignes détail du fret.

Les colonnes AE, AF(label), AG(sommes) sont répétées identiquement sur toutes les 2350 lignes (formule relative type =SUM(Zn:ABn) pour AE), donc les vraies "constantes" de contrôle ne sont qu'aux lignes 1-9.

b) Feuille "TCD" (A1:Z2210)
TCD natif (Tableau croisé dynamique1) :

name=Tableau croisé dynamique1, location=E1:R2210, firstHeaderRow=1, firstDataRow=2, firstDataCol=4
source = 'Facture Chronopost'!B1:AD1048576
rowFields = index [10, 17, 13, 15] → Numero LT, Produit, TVA, Zone Tarifaire (4 champs de ligne imbriqués)
colFields = [27] → Catégories
dataFields = [('Somme de Montant HT', 18, 'sum')] (champ 18 = Montant HT)
cacheFields complet : ['No Facture','Sous-compte','Date LT','Code postal depart','Code postal arrivee','Pays depart','Pays arrivee','Ref Destinataire','Ref Expediteur','No Groupage tarifaire','Numero LT','Groupage','Type prestation','TVA','Observations','Zone Tarifaire','Poids','Produit','Montant HT','Raison sociale','Raison sociale 2','Zoning 2shop','Gazole %','frêt','sureté + eco','gazole','hors gazole','Catégories']
Le TCD éclate donc chaque colis par Numero LT/Produit/TVA/Zone (lignes) x Catégories (colonnes : Adresse/Assurance/Colis volumineux/Corse/Droits et taxes/Frais facturation/Frêt/Gazole/Zones éloignées/(vide)/Sureté+eco), sortie en E2:S2210+.

Colonnes de calcul manuel à côté (ligne 3 = première ligne de données) :

S3 : =_xlfn.XLOOKUP(E3,'Facture Chronopost'!L:L,'Facture Chronopost'!AA:AA,"") — récupère le taux sureté+eco (AA) depuis Facture Chronopost via le Numero LT.
T3 : =IF(COUNTIF('Bibliothèque transporteurs'!C:C,F3)=0,"inconnu",_xlfn.XLOOKUP(F3,'Bibliothèque transporteurs'!C:C,'Bibliothèque transporteurs'!A:A)) — mode envoi ERP depuis le code produit Chrono (F = Produit).
U3 : =IF(OR(F3="6C",F3="6B"),_xlfn.XLOOKUP(E3,'Facture Chronopost'!L:L,'Facture Chronopost'!X:X),IF(OR(F3=17,F3=44),F3&"_"&H3,IF(COUNTIF('Bibliothèque transporteurs'!C:C,F3)=0,"inconnu",_xlfn.XLOOKUP(F3,'Bibliothèque transporteurs'!C:C,'Bibliothèque transporteurs'!B:B)))) — zone ERP : cas spécial 2shop (6B/6C → lookup Zoning), cas spécial produits 17/44 (concat produit_zone), sinon lookup bibliothèque.
V3 : =IF(COUNTIF('Bibliothèque transporteurs'!C:C,F3)=0,"inconnu",_xlfn.XLOOKUP(F3,'Bibliothèque transporteurs'!C:C,'Bibliothèque transporteurs'!E:E)) — nom Transporteur (CHRONOPOST / CHRONO_2SHOP / CHRONO_SHOP_KERSUN).
Ces 4 colonnes (S,T,U,V) sont remplies sur toutes les lignes de données du TCD (lignes 3 à 2209/2210, y compris les lignes "pool" 3-9 sans ID client).

A, B, C, W, Y, Z ne démarrent qu'à la ligne 10 (les lignes 3-9 sont le pool CAPI1/CAPI2/CAPN1/CAPN2/CAPO1/ECORI/ECORN sans "ID client" associé, donc pas de ventilation par colis nécessaire). Formules ligne 10 :

A10 (Total GO) : =C10/SUM(C:C)*(SUM(C:C)+$X$2) — répartition proportionnelle du pool gazole (X2) sur le "Total hors GO" de la ligne.
B10 (Total avec CAP+ECO) : =C10/SUM(C:C)*(SUM(C:C)+SUM($N$3:$O$7))+S10 — répartition proportionnelle du pool "Frais facturation + Frêt" (N3:O7) plus la part sûreté+eco (S10).
C10 (Total hors GO) : =SUM(I10:Q10)-N10-P10 — somme des colonnes catégories (I à Q) moins Frais facturation (N) et Gazole (P).
D10 = ID client (sortie native du TCD, ex. 6235).
W10 : =IF(T10="FR_2SHOP",O10,IF(O10=0,0,ROUNDUP(O10/C10*B10,2))) — recalcul du fret "corrigé" (répartit B10 au prorata du fret O10/C10), sauf cas spécial mode envoi "FR_2SHOP" où on garde O10 tel quel.
Y10 : =ROUNDUP(B10,2)-M10-L10-N10-J10-K10-I10-Q10 — formule alternative de vérification ("Formule Test à vérifier").
Z10 : =IF(Y10=W10,"OK","Non") — comparaison des deux formules (colonne "Vérif avec ancienne formule"). Beaucoup de "Non" observés (arrondis différents), un "OK" en ligne 2209.
X2 (cellule unique, pas répétée par ligne) : =SUM(N$3:N$1048576)+SUM(P$3:P$1048576)+SUM(O3:O13) = 3192.3 — total Frais facturation + Gazole + Frêt des lignes pool(3:13).
Ligne 2210 = ligne "(vide)" (catégorie du TCD pour les lignes sans classification).

c) Feuille "Fichier import" (A1:W2208, 23 colonnes ERP)
En-têtes ligne 1 : Transporteur, Date validité tarif, Réf.1, Réf. 2, Id client, N° Tracking, Nom, E / P, Pays, Zone, Nbr Colis, Poids, mode envoi, TVA, Droits et taxes, Assurance, Zones éloignées, Colis volumineux, Adresses, Frêt, plus-value BtoC, gazole, 2SHOP : ne pas mettre de gazole.

Formules ligne 2 (correspond à TCD ligne 3, décalage de +1) :

A2 (Transporteur) : =TCD!V3
B2 (Date validité tarif) : valeur en dur 01/05/2026 (première ligne), puis =B2 pour les lignes suivantes (recopie la même date sur tout le fichier)
C2, D2, E2, G2 (Réf.1, Réf.2, Id client, Nom) : vides
F2 (N° Tracking) : =TCD!E3 (Numero LT)
H2 (E/P) : =IF(RIGHT(J2,3)="BTB","E","P") — dérivé du suffixe de la Zone (colonne J de ce même fichier)
I2 (Pays) : =LOOKUP(F2,'Facture Chronopost'!L:L,'Facture Chronopost'!H:H) — lookup Pays arrivée depuis Facture Chronopost via N° Tracking
J2 (Zone) : =TCD!U3
K2 (Nbr Colis) : valeur en dur 1, puis =K2 pour lignes suivantes
L2 (Poids) : =_xlfn.XLOOKUP(F2,'TCD poids'!A:A,'TCD poids'!C:C) — lookup direct dans TCD poids (colonne C = Poids arrondi) via N° Tracking
M2 (mode envoi) : =TCD!T3
N2 (TVA) : =IF(I2="",0.2,LOOKUP(I2,'Pays TVA'!A:A,'Pays TVA'!B:B))
O2 (Droits et taxes) : =IF(TCD!M3="","",TCD!M3)
P2 (Assurance) : =IF(TCD!J3="","",TCD!J3)
Q2 (Zones éloignées) : =IF(AND(TCD!L3="",TCD!Q3=""),"",IF(SUM(TCD!L3,TCD!Q3)>15,29,9.5)) — Corse (L) + Zones éloignées (Q) du TCD, seuil forfaitaire à 15€ pour choisir entre 29 et 9.5
R2 (Colis volumineux) : =_xlfn.XLOOKUP(F2,TCD!E:E,TCD!K:K,"")
S2 (Adresses) : =IF(TCD!I3="","",IF(TCD!I3>8.5,17,8.5)) — seuil forfaitaire sur colonne Adresse (I) du TCD
T2 (Frêt) : =TCD!Y3 ← le Frêt final vient de la colonne Y du TCD (formule de vérif "Test à vérifier"), pas de W !
U2, V2 (plus-value BtoC, gazole) : vides
W2 (2SHOP note) : texte "dans les coûts de revient" en ligne 2, "Renseigner 'Pays'" en ligne 4 — commentaires ponctuels, pas une formule répétée.
Confirmation : Poids vient de TCD poids (XLOOKUP colonne C = Poids arrondi ROUNDUP). Mode envoi et Zone viennent de TCD (colonnes T et U). Fret vient de TCD!Y (PAS W). Le décalage systématique est ligne_import N ↔ TCD ligne N+1.

d) Feuille "Zoning 2shop" (A1:D25)
Deux mini-tables pays→zone côte à côte :

Table 6B : colonnes A (code pays) / B (zone numérique 1-24), ex. AT=1, BE=2, BG=3, CH=4, CZ=5, DE=6, DK=7, EE=8, ES=9, FI=8, HR=8, HU=8, IE=7, IT=14, LT=15, LU=6, LV=8, NL=6, PL=5, PT=9, RO=3, SE=15, SI=7, SK=24. En-tête A1='6B'.
Table 6C : colonnes C (code pays) / D (zone numérique 25-43), mêmes 24 pays, ex. AT=25, BE=26, BG=27, CH=28, CZ=29, DE=30, DK=25, EE=27, ES=33, FI=27, HR=27, HU=28, IE=37, IT=33, LT=39, LU=30, LV=27, NL=42, PL=43, PT=33, RO=27, SE=39, SI=25, SK=25. En-tête C1='6C'.
e) Feuille "Catégories" (33 lignes de mapping utiles, A1:B33 — la feuille va jusqu'à G38 mais G est une colonne parasite sans lien logique)
Mapping complet Type prestation (colonne A) → Catégorie (colonne B) :


Assurance                              → Assurance
CAP RouteCAP ROUTE                     → Gazole
Correction d'adresse                   → Adresse
Eco responsablePar                     → Frêt
ESD                                    → Gazole
Frais de gestion                       → Frais facturation
Participation Eco-Responsable          → Frêt
Supp Retour Expediteur Europe          → Frêt
Supp Retour Expediteur Inter           → Frêt
Supp Zone Internationale Eloignee      → Zones éloignées
Supplement Annonce incomplète          → Adresse
Supplement Corse                       → Zones éloignées
Supplement Corse 18h                   → Corse
Supplement domicile prive              → Zones éloignées
Supplement Douane Zone C2              → Droits et taxes
Supplement Douane Zone C4              → Droits et taxes
Supplement Etiquette Non Conforme      → Frêt
Supplement Forfait Expedition          → Frêt
Supplement GT                          → Frêt
Supplement Hors Norme                  → Colis volumineux
Supplement Manutention                 → Colis volumineux
Supplement Retour Expediteur           → Frêt
Supplement Retrait Bureau              → Zones éloignées
Surcharge Carburant Aérien             → Gazole
Surcharge Carburant Routier            → Gazole
Surcharge Facture Papier               → Frais facturation
Sûreté colis                           → Frêt
Traitement Réacheminement              → Frêt
Traitement SAV complémentaire          → Adresse
Transport                              → Frêt
Transport encompte                     → Frêt
Sûreté étendue                         → Frêt
Zones Difficiles d'accès               → Zones éloignées
Important : Sûreté colis est classée en catégorie "Frêt" (pas "Sureté + eco"), et Participation Eco-Responsable aussi en "Frêt". Ce sont des libellés statiques (LOOKUP exact-match, colonne A triée par ordre alphabétique — LOOKUP nécessite un tri croissant, confirmé ici).

f) Feuille "Contrôle pdf" (A3:D12)
TCD natif (Tableau croisé dynamique6) : name=Tableau croisé dynamique6, location=A3:B11, source='Facture Chronopost'!B1:T1048576, rowFields=[0] (No Facture), dataFields=[('Somme de Montant HT', 18, 'sum')].

Table de contrôle (colonnes C/D saisies à la main), pour le mois modèle (mai/juin 2026, PAS le même mois que les 7 PDF du dossier) :


No Facture   | Somme Montant HT (B) | pdf (C, saisi main) | écarts (D=B-C)
13617721     | 4555.12               | 4555.12              | 0
13617722     | 21.68                 | 21.68                | 0
13617723     | 6688.75               | 6688.75              | 0
13624019     | 1487.32               | 1487.32              | ~0 (arrondi flottant)
13624020     | 142.46                | 142.46               | 0
13624021     | 8246.49               | 8246.49              | ~0 (arrondi flottant)
(vide)       |                       |                      | 0
Total général| 21141.82              | =SUM(C4:C10)=21141.82| ~0
Attention importante : ces numéros de facture (13617721-23, 13624019-21) ne correspondent pas aux 7 PDF présents dans le dossier (13655988-90 et 13662501-04). Le classeur modèle "2026_06_Facture Chronopost.xlsx" représente donc un mois différent de celui des PDF/factures brutes fournies — cohérent avec le pattern déjà observé pour DPD/BLS où le classeur "modèle" sert de gabarit à cloner, indépendant du mois réellement traité.

Fichiers 2 et 3 — factures brutes reçues
Les deux fichiers ont un format identique : feuille unique Données, titre en A1, Compte : XXXXXXXX - LA RUCHE LOGISTIQUE en A2, ligne 3 vide, en-tête exact en ligne 4 :
No Facture | Sous-compte | Date LT | Code postal depart | Code postal arrivee | Pays depart | Pays arrivee | Ref Destinataire | Ref Expediteur | No Groupage tarifaire | Numero LT | Groupage | Type prestation | TVA | Observations | Zone Tarifaire | Poids | Produit | Montant HT | Raison sociale (20 colonnes A-T).

facture_chronopost_51291303_202606.xlsx (sous-compte standard CHRONOPOST) : 1519 lignes de données (5 à 1523). 3 blocs de lignes spéciales (un par No Facture : 13655988, 13655989, 13655990) :

13655988 (ligne ~1024-1026) : ECORN/Participation Eco-Responsable=95.3, SURTN/Sûreté colis=381.2, CAPN2/Surcharge Carburant Routier=921.7
13655989 (ligne ~1035-1037) : ECORN=0.8, SURTN=3.2, CAPN2=8.26 (petit sous-compte Foutas)
13655990 (ligne ~1513-1523) : bloc complet avec variantes I/N/O — ECORI=3.4, ECORN=38.8, ECORO=0.1, SURTI=13.6, SURTN=155.2, SURTO=0.4, CAPI1=48.73 (Aérien), CAPI2=40.66 (Routier), CAPN1=2583.56 (Aérien), CAPN2=8.47 (Routier), CAPO1=18.94 (Aérien)
Total 12 lignes spéciales, répartition : 5×Participation Eco-Responsable, 5×Sûreté colis, 4×Surcharge Carburant Routier, 3×Surcharge Carburant Aérien.

facture_chronopost_65481903_202606.xlsx (sous-compte CHRONO_2SHOP) : 2194 lignes de données (5 à 2198). 4 blocs de lignes spéciales (No Facture 13662501, 502, 503, 504), mais motif systématiquement identique et plus simple : toujours exactement un seul triplet ECORN (Participation Eco-Responsable) / SURTN (Sûreté colis) / CAPN2 (Surcharge Carburant Routier) par facture — jamais de Surcharge Carburant Aérien (CAPI*), ni de variantes I/O (ECORI/ECORO/SURTI/SURTO/CAPI/CAPO). Confirmé par le comptage exhaustif : {'Participation Eco-Responsable': 4, 'Sûreté colis': 4, 'Surcharge Carburant Routier': 4} sur 2194 lignes.

Conclusion clé : le motif des lignes spéciales diffère selon le sous-compte — le contrat standard (51291303) peut avoir des envois aériens (d'où CAPI/ECORI/SURTI et CAPO/ECORO/SURTO en plus de N), tandis que le contrat 2Shop (65481903) n'a que du routier national (uniquement le triplet N). Le code du carrier devra gérer dynamiquement 1 à 3 triplets par facture selon ce qui est présent, pas un nombre fixe.

PDF (7 fichiers CHRONOPOST_*.pdf)
Format texte structuré, facilement parseable par regex (comme Mondial Relay/DPD, pas un visuel complexe comme BLS) : texte natif extractible via pypdf sans OCR, tableau tabulé avec labels clairs (Compte, Facture, Date, Net à payer, Code Taux H.T. T.V.A. T.T.C., TOTAL FACTURE).

PDF	Facture	Sous-compte	Total HT imprimé (TOTAL FACTURE)	Net à payer TTC	Pages
CHRONOPOST_13655988_51291303.pdf	13655988	51291303 (compte principal)	8290,14 €	9948,17 €	17
CHRONOPOST_13655989_51291303.pdf	13655989	51291303 — Sous-compte 002 BY FOUTAS	73,86 €	88,63 €	3
CHRONOPOST_13655990_51291303.pdf	13655990	51291303 (compte principal, contient de l'aérien EXPRE)	9979,94 €	11700,45 €	10
CHRONOPOST_13662501_65481903.pdf	13662501	65481903 (compte principal)	31,13 €	37,36 €	3
CHRONOPOST_13662502_65481903.pdf	13662502	65481903 — Sous-compte 002 KER SUN	2487,72 €	2985,26 €	13
CHRONOPOST_13662503_65481903.pdf	13662503	65481903 (compte principal, produit 2SE, retours PT)	74,81 €	89,77 €	3
CHRONOPOST_13662504_65481903.pdf	13662504	65481903 — Sous-compte 002 KER SUN	12117,32 €	14540,78 €	22
Structure commune de chaque PDF :

Page 1-2 (ou seule page 1 si court) : récapitulatif journalier (Date / Objets / Expéditions / Poids / Montant HT), puis lignes "Redevance sûreté N colis à X EUR", "Participation eco-responsable N colis à X EUR", "Surcharge Carburant Routier/Aérien : X % sur montant de Y EUR = Z", puis bloc TVA (Code Taux H.T. T.V.A. T.T.C., ligne(s) par code I/N/O, puis TOTAL FACTURE).
Pages suivantes : détail ligne par ligne (N° objet | Départ | Arrivée | Poids | Montant HT | Obs | Produit | TVA), avec sous-totaux "Report" en cumul et reprises de "Total sous-compte"/"Mois MM/AAAA"/"Total général" en fin de section — cohérent avec la logique Fichier/TCD du classeur modèle (regroupement par No Facture, avec sous-compte identifié par le texte "Sous-compte : XXX NOM" quand applicable).
Le taux de surcharge carburant est imprimé en clair dans le texte (ex. "12,75 %" pour le contrat standard routier, "43,75 %" pour l'aérien standard, "3,25 %" pour le contrat 2Shop routier) — cohérent avec AH8=0.2315/AI8='routier' et AH9=0.4375/AI9='aérien' du modèle (le 0.2315 du modèle de juin ne correspond à aucun taux vu ici, confirmant que le taux carburant est mensuel/variable et lu directement depuis la facture source, pas une constante figée dans le classeur).
Fichiers d'inspection (scratchpad, pour référence, non livrables projet)
Scripts et sorties texte dans C:\Users\MALAKO~1\AppData\Local\Temp\claude\c--Users-Malak-OUJDID-Desktop-Automatisation-Facture\ae1f3bd5-892c-44a1-9577-61219a8de09f\scratchpad\ : inspect1_overview.py à inspect10_pdf_totals.py, et leurs sorties out_overview.txt, out_facture_chronopost.txt, out_tcd_poids.txt, out_tcd.txt, out_tcd_detail.txt, out_tcd_more.txt, out_fichier_import.txt, out_zoning_2shop.txt, out_categories.txt, out_controle_pdf.txt, out_support_sheets.txt, out_raw_51291303.txt, out_raw_65481903.txt, out_pdfs.txt, out_pdf_totals.txt. Aucun fichier du projet n'a été modifié — inspection en lecture seule uniquement.

Excellent, l'inspection est très complète et révèle des points importants qui changent la donne par rapport à ce qu'on pensait. Je synthétise avant de coder.

Découvertes clés qui corrigent/précisent les hypothèses précédentes :

Le taux de gasoil (AH8/AH9, pas AG8/AG9 comme dans mes notes) vient en fait de la facture reçue elle-même (imprimé en clair dans le PDF, ex. "12,75%" routier / "43,75%" aérien) — ce n'est PAS une constante fixe à chercher sur le site Chronopost chaque mois comme supposé. C'est une donnée dynamique par facture.

Le motif des lignes CAP/ECO/SUR diffère selon le sous-compte : le contrat standard peut avoir 1 à 3 triplets (aérien I / national N / outre-mer O), le contrat 2SHOP n'a toujours qu'un seul triplet N. Le code doit être dynamique, pas supposer un nombre fixe.

Le Frêt final vient de TCD!Y (la colonne "Formule Test à vérifier", calcul résiduel) et PAS de TCD!W (le calcul principal) — confirmé formule par formule dans "Fichier import"!T2 = =TCD!Y3. C'est contre-intuitif mais formellement vérifié.

Le classeur modèle (2026_06_...) est en fait un gabarit d'un autre mois que les PDF/factures fournis (nos de facture différents) — comme pour DPD/BLS, le modèle sert juste de gabarit de structure/formules, pas de données à comparer.

Les PDF sont bien structurés et parseables par regex (comme Mondial Relay/DPD).