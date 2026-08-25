# Décisions importantes

## FedEx — colonnes CSV "en USD" trompeuses (2026-08-19)

Le CSV brut FedEx (export portail, en-tête FR) contient deux jeux de colonnes de montant :
- `Montant des frais de transport de l'envoi en USD` / `Montant de la remise de l'envoi en USD` (colonnes ~17/20) : montants en USD, **non utilisés** par le classeur réel.
- `Devise de facturation des frais de transport de l'envoi` / `Devise de facturation de la remise de l'envoi` (colonnes ~57/60) : noms trompeurs (« Devise de facturation »), mais ce sont en réalité des **montants en EUR** — le vrai code devise est dans la colonne suivante (`Code de la devise de facturation`).

Le classeur fait-main (`Shipment Detail!A = BK+BN`) utilise les colonnes EUR. Vérifié sur un cas réel (tracking `381547998902`, juin 2026) : USD `539.82-458.09=81.73` vs EUR `465.03-394.62=70.41` = valeur réelle du classeur modèle.

**Conséquence** : `facturation-app/src/carriers/fedex/index.js` et `automatisation/finaliser_fedex.py` doivent résoudre `iFret`/`iRemise`/`iDroitsTaxes` sur les colonnes EUR (« Devise de facturation ... »), jamais sur les colonnes « ... en USD ». Validé sur juin 2026 réel après correction : 45515,83 € de Fret, 0 écart sur 3607/3607 lignes.

## FedEx — factures d'un autre mois de facturation dans le CSV brut (question ouverte, 2026-08-19)

Le CSV brut FedEx contient parfois des lignes rattachées à une facture dont la quasi-totalité des expéditions datent d'un mois différent du mois cible (ex. juin 2026 : une facture de 392 lignes datée mi-mai, absente du fichier livré réel bien qu'émise le 01/06/2026).

Aucune règle simple (ni date d'expédition, ni majorité de lignes hors mois) ne reproduit exactement les exclusions constatées dans le fichier réel de juin (une facture 100% hors-mois est incluse, une autre avec des lignes dans le mois cible est exclue) — seule l'appartenance réelle à l'ensemble des factures du mois (`Bilan factures`) est fiable, information non déductible du CSV seul.

**Décision utilisateur (2026-08-19)** : garder les lignes qui ont été **facturées** au mois cible, même si expédiées un autre mois — mais cette règle contredit un cas réel observé (facture 634313590, facturée le 01/06 mais absente du fichier réel de juin).

**État actuel du carrier** : seule règle certaine appliquée automatiquement = exclusion des factures 100% Droits & Taxes (aucun fret, ajustement douanier pur). Pour toute facture sans aucune ligne dans le mois cible, le carrier **alerte sans exclure** — à trancher au cas par cas avec le pôle transport avant validation finale d'un mois donné.

## FedEx — carrier construit et validé (2026-08-19)

`facturation-app/src/carriers/fedex/` (Node) + `automatisation/finaliser_fedex.py` (clonage Excel COM, 6 feuilles) inscrits dans `facturation-app/src/registry.js` (`status: 'ready'`).

Validé sur juin 2026 réel : 3607/3607 lignes communes avec le fichier livré, 0 écart sur Transporteur/EP/Pays/Zone/Mode/TVA/Fret, dans les deux implémentations (Node et Python).

Piège PivotCache (même famille que DPD/Geodis/Mondial Relay/Chronopost/TNT) : les 3 PivotTables du modèle (TCD, Bilan factures, Bilan clients) ont un cache source étroit/figé — redirigé vers la vraie plage large de `Shipment Detail` à chaque génération.

Reste à faire : validation sur un 2e mois réel, mémoire `fedex_carrier_construit.md`.

## UPS — carrier construit et validé (2026-08-20)

`facturation-app/src/carriers/ups/` (Node) + `automatisation/finaliser_ups.py` (clonage Excel COM, 16 feuilles) inscrits dans `facturation-app/src/registry.js` (`status: 'ready'`).

Validé sur juin 2026 réel (`2026_06_UPS_Import.csv`, fichier livré) : 8719/8719 trackings communs. Fret, Colis volumineux, Droits et taxes, plus-value BtoC, Adresses : **exacts à 100%**. TVA : exacte à 99,99% (1 écart résiduel).

**Bugs trouvés et corrigés en cours de construction** :
- Colis volumineux : la formule modèle applique un barème par palier sur le **montant réel** du poste "Colis volumineux" facturé par UPS (`TCD!H`), pas sur le poids du colis — confusion initiale corrigée.
- TVA : teste la présence du poste "TVA" réel (catégorie UPS, `codeClasse="TAX"`), pas une liste de pays codée en dur.
- Zones éloignées : mauvaise colonne TCD référencée initialement (`Droits et taxes` au lieu de `Zones éloignées`) — 8719/8719 lignes forfaitées à tort avant correction.
- Zone : `Pays="FR"` force toujours `"France"` pour une zone native courte (1-2 chiffres), même non nulle — règle non triviale de la formule modèle (`IF(LEN(C)>2,C,IF(L="FR","France",...))`).
- Cascade Adresse/plus-value BtoC prioritaire sur Frêt/TVA : les colonnes `'ST SV'!Q:Q` et `D:D` (jamais vides contrairement à l'hypothèse initiale) forcent la catégorie "Adresse"/"plus-value BtoC" avant même de tester `codeClasse="FRT"`.
- **Piège ColField réordonné** (même famille que TNT/Chronopost) : le ColField natif du TCD "Catégorie" n'affiche que les postes réellement présents dans le mois traité — un mois avec un poste supplémentaire (ex. "code inconnu") décale toutes les colonnes suivantes. Fix : résolution dynamique des lettres de colonnes TCD par nom d'en-tête dans `finaliser_ups.py`, jamais en dur.
- Zones éloignées = 0€ facturé (ex. code "ESP" à 0,00€) déclenchait à tort le forfait 40€ (`TCD!col=""` ne capture pas `0`) — fix : traiter aussi `0` comme absence de charge.
- Trackings sans aucune charge facturable (ligne "INF"/retour indélivrable isolée) : gardés dans `Facture UPS` (données brutes complètes), exclus uniquement de `Fichier import` (décision utilisateur 2026-08-20) — confirmé 0 ligne totalement vide sur 8736 dans le fichier réel de juin.

**Limitation connue non résolue** (avant ce dernier fix, observé sur test réel) : ~9% des trackings peuvent avoir un écart "Zones éloignées" dû à l'ambiguïté PivotTable Excel entre case vide et case à 0 — fix appliqué (`TCD!col=0` traité comme absence) mais pas re-testé en conditions réelles complètes (~10 min par run) avant la mise en prod. **À vérifier sur le premier mois réel traité en production.**

**E/P** : dépend fortement de la disponibilité des trackings dans l'export WMS partagé (`Automatisation/AAAA MM - Export expéditions_brut.xlsx`, mois cible + mois-1) — taux de correspondance variable selon les mois (91,8% observé sur le test réel), repli sur `plus-value BtoC` puis `'P'` par défaut sinon.

Reste à faire : mémoire `ups_carrier_construit.md`, vérification du "Mode envoi" résiduel (5 écarts/8719, cas limite non investigué), surveillance du taux de correspondance E/P sur les premiers mois réels.

## Champ "Mois de facturation" rendu autoritaire (2026-08-20)

Remontée pôle transport : sur UPS, la "date validité" du fichier import est parfois fausse (décale le mois de juin au lieu de juillet) — cause : `moisCible` auto-détecté par majorité sur la colonne "Date de la facture" du CSV brut, peu fiable.

**Décision utilisateur** : le champ `<select>` "Mois de facturation" de l'UI (`public/index.html`), jusqu'ici **informatif uniquement** (décision du 2026-08-14, jamais transmis au serveur), devient la **source de vérité de la période**, envoyé au serveur et prioritaire sur l'auto-détection — décision explicitement étendue à tous les transporteurs, pas seulement UPS.

**Implémentation** :
- `public/index.html` (`run()`) : envoie `period` (format `AAAA_MM`) dans le FormData.
- `server.js` (`/api/process`) : parse `req.body.period`, construit `{ formatted: "AAAA_MM", compact: "AAAAMM" }`, l'utilise pour le nommage de sortie (`period = periodOverride.formatted || result.period || 'export'`), et le transmet en 2e argument à `carrier.process(files, { period })`.
- Chaque carrier garde la signature `process(files)` et ignore ce 2e argument tant qu'il n'a pas été adapté — **aucune régression** sur les carriers non modifiés.
- Seul **UPS** (`src/carriers/ups/index.js`) a été adapté pour l'instant : `opts.period.compact` prioritaire sur `moisCible` auto-détecté (majorité sur "Date de la facture"), avec repli sur l'ancien comportement si le champ n'est pas fourni.

**Reste à faire** : étendre la prise en compte d'`opts.period` aux autres carriers si un besoin similaire remonte (aucun autre carrier n'a signalé ce bug à ce jour) ; valider le fix UPS sur un cas réel où l'ancienne auto-détection se trompait.

## BLS — navette redevient refacturable au client (2026-08-24)

Remontée pôle transport : "étendre formule Frêt (ne pas supprimer pour navette)".

Contexte : depuis le 2026-08-12, les lignes "navette" (trajet interne Longvic↔Créancey, non affrété via AffreTrans) étaient volontairement mises à `Frêt=0` dans le fichier import ERP — décision de l'époque : "PAS refacturable au client", même si le montant réel payé à BLS était conservé pour la réconciliation PDF.

**Décision utilisateur (2026-08-24)** : cette règle est inversée — la navette est désormais **refacturée au client** avec son montant réel (`rec.montantHt`), comme toute autre ligne. `Poids`/`Nbr Colis` restent forfaitaires (13200 kg / 33, pas de jointure AffreTrans possible pour ce trajet interne).

**Implémentation** :
- `facturation-app/src/carriers/bls/index.js` : `Fret: rec.isNavette ? 0 : rec.montantHt` → `Fret: rec.montantHt`.
- `Automatisation/finaliser_bls.py` : suppression du gel post-FillDown qui figeait la cellule Frêt (colonne T) à 0 pour les lignes navette — la formule `='Factures BLS'!J{n}` (déjà correcte, montant réel) s'applique désormais aussi à la navette.
- Messages d'info mis à jour des deux côtés (Node + Python) pour refléter "Frêt refacturé au client" au lieu de "Frêt=0".
- `validate()` continue d'exclure les lignes navette (filtre `_isNavette`) pour éviter du bruit sur `NbrColis`/`Poids` forfaitaires — n'a plus d'impact sur Frêt puisque ce n'est plus un 0 artificiel.

**Non testé** : pas de re-génération sur un fichier réel après ce changement (juin/juillet 2026 avaient déjà été validés avec l'ancienne règle Frêt=0) — à vérifier au prochain mois traité.

## FedEx — extraction "Zones éloignées" PDF simplifiée au mot-clé générique "zone" (2026-08-24)

Remontée pôle transport : "recherche pdf poids, zone éloigné à mettre dans le fichier import csv, chercher que le mot-clé 'zone'".

Le carrier extrayait déjà "Zones éloignées" depuis le PDF via 2 patterns précis ("Frais de traitement des importations..." et "Supplément pour Livraison Hors Zone", ajoutés le 2026-08-20 suite au mail "Oubli facturation surplus Fedex janvier à mai 2026"). Simplifié en un seul pattern générique cherchant le mot "zone" (insensible à la casse) dans `facturation-app/src/carriers/fedex/index.js` (`extractSupplements`) et `Automatisation/finaliser_fedex.py` (`kw_pats`).

**Découverte en testant sur les 9 PDF réels disponibles** (`Transporteurs/Fedex/FEDEX_*.pdf`) : le nouveau pattern générique capture un libellé **jamais géré jusqu'ici** — `"Livraison hors zone Tiers C"` (précédé de la ligne "ExpéditeurDestinataire") — présent en grand nombre (14 à 27 occurrences par PDF sur plusieurs fichiers testés), en plus des 2 libellés déjà couverts. Contexte PDF brut vérifié (`Charges / Montant(EUR)` ... `19,00 / ExpéditeurDestinataire / Livraison hors zone Tiers C`) : le montant est correctement associé à ce libellé, pas un artefact d'extraction. C'est donc un **vrai surplus supplémentaire non facturé jusqu'ici**, au-delà de ce qui avait été identifié le 2026-08-20.

**Non traité dans ce changement** : le "Poids" mentionné dans le même feedback pôle transport n'a pas été touché — l'utilisateur a confirmé que ce point est hors scope de cette demande (feedback ambigu, tranché en faveur de la seule simplification "zone").

**À faire** : valider l'impact chiffré de ce surplus nouvellement capté sur un mois réel (comparer aux montants déjà réconciliés PDF), informer le pôle transport si le delta est significatif.

**Complément (2026-08-25)** : le pôle transport a précisé les 2 mots-clés attendus pour "Zones éloignées" : **"zone"** (déjà en place) **et "Frais de traitement des importations aux États-Unis"** (libellé précis, ne contient pas "zone" donc pas capté par le 1er pattern — ajouté en complément, les deux patterns s'additionnent par tracking s'ils matchent des lignes distinctes). Modifié dans `facturation-app/src/carriers/fedex/index.js` (`kwPats.ZonesEloignees`) et `Automatisation/finaliser_fedex.py` (`kw_pats["ZonesEloignees"]`).

**Bug trouvé et corrigé (2026-08-25)** : premier ajout testé sur les 9 PDF disponibles à ce moment (juin/juillet 2026, racine `Transporteurs/Fedex/`) sans aucune occurrence du libellé — pattern non vérifié sur un cas réel. Le pôle transport a signalé 2 bordereaux réels (`382198649394` et `382408298943`, PDF `Transporteurs/Fedex/juillet 2026/FEDEX_634373761_781542172.pdf`, 2,40€ chacun) où le pattern ne matchait pas. Cause : le PDF **tronque le libellé à "...aux États-Uni" (sans le "s" final)** — probablement une troncature d'affichage du document FedEx lui-même, pas un artefact d'extraction. Fixé en rendant le "s" optionnel (`[ÉE]tats-Unis?`). Revérifié sur ce PDF réel : les 2 occurrences sont maintenant capturées avec le bon montant (2,40€ chacune).

## FedEx — "Colis volumineux" : 2e libellé PDF "Supplément pour manutention supplémentaire :poids" (2026-08-25)

Le pôle transport a précisé que "Colis volumineux" doit aussi capter le supplément **poids**, pas seulement dimension : "les mots-clés 'dimension' et aussi 'poids'". Vérifié sur bordereau réel `873870075714` (PDF `Transporteurs/Fedex/juillet 2026/FEDEX_634374150_200720433.pdf`) : le même bordereau porte **55,00€ "Charge pour dépassement de dimension"** ET **18,00€ "Supplément pour manutention supplémentaire :poids"**, 2 lignes distinctes qui doivent s'additionner (73€ total).

Ajouté un 2e pattern dans `kwPats.ColisVolumineuxPdf` (`facturation-app/src/carriers/fedex/index.js`) et `kw_pats["ColisVolumineuxPdf"]` (`Automatisation/finaliser_fedex.py`) : `Supplément pour manutention supplémentaire\s*:?\s*poids` (insensible à la casse). Le mécanisme d'agrégation existant (`acc[poste] = round2(acc[poste] + montant)`) additionne déjà correctement plusieurs matches sur le même poste — aucun changement nécessaire côté agrégation. Testé sur le PDF réel : 55€ + 18€ = 73€ capturés correctement pour ce tracking.

**Total "Colis volumineux" possible désormais** : règle Excel (longueur > 60cm → 10€) + "Charge pour dépassement de dimension" (PDF) + "Supplément pour manutention supplémentaire :poids" (PDF) — 3 composantes indépendantes qui s'additionnent.

## Colissimo/FedEx — fichier import CSV/XLSX reconstruit depuis le classeur (2026-08-24)

Remontée pôle transport : "le fichier import de quelques transporteurs en CSV déconne alors que la feuille Import CSV du fichier de facture est correcte, donc il faut copier-coller depuis le fichier de la facture, coller en valeur et vérifier tout ce qui est à vérifier." Périmètre confirmé par l'utilisateur : **Colissimo et FedEx uniquement** (pas les 10 autres carriers avec finaliseur, pour rester dans la règle "ne pas modifier plusieurs transporteurs pour un problème local").

**Cause racine** : le fichier import CSV/XLSX livré était calculé **indépendamment en JS** (`result.importRows`, logique dupliquée du classeur) et écrit **avant** la génération du classeur final par le finaliseur Python (Excel COM, qui applique les vraies formules du modèle : XLOOKUP, TCD...). Les deux chemins de calcul pouvaient diverger.

**Implémentation** :
- `Automatisation/finaliser_colissimo.py` / `finaliser_fedex.py` : après `wb.Save()`/`wb.Close()`, le classeur généré est **rouvert en lecture seule** et l'onglet Import ("Import CSV" / "Import ERP") est relu via `Range.Value` (COM, valeurs déjà calculées) et exporté dans un CSV intermédiaire `<sortie>_import_valeurs.csv`, signalé au appelant via `EXPORT_IMPORT_VALEURS:<chemin>` sur stdout.
- `facturation-app/src/core/excelOut.js` : nouvelle fonction `readImportRowsFromValuesCsv()` qui relit ce CSV et reconstruit les objets `importRows` (22 colonnes standard, dates converties en JJ/MM/AAAA).
- `facturation-app/server.js` : l'étape "1) Import ERP" (écriture CSV/XLSX) est déplacée **après** l'étape "2) Classeur" ; capture `EXPORT_IMPORT_VALEURS:` dans le stdout du finaliseur ; si `carrier.importFromWorkbook` est vrai et le fichier de valeurs existe, il remplace `result.importRows` pour l'écriture finale (repli sur l'ancien comportement sinon).
- `carriers/colissimo/index.js` / `carriers/fedex/index.js` : nouveau flag `importFromWorkbook: true`.

**3 bugs trouvés et corrigés pendant la validation (tests HTTP end-to-end sur juillet 2026 réel pour Colissimo, juin 2026 réel pour FedEx)** :
1. **Valeur d'erreur COM brute** : lire `Range.Value` juste avant `Save()`, même après un `Calculate()` explicite, pouvait renvoyer une cellule encore en état transitoire (erreur COM du type `-2146826246` = `0x800a07fa`, écrite telle quelle comme un "tracking" absurde). Fix : lire **après** `Save()`/`Close()`, sur le classeur **rouvert** (même état que ce que le pôle transport voit en ouvrant le fichier) — garde un filet de sécurité résiduel (`clean()`) au cas où.
2. **Ligne fantôme (item de cache PivotTable `"(blank)"` légitime)** : le TCD (RowField "Tracking") peut contenir un item `"(blank)"` réel quand le CSV brut a au moins une ligne sans tracking (ex. ajustement/note de crédit sans envoi associé) — ce item se propageait jusqu'à l'onglet Import. Le carrier Node (`index.js`) exclut déjà ce cas (`if (!tracking) continue;`) ; même filtre appliqué à la relecture Python pour rester cohérent.
3. **EXCEL.EXE orphelin verrouillant le fichier de sortie** : une exception entre l'ouverture en lecture du classeur (étape de relecture) et sa fermeture laissait le process Excel bloqué en arrière-plan, empêchant toute régénération suivante du même mois ("fichier ouvert dans Excel"). Fix : `wbRead.Close()` déplacé dans un bloc `finally` imbriqué, garanti quoi qu'il arrive.

**Écart de comptage résiduel (mineur, documenté)** : `result.lignes` (affiché à l'écran, basé sur `result.importRows.length` calculé en JS) peut différer de 1 par rapport au nombre réel de lignes du fichier livré (basé sur le classeur) — observé sur FedEx juin 2026 (358 vs 357). Pas d'impact sur le contenu livré, uniquement sur le compteur affiché — à surveiller si l'écart grandit sur un futur mois.

**Validé end-to-end** : Colissimo juillet 2026 réel (3187 lignes, 0 warning), FedEx juin 2026 réel (357 lignes livrées, avertissements PDF/E-P déjà connus et sans rapport avec ce changement).

## UPS — lignes sans identification supprimées de "Facture UPS" (2026-08-25)

Demande utilisateur : les lignes du CSV brut UPS Billing sans "Numéro de suivi" (Tracking) **ET** sans "Numéro de référence 1" (Ref1, colonne T "Numero de reference 1 de l'envoi") doivent être supprimées, pas seulement du fichier import mais de l'onglet "Facture UPS" lui-même (contrairement au filtre "Montant=0" d'autres transporteurs qui ne touche que l'import).

**Implémentation** : `facturation-app/src/carriers/ups/index.js` et `Automatisation/finaliser_ups.py` — même filtre appliqué dans la boucle qui construit `lignesRetenues`/`lignes_retenues` (déjà utilisée pour exclure les colis 1Z79), avant que ces lignes n'alimentent "Facture UPS" et l'agrégation par tracking. `ref1` (colonne T) ajouté à la table `COL` côté Python (absente jusqu'ici, seul le Node l'avait déjà repérée).

**Testé sur les 56 CSV réels de juin 2026** (`Transporteurs/UPS/csv/`, 60973 lignes brutes) : **10467 lignes (~17%) supprimées** par ce filtre, pour un montant NET total de seulement **14,26€** — confirmé que ce sont des lignes techniques/administratives sans vraie valeur financière (codes description dominants : "01", "OSW", "OFW", "PDS" — pas des postes de frais réels), pas des charges facturables perdues silencieusement. Cohérent avec la demande.

**Complément (2026-08-25)** : Ref1 vide ne suffit pas — UPS renseigne parfois une valeur non-informative à la place d'un vrai vide. Confirmé sur données réelles : `"."` (9 occurrences sur juin 2026) et `"null"` en texte littéral (3 occurrences, toutes avec tracking vide). Les deux valeurs sont désormais traitées comme équivalentes à "Ref1 vide" via une fonction `refVide()`/`ref_vide()` (Node/Python) : `!v || /^\.+$/.test(v) || v.toLowerCase() === 'null'`.

## UPS — "Demande avoir" (colis 1Z79) automatisé (2026-08-25)

Investigation lancée suite à un écart de volume massif entre le fichier généré par l'app et le fichier fait-main de juillet 2026 (`Transporteurs/UPS/juillet 2026/`) : 40382 lignes vs 25346 lignes dans "Facture UPS". Diagnostic en 2 temps :
1. La quasi-totalité des factures avaient un écart mineur (1-4 lignes) explicable par le filtre "sans identification" ci-dessus, sauf une facture géante (`202600782885`, +13513 lignes) où l'écart correspondait exactement aux lignes à **Montant net = 0** — non lié à "Demande avoir", traité séparément (pas encore corrigé à ce jour, cf. section précédente pour le contexte).
2. Une facture (`202600793898`, 100% de lignes en tracking `1Z79...`) était **totalement absente** du fichier app (0/49 lignes) alors que le fait-main en gardait 49 — révélant que le fait-main **garde les 1Z79 dans "Facture UPS"**, contrairement à ce que documentait le code jusqu'ici ("EXCLU de l'import", ambigu — en réalité seul le fichier import les exclut, pas le classeur).

**Décision utilisateur** : les colis 1Z79 (retour viticulteur chez La Ruche) doivent :
- Rester dans "Facture UPS" (données brutes, comme le fait-main).
- Être exclus du fichier import ERP (comme avant).
- Être reportés automatiquement dans l'onglet **"Demande avoir"** du classeur (déjà présent dans le modèle, jusqu'ici rempli à la main uniquement) : colonnes Tracking / Nb colis / Montant / Cause remplies, Factures / Poids / Mode livraison laissés vides (jamais renseignés même dans le fait-main — saisie manuelle du pôle transport).

**Règle du Montant** (2 itérations, corrigée après tests sur données réelles) :
- 1ère tentative : `codeClasse="FRT"` strict — **0/3 cas d'écart réexaminés ne correspondait**.
- **Règle finale confirmée** : somme de **toutes** les lignes du tracking **sauf** TVA (`codeClasse="TAX"`) et Taxe gazole (`codeClasse="FSC"`) — validé exact sur 13/14 trackings réels de juillet 2026 (12 corrects, 1 résiduel `1Z79A7T06819992295` à 14,29€ calculé vs 23,02€ dans le fait-main, aucune combinaison de lignes brutes ne l'explique — probable ajustement manuel isolé du pôle transport).

**Règle du Nb colis** : fixé à **1** par tracking — aucune formule de somme sur les colonnes brutes (Nombre de colis, filtré ou non) ne collait au fait-main ; `1` est la valeur observée sur 13/14 trackings réels testés.

**Implémentation** :
- `facturation-app/src/carriers/ups/index.js` : `demandesAvoir1Z79` (Map tracking→montant) construite dans la même boucle que le filtre Ref1/1Z79 ; les 1Z79 restent dans `lignesRetenues` (donc dans "Facture UPS" via le finaliseur) mais sont explicitement exclus de `parTracking` (donc de `importRows`). Nouveau champ `demandesAvoir` exposé dans le retour de `process()` (traçabilité — le Node n'écrit pas le classeur).
- `Automatisation/finaliser_ups.py` : même logique (`demandes_avoir_1z79`). Le TCD PivotTable "TCD" (source redirigée vers toute la plage "Facture UPS") aurait automatiquement réintégré les 1Z79 dans "Fichier import" — **masqués via `PivotItems().Visible = False`** sur le champ RowField "Numéro de suivi" (volume négligeable : 14-15 trackings distincts/mois, pas d'impact perf). L'onglet "Demande avoir" est purgé (lignes du modèle cloné) puis réécrit en une seule opération `Range.Value` juste avant `wb.Save()`.

**Validé end-to-end sur juillet 2026 réel** (52 CSV, 49485 lignes brutes) : 14 colis 1Z79 détectés, masqués du TCD, "Fichier import" confirmé à 0 ligne 1Z79 (contre 58 lignes présentes dans "Facture UPS", cohérent avec les données brutes). "Demande avoir" : 12/14 trackings exacts au centime près face au fichier fait-main.

**Écart connu non résolu** : 5 lignes du fait-main dans "Demande avoir" ont un tracking `1ZA1912W...` (préfixe différent de `1Z79`, rattachées au compte UPS `0000A1912W`) — non couvertes par le filtre actuel (basé sur le préfixe tracking uniquement). Mis en pause sur demande explicite de l'utilisateur ("je dois vérifier ce compte d'abord") — ne pas étendre la règle sans confirmation.
