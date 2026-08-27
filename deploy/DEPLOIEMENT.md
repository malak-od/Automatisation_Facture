# Déploiement — Facturation (PRODUCTION UNIQUEMENT)

> SSO : `authentik/AUTHENTIK.md`. Forward-auth NPM : `deploy/npm-facturation-advanced.conf`.
> Ce fichier décrit **où** poser quoi, et **avec quels noms**.

Cette app ne suit pas le schéma habituel du socle commun : elle **pilote Excel
via COM** (`pywin32`), donc elle **ne peut pas** tourner sur la VM Linux
`192.168.5.200` comme `module-transport` ou `gestion-du-temps` (démonstration
détaillée au §9). Elle tourne sur une **VM Windows dédiée avec Excel installé**,
dans le VLAN serveurs.

Elle ne suit pas non plus le socle sur le SSO : les autres apps portent l'OIDC
dans leur code, celle-ci délègue tout à un **forward-auth porté par NPM**
(`AUTHENTIK.md`). Il n'y a **aucun nginx intermédiaire** — NPM *est* nginx, et
il tape directement la VM Windows, conformément au §2.d du socle commun.

```
navigateur
   │ https://facturation.intra.laruche-logistique-france.fr
   ▼
UniFi / UDM ......... résolution DNS interne  ──> 192.168.5.11
   ▼
NPM (192.168.5.11) .. TLS du domaine + forward-auth (onglet Advanced)
   │                     └─ auth_request ──> Authentik (192.168.5.55:9443)
   │ http://192.168.5.74:4000
   ▼
node server.js ──> python finaliser_*.py ──> Excel (COM)
```

---

## 0. Nomenclature — les noms à utiliser partout

**Périmètre de l'app :** `facturation` est **l'application de facturation, tous
périmètres confondus**. Le module transporteurs en est le premier ; les suivants
(clients, fournisseurs…) viendront **dedans**, sous le même domaine, le même
provider Authentik, le même service et le même port. Ne rien créer « par
périmètre » : un seul objet dans chaque console, un groupe Authentik par
périmètre si un cloisonnement devient nécessaire (`AUTHENTIK.md` §4).

**Un seul nom par objet, identique dans les cinq consoles.** Les écarts de nom
sont la première cause de configuration introuvable six mois plus tard.

| Objet | Où | **Nom exact** |
|---|---|---|
| Domaine de l'app | partout | `facturation.intra.laruche-logistique-france.fr` |
| Enregistrement DNS | 🛜 UniFi | hôte `facturation` dans `intra.laruche-logistique-france.fr` → `192.168.5.11` |
| Proxy Host | 🔀 NPM | `facturation.intra.laruche-logistique-france.fr` → `192.168.5.74:4000` |
| Forward-auth | 🔀 NPM | onglet **Advanced** du Proxy Host — référence : `deploy/npm-facturation-advanced.conf` |
| Règles pare-feu | 🪟 VM app | `facturation - nginx uniquement` (autorise `.11`) · `facturation - blocage direct` |
| Provider | 🔑 Authentik | `facturation-web` |
| Application (slug) | 🔑 Authentik | `facturation` |
| Groupes | 🔑 Authentik | `facturation-administrateur` · `facturation-operateur` |
| VM hôte de l'app | 🖥️ hyperviseur | `CRE1-SV-EXEL-01` — `192.168.5.74`, IP **fixe**, VLAN serveurs |
| Accès RDP à la VM | 🪟 VM app | `192.168.5.74:2547` — port **non standard**, le `3389` est fermé |
| Tâche planifiée | 🪟 VM app | `facturation` — **tâche**, pas un service Windows (§1.c) |
| Dépôt GitLab | 🦊 GitLab | `developpement/automatisation-facturation` |

> Le dépôt garde son nom historique `automatisation-facturation` : renommer un
> projet GitLab **casse l'assignation du runner** (piège documenté au §2.a du
> socle commun). C'est le seul endroit où le nom diverge, et c'est volontaire.

**Ports.** Le tableau du §1 du socle commun attribue les ports **sur la VM
`.200`**, pour éviter les collisions entre apps qui y cohabitent. Facturation
n'y tourne pas : elle est seule sur `192.168.5.74` et y écoute le **4000**, qui
ne relève donc d'aucune convention. Le `3500` avait été réservé du temps où un
nginx intermédiaire était envisagé sur `.200` ; ce montage a été abandonné, et
`3500` reste **libre** pour une autre app.

---

## 1. 🪟 VM Windows — créer et préparer l'hôte de l'app

L'app est hébergée sur une **VM Windows dédiée, dans le VLAN serveurs** — pas
sur une station de travail. Ce choix règle trois problèmes d'un coup : l'IP
n'est plus distribuée en DHCP, le flux depuis NPM reste intra-VLAN, et la
facturation ne dépend plus d'un poste que quelqu'un peut éteindre ou déplacer.

### a. Gabarit de la VM

| | Valeur |
|---|---|
| Nom | `CRE1-SV-EXEL-01` |
| IP | **`192.168.5.74`** — fixe, dans `192.168.5.0/24` |
| OS | **Windows Server 2025** — édition **Desktop Experience** obligatoire (Excel ne s'installe pas sur Server Core) |
| Accès admin | **RDP sur le port `2547`**, pas le `3389` — NLA activé |
| Excel | **Office LTSC 2024**, licence en volume — c'est lui qui recalcule les TCD (§9). À préférer à M365 Apps, qui exigerait l'activation par ordinateur partagé et le rôle RDS |
| Ressources | **8 vCPU / 16 Go** recommandés ; 4 vCPU / 8 Go est un plancher. Le modèle UPS porte **201 793 formules `XLOOKUP`** (47 Mo) et Colissimo 113 928 : le moteur de calcul d'Excel est **multithread**, donc les cœurs réduisent directement la durée de génération |

> ⚠️ **La contrainte qui structure tout le reste : COM exige une session
> Windows interactive.** Excel refuse de démarrer sans profil utilisateur, donc
> le service ne peut pas tourner en `LocalSystem`. Il faut un compte de service
> local avec **autologon** et une session maintenue ouverte. C'est la partie
> fragile du montage — et la raison pour laquelle Microsoft ne supporte pas
> officiellement l'automatisation Office côté serveur. En pratique cela tourne,
> à condition de surveiller les `EXCEL.EXE` orphelins (§8).

> ✅ **Compatibilité des modèles avec LTSC 2024, vérifiée en août 2026.** Les 12
> classeurs n'emploient que deux fonctions modernes — `XLOOKUP` (11 modèles) et
> `_xlws.FILTER` (UPS) — toutes deux présentes depuis LTSC 2021. Aucune fonction
> réservée à M365 (`GROUPBY`, `PIVOTBY`, fonctions regex…), donc aucun risque de
> `#NOM?` silencieux dans une facture. À refaire si un modèle est remplacé par
> une version retravaillée sur un poste sous M365.

### b. Prérequis logiciels

Node 20+ et Python 3 accessible en `python` (voir `README.md`), puis :

```powershell
pip install -r automatisation/requirements.txt
cd facturation-app
npm ci
```

### c. Démarrage automatique — tâche planifiée, **pas** un service Windows

L'app doit démarrer seule au boot, sans terminal laissé ouvert. Le réflexe
serait d'en faire un service Windows (NSSM). **C'est un piège ici**, et il
annule tout le travail du §1.a.

> 🪤 **Un service Windows s'exécute toujours en session 0**, y compris quand on
> lui affecte un compte utilisateur : il ne rejoint pas la session interactive
> de ce compte, même ouverte par autologon. Or Excel COM a besoin d'un bureau
> interactif. Un service NSSM fonctionnerait donc en test manuel puis se
> bloquerait en production, sans erreur ni journal — le mode de panne le plus
> coûteux à diagnostiquer de tout ce document.

La tâche planifiée déclenchée **à l'ouverture de session** place au contraire
le processus dans la session interactive de l'autologon, celle où Excel est
configuré.

```powershell
$action  = New-ScheduledTaskAction -Execute "C:\Program Files\nodejs\node.exe" `
             -Argument "server.js" `
             -WorkingDirectory "C:\Projets\automatisation-facturation\facturation-app"
$trigger = New-ScheduledTaskTrigger -AtLogOn -User "Service-Excel"
$set     = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries `
             -DontStopIfGoingOnBatteries -ExecutionTimeLimit 0 `
             -RestartCount 3 -RestartInterval (New-TimeSpan -Minutes 1)
Register-ScheduledTask -TaskName "facturation" -Action $action -Trigger $trigger `
  -Settings $set -User "Service-Excel" -RunLevel Limited
```

- `-RunLevel Limited` : l'app n'a besoin d'aucun privilège. Surtout pas `Highest`.
- `-ExecutionTimeLimit 0` : sans quoi la tâche serait tuée au bout de 3 jours.
- **« N'exécuter que si l'utilisateur est connecté »** doit rester actif : c'est
  ce réglage qui garantit la session interactive.
- Pas de `PYTHONPATH` à poser : la tâche hérite de l'environnement de la
  session, donc `%APPDATA%` est correct et `pywin32` est trouvé même quand il
  est installé dans le profil (`pip install --user`).

Contrepartie : l'app ne démarre qu'après l'autologon, soit quelques secondes de
plus après un redémarrage. Sans conséquence pour un usage humain.

**Vérification après redémarrage :**

```powershell
Get-ScheduledTask facturation | Get-ScheduledTaskInfo
Get-NetTCPConnection -LocalPort 4000 -State Listen
```

`LastTaskResult = 267011` (`0x41303`, *task has not run*) n'est pas une erreur :
c'est l'état normal tant qu'aucune ouverture de session n'a eu lieu depuis
l'enregistrement. `267009` signifie que la tâche est en cours d'exécution.

> ⚠️ L'autologon est le préalable : sans lui, aucune session ne s'ouvre au
> démarrage, la tâche ne se déclenche jamais et l'app ne démarre pas. Vérifier
> `AutoAdminLogon = 1` et `DefaultUserName = Service-Excel` sous
> `HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon`.

> 🪤 `server.js` **auto-incrémente le port** si celui-ci est occupé (jusqu'à
> +10, cf. `start()` en fin de fichier). Pratique en développement, piège en
> production : si le `4000` est pris au démarrage du service, l'app part sur le
> `4001`, l'`upstream` nginx pointe dans le vide et le navigateur reçoit un
> `502` que rien n'explique. Vérifier après chaque démarrage du service :
> `Get-NetTCPConnection -LocalPort 4000 -State Listen`.

### d. Récupérer le code **et les modèles Excel**

Le chemin du clone est porteur, ce n'est pas un détail cosmétique : chaque
transporteur code son modèle en dur en `../Transporteurs/…`, relatif à
`facturation-app/` (ex. `src/carriers/bls/index.js:403`, et les 11 autres sur le
même schéma). Cloner ailleurs, ou ne cloner que `facturation-app/`, casse les
12 transporteurs au premier clic.

```powershell
git clone git@gitlab-ssh.intra.laruche-logistique-france.fr:developpement/automatisation-facturation.git `
  C:\Projets\automatisation-facturation
cd C:\Projets\automatisation-facturation
git checkout prod
```

- Le clone pèse **~225 Mo** : les 12 modèles de `Transporteurs/` sont versionnés
  dans le dépôt (pas de git-lfs), dont le classeur UPS de 48 Mo à lui seul.
- Prévoir une **clé de déploiement** GitLab sur la VM, sinon le `git clone` en
  SSH demandera une authentification interactive.
- `uploads/` et `outputs/` sont créés automatiquement au démarrage
  (`server.js:25-26`) : rien à créer à la main, mais le compte de service doit
  pouvoir **écrire** dans `facturation-app\`.

> 🪤 **Excel et l'emplacement approuvé — le piège qui ne produit aucune
> erreur.** Les modèles contiennent des TCD. Si
> `C:\Projets\automatisation-facturation` n'est pas déclaré **emplacement
> approuvé** dans le Centre de gestion de la confidentialité d'Excel, le
> classeur s'ouvre en mode protégé et l'automatisation COM reste bloquée sur une
> boîte de dialogue que personne ne voit : pas d'exception côté Python, pas de
> message côté Node, juste une génération qui ne finit jamais. Le réglage est
> **par profil utilisateur** (HKCU) : il doit être posé sous le **compte de
> service** du §1.a, pas sous ton compte d'administration.

## 2. 🛜 UniFi — DNS et flux réseau

1. **Enregistrement DNS interne** — Settings → Networks → *DNS* (ou le serveur
   DNS interne selon la configuration du site) :
   `facturation` → **`192.168.5.11`** (l'IP de NPM, pas celle de la VM app :
   c'est NPM qui porte le certificat du domaine).
2. **IP fixe de la VM** : `192.168.5.74` doit être fixe (configurée dans la
   VM, ou réservée sur la MAC côté DHCP). NPM la code en dur dans son Proxy
   Host — une IP qui bouge, c'est un `502` un matin sans que personne n'ait
   touché à la configuration.

> ✅ **Aucune règle de pare-feu UniFi n'est nécessaire** (vérifié en août 2026).
> NPM (`192.168.5.11`) et la VM app (`192.168.5.74`) sont dans le même `/24` :
> leur trafic est commuté et ne traverse jamais la passerelle, donc aucune règle
> ne le voit. La règle inter-VLAN décrite dans les versions précédentes de ce
> document datait de l'hébergement sur une station en `192.168.1.x`.

> Cette app n'a **pas** de base PostgreSQL : pas de règle `5432` à créer,
> contrairement à la check-list du socle commun.

## 3. 🔀 NPM (192.168.5.11:81) — Proxy Host **et** forward-auth

> 🪤 **Il n'y a pas de nginx sur `192.168.5.200`** — vérifié en août 2026,
> `command -v nginx` y répond `ABSENT`. Le socle commun (§2.d) prévoit NPM
> comme unique reverse-proxy, tapant directement le port de l'app ; aucune app
> n'a de nginx local. Les versions précédentes de ce document décrivaient un
> site nginx sur `.200` : ce montage n'a jamais existé.

**Onglet Details**

| Champ | Valeur |
|---|---|
| Domain Names | `facturation.intra.laruche-logistique-france.fr` |
| Scheme | `http` |
| Forward Hostname / IP | **`192.168.5.74`** |
| Forward Port | **`4000`** |
| Websockets Support | ON |
| **Cache Assets** | **OFF** |

> ⚠️ **Cache Assets doit rester désactivé.** Derrière un `auth_request`, une
> réponse servie depuis le cache échappe à la vérification d'authentification —
> et l'app diffuse des **classeurs de facturation générés** (`.xlsx`, `.csv`).
> Les mettre en cache, c'est risquer qu'un utilisateur reçoive la facture d'un
> autre.

**Onglet SSL** — certificat du domaine, **Force SSL** activé. Sans HTTPS forcé,
Authentik pose ses cookies sur une session HTTP et le SSO devient erratique.

**Onglet Advanced** — coller le contenu de `deploy/npm-facturation-advanced.conf`.
Il porte à la fois les timeouts (900 s, sans quoi les générations longues sont
coupées à ~60 s, cf. socle commun §2.d) et tout le forward-auth Authentik.

> 🪤 **Les en-têtes `X-authentik-*` n'atteignent pas l'app** avec ce montage. En
> nginx, les `proxy_set_header` d'un niveau supérieur sont ignorés dès qu'une
> `location` en définit un seul — et le `location /` généré par NPM en définit
> plusieurs. Sans conséquence aujourd'hui : `server.js` ignore ces en-têtes.
> Mais pour tracer qui a généré quelle facture (`AUTHENTIK.md` §5), il faudra
> reprendre la main sur `location /` via l'onglet **Custom Locations**.

> 🪤 **Cette configuration n'est pas versionnée** : c'est un champ texte dans
> une interface web. Pas de `nginx -t`, pas de git, et elle disparaît si
> quelqu'un recrée le Proxy Host. `deploy/npm-facturation-advanced.conf` en est
> la copie de référence — **le tenir à jour à chaque modification dans NPM**.

## 5. 🪟 Pare-feu Windows — l'étape à ne pas sauter

Le SSO est appliqué par NPM, **pas** par l'application. Tant que la VM Windows
répond à tout le LAN sur le port 4000, il suffit de taper
`http://192.168.5.74:4000` pour entrer sans authentification, et de forger
soi-même les en-têtes `X-authentik-*`. **Ce n'est pas théorique** : constaté en
août 2026 depuis un poste utilisateur, `HTTP 200` sans la moindre
identification, avant la pose de ces règles.

```powershell
New-NetFirewallRule -DisplayName "facturation - nginx uniquement" `
  -Direction Inbound -Protocol TCP -LocalPort 4000 -Action Allow `
  -RemoteAddress 192.168.5.11
New-NetFirewallRule -DisplayName "facturation - blocage direct" `
  -Direction Inbound -Protocol TCP -LocalPort 4000 -Action Block
```

`192.168.5.11` est l'IP de NPM : c'est le **seul** client légitime du port 4000.
Les tests locaux sur la VM (`http://localhost:4000`) ne traversent pas le
pare-feu entrant et continuent de fonctionner.

Variante plus sûre si l'app n'a rien à faire du reste du réseau : la faire
écouter uniquement en local (`HOST=127.0.0.1`) et poser un tunnel — mais en
l'état `server.js` n'expose pas de variable d'écoute, la règle de pare-feu
ci-dessus est la voie directe.

## 6. Authentik

Voir `authentik/AUTHENTIK.md` — Proxy Provider `facturation-web` en mode
*Forward auth (single application)*, rattaché à l'outpost embarqué, avec un
*Group binding* sur l'application (sans lui, **tout compte Authentik entre**).

## 7. Vérifications après déploiement

```bash
# 🔀 depuis NPM (.11) — l'app Windows repond a travers le reseau
curl -sS -o /dev/null -w "%{http_code}\n" http://192.168.5.74:4000/api/carriers   # 200
# depuis tout autre poste, ce meme curl DOIT echouer (pare-feu §5)

# 🐧 l'outpost Authentik répond (401 SANS cookie = correct)
# Les 3 en-tetes X-* sont indispensables : sans X-Original-URL l'outpost renvoie
# 500 meme quand tout est correct (cf. AUTHENTIK.md §6).
curl -sk -o /dev/null -w "%{http_code}\n" \
  -H "Host: facturation.intra.laruche-logistique-france.fr" \
  -H "X-Original-URL: https://facturation.intra.laruche-logistique-france.fr/" \
  -H "X-Forwarded-Proto: https" \
  -H "X-Forwarded-Host: facturation.intra.laruche-logistique-france.fr" \
  https://192.168.5.55:9443/outpost.goauthentik.io/auth/nginx                     # 401

# 🌐 depuis n'importe quel poste — NPM redirige vers le SSO au lieu de l'app
curl -sk -o /dev/null -w "%{http_code} %{redirect_url}\n" \
  https://facturation.intra.laruche-logistique-france.fr/
#    -> 302 vers /outpost.goauthentik.io/start?rd=/
```

Puis, en **navigation privée**, ouvrir
`https://facturation.intra.laruche-logistique-france.fr` : redirection Authentik,
identification, retour sur l'app, et générer une facture de bout en bout (c'est
le seul test qui valide aussi les timeouts).

## 8. Réflexes de diagnostic

| Symptôme | Première chose à vérifier |
|---|---|
| **502 Bad Gateway** | La tâche `facturation` tourne-t-elle ? Autologon actif ? Règle de pare-feu §5 trop stricte ? |
| **500 au lieu du login** | Provider non rattaché à l'outpost (`AUTHENTIK.md` §3) |
| **504 après ~60 s** | Timeout NPM non relevé (§4) — nginx seul ne suffit pas |
| **413 à l'envoi des factures** | `client_max_body_size` absent côté NPM (§4) |
| **On entre sans mot de passe** | Pare-feu Windows §5 absent, ou accès direct au `:4000` |
| **N'importe qui entre après login** | Aucun *Group binding* sur l'application |
| **« fichier ouvert dans Excel »** | Process `EXCEL.EXE` orphelin sur la VM Windows |
| **502 alors que la tâche tourne** | Port auto-incrémenté : l'app écoute sur 4001+ et non 4000 (§1.c) |
| **Génération qui ne finit jamais, sans erreur** | Excel en mode protégé : emplacement non approuvé pour le compte de service (§1.d) |
| **`FileNotFoundError` sur un modèle** | Dépôt cloné ailleurs que `C:\Projets\automatisation-facturation` (§1.d) |

## 9. Pourquoi l'app ne peut pas tourner sur la VM Linux `.200`

Question qui revient à chaque revue d'infrastructure. La réponse, vérifiée en
août 2026 :

- les **12** finaliseurs de `automatisation/` importent tous `win32com` — aucun
  n'utilise `openpyxl` seul ;
- ils ne se contentent pas d'écrire des cellules : ils appellent `RefreshAll()`,
  `Calculate()` et `CalculateUntilAsyncQueriesDone()` ;
- **11 des 12 modèles** de `Transporteurs/` contiennent des tableaux croisés
  dynamiques (de 6 à 18 entrées `pivotCache` par classeur ; seul
  `2026_06_Delivengo_LPPAQ.xlsx` n'en a aucune).

Ces TCD ne sont recalculés que par le moteur de calcul d'Excel. `openpyxl`
n'évalue aucune formule et ne rafraîchit aucun pivot ; LibreOffice headless
recalcule les formules mais rend mal les TCD. Sur des **factures**, un écart de
calcul silencieux est le pire mode de panne imaginable.

Déplacer l'app sur `.200` ne serait donc pas un changement de configuration,
mais la réécriture des 12 finaliseurs sans Excel. La VM `.200` garde son rôle :
nginx + forward-auth, rien d'autre.
