# Déploiement — Facturation (PRODUCTION UNIQUEMENT)

> SSO : `authentik/AUTHENTIK.md`. Configuration nginx : `deploy/nginx-facturation.conf`.
> Ce fichier décrit **où** poser quoi, et **avec quels noms**.

Cette app ne suit pas le schéma habituel du socle commun : elle **pilote Excel
via COM** (`pywin32`), donc elle **ne peut pas** tourner sur la VM Linux
`192.168.5.200` comme `module-transport` ou `gestion-du-temps` (démonstration
détaillée au §9). Elle tourne sur une **VM Windows dédiée avec Excel installé**,
dans le VLAN serveurs ; la VM `.200` ne fait que du reverse-proxy et porte le SSO.

```
navigateur
   │ https://facturation.intra.laruche-logistique-france.fr
   ▼
UniFi / UDM ......... résolution DNS interne  ──> 192.168.5.11
   ▼
NPM (192.168.5.11) .. TLS du domaine, Proxy Host
   │ http://192.168.5.200:3500
   ▼
nginx (VM 192.168.5.200) ── auth_request ──> Authentik (192.168.5.55)
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
| Règle de pare-feu | 🛜 UniFi | `ALLOW-facturation-nginx-vers-windows-4000` |
| Proxy Host | 🔀 NPM | `facturation.intra.laruche-logistique-france.fr` → `192.168.5.200:3500` |
| Fichier de site | 🐧 VM `.200` | `/etc/nginx/sites-available/facturation` |
| Journaux nginx | 🐧 VM `.200` | `/var/log/nginx/facturation.{access,error}.log` |
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

**Ports réservés pour cette app** (tableau §1 du socle commun) : **3500** en
production. `3400` était déjà pris ; `3501` reste libre si une recette est créée
un jour. Le `4000` de la VM Windows n'est **pas** un port de la convention :
c'est le port local de Node sur la VM Windows, jamais exposé au-delà de la
VM `.200`.

---

## 1. 🪟 VM Windows — créer et préparer l'hôte de l'app

L'app est hébergée sur une **VM Windows dédiée, dans le VLAN serveurs** — pas
sur une station de travail. Ce choix règle trois problèmes d'un coup : l'IP
n'est plus distribuée en DHCP, le flux vers nginx redevient intra-VLAN, et la
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
   VM, ou réservée sur la MAC côté DHCP). L'`upstream` nginx la code en dur —
   une IP qui bouge, c'est un `502` un matin sans que personne n'ait touché à
   la configuration.
3. **Règle de pare-feu** `ALLOW-facturation-nginx-vers-windows-4000` :
   autoriser **`192.168.5.200` → `192.168.5.74` TCP 4000**. C'est le seul
   flux entrant nécessaire côté VM Windows. La VM étant dans le même VLAN que
   nginx, c'est une règle **intra-VLAN** — et non plus une règle de routage
   inter-VLAN comme du temps de l'hébergement sur station de travail.
4. Le flux `192.168.5.200 → 192.168.5.55` (Authentik) doit être ouvert pour le
   `auth_request` — il l'est déjà pour les autres apps hébergées sur `.200`.

> Cette app n'a **pas** de base PostgreSQL : pas de règle `5432` à créer,
> contrairement à la check-list du socle commun.

## 3. 🐧 VM 192.168.5.200 — nginx

```bash
sudo cp deploy/nginx-facturation.conf /etc/nginx/sites-available/facturation
sudo ln -s /etc/nginx/sites-available/facturation /etc/nginx/sites-enabled/
# ⚠️ renseigner l'IP fixe de la VM Windows dans le bloc `upstream facturation_app`
sudo nano /etc/nginx/sites-available/facturation
sudo nginx -t && sudo systemctl reload nginx
```

Vérifier que le module d'authentification déléguée est bien compilé :

```bash
nginx -V 2>&1 | grep -o with-http_auth_request_module   # doit répondre
```

## 4. 🔀 NPM (192.168.5.11:81) — Proxy Host

- **Domain Names** : `facturation.intra.laruche-logistique-france.fr`
- **Forward Hostname/IP** : `192.168.5.200` — **Port : `3500`** — scheme `http`
- **Websockets Support** : ON
- **SSL** : certificat du domaine, *Force SSL* activé
- **Advanced** — obligatoire, sinon les générations longues sont coupées :

```
proxy_read_timeout 900s;
proxy_send_timeout 900s;
client_max_body_size 200m;
```

> 🪤 Le piège des ~60 s du socle commun s'applique ici de plein fouet : une
> génération Excel dépasse largement la minute sur un gros transporteur. Le
> timeout doit être relevé **aux deux étages** (NPM *et* nginx), sinon c'est
> l'étage le plus bas qui coupe.

## 5. 🪟 Pare-feu Windows — l'étape à ne pas sauter

Le SSO est appliqué par nginx, **pas** par l'application. Tant que la machine
Windows répond à tout le LAN sur le port 4000, il suffit de taper
`http://192.168.5.74:4000` pour entrer sans authentification, et de forger
soi-même les en-têtes `X-authentik-*`.

```powershell
New-NetFirewallRule -DisplayName "facturation - nginx uniquement" `
  -Direction Inbound -Protocol TCP -LocalPort 4000 -Action Allow `
  -RemoteAddress 192.168.5.200
New-NetFirewallRule -DisplayName "facturation - blocage direct" `
  -Direction Inbound -Protocol TCP -LocalPort 4000 -Action Block
```

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
# 🐧 VM .200 — l'app Windows répond bien à travers le réseau
curl -sS -o /dev/null -w "%{http_code}\n" http://192.168.5.74:4000/api/carriers   # 200

# 🐧 VM .200 — l'outpost Authentik répond (401 SANS cookie = correct)
curl -sk -o /dev/null -w "%{http_code}\n" \
  -H "Host: facturation.intra.laruche-logistique-france.fr" \
  https://192.168.5.55:9443/outpost.goauthentik.io/auth/nginx                     # 401

# 🐧 VM .200 — nginx redirige bien vers le SSO au lieu de servir l'app
curl -sS -o /dev/null -w "%{http_code} %{redirect_url}\n" \
  -H "Host: facturation.intra.laruche-logistique-france.fr" \
  http://127.0.0.1:3500/                    # 302 vers /outpost.goauthentik.io/start
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
