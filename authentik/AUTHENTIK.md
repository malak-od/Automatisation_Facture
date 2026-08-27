# Authentik — authentification de Facturation

`facturation` est **l'application de facturation de la maison, tous périmètres
confondus**. Le module **transporteurs** en est le premier ; les périmètres
suivants (clients, fournisseurs…) viendront **dans cette même application** —
même domaine, même provider, même service. Il n'y a donc **qu'un seul objet
Authentik à créer**, pas un par périmètre.

L'application délègue TOUTE l'authentification à **Authentik**, mais **pas de la
même façon que `gestion-du-temps`** : ici, aucune ligne de code applicatif n'est
concernée.

| | `gestion-du-temps` | **`facturation`** |
|---|---|---|
| Type de provider | OAuth2/OpenID (client public) | **Proxy Provider** (forward-auth) |
| Qui fait le flux ? | l'app Flutter, l'API vérifie le jeton | **nginx** (`auth_request`), avant l'app |
| Code applicatif | `server/lib/auth.js` (JWKS, RS256) | **aucun** — `server.js` est inchangé |
| Identité côté app | claims du jeton | en-têtes HTTP `X-authentik-*` |

**Pourquoi ce choix :** l'app pilote Excel via COM et n'a jamais eu d'auth. Y
greffer un BFF OIDC toucherait le chemin critique de génération des factures ;
le forward-auth met la barrière **avant** Node et laisse le code intact.

> ⚠️ Conséquence à ne pas oublier : la sécurité repose entièrement sur le fait
> que **personne ne peut joindre la VM Windows en direct** sur le port
> 4000. Voir `deploy/DEPLOIEMENT.md` §5 (règle de pare-feu Windows). Sans elle,
> le SSO se contourne en tapant `http://192.168.5.74:4000`.

---

## 1. Créer le Proxy Provider — 🔑 console `auth.intra…`

**Applications → Providers → Create → Proxy Provider**

- **Name** : `facturation-web`
- **Authorization flow** : `default-provider-authorization-explicit-consent`
  (ou `implicit-consent` si on ne veut pas d'écran de consentement)
- **Mode** : **Forward auth (single application)**
  — ⚠️ **pas** « Proxy », **pas** « Forward auth (domain level) ». Le mode
  *domain level* protégerait tous les `*.intra…` d'un coup avec un seul cookie ;
  ce n'est pas ce qu'on veut pour une app isolée.
- **External host** : `https://facturation.intra.laruche-logistique-france.fr`
  — c'est cette valeur qu'Authentik compare à l'en-tête `Host` reçu de nginx
  pour retrouver le provider. **Une faute ici = 400/404 sur `/auth/nginx`.**
- **Token validity** : laisser la valeur par défaut (`hours=24`).

## 2. Créer l'Application

**Applications → Applications → Create**

| Champ | Valeur |
|---|---|
| Name | `Facturation` |
| Slug | `facturation` |
| Provider | `facturation-web` |

Le slug pilote l'entrée dans le portail « Mes applications » — c'est par là que
les utilisateurs entreront (§4 du socle commun). **Une seule entrée de portail**,
quel que soit le nombre de périmètres de facturation à l'intérieur.

## 3. Rattacher le provider à l'outpost

**Applications → Outposts → `authentik Embedded Outpost` → Edit → Applications**
→ cocher l'application ci-dessus → Update.

Dans le même écran, vérifier le champ **Configuration** de l'outpost :

```yaml
authentik_host: https://auth.intra.laruche-logistique-france.fr/
authentik_host_insecure: false
authentik_host_browser: https://auth.intra.laruche-logistique-france.fr/
```

> 🪤 **`authentik_host` est le piège propre à notre montage.** nginx joint
> Authentik par son **IP** (`192.168.5.55`) avec un `Host:` qui vaut
> `facturation.intra…`. Si `authentik_host` est vide, l'outpost construit
> l'URL de redirection à partir de ce qu'il reçoit et **renvoie le navigateur
> vers une adresse injoignable** (l'IP interne, ou `facturation.intra…/if/flow/`
> qui n'existe pas). Il doit contenir l'URL **publique** d'Authentik, celle que
> le navigateur sait ouvrir.

> 🪤 **Piège n°1 de ce mode :** un provider créé mais **non rattaché à un
> outpost** répond `404` sur `/outpost.goauthentik.io/auth/nginx`. nginx traduit
> ça en `500` côté navigateur (`auth_request` n'accepte que 2xx/401/403), et
> rien dans le journal ne dit « outpost ». Vérifier ce rattachement en premier.

## 4. Groupes (rôles et périmètres)

**Directory → Groups → Create**, préfixés par l'app pour ne pas déborder sur les
autres (convention du socle commun §2.e) :

- `facturation-administrateur`
- `facturation-operateur`

Puis **restreindre l'accès** : sur l'Application → **Policy / Group / User
Bindings** → **Bind existing group** → ajouter les deux groupes. **Sans binding,
tout compte Authentik valide entre dans l'app** — le forward-auth vérifie
seulement qu'une session existe.

> **Quand d'autres périmètres arriveront** (clients, fournisseurs…), ne pas
> créer une nouvelle application Authentik : ajouter un groupe par périmètre —
> `facturation-transporteurs`, `facturation-clients`, … — et laisser l'app
> filtrer ce qu'elle affiche d'après `X-authentik-groups` (§5). Le SSO reste un
> seul point d'entrée ; c'est le code qui découpe les périmètres.

Les rôles ne sont **pas** exploités par le code aujourd'hui : l'app ne distingue
ni administrateur ni opérateur, et ne montre qu'un périmètre. Les groupes
servent pour l'instant de **liste d'accès** (qui entre / qui n'entre pas).

## 5. Ce que l'app reçoit (et n'utilise pas encore)

nginx transmet à Node, sur chaque requête :

```
X-authentik-username : jdupont
X-authentik-email    : j.dupont@laruche-logistique-france.fr
X-authentik-name     : Jean Dupont
X-authentik-groups   : facturation-operateur|facturation-transporteurs|…
X-authentik-uid      : <identifiant stable>
```

`facturation-app/server.js` les ignore. Ces en-têtes sont la seule source
d'identité disponible pour les deux évolutions déjà prévisibles :

1. **tracer qui a généré quelle facture** — lire `req.get('x-authentik-username')`
   dans le handler `/api/process` ;
2. **n'afficher que les périmètres autorisés** — filtrer d'après
   `req.get('x-authentik-groups')`.

Ni l'une ni l'autre ne demande de changement d'infrastructure.

## 6. Vérifications

```bash
# 🐧 VM APP (192.168.5.200) — l'outpost répond (401 attendu SANS cookie)
curl -sk -o /dev/null -w "%{http_code}\n" \
  -H "Host: facturation.intra.laruche-logistique-france.fr" \
  https://192.168.5.55:9443/outpost.goauthentik.io/auth/nginx
# -> 401 = provider trouvé, outpost OK (c'est le bon résultat)
# -> 404 = provider non rattaché à l'outpost (§3) ou External host erroné (§1)
# -> connexion refusée = port 9443 fermé depuis la VM : essayer le port 9000
#    (HTTP) et ajuster le proxy_pass de deploy/nginx-facturation.conf
```

Puis, en **navigation privée**, ouvrir
`https://facturation.intra.laruche-logistique-france.fr` :
redirection vers Authentik → identification → retour sur l'app.

## 7. Pièges connus

1. **`500` au lieu de la page de login** → `auth_request` a reçu autre chose
   qu'un 2xx/401/403 : provider non rattaché à l'outpost (§3), ou `Host` qui ne
   correspond à aucun *External host* (§1).
2. **Boucle de redirection infinie** → le `location /outpost.goauthentik.io` est
   lui aussi passé par `auth_request` : il doit rester **hors** du bloc protégé
   (c'est le cas dans le `.conf` fourni, ne pas le déplacer sous `location /`).
3. **Requête vers Authentik via `auth.intra…`** → NPM route d'après le `Host`,
   et le `Host` envoyé ici est `facturation.intra…` : la requête n'arriverait
   jamais. Taper **le serveur Authentik en direct** (`192.168.5.55`), comme le
   fait le `.conf`.
4. **Déconnexion sans effet** → il faut clore la session Authentik du
   navigateur (`/outpost.goauthentik.io/sign_out`), sinon la session SSO
   reconnecte immédiatement (même piège que le socle commun §4).
5. **Génération longue coupée à 60 s** → ce n'est pas Authentik : c'est le
   timeout de NPM (onglet Advanced) puis celui de nginx. Voir `deploy/DEPLOIEMENT.md` §4.
6. **Tout le monde entre** → aucun *Group binding* posé sur l'Application (§4).
