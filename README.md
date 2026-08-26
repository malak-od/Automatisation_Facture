# Facturation Transporteurs — déploiement

Branche `prod` : contient uniquement le code applicatif et les 12 modèles Excel
nécessaires pour faire tourner l'application. Pas de documentation métier, pas
de données réelles, pas de vidéos process — tout ça reste sur la branche `main`.

## Structure

```
facturation-app/       Serveur web (Node.js) + interface
Automatisation/         Scripts Python (un finaliseur par transporteur, pilotage Excel via COM)
Transporteurs/<Nom>/    1 classeur Excel modèle par transporteur (2026_06_Facture <Nom>.xlsx)
```

Le code utilise des **chemins relatifs** entre ces 3 dossiers (`../automatisation/...`,
`../Transporteurs/...` depuis `facturation-app/`) : ils doivent rester frères,
à la même profondeur, dans cette disposition exacte.

> Note : le dossier `Automatisation/` est suivi par git en minuscule
> (`automatisation/`) alors qu'il apparaît en majuscule sur Windows (système de
> fichiers insensible à la casse). Aucun impact sur Windows, mais à surveiller
> en cas de déploiement futur sur un système sensible à la casse.

## Prérequis machine

- **Windows** — pas de déploiement Linux/Docker possible en l'état.
- **Microsoft Excel installé** sur la machine serveur, avec une licence active.
  Les finaliseurs pilotent Excel via COM (`pywin32`) : ce n'est pas une
  bibliothèque de génération de fichier autonome, Excel doit réellement
  tourner sur la machine.
- **Node.js** (testé en v24 ; une v20+ récente devrait convenir).
- **Python**, accessible via la commande `python` dans le PATH système
  (pas `python3`). Testé en 3.14.

## Installation

```
pip install -r Automatisation/requirements.txt
cd facturation-app
npm ci
npm start
```

Ouvrir `http://localhost:4000` (port configurable via la variable
d'environnement `PORT`).

## Points d'attention

- **Pas de rechargement à chaud** : toute mise à jour de code nécessite
  d'arrêter puis relancer `npm start`.
- **Excel doit être libre pendant les générations.** Si un classeur de sortie
  est déjà ouvert (ou qu'un process `EXCEL.EXE` orphelin traîne en tâche de
  fond suite à un plantage), la génération échoue avec un message explicite
  — vérifier le Gestionnaire des tâches, pas seulement les fenêtres visibles.
- Un **service Windows** (NSSM, ou tâche planifiée au démarrage) est
  recommandé pour que l'app survive à une fermeture de session, plutôt que
  de dépendre d'un terminal laissé ouvert.
- Aucune gestion HTTPS/authentification intégrée. Si l'app doit être
  accessible au-delà du poste local, prévoir un reverse proxy (IIS/nginx)
  avec accès restreint au réseau interne.

## Mettre à jour cette branche

`prod` diverge volontairement de `main` (moins de fichiers) : une fusion
classique (`git merge main`) réimporterait tout ce qui a été retiré. Pour
répercuter une évolution de `main` ici, il faut la reporter manuellement
(cherry-pick des fichiers concernés) plutôt qu'un merge direct.
