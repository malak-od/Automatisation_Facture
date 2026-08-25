# Automatisation Facturation

## Contexte du projet

Ce repository contient un système d'automatisation de la facturation pour plusieurs transporteurs.

Les transporteurs sont organisés dans `Transporteurs/`.

Le projet contient notamment :
- des scripts d'automatisation ;
- des traitements spécifiques aux transporteurs ;
- une application de facturation ;
- des notebooks ;
- de la documentation métier et technique.

## Règle principale

Ne jamais analyser l'ensemble du repository par défaut.

Pour chaque tâche, déterminer d'abord le périmètre nécessaire puis travailler uniquement sur ce périmètre.

## Exploration

Avant de lire des fichiers :

1. Identifier le transporteur ou le module concerné.
2. Identifier les fichiers potentiellement concernés.
3. Lire uniquement les fichiers nécessaires.
4. Ne pas explorer les autres transporteurs sans raison.

Si la tâche concerne `Transporteurs/DHL/`, ne pas analyser automatiquement `Chronopost`, `UPS`, `DPD`, etc.

## Modifications

- Modifier uniquement ce qui est nécessaire.
- Ne pas refactoriser du code hors périmètre.
- Ne pas modifier plusieurs transporteurs pour résoudre un problème local.
- Préserver le comportement existant.
- Ne pas créer de nouvelles abstractions sans nécessité.

## Subagents

Les subagents doivent être utilisés avec parcimonie.

Avant de créer un subagent, déterminer si la tâche peut être réalisée directement.

Pour une tâche simple :
→ aucun subagent.

Si un subagent est nécessaire :
- utiliser le minimum possible ;
- lui donner une mission unique et précise ;
- limiter son périmètre de fichiers ;
- ne pas lui demander d'analyser tout le repository ;
- ne pas créer plusieurs subagents pour obtenir plusieurs analyses similaires ;
- lui demander un résultat court et directement exploitable.

Ne jamais lancer automatiquement un subagent pour explorer le projet.

## Transporteurs

Chaque transporteur doit être considéré comme un périmètre indépendant sauf lorsqu'une dépendance commune est explicitement identifiée.

Lorsqu'une modification concerne un transporteur spécifique, ne pas modifier les autres transporteurs sans justification.

Avant de modifier une règle métier, consulter la documentation correspondante lorsqu'elle existe.

## Tests

Après une modification :

1. lancer les tests directement concernés ;
2. vérifier le comportement modifié ;
3. ne lancer la suite complète que si nécessaire.

Ne pas lancer automatiquement tous les tests après chaque modification.

## Documentation

Les informations importantes concernant le projet doivent être conservées dans `Documentation/`.

Ne pas utiliser l'historique du chat comme unique source de connaissance du projet.

Si une décision importante est prise, la documenter dans `Documentation/DECISIONS.md`.

## Contexte

Si la conversation devient très longue, utiliser `/compact`.

Avant un compact, conserver dans la documentation :
- l'état actuel ;
- les fichiers modifiés ;
- les décisions importantes ;
- les problèmes restants.

Si la tâche change complètement, utiliser `/clear`.

## Fin de tâche

Lorsque la demande est terminée :

- indiquer brièvement les fichiers modifiés ;
- indiquer les tests effectués ;
- indiquer les éventuels problèmes restants.

Ne pas continuer automatiquement avec d'autres améliorations.

Ne pas effectuer de tâche non demandée.