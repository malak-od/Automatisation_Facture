# Logos transporteurs

Pour afficher le vrai logo d'un transporteur sur la page d'accueil, dépose ici
un fichier `<id>.png` (fond transparent ou blanc conseillé, ~200x200px).

`<id>` = l'identifiant du transporteur dans `src/registry.js` :

kuehne, delivengo, dpd, ups, geodis, gls, colissimo, chronopost, fedex,
tnt, mondial_relay, lettres, bls

Aucune modification de code n'est nécessaire : si `logos/<id>.png` existe, il
remplace automatiquement le badge coloré généré par défaut. S'il est absent,
le badge (couleur de marque approximative + initiales) reste affiché.
