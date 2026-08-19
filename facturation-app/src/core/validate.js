// Validation automatique = la check-list "Pour tous les transporteurs" (FACTURATION EXCEL.docx).
// Commune a TOUS les transporteurs. Ne bloque pas.
//  - alerts : lignes a VERIFIER (vraies anomalies)
//  - infos  : lignes NORMALES a connaitre (affretement / navette : zone volontairement vide,
//             pour que l'ERP n'aille PAS chercher la grille tarifaire).
function validate(importRows, opts = {}) {
  const alerts = [];
  const infos = [];
  importRows.forEach((o, idx) => {
    const l = idx + 2; // ligne 1 = entete
    const t = o.Tracking || '(vide)';

    // -- Champs obligatoires sur toutes les lignes --
    if (!o.Transporteur) alerts.push(`L${l} ${t}: TRANSPORTEUR manquant`);
    // E/P manquant : pas d'alerte affichee (bruit trop frequent, notamment Geodis),
    // mais la ligne reste traitee normalement -- rien d'autre a changer, validate()
    // ne modifie jamais importRows.
    if (!o.DateValidite) alerts.push(`L${l} ${t}: DATE manquante`);
    if (!o.Tracking) alerts.push(`L${l}: N° TRACKING manquant`);
    if (!o.Pays) alerts.push(`L${l} ${t}: PAYS manquant`);

    // -- Zone --
    const zoneHorsGrille = o.Zone.includes('0France00') || o.Zone.includes('0étranger00');
    // Zone deja signalee comme non resolue PAR LE CARRIER LUI-MEME (son propre
    // warnings.push() a la construction de la ligne, ex. Colissimo/Mondial Relay
    // "zone inconnue", DPD "inconnu") -- pas la peine de dupliquer l'alerte ici,
    // route en info comme le pattern hors-grille Kuehne.
    const zoneNonResolue = /^zone inconnue$/i.test(o.Zone) || /^inconnu$/i.test(o.Zone);
    if (!o.Zone) {
      alerts.push(`L${l} ${t}: ZONE manquante`);
    } else if (zoneHorsGrille) {
      if (o._horsGrille) {
        // Affretement ou navette : zone vide VOULUE (pas de grille tarifaire) -> normal
        const type = o._metier === 'AFFR' ? 'affrètement' : 'navette';
        infos.push(`L${l} ${t}: ${type} — zone volontairement vide (${o.Zone}), hors grille`);
      } else {
        alerts.push(`L${l} ${t}: ZONE inconnue (${o.Zone}) - dept non trouve`);
      }
    } else if (zoneNonResolue) {
      infos.push(`L${l} ${t}: zone/mode non résolu(e) (${o.Zone}) — déjà signalé par le transporteur, à compléter dans la table de correspondance`);
    } else if (o.Zone.includes('étranger')) {
      alerts.push(`L${l} ${t}: ZONE etranger (${o.Zone})`);
    }

    // -- Quantites --
    if (!o.NbrColis) alerts.push(`L${l} ${t}: COLIS = 0`);
    if (!o.Poids) alerts.push(`L${l} ${t}: POIDS = 0`);
    else if (!opts.skipPoidsDecimal && Math.round(o.Poids * 10) !== o.Poids * 10) alerts.push(`L${l} ${t}: POIDS ${o.Poids} a plus d'1 decimale`);

    // -- Montants (fret != 0, sauf si un autre poste porte le montant) --
    const autres = o.DroitsTaxes || o.Assurance || o.ZonesEloignees || o.ColisVolumineux || o.Adresses || o.PlusValueB2C;
    if (!o.Fret && !autres) alerts.push(`L${l} ${t}: FRET = 0 (ligne sans montant)`);
  });
  return { alerts, infos };
}

module.exports = { validate };
