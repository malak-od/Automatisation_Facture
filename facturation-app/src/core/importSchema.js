// Schéma du fichier d'import ERP : 23 colonnes, ordre et libellés exacts.
// `key` = nom interne (dans les objets ligne) ; `label` = en-tête ERP ; `num` = colonne numérique.
const IMPORT_COLUMNS = [
  { key: 'Transporteur',    label: ' Transporteur ',      num: false },
  { key: 'DateValidite',    label: 'Date validité tarif',  num: false },
  { key: 'Ref1',            label: 'Réf.1',                num: false },
  { key: 'Ref2',            label: 'Réf. 2',               num: false },
  { key: 'IdClient',        label: 'Id client',            num: false },
  { key: 'Tracking',        label: 'N° Tracking',          num: false },
  { key: 'Nom',             label: 'Nom',                  num: false },
  { key: 'EP',              label: 'E / P',                num: false },
  { key: 'Pays',            label: 'Pays',                 num: false },
  { key: 'Zone',            label: ' Zone ',               num: false },
  { key: 'NbrColis',        label: ' Nbr Colis ',          num: true  },
  { key: 'Poids',           label: ' Poids ',              num: true  },
  { key: 'Mode',            label: 'mode envoi',           num: false },
  { key: 'TVA',             label: ' TVA ',                num: true  },
  { key: 'DroitsTaxes',     label: ' Droits et taxes ',    num: true  },
  { key: 'Assurance',       label: ' Assurance ',          num: true  },
  { key: 'ZonesEloignees',  label: ' Zones éloignées ',    num: true  },
  { key: 'ColisVolumineux', label: ' Colis volumineux ',   num: true  },
  { key: 'Adresses',        label: ' Adresses ',           num: true  },
  { key: 'Fret',            label: '    Frêt    ',         num: true  },
  { key: 'PlusValueB2C',    label: ' plus-value BtoC ',    num: true  },
  { key: 'TaxeGasoil',      label: 'Gazole',               num: false },
  { key: 'NbColis',         label: 'Nb Colis',             num: false },
];

module.exports = { IMPORT_COLUMNS };
