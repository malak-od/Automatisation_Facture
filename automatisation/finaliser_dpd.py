#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
FINALISEUR DPD — produit un fichier d'import ERP depuis les CSV DPD.
Usage :
  python finaliser_dpd.py --csv <fichier1.csv> [<fichier2.csv> ...] --output <sortie.csv>

Le script lit les fichiers DPD CSV en encodage latin-1, filtre les lignes de détail
(Type (Slave Export) numérique), normalise les numéros de colis, et mappe les
surcharges vers les postes de l'import ERP.
"""

import argparse
import csv
import datetime
import os
import re
import unicodedata

IMPORT_COLUMNS = [
    'Transporteur', 'DateValidite', 'Ref1', 'Ref2', 'IdClient', 'Tracking',
    'Nom', 'EP', 'Pays', 'Zone', 'NbrColis', 'Poids', 'Mode', 'TVA',
    'DroitsTaxes', 'Assurance', 'ZonesEloignees', 'ColisVolumineux',
    'Adresses', 'Fret', 'PlusValueB2C', 'TaxeGasoil', 'NbColis',
]

# Champs DPD d'intérêt, normalisés par remove_accents
DPD_FIELD_PATTERNS = {
    'type_slave_export': ['type (slave export)', 'type slave export'],
    'date_facture': ['date de facture', 'date facture'],
    'valable_du': ['valable du'],
    'no_de_compte': ['no de compte', 'numero de compte'],
    'nom_destinataire': ['nom destinataire', 'nom destinataire'],
    'code_pays_destinataire': ['code pays destinataire', 'code pays destinataire'],
    'n_colis': ['n° colis', 'n col', 'numero de colis', 'n colis'],
    'dpd_id': ['dpd id'],
    'poids': ['poids'],
    'nombre_de_colis': ['nombre de colis', 'nbre colis', 'nbr colis'],
    'prix_transport': ['prix transport'],
    'indexation_gasoil': ['indexation gasoil', 'indexation gazole'],
    'indexation_kerosene': ['indexation kerosene', 'indexation kero', 'indexation kerosene'],
    'supplement_dpd_secure': ['supplment dpd secure', 'supplment dpd scure', 'supplement dpd secure'],
    'supplement_predict': ['supplment predict', 'supplement predict'],
    'supplement_consign': ['supplment dpd consigne', 'supplement dpd consigne'],
    'supplement_ile_montagne': ['supplment le et montagne', 'supplement ile et montagne'],
    'cout_vd': ['cot de la vd', 'coût de la vd', 'cout de la vd'],
    'taxe_collection_request': ['taxe collection request'],
    'taxe_fixe': ['taxe fixe'],
    'taxe_consolidation': ['taxe consolidation'],
    'participation_surete': ['participation surete', 'participation suret'],
    'contribution_logistique': ['contribution logistique responsable'],
    'frais_tenue_compte': ['frais de tenue de compte'],
    'taxe_triangular': ['taxe triangulaire', 'taxe triangulaire'],
    'colis_refacture': ['colis refactur', 'colis refacture'],
}

# Si un champ de surcharge n'est pas mappé explicitement, on le laisse dans Fret.
MAPPED_FIELDS = {
    'prix_transport', 'indexation_gasoil', 'indexation_kerosene',
    'supplement_dpd_secure', 'supplement_predict', 'supplement_consign',
    'supplement_ile_montagne', 'cout_vd', 'taxe_collection_request',
    'taxe_fixe', 'taxe_consolidation', 'participation_surete',
    'contribution_logistique', 'frais_tenue_compte',
    'taxe_triangular', 'colis_refacture',
}


def normalize_text(value):
    if value is None:
        return ''
    if not isinstance(value, str):
        value = str(value)
    value = unicodedata.normalize('NFKD', value)
    value = ''.join(ch for ch in value if not unicodedata.combining(ch))
    return value.strip().lower()


def parse_amount(value):
    if value is None:
        return 0.0
    text = normalize_text(value).replace(' ', '').replace('\u00a0', '').replace(',', '.')
    text = re.sub(r'[^0-9.+-]', '', text)
    if not text:
        return 0.0
    try:
        return float(text)
    except ValueError:
        return 0.0


def parse_int(value):
    if value is None:
        return 0
    text = normalize_text(value).replace(' ', '').replace('\u00a0', '')
    if not text:
        return 0
    try:
        return int(float(text))
    except ValueError:
        return 0


def build_header_map(header):
    normalized = [normalize_text(h) for h in header]
    return {name: idx for idx, name in enumerate(normalized)}


def locate_column(header_map, patterns):
    for pattern in patterns:
        norm = normalize_text(pattern)
        for name, idx in header_map.items():
            if norm == name or norm in name:
                return idx
    return None


def find_dpd_columns(header):
    header_map = build_header_map(header)
    cols = {}
    for field, patterns in DPD_FIELD_PATTERNS.items():
        cols[field] = locate_column(header_map, patterns)
    return cols


def first_day_of_month(date_str):
    if not date_str:
        return ''
    text = normalize_text(date_str).replace('.', '/').replace('-', '/').replace(' ', '')
    for fmt in ('%d/%m/%Y', '%Y/%m/%d', '%d/%m/%y'):
        try:
            dt = datetime.datetime.strptime(text, fmt)
            return dt.replace(day=1).strftime('%d/%m/%Y')
        except ValueError:
            pass
    return ''


def read_csv_rows(csv_paths):
    rows = []
    for csv_path in csv_paths:
        if os.path.isdir(csv_path):
            continue
        with open(csv_path, encoding='latin-1', newline='') as f:
            reader = csv.reader(f, delimiter=';')
            try:
                header = next(reader)
            except StopIteration:
                continue
            columns = find_dpd_columns(header)
            for row in reader:
                if len(row) < len(header):
                    row += [''] * (len(header) - len(row))
                rows.append((row, columns))
    return rows


def is_detail_row(row_tuple):
    row, columns = row_tuple
    idx = columns.get('type_slave_export')
    if idx is None:
        return False
    value = row[idx] if idx < len(row) else ''
    return str(value).strip().isdigit()


def row_value(row, idx):
    return row[idx] if idx is not None and idx < len(row) else ''


def build_import_record(row_tuple):
    row, columns = row_tuple
    tracking = row_value(row, columns.get('n_colis')) or row_value(row, columns.get('dpd_id'))
    tracking = re.sub(r'[^A-Za-z0-9]', '', str(tracking))
    date_facture = row_value(row, columns.get('date_facture')) or row_value(row, columns.get('valable_du'))
    date_validite = first_day_of_month(date_facture)
    pays = row_value(row, columns.get('code_pays_destinataire'))
    client = row_value(row, columns.get('no_de_compte'))
    nom = row_value(row, columns.get('nom_destinataire'))
    colis_refacture = parse_amount(row_value(row, columns.get('colis_refacture')))
    prix_transport = parse_amount(row_value(row, columns.get('prix_transport')))
    # Taxe gasoil DPD = uniquement le gasoil routier (le kerosene, surcharge aerienne,
    # n'est pas un "taxe gasoil" au sens du contrat DPD -> reste dans Fret).
    gazole = parse_amount(row_value(row, columns.get('indexation_gasoil')))
    kerosene = parse_amount(row_value(row, columns.get('indexation_kerosene')))
    assurance = parse_amount(row_value(row, columns.get('supplement_dpd_secure')))
    adresses = parse_amount(row_value(row, columns.get('supplement_predict'))) + parse_amount(row_value(row, columns.get('supplement_consign')))
    zones_eloignees = parse_amount(row_value(row, columns.get('supplement_ile_montagne')))
    plus_value = parse_amount(row_value(row, columns.get('frais_tenue_compte')))
    droits_taxes = (
        parse_amount(row_value(row, columns.get('cout_vd'))) +
        parse_amount(row_value(row, columns.get('taxe_collection_request'))) +
        parse_amount(row_value(row, columns.get('taxe_fixe'))) +
        parse_amount(row_value(row, columns.get('taxe_consolidation'))) +
        parse_amount(row_value(row, columns.get('participation_surete'))) +
        parse_amount(row_value(row, columns.get('contribution_logistique'))) +
        parse_amount(row_value(row, columns.get('taxe_triangular')))
    )
    # Fret = prix transport + kerosene (surcharge aerienne, hors "taxe gasoil" DPD)
    fret = prix_transport + kerosene
    # Regle facturation Excel (p.8) : si le fret ne contient que les frais fixes
    # (frais de dossier 0,10 seul, ou 0,10 + taxe eco 0,86 + taxe surete 0,27 = 1,23),
    # le deplacer en Adresses plutot que de le laisser en Fret.
    if abs(fret - 0.10) < 0.005 or abs(fret - 1.23) < 0.005:
        adresses += fret
        fret = 0.0

    return {
        'Transporteur': 'DPD',
        'DateValidite': date_validite,
        'Ref1': row_value(row, columns.get('dpd_id')),
        'Ref2': '',
        'IdClient': client,
        'Tracking': tracking,
        'Nom': nom,
        'EP': 'P',
        'Pays': pays,
        'Zone': '',
        'NbrColis': str(parse_int(row_value(row, columns.get('nombre_de_colis')))),
        'Poids': str(parse_amount(row_value(row, columns.get('poids')))),
        'Mode': 'ST',
        'TVA': '0,2',
        'DroitsTaxes': f'{droits_taxes:.2f}',
        'Assurance': f'{assurance:.2f}',
        'ZonesEloignees': f'{zones_eloignees:.2f}',
        'ColisVolumineux': f'{colis_refacture:.2f}',
        'Adresses': f'{adresses:.2f}',
        'Fret': f'{fret:.2f}',
        'PlusValueB2C': f'{plus_value:.2f}',
        'TaxeGasoil': f'{gazole:.2f}',
        'NbColis': str(parse_int(row_value(row, columns.get('nombre_de_colis')))),
    }


def write_import_csv(output_path, records):
    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    with open(output_path, 'w', encoding='utf-8', newline='') as f:
        writer = csv.DictWriter(f, fieldnames=IMPORT_COLUMNS, delimiter=';', extrasaction='ignore')
        writer.writeheader()
        for record in records:
            writer.writerow(record)


def parse_args():
    parser = argparse.ArgumentParser(description='Génère un import ERP DPD à partir des CSV DPD.')
    parser.add_argument('--csv', nargs='+', required=True, help='Fichiers CSV DPD à traiter')
    parser.add_argument('--output', required=True, help='Fichier CSV de sortie pour l import ERP')
    return parser.parse_args()


def main():
    args = parse_args()
    rows = read_csv_rows(args.csv)
    detail_rows = [row for row in rows if is_detail_row(row)]
    records = [build_import_record(row) for row in detail_rows]
    write_import_csv(args.output, records)
    print(f'Généré {len(records)} lignes dans {args.output}')


if __name__ == '__main__':
    main()
