#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
FINALISEUR COLISSIMO — produit "AAAA_MM_Facture Colissimo.xlsx" A L'IDENTIQUE
du fichier fait a la main (memes feuilles, memes formules, memes TCD/pivots,
meme mise en forme), en partant du fichier final comme MODELE et en n'y
remplacant que les donnees.

Classeur reel (9 feuilles) :
  Table de correspondance, Modes envois, Pays, Poids : tables de reference
    figees (Poids est lui-meme un TCD source sur Facture Colissimo, recalcule
    par RefreshAll()).
  Facture Colissimo : 1 ligne = 1 CHARGE brute des 2 CSV (plusieurs lignes
    par colis : Transport, Remise, CAE=taxe gazole, Supplements...). Colonnes
    A-G = FORMULES (Pays/Mode envoi/Zone/Concatener/Prefixe/Pays/Categorie) ;
    colonnes H+ = donnees brutes des CSV. Les 2 CSV ("Prestations au colis"
    26 col. et "Frais de douane" 20 col.) sont colles A LA SUITE dans le meme
    bloc de colonnes H+ (leurs 7 premieres colonnes N de ligne->Produit sont
    communes ; P-Z restent vides pour les lignes douane, confirme sans
    decalage dangereux -- cf. Documentation/
    Video_Colissimo_1_Preparation_fichier_import.md, Partie 5).
  TCD : tableau croise dynamique (source = Facture Colissimo!A:AD), 1 ligne
    par N colis (tracking), colonnes E-L = les 8 postes ERP (Adresse,
    Assurance, Colis volumineux, Droits et taxes, Fret, plus-value BtoC,
    Taxe gazole, Zones eloignees). Colonne A "ID Client" : PAS un champ du
    pivot -- une colonne juxtaposee, valeurs saisies a la main (jointure
    externe, ex. '6739'/'7027'), PURGEE ici a chaque generation (cf. fonction
    main(), etape 3) pour que la personne qui la renseigne reparte d'une
    colonne vide alignee sur le bon mois -- meme regle que "supprimer les Id
    clients" (Documentation/FACTURATION EXCEL.docx, "Pour tous les
    transporteurs") et que Chronopost/BLS.
  Bilan Factures / Bilan Client : TCD sources sur Facture Colissimo / TCD.
    "Bilan Factures" = 1 ligne par N facture, col B = Somme de Total HT ;
    reconciliation PDF (video process) : on colle le "TOTAL HT" du PDF en
    colonne D (indemnisation/avoir colonnes E/F restent a la main, non
    extractibles automatiquement d'un PDF simple).
  Import CSV : 1 ligne par colis (=TCD!D{n+1}), formules XLOOKUP vers
    Facture Colissimo/Poids/TCD -- son nombre de lignes suit le TCD (colis
    UNIQUES), pas le nombre de lignes CSV brutes -> etendu APRES RefreshAll().

Necessite : Windows + Excel + pywin32 + pypdf/pdfplumber (reconciliation PDF,
optionnelle).
Usage : python finaliser_colissimo.py "<modele.xlsx>" "<sortie.xlsx>" \
        --presta <csv1> [<csv2> ...] [--douane <csv1> ...] [--pdf <pdf1> ...]
"""
import sys, os, shutil, csv, re


# Table statique NOM PAYS (francais, style administratif -- majuscules, sans accents,
# meme convention que les 66 entrees deja validees de config.json/feuille "Pays") -> code
# ISO 3166-1 alpha-2. Couvre les 249 pays/territoires ISO (verifie contre pycountry
# 2026-08-17). Sert de FILET DE SECURITE uniquement : la feuille "Pays" du classeur (source
# de verite, confirmee par le pole transport) est TOUJOURS consultee en premier -- cette
# table n'intervient que si un nom y est absent, pour eviter un ajout manuel a chaque
# nouveau pays desservi par Colissimo (demande utilisateur 2026-08-17, suite a la decouverte
# de "ILE DE MAN"/"SINGAPOUR" manquants sur juillet 2026).
ISO_PAYS_FR = {
    "AD": "ANDORRE", "AE": "EMIRATS ARABES UNIS", "AF": "AFGHANISTAN", "AG": "ANTIGUA-ET-BARBUDA",
    "AI": "ANGUILLA", "AL": "ALBANIE", "AM": "ARMENIE", "AO": "ANGOLA", "AQ": "ANTARCTIQUE",
    "AR": "ARGENTINE", "AS": "SAMOA AMERICAINES", "AT": "AUTRICHE", "AU": "AUSTRALIE", "AW": "ARUBA",
    "AX": "ILES ALAND", "AZ": "AZERBAIDJAN", "BA": "BOSNIE-HERZEGOVINE", "BB": "BARBADE",
    "BD": "BANGLADESH", "BE": "BELGIQUE", "BF": "BURKINA FASO", "BG": "BULGARIE", "BH": "BAHREIN",
    "BI": "BURUNDI", "BJ": "BENIN", "BL": "SAINT-BARTHELEMY", "BM": "BERMUDES", "BN": "BRUNEI",
    "BO": "BOLIVIE", "BQ": "BONAIRE SAINT-EUSTACHE ET SABA", "BR": "BRESIL", "BS": "BAHAMAS",
    "BT": "BHOUTAN", "BV": "ILE BOUVET", "BW": "BOTSWANA", "BY": "BIELORUSSIE", "BZ": "BELIZE",
    "CA": "CANADA", "CC": "ILES COCOS", "CD": "CONGO (RD)", "CF": "REPUBLIQUE CENTRAFRICAINE",
    "CG": "CONGO", "CH": "SUISSE", "CI": "COTE D'IVOIRE", "CK": "ILES COOK", "CL": "CHILI",
    "CM": "CAMEROUN", "CN": "CHINE", "CO": "COLOMBIE", "CR": "COSTA RICA", "CU": "CUBA",
    "CV": "CAP-VERT", "CW": "CURACAO", "CX": "ILE CHRISTMAS", "CY": "CHYPRE",
    "CZ": "REPUBLIQUE TCHEQUE", "DE": "ALLEMAGNE", "DJ": "DJIBOUTI", "DK": "DANEMARK",
    "DM": "DOMINIQUE", "DO": "REPUBLIQUE DOMINICAINE", "DZ": "ALGERIE", "EC": "EQUATEUR",
    "EE": "ESTONIE", "EG": "EGYPTE", "EH": "SAHARA OCCIDENTAL", "ER": "ERYTHREE", "ES": "ESPAGNE",
    "ET": "ETHIOPIE", "FI": "FINLANDE", "FJ": "FIDJI", "FK": "ILES MALOUINES", "FM": "MICRONESIE",
    "FO": "ILES FEROE", "FR": "FRANCE", "GA": "GABON", "GB": "ROYAUME-UNI", "GD": "GRENADE",
    "GE": "GEORGIE", "GF": "GUYANE FRANCAISE", "GG": "GUERNESEY", "GH": "GHANA", "GI": "GIBRALTAR",
    "GL": "GROENLAND", "GM": "GAMBIE", "GN": "GUINEE", "GP": "GUADELOUPE",
    "GQ": "GUINEE EQUATORIALE", "GR": "GRECE", "GS": "GEORGIE DU SUD ET SANDWICH DU SUD",
    "GT": "GUATEMALA", "GU": "GUAM", "GW": "GUINEE-BISSAU", "GY": "GUYANA", "HK": "HONG KONG",
    "HM": "ILES HEARD-ET-MACDONALD", "HN": "HONDURAS", "HR": "CROATIE", "HT": "HAITI",
    "HU": "HONGRIE", "ID": "INDONESIE", "IE": "IRLANDE", "IL": "ISRAEL", "IM": "ILE DE MAN",
    "IN": "INDE", "IO": "TERRITOIRE BRITANNIQUE DE L'OCEAN INDIEN", "IQ": "IRAK", "IR": "IRAN",
    "IS": "ISLANDE", "IT": "ITALIE", "JE": "JERSEY", "JM": "JAMAIQUE", "JO": "JORDANIE",
    "JP": "JAPON", "KE": "KENYA", "KG": "KIRGHIZISTAN", "KH": "CAMBODGE", "KI": "KIRIBATI",
    "KM": "COMORES", "KN": "SAINT-CHRISTOPHE-ET-NIEVES", "KP": "COREE DU NORD",
    "KR": "COREE DU SUD", "KW": "KOWEIT", "KY": "ILES CAIMANS", "KZ": "KAZAKHSTAN", "LA": "LAOS",
    "LB": "LIBAN", "LC": "SAINTE-LUCIE", "LI": "LIECHTENSTEIN", "LK": "SRI LANKA",
    "LR": "LIBERIA", "LS": "LESOTHO", "LT": "LITUANIE", "LU": "LUXEMBOURG", "LV": "LETTONIE",
    "LY": "LIBYE", "MA": "MAROC", "MC": "MONACO", "MD": "MOLDAVIE", "ME": "MONTENEGRO",
    "MF": "SAINT-MARTIN", "MG": "MADAGASCAR", "MH": "ILES MARSHALL", "MK": "MACEDOINE DU NORD",
    "ML": "MALI", "MM": "MYANMAR", "MN": "MONGOLIE", "MO": "MACAO",
    "MP": "ILES MARIANNES DU NORD", "MQ": "MARTINIQUE", "MR": "MAURITANIE", "MS": "MONTSERRAT",
    "MT": "MALTE", "MU": "MAURICE", "MV": "MALDIVES", "MW": "MALAWI", "MX": "MEXIQUE",
    "MY": "MALAISIE", "MZ": "MOZAMBIQUE", "NA": "NAMIBIE", "NC": "NOUVELLE-CALEDONIE",
    "NE": "NIGER", "NF": "ILE NORFOLK", "NG": "NIGERIA", "NI": "NICARAGUA", "NL": "PAYS-BAS",
    "NO": "NORVEGE", "NP": "NEPAL", "NR": "NAURU", "NU": "NIUE", "NZ": "NOUVELLE-ZELANDE",
    "OM": "OMAN", "PA": "PANAMA", "PE": "PEROU", "PF": "POLYNESIE FRANCAISE",
    "PG": "PAPOUASIE-NOUVELLE-GUINEE", "PH": "PHILIPPINES", "PK": "PAKISTAN", "PL": "POLOGNE",
    "PM": "SAINT-PIERRE-ET-MIQUELON", "PN": "ILES PITCAIRN", "PR": "PORTO RICO",
    "PS": "PALESTINE", "PT": "PORTUGAL", "PW": "PALAOS", "PY": "PARAGUAY", "QA": "QATAR",
    "RE": "REUNION", "RO": "ROUMANIE", "RS": "SERBIE", "RU": "RUSSIE", "RW": "RWANDA",
    "SA": "ARABIE SAOUDITE", "SB": "ILES SALOMON", "SC": "SEYCHELLES", "SD": "SOUDAN",
    "SE": "SUEDE", "SG": "SINGAPOUR", "SH": "SAINTE-HELENE", "SI": "SLOVENIE",
    "SJ": "SVALBARD ET JAN MAYEN", "SK": "SLOVAQUIE", "SL": "SIERRA LEONE", "SM": "SAINT-MARIN",
    "SN": "SENEGAL", "SO": "SOMALIE", "SR": "SURINAME", "SS": "SOUDAN DU SUD",
    "ST": "SAO TOME-ET-PRINCIPE", "SV": "EL SALVADOR",
    "SX": "SAINT-MARTIN (PARTIE NEERLANDAISE)", "SY": "SYRIE", "SZ": "ESWATINI",
    "TC": "ILES TURQUES-ET-CAIQUES", "TD": "TCHAD", "TF": "TERRES AUSTRALES FRANCAISES",
    "TG": "TOGO", "TH": "THAILANDE", "TJ": "TADJIKISTAN", "TK": "TOKELAU",
    "TL": "TIMOR ORIENTAL", "TM": "TURKMENISTAN", "TN": "TUNISIE", "TO": "TONGA", "TR": "TURQUIE",
    "TT": "TRINITE-ET-TOBAGO", "TV": "TUVALU", "TW": "TAIWAN", "TZ": "TANZANIE", "UA": "UKRAINE",
    "UG": "OUGANDA", "UM": "ILES MINEURES ELOIGNEES DES ETATS-UNIS", "US": "ETATS-UNIS",
    "UY": "URUGUAY", "UZ": "OUZBEKISTAN", "VA": "VATICAN",
    "VC": "SAINT-VINCENT-ET-LES-GRENADINES", "VE": "VENEZUELA", "VG": "ILES VIERGES BRITANNIQUES",
    "VI": "ILES VIERGES DES ETATS-UNIS", "VN": "VIETNAM", "VU": "VANUATU",
    "WF": "WALLIS-ET-FUTUNA", "WS": "SAMOA", "YE": "YEMEN", "YT": "MAYOTTE",
    "ZA": "AFRIQUE DU SUD", "ZM": "ZAMBIE", "ZW": "ZIMBABWE",
}


def normalize_pays(s):
    """Normalise un nom de pays pour comparaison : majuscules, accents retires, espaces
    autour du tiret/apostrophe uniformises. Tolere les variantes mineures (accents mal
    encodes type '�', tirets simples/doubles) sans exiger une correspondance caractere
    pres -- objectif : ne pas re-ajouter 'ROYAUME-UNI' juste parce que le CSV a
    'royaume uni' ou 'ROYAUME  UNI'."""
    if not s:
        return ""
    s = s.strip().upper()
    for a, b in (("É", "E"), ("È", "E"), ("Ê", "E"), ("À", "A"), ("Ô", "O"), ("Î", "I"), ("Ï", "I"), ("Ç", "C"), ("Ù", "U"), ("Û", "U"), ("�", "E")):
        s = s.replace(a, b)
    s = re.sub(r"[\s-]+", " ", s).strip()
    return s


def resolve_missing_pays(pays_ws, pays_seen, xlUp):
    """Ajoute automatiquement a la feuille 'Pays' les noms de pays presents dans les
    donnees brutes du mois mais absents de la feuille -- SEULEMENT si le nom matche un
    pays de la table ISO_PAYS_FR (filet de securite, pas d'invention de code). Retourne
    la liste des ajouts (nom, code) pour log/warning ; les noms non reconnus meme dans
    ISO_PAYS_FR restent "Pays a creer" comme avant (aucun changement de comportement pour
    ces cas -- toujours un ajout MANUEL requis, comme documente dans la video process)."""
    existing = {}
    last_row = pays_ws.Cells(pays_ws.Rows.Count, 1).End(xlUp).Row
    for r in range(2, last_row + 1):
        name = pays_ws.Cells(r, 1).Value
        if name:
            existing[normalize_pays(str(name))] = True

    iso_by_norm = {normalize_pays(name): (name, code) for code, name in ISO_PAYS_FR.items()}

    added = []
    row = last_row
    for raw_name in sorted(pays_seen):
        norm = normalize_pays(raw_name)
        if not norm or norm in existing:
            continue
        match = iso_by_norm.get(norm)
        if not match:
            continue  # pas dans la table ISO -- reste "Pays a creer", ajout manuel requis
        canonical_name, code = match
        row += 1
        pays_ws.Cells(row, 1).Value = canonical_name
        pays_ws.Cells(row, 2).Value = code
        existing[norm] = True
        added.append((canonical_name, code))
    return added


def resolve_missing_modes_envois(modes_ws, prefix_pays_seen, xlUp):
    """Ajoute automatiquement a la feuille 'Modes envois' les combinaisons
    prefixe+pays presentes dans les donnees brutes mais absentes -- en DEVINANT la
    zone/mode par la valeur la PLUS FREQUENTE deja utilisee pour ce meme prefixe
    (demande utilisateur 2026-08-17 : "deviner par le prefixe existant le plus proche").
    Contrairement au code ISO d'un pays, la zone/mode depend de regles tarifaires
    Colissimo non deductibles -- cette valeur est donc TOUJOURS une estimation, jamais
    certaine a 100%, d'ou le warning systematique retourne pour chaque ajout (jamais
    silencieux). Retourne la liste (concat, prefixe, pays, zone, mode) ajoutee."""
    last_row = modes_ws.Cells(modes_ws.Rows.Count, 1).End(xlUp).Row
    existing_concats = set()
    zone_mode_by_prefix = {}  # prefixe -> {(zone,mode): count}
    for r in range(2, last_row + 1):
        concat = modes_ws.Cells(r, 1).Value
        prefixe = modes_ws.Cells(r, 2).Value
        zone = modes_ws.Cells(r, 4).Value
        mode = modes_ws.Cells(r, 6).Value
        if concat:
            existing_concats.add(str(concat))
        if prefixe and zone and mode:
            zone_mode_by_prefix.setdefault(str(prefixe), {})
            pair = (str(zone), str(mode))
            zone_mode_by_prefix[str(prefixe)][pair] = zone_mode_by_prefix[str(prefixe)].get(pair, 0) + 1

    added = []
    row = last_row
    for prefixe, pays_code in sorted(prefix_pays_seen):
        concat = prefixe + pays_code
        if concat in existing_concats:
            continue
        candidates = zone_mode_by_prefix.get(prefixe)
        if not candidates:
            continue  # prefixe jamais vu -- aucune base pour deviner, reste "zone inconnue"
        (zone, mode), _ = max(candidates.items(), key=lambda kv: kv[1])
        row += 1
        modes_ws.Cells(row, 1).Value = concat
        modes_ws.Cells(row, 2).Value = prefixe
        modes_ws.Cells(row, 3).Value = pays_code
        modes_ws.Cells(row, 4).Value = zone
        modes_ws.Cells(row, 5).Value = 0
        modes_ws.Cells(row, 6).Value = mode
        existing_concats.add(concat)
        added.append((concat, prefixe, pays_code, zone, mode))
    return added


def extract_pays_and_prefixes(hdr, rows):
    """Extrait, depuis les lignes brutes deja alignees (load_rows), l'ensemble des noms
    de pays destination distincts (pour resolve_missing_pays) et des couples
    (prefixe tracking, nom pays destination) distincts (pour deduire, apres resolution
    ISO, les combinaisons prefixe+code a verifier dans 'Modes envois')."""
    idx_pays_dest = hdr.index("Pays Destination") if "Pays Destination" in hdr else None
    idx_n_colis = hdr.index("N colis") if "N colis" in hdr else None
    pays_seen = set()
    prefix_pays_names = set()
    if idx_pays_dest is None or idx_n_colis is None:
        return pays_seen, prefix_pays_names
    for r in rows:
        pays_name = r[idx_pays_dest].strip() if idx_pays_dest < len(r) and r[idx_pays_dest] else ""
        tracking = r[idx_n_colis].strip() if idx_n_colis < len(r) and r[idx_n_colis] else ""
        if pays_name:
            pays_seen.add(pays_name)
        if pays_name and len(tracking) >= 2:
            prefix_pays_names.add((tracking[:2], pays_name))
    return pays_seen, prefix_pays_names


def coerce(s):
    """Nombre si la cellule est purement numerique (virgule decimale), sinon texte, sinon vide."""
    if s is None or s == "":
        return None
    if re.fullmatch(r"-?\d+(,\d+)?", s):
        return float(s.replace(",", "."))
    return s


def load_rows(presta_paths, douane_paths):
    """Lit les CSV Colissimo et les REALIGNE PAR NOM DE COLONNE sur le referentiel
    "Prestations au colis" (26 colonnes) -- PAS par position brute. BUG TROUVE
    2026-08-14 (2e test) : "Prestations au colis" a "Total HT"/"Code charge" en
    position 20/23, "Frais de douane" (20 col. seulement) les a en position 14/17
    (colonnes Droit de douane/TVA importation/Octroi de mer/Frais de gestion/Autres
    taxes en moins que Zone/Pays/Poids/dimensions cote presta) -- coller les lignes
    douane brutes telles quelles decalait leur Total HT/Code charge de 6 colonnes
    vers la gauche par rapport a ce que "Facture Colissimo" attend en position fixe,
    laissant ces lignes hors de toute categorie (0€ de "Droits et taxes" au lieu de
    879,85€, ecart confirme empiriquement contre le fichier fait a la main de juin
    2026). Colonnes P-Z (Zone/Pays/Poids/dimensions, absentes du CSV douane) restent
    VIDES pour les lignes douane -- confirme deja sans risque (cf. docstring en tete
    de fichier)."""
    def read_csv_dict(path):
        with open(path, encoding="cp1252", newline="") as fh:
            data = list(csv.reader(fh, delimiter=";"))
        hdr = [h.strip() for h in data[0]]
        rows = [r for r in data[1:] if any(v != "" for v in r)]
        return hdr, rows

    presta_hdr, presta_rows = None, []
    for f in presta_paths:
        h, r = read_csv_dict(f)
        if presta_hdr is None:
            presta_hdr = h
        presta_rows.extend(r)
    if not presta_hdr:
        raise RuntimeError("Aucun CSV 'Prestations au colis' fourni -- impossible de determiner le referentiel de colonnes.")
    ncol = len(presta_hdr)

    rows = list(presta_rows)  # deja alignees (referentiel lui-meme)
    for f in douane_paths:
        h, r = read_csv_dict(f)
        idx_in_presta = [presta_hdr.index(name) if name in presta_hdr else None for name in h]
        for raw_row in r:
            aligned = [""] * ncol
            for i, val in enumerate(raw_row):
                if i < len(idx_in_presta) and idx_in_presta[i] is not None:
                    aligned[idx_in_presta[i]] = val
            rows.append(aligned)

    rows = [r + [""] * (ncol - len(r)) for r in rows]  # securite si une ligne presta est plus courte
    return presta_hdr, rows, ncol


def extract_pdf_total_ht(pdf_path):
    """1ere page de la facture Colissimo, bloc 'Votre recapitulatif de facture HT' --
    extrait 'TOTAL HT' + 'Indemnisations' + 'Avoirs' (libelles explicites en clair, ex.
    'TOTAL HT 23 220,68 �' / 'Indemnisations -341,78 �' / 'Avoirs -8,00 �' -- capture
    ecran utilisateur 2026-08-17, facture de juillet 2026). 'Avoirs' n'apparait pas tous
    les mois (absent en juin 2026) -- reste None si le libelle n'est pas trouve, PAS 0
    (0 serait une fausse info, alors qu'on ne sait juste pas)."""
    try:
        import pdfplumber
    except ImportError:
        return None, None, None, None
    try:
        with pdfplumber.open(pdf_path) as pdf:
            text = pdf.pages[0].extract_text() or ""
    except Exception:
        return None, None, None, None
    m_num = re.search(r"FACTURE\s*N[°\s]*([A-Z0-9]+)", text)
    # "€" du PDF Colissimo parfois mal encode (extrait comme "�" par pdfplumber selon
    # la police embarquee) -- "." accepte n'importe quel caractere de fin, pas seulement "€".
    m_total = re.search(r"TOTAL\s*HT\s*([\d\s\xa0]+,\d{2})\s*.", text)
    # Indemnisations/Avoirs : toujours negatifs dans le PDF (deduits du total) -- stockes
    # en VALEUR ABSOLUE dans "Bilan Factures" (confirme sur juin/juillet : G=D+E+F doit
    # redonner le total HT positif, cf. 'Somme total' = 'PDF HT'+'Indemnisation'+'Avoir').
    m_indem = re.search(r"Indemnisations\s*-?([\d\s\xa0]+,\d{2})\s*.", text)
    m_avoir = re.search(r"Avoirs\s*-?([\d\s\xa0]+,\d{2})\s*.", text)
    numero = m_num.group(1) if m_num else None

    def parse_montant(m):
        return float(m.group(1).replace("\xa0", "").replace(" ", "").replace(",", ".")) if m else None

    total = parse_montant(m_total)
    indemnisation = parse_montant(m_indem)
    avoir = parse_montant(m_avoir)
    return numero, total, indemnisation, avoir


def fill_reconciliation(wb, pdf_paths):
    """Onglet 'Bilan Factures' : colle 'TOTAL HT' (colonne D, 'PDF HT'), 'Indemnisations'
    (colonne E) et 'Avoirs' (colonne F) extraits du PDF -- places en LIGNE FIXE 4 (juste
    sous l'entete ligne 3), PAS sur la ligne du n. de facture (confirme identique sur les
    2 classeurs reels juin/juillet : D4/E4/F4, alors que la ligne CO0... est en ligne 5+).
    G4 = D4+E4+F4 (deja en formule dans le modele) redonne le total HT reel, comparable a
    'Somme de Total HT' calculee depuis les CSV bruts (colonne B)."""
    bf = wb.Sheets("Bilan Factures")
    xlUp = -4162
    lastRow = bf.Cells(bf.Rows.Count, 1).End(xlUp).Row
    numeros = set()
    for r in range(4, lastRow + 1):
        v = bf.Cells(r, 1).Value
        if v:
            numeros.add(str(v).strip())
    matched = 0
    for p in pdf_paths:
        numero, total, indemnisation, avoir = extract_pdf_total_ht(p)
        if numero is None or total is None:
            print(f"Reconciliation Colissimo : PDF ignore (numero de facture ou TOTAL HT introuvable) -> {os.path.basename(p)}")
            continue
        if numero not in numeros:
            print(f"Reconciliation Colissimo : facture {numero} (PDF {os.path.basename(p)}) absente de 'Bilan Factures'")
            continue
        bf.Cells(4, 4).Value = total
        if indemnisation is not None:
            bf.Cells(4, 5).Value = indemnisation
        if avoir is not None:
            bf.Cells(4, 6).Value = avoir
        matched += 1
        print(f"Reconciliation Colissimo : facture {numero} -> TOTAL HT PDF={total}, Indemnisation={indemnisation}, Avoir={avoir}")
    print(f"Reconciliation Colissimo : {matched}/{len(pdf_paths)} PDF apparies")


def retry(fn, tries=8, delay=0.6):
    import time
    last = None
    for _ in range(tries):
        try:
            return fn()
        except Exception as e:  # Excel occupe -> RPC_E_CALL_REJECTED : on reessaie
            last = e
            time.sleep(delay)
    raise last


def parse_args(rest):
    """--presta <csv...> [--douane <csv...>] [--pdf <pdf...>] [--period AAAA_MM]"""
    presta, douane, pdf_paths = [], [], []
    period = None
    cur = None
    for a in rest:
        if a == "--presta":
            cur = "presta"
        elif a == "--douane":
            cur = "douane"
        elif a == "--pdf":
            cur = "pdf"
        elif a == "--period":
            cur = "period"
        elif cur == "presta":
            presta.append(a)
        elif cur == "douane":
            douane.append(a)
        elif cur == "pdf":
            pdf_paths.append(a)
        elif cur == "period":
            period = a
            cur = None
    return presta, douane, pdf_paths, period


def main():
    modele, sortie = sys.argv[1], sys.argv[2]
    presta_paths, douane_paths, pdf_paths, period = parse_args(sys.argv[3:])
    shutil.copyfile(modele, sortie)  # on ne touche JAMAIS au modele

    hdr, rows, ncol = load_rows(presta_paths, douane_paths)
    n = len(rows)
    data = [[coerce(v) for v in r] for r in rows]  # bloc 2D a coller
    print(f"CSV Colissimo : {n} lignes ({len(presta_paths)} presta + {len(douane_paths)} douane), {ncol} colonnes")

    import win32com.client as win32
    xlUp = -4162
    xl = win32.DispatchEx("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    xl.AskToUpdateLinks = False
    try:
        wb = retry(lambda: xl.Workbooks.Open(os.path.abspath(sortie), UpdateLinks=0, ReadOnly=False))
        if wb is None:
            raise RuntimeError("Excel n'a pas pu ouvrir le fichier (deja ouvert ? verrouille ?)")

        # ---- 0) Resolution auto des "Pays a creer" / "zone inconnue" -- AVANT le collage
        #    des donnees, pour que les formules XLOOKUP A-G de "Facture Colissimo" trouvent
        #    directement les bonnes valeurs des le premier RefreshAll(). Demande utilisateur
        #    2026-08-17 (video process Colissimo_1, Partie 14 : la recherche croisee "mois
        #    precedent par tracking" y est filmee mais quasi-toujours infructueuse -- un
        #    tracking ne se repete jamais d'un mois a l'autre -- donc reproduite ici comme
        #    2 mecanismes distincts et plus fiables) :
        #    a) Pays : ajout automatique si le nom matche un pays connu de ISO_PAYS_FR
        #       (filet de securite complet 249 pays, jamais d'invention de code).
        #    b) Modes envois : ajout automatique en DEVINANT la zone/mode la plus frequente
        #       deja utilisee pour ce meme prefixe tracking -- TOUJOURS signale (jamais
        #       silencieux, cf. AJOUTS_PAYS/AJOUTS_MODES loggues plus bas) car ceci reste
        #       une estimation, pas une certitude (regles tarifaires Colissimo non
        #       deductibles depuis les donnees).
        pays_seen, prefix_pays_names = extract_pays_and_prefixes(hdr, rows)
        pays_ws = wb.Sheets("Pays")
        added_pays = resolve_missing_pays(pays_ws, pays_seen, xlUp)
        for name, code in added_pays:
            print(f"AJOUT_PAYS_AUTO:{name} -> {code} (ajoute automatiquement a la feuille 'Pays', table ISO -- a verifier)")

        # Traduit chaque (prefixe, nom pays) vu dans les CSV en (prefixe, code ISO), en
        # relisant la feuille Pays fraichement mise a jour -- necessaire pour construire les
        # concatenations "prefixe+code" attendues par "Modes envois".
        pays_name_to_code = {}
        pays_last_row = pays_ws.Cells(pays_ws.Rows.Count, 1).End(xlUp).Row
        for r in range(2, pays_last_row + 1):
            nm = pays_ws.Cells(r, 1).Value
            cd = pays_ws.Cells(r, 2).Value
            if nm and cd:
                pays_name_to_code[normalize_pays(str(nm))] = str(cd)
        prefix_pays_codes = set()
        for prefixe, pays_name in prefix_pays_names:
            code = pays_name_to_code.get(normalize_pays(pays_name))
            if code:
                prefix_pays_codes.add((prefixe, code))

        modes_ws = wb.Sheets("Modes envois")
        added_modes = resolve_missing_modes_envois(modes_ws, prefix_pays_codes, xlUp)
        for concat, prefixe, pays_code, zone, mode in added_modes:
            print(f"AJOUT_MODE_ENVOI_AUTO:{concat} -> zone={zone}, mode={mode} (DEDUIT du prefixe '{prefixe}' le plus frequent -- A VERIFIER, pas une certitude)")

        # ---- 1) Facture Colissimo : purge + collage des donnees brutes + formules A-G ----
        FIRST_RAW_COL = 8   # colonne H : debut des donnees brutes CSV
        LAST_FORMULA_COL = 7  # colonne G : derniere formule calculee (Categorie)
        fc = wb.Sheets("Facture Colissimo")
        if fc.AutoFilterMode:
            fc.AutoFilterMode = False
        lastRawCol = FIRST_RAW_COL + ncol - 1
        oldLast = fc.Cells(fc.Rows.Count, FIRST_RAW_COL).End(xlUp).Row
        newLast = 1 + n

        # PURGE UNIQUEMENT LES DONNEES BRUTES (colonnes H+), PAS les colonnes A-G formules --
        # BUG TROUVE 2026-08-14 (1er test) : ClearContents() sur TOUTE la plage A:lastRawCol
        # effacait aussi la ligne 2 des colonnes A-G, qui sert de MODELE au FillDown()
        # suivant -- resultat : FillDown() propageait des cellules VIDES sur toute la colonne,
        # donc Pays/Mode/Zone/Categorie restaient vides partout (TCD ensuite tout a 0).
        if oldLast >= 2:
            retry(lambda: fc.Range(fc.Cells(2, FIRST_RAW_COL), fc.Cells(oldLast, lastRawCol)).ClearContents())
        retry(lambda: setattr(fc.Range(fc.Cells(2, FIRST_RAW_COL), fc.Cells(newLast, lastRawCol)), "Value", data))
        retry(lambda: fc.Range(fc.Cells(2, 1), fc.Cells(newLast, LAST_FORMULA_COL)).FillDown())
        if newLast < oldLast:
            retry(lambda: fc.Range(fc.Cells(newLast + 1, 1), fc.Cells(oldLast, lastRawCol)).ClearContents())
        print(f"Facture Colissimo : {oldLast - 1} anciennes lignes -> {n} nouvelles")

        # Tri par colonne O "N colis" (tracking) -- demande utilisateur 2026-08-17, deja
        # l'ordre de fait dans les 2 classeurs reels (groupe les lignes d'un meme colis
        # ensemble). Porte UNIQUEMENT sur les lignes de donnees (2..newLast), colonnes A a
        # lastRawCol (A-G formules + H+ brutes) -- jamais la ligne d'entete.
        N_COLIS_COL = 15  # colonne O
        if newLast >= 2:
            rng_sort = fc.Range(fc.Cells(2, 1), fc.Cells(newLast, lastRawCol))
            fc.Sort.SortFields.Clear()
            fc.Sort.SortFields.Add2(Key=fc.Range(fc.Cells(2, N_COLIS_COL), fc.Cells(newLast, N_COLIS_COL)), SortOn=0, Order=1)
            fc.Sort.SetRange(rng_sort)
            fc.Sort.Header = 0  # xlNoGuess -- la plage ne contient pas la ligne d'entete
            retry(lambda: fc.Sort.Apply())

        # Liste des 8 postes ERP (colonne G), juste apres la derniere ligne de donnees --
        # PRESENTE dans les 2 classeurs reels (juin/juillet) mais purement decorative (aucune
        # formule/validation ne la referme, cf. verification 2026-08-17), remplacee par de
        # vraies donnees a chaque generation puisque "Facture Colissimo" grandit d'un mois a
        # l'autre -- sans reecriture ici, elle disparaissait des que newLast >= oldLast.
        POSTE_KEYS_ERP = [
            "Adresse", "Assurance", "Colis volumineux", "Droits et taxes",
            "Frêt", "plus-value BtoC", "Taxe gazole", "Zones éloignées",
        ]
        # Arriere-plan des 8 cellules (colonne G) = MEME couleur que la colonne O "N colis"
        # (demande utilisateur 2026-08-17, capture ecran) -- PAS la couleur heritee de
        # "Categorie" sur une vraie ligne de donnees. Repris depuis la ligne de donnees
        # juste au-dessus (colonne O), qui garde toujours cette couleur quel que soit le mois.
        n_colis_color = fc.Cells(newLast, 15).Interior.Color  # colonne O, derniere ligne de donnees
        for i, poste in enumerate(POSTE_KEYS_ERP):
            cell = fc.Cells(newLast + 1 + i, LAST_FORMULA_COL)
            cell.Value = poste
            cell.Interior.Color = n_colis_color

        # ---- 2) Rafraichir tous les TCD (Poids, TCD, Bilan Factures, puis Bilan Client qui
        #    depend de TCD -- RefreshAll() gere l'ordre de dependance) ----
        retry(lambda: wb.RefreshAll())
        try:
            xl.CalculateUntilAsyncQueriesDone()
        except Exception:
            pass
        xl.Calculate()

        # BUG TROUVE 2026-08-17 (donnees reelles de juillet 2026) : le champ colonne
        # "Categorie" du TCD a ShowAllItems=False par defaut -- si un mois n'a AUCUNE
        # ligne dans une categorie (ex. juillet : "Assurance"/"plus-value BtoC" absentes),
        # Excel fait DISPARAITRE la colonne correspondante au lieu de l'afficher a 0, ce qui
        # DECALE toutes les colonnes suivantes (H/I deviennent Fret/Taxe gazole au lieu de
        # Droits et taxes/Fret). Or "Import CSV" reference TCD!H:H, TCD!I:I... par LETTRE DE
        # COLONNE FIXE (formules copiees du modele de juin, jamais dynamiques) -- un decalage
        # silencieux fait alors remonter la mauvaise valeur dans la mauvaise colonne (ex.
        # "Droits et taxes" recevait la valeur de "Fret"/10 au lieu de la sienne). Fixe en
        # forcant ShowAllItems=True sur le champ "Categorie" : les 8 postes ERP restent
        # TOUJOURS aux memes colonnes E-L, a 0 si absents du mois, comme dans le fichier de
        # juillet fait a la main (verifie identique par nom de colonne).
        #
        # ATTENTION : ShowAllItems affiche aussi les items OBSOLETES restes dans le cache du
        # pivot (ex. "Surcharge de securite", present dans le cache du modele de juin mais
        # absent de la table de correspondance actuelle ET des CSV bruts -- probablement un
        # residu d'un mois anterieur jamais nettoye). Sans purge, cet item fantome cree une
        # 9e colonne qui decale "Zones eloignees" de L a M et casse les formules figees
        # "Total hors Gazole"/"Total + GO" (SUM(E:J,L) / SUM(E:L), ecrites pour 8 colonnes
        # fixes). MissingItemsLimit=0 purge TROP (il retire aussi "Assurance"/"plus-value
        # BtoC" les mois ou elles sont a 0, qui doivent pourtant rester visibles -- c'est
        # justement ce que ShowAllItems doit empecher) -- on retire donc UNIQUEMENT les items
        # qui ne font pas partie des 8 postes ERP legitimes (POSTE_KEYS_ERP), en gardant
        # MissingItemsLimit par defaut pour ne pas perdre les postes legitimes a 0.
        POSTE_KEYS_ERP = {
            "Adresse", "Assurance", "Colis volumineux", "Droits et taxes",
            "Frêt", "plus-value BtoC", "Taxe gazole", "Zones éloignées",
        }
        try:
            tcd_pt = wb.Sheets("TCD").PivotTables(1)
            for pf in tcd_pt.PivotFields():
                if pf.Orientation == 2:  # xlColumnField
                    for pi in list(pf.PivotItems()):
                        name = str(pi.Name)
                        if name not in POSTE_KEYS_ERP and name not in ("(blank)", "#N/A"):
                            try:
                                pi.Delete()
                                print(f"TCD : item obsolete retire du cache pivot -> {name!r}")
                            except Exception as e:
                                print(f"TCD : impossible de retirer l'item obsolete {name!r} :", e)
                    pf.ShowAllItems = True
            retry(lambda: wb.RefreshAll())
            xl.Calculate()
        except Exception as e:
            print("Avertissement : ShowAllItems sur le TCD a echoue :", e)

        # Colonnes B "Total hors Gazole"/C "Total + GO" du TCD : formules JUXTAPOSEES au
        # pivot (SUM(E:J,L) / SUM(E:L) par ligne), PAS des champs du TCD lui-meme -> le
        # RefreshAll() ci-dessus etend bien les colonnes D+ (pivot) au nouveau nombre de
        # lignes, mais B/C restent figees a l'ancien nombre de lignes du MODELE (juin,
        # 2768 lignes) -- BUG TROUVE 2026-08-17 (juillet, 3190 lignes) : B/C vides
        # au-dela de la ligne 2768, "Total hors Gazole"/"Total + GO" alors sous-estimes
        # de tout ce qui est au-dela (mais 'Total + GO', lui, restait juste par coincidence
        # car egal a la somme des colonnes E:L, verifiee separement plus bas). Meme regle que
        # "Import CSV" (etape 3 suivante) et "ID Client" (colonne A) : etendre par FillDown()
        # depuis la ligne 3 (modele) jusqu'a la nouvelle derniere ligne du TCD.
        tcd_bc = wb.Sheets("TCD")
        bcLast = tcd_bc.Cells(tcd_bc.Rows.Count, 4).End(xlUp).Row  # colonne D = tracking
        if bcLast >= 3:
            retry(lambda: tcd_bc.Range(tcd_bc.Cells(3, 2), tcd_bc.Cells(bcLast, 3)).FillDown())
            xl.Calculate()

        # ---- 3) Import CSV : TOUT en formules (pas de donnees brutes a coller) -- son
        #    nombre de lignes suit le nombre de colis UNIQUES du TCD recalcule (pas n, le
        #    nombre de charges brutes). On ne touche JAMAIS la ligne 2 (modele des formules
        #    pour le FillDown). F2 = TCD!D3 (1ere ligne de donnees du TCD, apres l'entete).
        LAST_COL_IMPORT = 21  # colonne U : derniere colonne utile de "Import CSV"
        tcd = wb.Sheets("TCD")
        tcdLast = tcd.Cells(tcd.Rows.Count, 4).End(xlUp).Row  # colonne D = tracking (etiquettes de lignes)
        nbColis = tcdLast - 2  # le TCD demarre a la ligne 3 (lignes 1-2 = entetes du pivot)
        impNewLast = 1 + nbColis
        print(f"Import CSV : {nbColis} colis uniques (TCD, lignes 3..{tcdLast})")

        # Colonne A "ID Client" du TCD : PAS un champ du pivot (RowField = "N colis", cf.
        # docstring) -- une colonne JUXTAPOSEE, valeurs litteralement saisies a la main en
        # face de chaque tracking (confirme sur le modele : '6739'/'7027'... alignes ligne a
        # ligne avec la colonne D). Le RefreshAll() ci-dessus recalcule le pivot (nouveaux
        # trackings du mois traite en D) mais NE TOUCHE JAMAIS la colonne A -> sans purge,
        # les anciens ID clients du modele restent affiches, DESALIGNES avec les nouveaux
        # trackings (pas juste vide -- une fausse info silencieuse). Purgee ici pour laisser
        # la personne qui renseigne les ID clients repartir d'une colonne vide alignee sur
        # le bon mois (decision utilisateur 2026-08-14, meme regle que BLS/Chronopost : le
        # TCD lui-meme, colonnes B/C/D.. (Total hors Gazole/Total + GO/tracking/postes),
        # reste intact et continuera de s'actualiser normalement par RefreshAll une fois les
        # ID clients ressaisis).
        if tcdLast >= 3:
            retry(lambda: tcd.Range(tcd.Cells(3, 1), tcd.Cells(tcdLast, 1)).ClearContents())

        imp = wb.Sheets("Import CSV")
        if imp.AutoFilterMode:
            imp.AutoFilterMode = False

        # Colonne B "Date validite tarif" : valeur LITTERALE (pas une formule) sur la ligne
        # modele (B2), copiee telle quelle par le FillDown() ci-dessous sur toutes les lignes
        # (B3+ = "=B2" dans le modele). BUG TROUVE 2026-08-17 (juillet 2026) : le modele clone
        # (juin) laisse B2 a l'ancienne date de juin -> sans correction, TOUT le mois traite
        # affiche encore le 01/06/2026 en "Date validite tarif", quel que soit le mois reel des
        # CSV. Le fichier de juillet fait a la main a B2 mis a jour au 01/07/2026 -- reproduit
        # ici depuis --period (AAAA_MM, deja calcule cote Node depuis la colonne "Date" du CSV).
        if period:
            m = re.fullmatch(r"(\d{4})_(\d{2})", period)
            if m:
                import datetime
                # Serial Excel (jours depuis 1899-12-30), PAS un datetime.datetime Python
                # passe tel quel a la cellule COM -- BUG TROUVE 2026-08-17 : le marshaling
                # win32com d'un datetime naif se fait via VT_DATE, reinterprete au retour
                # selon le fuseau local (UTC+2 ete FR) -> decalage de -2h, la date affichee
                # devenait "30/06/2026 22:00" au lieu de "01/07/2026 00:00". Le serial est
                # sans ambiguite de fuseau.
                excel_epoch = datetime.date(1899, 12, 30)
                serial = (datetime.date(int(m.group(1)), int(m.group(2)), 1) - excel_epoch).days
                retry(lambda: setattr(imp.Cells(2, 2), "Value", serial))

        impOldLast = imp.Cells(imp.Rows.Count, 6).End(xlUp).Row  # colonne F = N Tracking
        if impNewLast > impOldLast:
            retry(lambda: imp.Range(imp.Cells(2, 1), imp.Cells(impNewLast, LAST_COL_IMPORT)).FillDown())
        elif impNewLast < impOldLast:
            retry(lambda: imp.Range(imp.Cells(impNewLast + 1, 1), imp.Cells(impOldLast, LAST_COL_IMPORT)).ClearContents())
        xl.Calculate()  # recalcule Import CSV une fois etendu/reduit

        # ---- 4) Reconciliation PDF : TOTAL HT Colissimo vs Somme de Total HT (Bilan Factures) ----
        if pdf_paths:
            try:
                fill_reconciliation(wb, pdf_paths)
            except Exception as e:
                print("Reconciliation Colissimo ignoree :", e)
            xl.Calculate()

        retry(lambda: wb.Save())
        wb.Close(SaveChanges=True)
        print("OK -> " + sortie)
    finally:
        try:
            xl.Quit()
        except Exception:
            pass


if __name__ == "__main__":
    main()
