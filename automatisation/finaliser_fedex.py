#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
FINALISEUR FEDEX — produit "AAAA_MM_Facture Fedex.xlsx" A L'IDENTIQUE du
fichier fait a la main (6 feuilles : Zoning, Bilan factures, Shipment
Detail, TCD, Bilan clients, Import ERP), en partant du fichier existant
comme MODELE. Entree = CSV BRUT FedEx (export portail, en-tete FR, meme
fichier que celui lu par le carrier Node index.js).

Classeur reel :
  Shipment Detail : 1 ligne = 1 colis/envoi, donnees brutes CSV FedEx
    (colonnes F->BS, en-tete FR, resolues PAR NOM), colonnes A->E
    CALCULEES (A="Montant"=BK+BN, B="Clients"=ID client MANUEL (laisse
    vide a la generation, saisi apres coup, cf. regle transversale ID
    client), C="Categorie"=toujours vide (confirme juin 2026), D="E/P"
    =VALEUR LITERALE ("entreprise"/"particulier", calculee ICI via
    l'export WMS "expeditions_brut" m/m-1 -- meme mecanisme que Delivengo/
    Geodis/DPD, PAS une formule Excel : la colonne D du fichier reel de
    juin 2026 est 100% litterale sur 3607/3607 lignes), E="Mode d'envoi"
    =formule =IF(Service="FedEx Priority","ST","FICP").
  TCD : VRAI TCD natif (RowFields=Shipment Tracking Number+Recipient
    Country/Territory, DataFields=Somme de Montant+Somme de Shipment
    Rated Weight (Pounds)) + colonnes MANUELLES juxtaposees E (Poids kg,
    ROUNDUP*0.453592) / F (E/P, XLOOKUP vers Shipment Detail!D).
  Bilan factures : VRAI TCD (RowField=Invoice Number, DataFields=Fret
    HT+Total TTC) + colonne D "PDF" MANUELLE (collee/saisie), E
    "Ecart"=C-D.
  Bilan clients : VRAI TCD (RowField=Clients, DataField=Somme de
    Montant), source='Shipment Detail' (PAS 'TCD').
  Import ERP : formules PAR LIGNE (pas de decalage fixe comme TNT),
    referencant TCD par XLOOKUP(tracking) et Shipment Detail par
    XLOOKUP(tracking), PAS par position.

PIEGE CONNU (meme famille que DPD/Geodis/Mondial Relay/Chronopost/TNT) :
  les 3 PivotTables (TCD/Bilan factures/Bilan clients) du modele ont un
  PivotCache SOURCE ETROITE/FIGEE ('Shipment Detail'!C1:C70 ou C1:C2,
  constate via COM sur le modele de juin 2026) -- DOIT etre redirige
  vers la vraie plage large de 'Shipment Detail' a chaque generation,
  sinon le TCD reste fige sur les anciennes donnees du modele.

Detection "facture d'un autre mois" (RESOLUE 2026-08-20, cf. commentaire en
tete de facturation-app/src/carriers/fedex/index.js) : une facture du CSV
brut sans PDF correspondant parmi les PDF portail fournis est exclue
automatiquement (le pole transport ne recoit le PDF qu'une fois la facture
reellement due ce mois-ci) -- validee exacte sur juin 2026 (exclut les 2
memes factures que le fichier reel livre, Fret recalcule = 45515,83EUR
exactement). Ne s'applique que si au moins 1 PDF est fourni.

Necessite : Windows + Excel + pywin32.
Usage :
  python finaliser_fedex.py "<modele.xlsx>" "<sortie.xlsx>" --csv <csv1> [<csv2>...] [--brut <brutM.xlsx> [<brutM-1.xlsx>]] [--pdf <pdf1> [<pdf2>...]]
"""
import sys, os, shutil, re, csv


def normalize_header(h):
    return re.sub(r"\s+", " ", str(h or "")).strip()


def compare_key(h):
    s = normalize_header(h).lower()
    s = s.translate(str.maketrans("éèêëàâäùûüôöîïç’‘", "eeeeaaauuuooiic''"))
    s = re.sub(r"[.,]", "", s)
    return re.sub(r"\s+", " ", s).strip()


def col_index(header, name):
    """name : une chaine, ou une liste d'alias (ex. [libelle_FR, libelle_EN]) -- BUG
    TROUVE 2026-09-09 (meme piege que carriers/fedex/index.js cote Node) : un export
    FedEx en ANGLAIS n'a aucune colonne au libelle FR attendu, ce qui rendait i_tracking/
    i_invoice_date/etc. = None (plantage plus loin, ex. r[None])."""
    names = name if isinstance(name, (list, tuple)) else [name]
    for n in names:
        target = compare_key(n)
        for i, h in enumerate(header):
            if compare_key(h) == target:
                return i
    return None


def col_letter(idx0):
    import openpyxl
    return openpyxl.utils.get_column_letter(idx0 + 1)


def detect_delimiter_fedex(first_line):
    """Reprend detectDelimiter de carriers/fedex/index.js (compte , et ; HORS zones entre
    guillemets sur la 1ere ligne) -- plus robuste que l'ancien sample.count(';') > count(',')
    (celui-ci se trompait si un champ entre guillemets contenait lui-meme le delimiteur)."""
    commas = semicolons = 0
    in_quotes = False
    for c in first_line:
        if c == '"':
            in_quotes = not in_quotes
        elif not in_quotes and c == ",":
            commas += 1
        elif not in_quotes and c == ";":
            semicolons += 1
    return ";" if semicolons > commas else ","


def read_fedex_csv(path):
    """CSV brut FedEx (export portail, en-tete FR ou EN, UTF-8 -- confirme sur les CSV bruts
    reels de juin 2026, cf. carrier Node index.js). Separateur virgule ou point-virgule selon
    l'export (BUG TROUVE 2026-09-09, cf. index.js detectDelimiter : un CSV Shipment Detail
    anglais recu utilisait ";" -- csv.reader() par defaut sur "," aurait alors lu chaque ligne
    comme un seul champ).
    BUG TROUVE 2026-10-08 (classeur fige en plein Excel COM, access violation combase.dll
    pendant l'ecriture bulk Shipment Detail!Value=data_brut) : ce parseur utilisait le
    csv.reader() standard de Python, qui applique les regles CSV STRICTES (un guillemet
    fermant DOIT etre suivi d'un delimiteur/fin de ligne) -- le fix du 2026-09-09 (champs mal
    echappes par FedEx, cf. index.js) n'avait ete porte QUE cote JS (carrier Node), jamais ici.
    Sur un CSV reel avec un guillemet interne isole (ex. "\".OK", "9."-3COUR), csv.reader()
    fermait le champ a tort, desequilibrant TOUT LE RESTE DU FICHIER -> plusieurs lignes
    fusionnees en un seul champ geant (potentiellement >32767 caracteres, la limite d'une
    cellule Excel) -- ecrire CA en bulk via COM a fait planter Excel (combase.dll, 0xc0000005)
    plutot que de lever une exception Python normale (explique le traceback tronque, sans
    message : le crash n'est pas une exception Python standard). Remplace par un PORT FIDELE,
    caractere par caractere, de readFedexCsv (index.js) -- meme tolerance sur les guillemets
    isoles (un vrai guillemet fermant est TOUJOURS suivi d'un delimiteur/\\r/\\n/fin de texte,
    sinon il reste dans le texte du champ), meme detection de delimiteur hors guillemets."""
    with open(path, "rb") as f:
        raw = f.read()
    text = raw.decode("utf-8-sig")
    first_line = text.split("\n", 1)[0].rstrip("\r")
    delim = detect_delimiter_fedex(first_line)
    rows = []
    row = []
    field = []
    in_quotes = False
    i = 0
    n = len(text)
    while i < n:
        c = text[i]
        if in_quotes:
            if c == '"':
                if i + 1 < n and text[i + 1] == '"':
                    field.append('"')
                    i += 2
                    continue
                nxt = text[i + 1] if i + 1 < n else None
                if nxt in (delim, "\r", "\n", None):
                    in_quotes = False
                else:
                    field.append(c)
            else:
                field.append(c)
        elif c == '"':
            in_quotes = True
        elif c == delim:
            row.append("".join(field))
            field = []
        elif c == "\r":
            pass
        elif c == "\n":
            row.append("".join(field))
            rows.append(row)
            row = []
            field = []
        else:
            field.append(c)
        i += 1
    if field or row:
        row.append("".join(field))
        rows.append(row)
    if not rows:
        return None
    header = [normalize_header(h) for h in rows[0]]
    data = [r for r in rows[1:] if any(v not in (None, "") for v in r)]
    return header, data


def coerce(v):
    if v is None or v == "":
        return None
    if isinstance(v, (int, float)):
        return v
    s = str(v).strip()
    if re.fullmatch(r"-?0\d+", s):
        return v
    if re.fullmatch(r"-?\d+(,\d+)?", s):
        return float(s.replace(",", "."))
    if re.fullmatch(r"-?\d+\.\d+", s):
        return float(s)
    # BUG TROUVE 2026-10-08 (CSV reel de septembre, ligne 1075, colonne "Reference Notes Line
    # 1" = "=.!+23" -- note de reference libre saisie par un client FedEx) : une chaine ecrite
    # via COM .Value= qui commence par "=" declenche TOUJOURS le parseur de formule Excel
    # (jamais traitee comme texte litteral) -- si ce qui suit n'est PAS une formule valide
    # (ici ".!+23"), Excel renvoie une exception COM generique ("Une exception s'est
    # produite", AUCUN rapport reel avec une penurie de memoire malgre les apparences du code
    # d'erreur). Reproduit et confirme isolement (ws.Cells(...).Value = "=.!+23" plante seul,
    # sans rapport avec la taille du fichier -- testes aussi "+"/"-"/"@": ceux-ci n'ont PAS ce
    # risque, Excel les accepte toujours tels quels, donc pas touches ici pour eviter de
    # changer un comportement existant non casse, ex. "+33...": deja converti en nombre). Fix
    # standard Excel : un guillemet simple en tete FORCE le texte litteral (equivalent de
    # taper '=.!+23 au clavier) -- invisible a la lecture (.Value relit la chaine originale
    # sans l'apostrophe), applique ICI pour couvrir tous les points d'ecriture .Value=/
    # .Formula= de ce finaliseur qui passent par coerce().
    if s[:1] == "=":
        return "'" + v
    return v


# ---- E/P : export WMS "expeditions_brut" (meme mecanisme que Delivengo/Geodis/DPD) ----
X_BRUT_TRACK, X_BRUT_DES_PARTICULIER = 41, 16  # PRO_TRACKING (AP) / DES_PARTICULIER (Q)


def load_brut_ep(brut_paths):
    """Map {tracking (str) -> 'entreprise'/'particulier'}. Ordre des fichiers = priorite
    (mois courant, puis M-1) -- la PREMIERE occurrence gagne, meme regle que Delivengo."""
    import io
    import openpyxl
    m = {}
    for p in [bp for bp in brut_paths if bp]:
        # BUG TROUVE 2026-09-09 : openpyxl refuse un chemin SANS EXTENSION (InvalidFileException
        # "openpyxl does not support  file format") -- les fichiers uploades via multer sont
        # renommes en identifiant hexadecimal sans extension (meme piege deja corrige sur
        # BLS/UPS). Fix : ouvrir en binaire et passer un buffer memoire (openpyxl accepte un
        # objet fichier-like, detecte le format par SIGNATURE ZIP interne, pas par le nom).
        with open(p, "rb") as f:
            buf = io.BytesIO(f.read())
        wb = openpyxl.load_workbook(buf, read_only=True)
        ws = wb[wb.sheetnames[0]]
        for row in ws.iter_rows(min_row=2, values_only=True):
            if len(row) > X_BRUT_TRACK:
                t = str(row[X_BRUT_TRACK] or "").strip()
                v = str(row[X_BRUT_DES_PARTICULIER] or "").strip().lower()
                ep = "particulier" if v == "particulier" else ("entreprise" if v in ("entreprise", "point relais") else None)
                if t and ep and t not in m:
                    m[t] = ep
        wb.close()
    return m


def extract_pdf_infos(pdf_paths):
    """PDF FedEx ("Facture de frais de transport") : bloc labels/valeurs colle par
    pypdf ("No de Client:\\nNo de Facture:\\n...\\nMontant du\\n*****1234\\n634323394\\n...")
    + ligne "Total dû EUR <montant>" -- confirme exact sur les 9 PDF reels de juin 2026
    (1 a 195 pages)."""
    try:
        import pypdf
    except ImportError:
        return []
    results = []
    pat_inv = re.compile(
        r"No de Client\s*:\s*\nNo de Facture\s*:\s*\nDate de la facture\s*:\s*\n"
        r"Date d.[eé]ch[eé]ance\s*\nMontant d[uû]\s*\n\*+\d+\s*\n(\d+)\s*\n"
    )
    pat_total = re.compile(r"Total d[uû]\s*EUR\s*([\d.,]+)")
    # Bordereaux (1 par envoi/tracking) : pypdf aplatit tout sur 1 seule ligne texte (contrairement
    # a pdf-parse cote Node qui insere un retour a la ligne par cellule) -- pattern SANS \n entre
    # les groupes, cf. carrier Node index.js pour le meme controle avec un pattern adapte.
    pat_bordereau = re.compile(r"([\d.,]+)(\d{2}/\d{2}/\d{4})(\d{9,15})\s+([\d.,]+)\s+([\d.,]+)")
    # Supplements "surplus" (Zones eloignees/Colis volumineux PDF/Plus-value BtoC) : jamais
    # presents dans le CSV Shipment Detail (mail pole transport "Oubli facturation surplus
    # Fedex janvier a mai 2026", 2026-08-20, 7919,64EUR de surplus non factures sur 5 mois
    # faute d'automatisation) -- extraits par mots-cles, 1 bloc par bordereau (du debut du
    # tracking jusqu'au bordereau suivant), montant precede son libelle sur la meme ligne
    # texte (pypdf aplatit, contrairement a pdf-parse cote Node qui insere 1 retour a la ligne
    # par cellule -- meme logique de decoupage par bloc, pattern adapte). Colis volumineux PDF
    # s'ADDITIONNE a la regle Excel existante (longueur>60cm=10EUR), confirme pole transport
    # 2026-08-20 -- 2 mecanismes independants, pas un remplacement.
    pat_track_start = re.compile(r"\d{2}/\d{2}/\d{4}(\d{9,15})")
    kw_pats = {
        # Zones eloignees : libelle precis "Frais de traitement des importations aux
        # Etats-Unis" (precise par le pole transport 2026-08-25), cf. carrier Node index.js
        # (meme changement). BUG TROUVE 2026-08-25 : le PDF tronque le libelle a "...aux
        # Etats-Uni" (SANS le "s" final) -- "s" optionnel.
        # BUG TROUVE 2026-10-12 (signalement utilisateur, trackings 876739026110/876869350243/
        # 876993993158/877001078783/877072324270 : "Zone eloignee" = 21320EUR, notre PROPRE
        # code postal expediteur, 21320 Creancey) : le mot-cle generique "zone" (insensible a
        # la casse, 0-60 caracteres apres n'importe quel nombre) matchait a tort "21320
        # CREANCEY ZONE <nom de zone commerciale/industrielle du destinataire>" -- aucune
        # surtaxe reelle, juste l'adresse expediteur suivie d'un nom de zone d'activite cote
        # destinataire. Confirme : sur tous les PDF de test reels, CHAQUE occurrence de
        # "zone" est un faux positif de ce type (ou le libelle generique "Livraison hors
        # zone", lui aussi sans rapport avec une surtaxe dediee). Retire la regle generique --
        # garde uniquement le libelle precis ci-dessous, jamais de faux positif constate.
        "ZonesEloignees": [
            re.compile(r"([\d.,]+)[^\n]{0,60}Frais de traitement des importations aux [ÉE]tats-Unis?", re.IGNORECASE),
        ],
        # Colis volumineux (PDF) : "Charge pour dépassement de dimension" + "Supplément pour
        # manutention supplémentaire :poids" (precise par le pole transport 2026-08-25, cf.
        # carrier Node index.js, meme bordereau reel 873870075714, 2 lignes distinctes).
        "ColisVolumineuxPdf": [
            re.compile(r"([\d.,]+)[^\n]{0,60}Charge pour d[ée]passement de dimension"),
            re.compile(r"([\d.,]+)[^\n]{0,60}Suppl[ée]ment pour manutention suppl[ée]mentaire\s*:?\s*poids", re.IGNORECASE),
        ],
        "PlusValueB2C": [
            re.compile(r"([\d.,]+)[^\n]{0,60}Plusieurs pi[èe]ces"),
            re.compile(r"([\d.,]+)[^\n]{0,60}Demand Surcharge"),
        ],
    }
    for p in pdf_paths:
        try:
            text = "\n".join((page.extract_text() or "") for page in pypdf.PdfReader(p).pages)
        except Exception:
            continue
        m_inv = pat_inv.search(text)
        m_total = pat_total.search(text)
        bordereaux = [
            {"tracking": m.group(3), "date_envoi": m.group(2),
             "sous_total": float(m.group(5).replace(".", "").replace(",", "."))}
            for m in pat_bordereau.finditer(text)
        ]
        starts = [(m.start(), m.group(1)) for m in pat_track_start.finditer(text)]
        supplements = []
        for i, (idx, tracking) in enumerate(starts):
            next_idx = starts[i + 1][0] if i + 1 < len(starts) else len(text)
            block = text[idx:next_idx]
            for poste, pats in kw_pats.items():
                for pat in pats:
                    m = pat.search(block)
                    if m:
                        supplements.append({"tracking": tracking, "poste": poste,
                                             "montant": float(m.group(1).replace(".", "").replace(",", "."))})
        results.append({
            "file": os.path.basename(p),
            "numero_facture": m_inv.group(1) if m_inv else None,
            "total_du": float(m_total.group(1).replace(".", "").replace(",", ".")) if m_total else None,
            "bordereaux": bordereaux,
            "supplements": supplements,
        })
    return results


def parse_args(argv):
    """--csv <csv...> [--brut <xlsx...>] [--pdf <pdf...>] [--period AAAA_MM]"""
    modele, sortie = argv[1], argv[2]
    csvs, brut, pdfs, cur = [], [], [], None
    period = None
    i = 3
    while i < len(argv):
        a = argv[i]
        if a == "--csv":
            cur = "c"
        elif a == "--brut":
            cur = "b"
        elif a == "--pdf":
            cur = "p"
        elif a == "--period":
            i += 1
            period = argv[i] if i < len(argv) else None
            cur = None
        elif cur == "c":
            csvs.append(a)
        elif cur == "b":
            brut.append(a)
        elif cur == "p":
            pdfs.append(a)
        i += 1
    return modele, sortie, csvs, brut, pdfs, period


def retry(fn, tries=8, delay=0.6):
    import time
    last = None
    for _ in range(tries):
        try:
            return fn()
        except Exception as e:
            last = e
            time.sleep(delay)
    raise last


def main():
    modele, sortie, csv_paths, brut_paths, pdf_paths, period = parse_args(sys.argv)
    if not csv_paths:
        raise RuntimeError("Aucun CSV fourni (--csv <shipment_detail.csv> [...]).")
    shutil.copyfile(modele, sortie)  # on ne touche JAMAIS au modele

    # PDF lus TOT (contenu, PAS le nom de fichier -- server.js/multer renomme les fichiers
    # uploades en identifiants aleatoires sans extension/nom d'origine avant de les passer ici
    # -- bug constate 2026-08-20 : le nom "FEDEX_<invoice>_<compte>.pdf" n'existe QUE sur le
    # disque de l'utilisateur avant upload). Reutilise plus bas pour le "Total dû" (pas de 2e
    # lecture).
    pdf_infos_precoce = extract_pdf_infos(pdf_paths) if pdf_paths else []
    pdf_invoice_nums = {p["numero_facture"] for p in pdf_infos_precoce if p["numero_facture"]}

    # Supplements "surplus" par tracking, cumules sur tous les PDF (cf. extract_pdf_infos) --
    # ecrits en VALEUR (pas de formule Excel existante pour ces colonnes dans le modele, cf.
    # meme logique que index.js/carrier Node).
    supplements_par_tracking = {}
    for p in pdf_infos_precoce:
        for s in p.get("supplements", []):
            acc = supplements_par_tracking.setdefault(s["tracking"], {"ZonesEloignees": 0.0, "ColisVolumineuxPdf": 0.0, "PlusValueB2C": 0.0})
            acc[s["poste"]] = round(acc[s["poste"]] + s["montant"], 2)

    header = None
    rows = []
    for p in csv_paths:
        result = read_fedex_csv(p)
        if result is None:
            raise RuntimeError(f"Fichier vide ou illisible : {p}")
        h, d = result
        if header is None:
            header = h
        ncol = len(header)
        for r in d:
            rows.append((list(r) + [None] * ncol)[:ncol])
    n = len(rows)
    print(f"Entrée : {n} lignes x {len(header)} colonnes")

    i_tracking = col_index(header, ["Numéro de suivi de l'envoi", "Shipment Tracking Number"])
    i_service = col_index(header, ["Description du service", "Service Description"])
    # BUG TROUVE 2026-08-19 : les colonnes "... en USD" du CSV brut NE SONT PAS celles
    # utilisees par le classeur reel -- le vrai "Montant" (Shipment Detail!A=BK+BN) reference
    # les colonnes EUR "Devise de facturation des frais de transport/de la remise" (noms
    # trompeurs : ce sont des MONTANTS EUR, pas des codes devise -- le code devise est dans
    # la colonne suivante "Code de la devise de facturation"). Reconfirme sur un cas reel
    # (tracking 381547998902) : USD 539.82-458.09=81.73 vs EUR 465.03-394.62=70.41 = valeur
    # reelle du classeur modele. Meme correction que index.js (carrier Node).
    i_fret = col_index(header, ["Devise de facturation des frais de transport de l'envoi", "Shipment Freight Charge Billed Currency"])
    i_remise = col_index(header, ["Devise de facturation de la remise de l'envoi", "Shipment Discount Billed Currency"])
    i_taxe = col_index(header, ["Devise de facturation des droits de douane et taxes de l'envoi", "Shipment Duty And Tax Charge Billed Currency"])
    i_invoice_num = col_index(header, ["Numéro de facture", "Invoice Number"])
    i_mois = col_index(header, ["Mois de facturation (aaaamm)", "Invoice Month (yyyymm)"])
    i_invoice_date = col_index(header, ["Date de facturation (mm/jj/aaaa)", "Invoice Date (mm/dd/yyyy)"])

    # Controle "bordereau facture par FedEx mais absent du CSV" (2026-08-20, cas reel juillet :
    # facture 634393398, ecart 83,81EUR -- un envoi de NOVEMBRE 2025 apparaissait dans le PDF
    # mais n'a jamais ete reporte dans le CSV Shipment Detail du mois). Meme logique que
    # index.js (carrier Node) -- compare TOUS les trackings du PDF a TOUS les trackings du CSV
    # brut (avant tout filtrage Droits&Taxes/sans PDF).
    if i_tracking is not None:
        import datetime as _dt
        csv_trackings = {str(r[i_tracking] or "").strip() for r in rows if r[i_tracking]}
        for p in pdf_infos_precoce:
            manquants = [b for b in p.get("bordereaux", []) if b["tracking"] not in csv_trackings]
            if manquants:
                total_manquant = round(sum(b["sous_total"] for b in manquants), 2)
                # BUG TROUVE 2026-10-12 (signalement utilisateur, "alertes PDF illisibles") :
                # le detail listait TOUS les bordereaux manquants sans limite (jusqu'a 71 sur
                # un cas reel), chacun avec date+montant -- une phrase geante, illisible a
                # l'ecran. Reecrit en resume compact (nombre, plage de dates, total) +
                # echantillon limite a 5 trackings -- meme principe d'echantillonnage que
                # carriers/fedex/index.js (deja applique le 2026-09-09 pour la meme raison,
                # jamais porte ici cote Python), juste plus court (5 au lieu de 20, le detail
                # complet par tracking etant deja moins utile qu'un resume clair).
                dates = []
                for b in manquants:
                    try:
                        dates.append(_dt.datetime.strptime(b["date_envoi"], "%d/%m/%Y"))
                    except (ValueError, TypeError):
                        pass
                dates.sort()
                if dates:
                    d_min, d_max = dates[0].strftime("%d/%m"), dates[-1].strftime("%d/%m")
                    plage = f"du {d_min} au {d_max}" if d_min != d_max else f"le {d_min}"
                else:
                    plage = "date inconnue"
                echantillon = ", ".join(b["tracking"] for b in manquants[:5])
                suite = f" (+{len(manquants) - 5} autre(s))" if len(manquants) > 5 else ""
                # BUG TROUVE 2026-08-20 : le caractere '~' (approximation) etait a l'origine
                # U+2248 ("≈") -- absent du charset cp1252 utilise par la console Windows lors
                # de l'appel subprocess depuis server.js, provoquant un UnicodeEncodeError SUR
                # CE print() precis (les accents francais eux passent, deja presents dans cp1252)
                # -- crash masque en cascade cote Node (writeWorkbook, "Cannot read properties
                # of undefined (reading 'Droits et taxes')"), exactement le meme piege que le
                # bug PermissionError/EBUSY deja documente sur TNT.
                print(f"AVERTISSEMENT: PDF {p['file']}" + (f" (n° {p['numero_facture']})" if p['numero_facture'] else "")
                      + f" : {len(manquants)} bordereau(x) facturé(s) par FedEx, absent(s) du CSV Shipment Detail, envoyé(s) {plage}, total ~{total_manquant:.2f} EUR"
                      + f" — ex. {echantillon}{suite} — probable envoi d'un mois antérieur jamais reporté, à vérifier avec le pôle transport.")

    # Filtrage "Droits & Taxes" (seule regle certaine, cf. docstring index.js) : une facture
    # dont TOUTES les lignes ont un fret nul et une taxe non nulle est un ajustement
    # douanier pur -- exclue de Shipment Detail.
    # BUG TROUVE 2026-10-12 (CSV reel de septembre, 2955 lignes) : float() plantait avec
    # "could not convert string to float: 'EUR'" -- ces colonnes "Devise de facturation ..."
    # contiennent normalement un MONTANT malgre leur nom trompeur (cf. commentaire plus haut,
    # confirme sur plusieurs mois reels), mais au moins 1 ligne de ce fichier porte bien un
    # vrai code devise litteral ("EUR") a cet emplacement -- cas limite non rencontre avant.
    # Fix defensif : une valeur non convertible est traitee comme 0 (comme une case vide)
    # plutot que de faire planter tout le traitement, avec un avertissement pour rester visible.
    n_taxe_invalide = 0
    def montant_sur(r, idx):
        nonlocal n_taxe_invalide
        if idx is None or r[idx] in (None, ""):
            return 0.0
        try:
            return float(str(r[idx]).replace(",", "."))
        except ValueError:
            n_taxe_invalide += 1
            return 0.0
    factures_fret = {}
    factures_taxe = {}
    if i_invoice_num is not None:
        for r in rows:
            nf = str(r[i_invoice_num] or "").strip()
            if not nf:
                continue
            fret = montant_sur(r, i_fret)
            taxe = montant_sur(r, i_taxe)
            factures_fret[nf] = factures_fret.get(nf, 0.0) + fret
            factures_taxe[nf] = factures_taxe.get(nf, 0.0) + taxe
    if n_taxe_invalide:
        print(f"AVERTISSEMENT: {n_taxe_invalide} valeur(s) non numérique(s) rencontrée(s) dans les colonnes Frêt/Taxe (ex. \"EUR\" au lieu d'un montant) -- traitée(s) comme 0, à vérifier sur le CSV source.")
    factures_droits_taxes = {nf for nf in factures_fret if factures_fret[nf] == 0 and factures_taxe.get(nf, 0) != 0}
    if factures_droits_taxes and i_invoice_num is not None:
        avant = len(rows)
        rows = [r for r in rows if str(r[i_invoice_num] or "").strip() not in factures_droits_taxes]
        print(f"Facture(s) 100% Droits & Taxes exclue(s) : {sorted(factures_droits_taxes)} ({avant - len(rows)} ligne(s)).")
        n = len(rows)

    # Facture absente des PDF portail fournis (regle utilisateur 2026-08-20, meme logique que
    # index.js) : si au moins un PDF est fourni, un numero de facture du CSV brut absent de
    # TOUS les PDF fournis (numero lu dans leur CONTENU, cf. pdf_invoice_nums plus haut) est
    # tres probablement deja facture un mois precedent -- exclue automatiquement (confirme sur
    # le cas reel 634313590 de juin 2026).
    if pdf_invoice_nums and i_invoice_num is not None:
        # factures_fret contient TOUTES les factures vues initialement (y compris celles deja
        # exclues ci-dessus par la regle Droits&Taxes) -- exclure factures_droits_taxes pour
        # ne pas lister deux fois la meme facture dans des messages differents (elle est deja
        # sortie de "rows", donc deja absente de facon certaine, pas besoin de la re-signaler
        # comme "sans PDF").
        factures_sans_pdf = {nf for nf in factures_fret if nf not in pdf_invoice_nums} - factures_droits_taxes
        if factures_sans_pdf:
            avant = len(rows)
            rows = [r for r in rows if str(r[i_invoice_num] or "").strip() not in factures_sans_pdf]
            print(f"Facture(s) absente(s) des PDF portail fournis (probablement deja facturee(s) un mois precedent) exclue(s) : {sorted(factures_sans_pdf)} ({avant - len(rows)} ligne(s)) -- a confirmer avec le pole transport.")
            n = len(rows)

    # Mois cible = choix utilisateur (--period, source de verite -- meme pattern que
    # finaliser_ups.py, BUG TROUVE 2026-08-26 sur UPS : "--period n'etait jamais transmis au
    # finaliseur, Date validite tarif restait calculee par majorite, peu fiable"). Repli sur le
    # mois majoritaire si absent.
    # BUG TROUVE 2026-10-12 (signalement utilisateur, FedEx : "Date validite tarif" figee au
    # 01/08 au lieu du 01/09 attendu) : le CSV Shipment Detail "de septembre" contenait en
    # realite 3 mois de facturation melanges (verifie sur le vrai fichier : 2962 lignes aout,
    # 2376 lignes septembre, 696 lignes octobre) -- le mois MAJORITAIRE (aout, le plus de
    # lignes) n'est pas forcement le mois CIBLE choisi par l'utilisateur au moment de
    # l'upload. Le carrier Node (index.js, computeFinalizerArgs) ne transmettait jamais le
    # --period choisi dans le formulaire -- corrige en parallele, meme fix que UPS.
    import datetime as _dt
    from collections import Counter
    EXCEL_EPOCH = _dt.datetime(1899, 12, 30)
    date_validite_serial = None
    mois_cible = None
    if period:
        m = re.fullmatch(r"(\d{4})_(\d{2})", period)
        if m:
            mois_cible = f"{m.group(1)}{m.group(2)}"
            date_validite_serial = (_dt.datetime(int(m.group(1)), int(m.group(2)), 1) - EXCEL_EPOCH).days
    if mois_cible is None and (i_mois is not None or i_invoice_date is not None):
        compte = Counter()
        for r in rows:
            v = str(r[i_mois] or "").strip() if i_mois is not None else ""
            if not re.match(r"^\d{6}$", v) and i_invoice_date is not None:
                m = re.match(r"^(\d{1,2})/(\d{1,2})/(\d{4})$", str(r[i_invoice_date] or "").strip())
                if m:
                    v = f"{m.group(3)}{int(m.group(1)):02d}"
            if re.match(r"^\d{6}$", v):
                compte[v] += 1
        if compte:
            mois_maj = compte.most_common(1)[0][0]
            annee, mois = int(mois_maj[:4]), int(mois_maj[4:6])
            date_validite_serial = (_dt.datetime(annee, mois, 1) - EXCEL_EPOCH).days

    # E/P : export WMS "expeditions_brut" m/m-1 -- repli "particulier" si tracking absent
    # (meme defaut le plus sur que le carrier Node, cf. index.js).
    ep_map = load_brut_ep(brut_paths) if brut_paths else {}
    if not brut_paths:
        print("AVERTISSEMENT: aucun export 'expeditions_brut' fourni -> toutes les lignes classées 'particulier' par défaut (E/P).")
    n_ep_defaut = 0
    ep_values = []
    for r in rows:
        t = str(r[i_tracking] or "").strip() if i_tracking is not None else ""
        ep = ep_map.get(t)
        if not ep:
            ep = "particulier"
            n_ep_defaut += 1
        ep_values.append(ep)
    if n_ep_defaut:
        print(f"{n_ep_defaut} ligne(s) sans correspondance dans l'export WMS (E/P) — classée(s) 'particulier' par défaut.")

    data_brut = [[coerce(v) for v in r] for r in rows]
    # BUG TROUVE 2026-10-12 (signalement utilisateur : tous les trackings finissent par ".0"
    # dans le CSV livre) : coerce() convertit tout champ purement numerique en float -- correct
    # pour un montant, mais le "Numero de suivi" FedEx est un IDENTIFIANT (confirme "TOUJOURS
    # numerique ... jamais de zero de tete" plus bas), jamais une quantite. DEUX conversions
    # en nombre s'enchainent : (1) coerce() cote Python, (2) Excel lui-meme reconvertit une
    # chaine purement numerique en nombre a l'ecriture .Value= (verifie : ecrire "876739026110"
    # en texte Python donne malgre tout un float cote Excel). Un seul fix cote Python (sans
    # toucher coerce(), qui reste correct pour les vrais montants) ne suffit donc pas -- fix
    # retenu : prefixe d'une APOSTROPHE (convention Excel standard "force texte", confirmee
    # invisible a la lecture .Value), applique ICI (apres coerce(), donc jamais recoerce) sur
    # la SEULE colonne tracking (demande explicite : pas de changement sur les autres champs
    # numeriques non signales, ex. montants/poids qui doivent rester des nombres).
    if i_tracking is not None:
        for ri, r in enumerate(data_brut):
            raw = rows[ri][i_tracking]
            data_brut[ri][i_tracking] = ("'" + str(raw).strip()) if raw not in (None, "") else None
    # Garde-fou (2026-10-08) : une cellule Excel accepte au plus 32767 caracteres -- un champ
    # plus long (ex. un cas de corruption CSV non encore rencontre, meme categorie que le bug
    # de parsing corrige plus haut) ferait planter Excel (COM access violation) a l'ecriture
    # bulk ci-dessous, AVANT toute exception Python lisible. Tronque defensivement (valeur
    # visible mais signalee) plutot que de laisser Excel crasher silencieusement.
    n_tronque = 0
    for r in data_brut:
        for i, v in enumerate(r):
            if isinstance(v, str) and len(v) > 32000:
                r[i] = v[:32000] + " [TRONQUE]"
                n_tronque += 1
    if n_tronque:
        print(f"AVERTISSEMENT: {n_tronque} champ(s) CSV anormalement long(s) (>32000 caractères, probable corruption source) tronqué(s) pour éviter un plantage Excel -- à vérifier manuellement.")
    pdfs = pdf_infos_precoce
    pdf_total_du_par_facture = {p["numero_facture"]: p["total_du"] for p in pdfs if p["numero_facture"] and p["total_du"] is not None}

    import win32com.client as win32
    xlUp = -4162
    xlDatabase = 1
    xl = win32.DispatchEx("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    xl.AskToUpdateLinks = False
    try:
        wb = retry(lambda: xl.Workbooks.Open(os.path.abspath(sortie), UpdateLinks=0, ReadOnly=False))
        if wb is None:
            raise RuntimeError("Excel n'a pas pu ouvrir le fichier (déjà ouvert ? verrouillé ?)")

        # BUG TROUVE 2026-10-08 (signalement utilisateur, com_error -2147024882 = 0x8007000E =
        # E_OUTOFMEMORY cote COM, reproduit en direct sur le vrai CSV de septembre -- 2376
        # lignes x 66 colonnes SEULEMENT, aucun champ anormalement long -- plante systematiquement
        # au 3e paquet de 500 lignes, meme en CLI hors serveur/navigateur) : Calculation restait
        # en AUTOMATIQUE (jamais mis en Manuel ici, contrairement au finaliseur UPS deja corrige
        # le 2026-10-07 pour la meme classe de probleme) -- chaque ecriture .Value= declenchait un
        # recalcul COMPLET du classeur (formules A-E de "Shipment Detail" + TCD/Bilan clients qui
        # en dependent), le cout cumule explosant a mesure que la plage de donnees grandit. Fix
        # identique a UPS : Calculation=Manuel + ScreenUpdating=False, DEFINI APRES Workbooks.Open
        # (jamais avant -- meme piege COM deja documente sur UPS, "Impossible de definir la
        # propriete Calculation de la classe Application" si tente avant qu'un classeur soit ouvert).
        xlCalculationManual = -4135
        xl.Calculation = xlCalculationManual
        xl.ScreenUpdating = False

        # 1) "Shipment Detail" : purge + collage donnees brutes a partir de la colonne F
        #    (A-E = calculees : Montant/Clients/Categorie/E-P/Mode d'envoi).
        ws = wb.Sheets("Shipment Detail")
        if ws.AutoFilterMode:  # meme piege que Chronopost/TNT : AutoFilter actif -> ecritures ignorees
            ws.AutoFilterMode = False
        FIRST_RAW_COL = 6  # colonne F
        LAST_RAW_COL = FIRST_RAW_COL - 1 + len(header)
        LAST_CALC_COL = 5  # A->E
        oldLast = ws.Cells(ws.Rows.Count, FIRST_RAW_COL).End(xlUp).Row
        newLast = 1 + n
        maxLast = max(oldLast, newLast, 2)

        purgeUntil = maxLast + 200
        retry(lambda: ws.Range(ws.Cells(2, 1), ws.Cells(purgeUntil, max(LAST_RAW_COL, LAST_CALC_COL))).ClearContents())

        # "Numéro de suivi" : TOUJOURS numerique dans le brut reel FedEx (jamais de zero de
        # tete constate sur 3607 lignes juin 2026) -- pas de force NumberFormat texte ici,
        # contrairement a TNT (tracking 16 chiffres avec zeros de tete).
        # BUG TROUVE 2026-10-08 (CSV reel de septembre, ligne 1075 : "Reference Notes Line 1"
        # = "=.!+23") : la VRAIE cause du pywintypes.com_error ("Une exception s'est produite")
        # etait une chaine commencant par "=" (cf. coerce() plus haut, qui l'echappe desormais
        # d'une apostrophe) -- PAS une histoire de memoire/taille malgre le code d'erreur
        # trompeur. Diagnostique par recherche dichotomique sur le vrai fichier (confirme :
        # memes dimensions en donnees synthetiques = aucun probleme, seule la vraie ligne
        # plantait, isolee hors de tout le reste du pipeline). Le decoupage par PAQUETS
        # ci-dessous n'etait pas la correction necessaire, mais reste une securite utile
        # (meme principe que UPS) : si un futur cas similaire apparait, l'erreur touche moins
        # de lignes a la fois et reste facile a isoler.
        CHUNK_ROWS = 500
        for debut in range(0, len(data_brut), CHUNK_ROWS):
            paquet = data_brut[debut:debut + CHUNK_ROWS]
            r0 = 2 + debut
            r1 = r0 + len(paquet) - 1
            retry(lambda p=paquet, a=r0, b=r1: ws.Range(ws.Cells(a, FIRST_RAW_COL), ws.Cells(b, LAST_RAW_COL)).__setattr__("Value", p))
        if newLast < oldLast:
            retry(lambda: ws.Range(ws.Cells(newLast + 1, 1), ws.Cells(oldLast, max(LAST_RAW_COL, LAST_CALC_COL))).ClearContents())

        # D (E/P) ecrite en VALEUR LITERALE (calculee ci-dessus via l'export WMS) -- PAS une
        # formule, reproduit l'etat reel du fichier de juin 2026 (100% litteral sur 3607/3607
        # lignes, jamais de XLOOKUP vers un fichier externe). B (Clients) laissee vide
        # (saisie manuelle post-generation, meme regle transversale que BLS/Chronopost/
        # Colissimo). C (Categorie) laissee vide (purge suffit, toujours vide en reel).
        col_d_vals = [[v] for v in ep_values]
        retry(lambda: ws.Range(ws.Cells(2, 4), ws.Cells(newLast, 4)).__setattr__("Value", col_d_vals))

        col_service = col_letter(FIRST_RAW_COL - 1 + i_service) if i_service is not None else None
        col_bk = col_letter(FIRST_RAW_COL - 1 + i_fret) if i_fret is not None else None
        col_bn = col_letter(FIRST_RAW_COL - 1 + i_remise) if i_remise is not None else None

        if col_bk and col_bn:
            formula_a = f"={col_bk}{{row}}+{col_bn}{{row}}"
            retry(lambda t=formula_a: ws.Range(ws.Cells(2, 1), ws.Cells(newLast, 1)).__setattr__("Formula", [[t.format(row=r)] for r in range(2, newLast + 1)]))
        else:
            print("AVERTISSEMENT: 'Montant des frais de transport'/'remise' introuvable(s) -> colonne A (Montant) non reconstruite.")

        if col_service:
            formula_e = f'=IF({col_service}{{row}}="FedEx Priority","ST","FICP")'
            retry(lambda t=formula_e: ws.Range(ws.Cells(2, 5), ws.Cells(newLast, 5)).__setattr__("Formula", [[t.format(row=r)] for r in range(2, newLast + 1)]))
        else:
            print("AVERTISSEMENT: 'Description du service' introuvable -> colonne E (Mode d'envoi) non reconstruite.")

        ws.Range(ws.Cells(1, 1), ws.Cells(newLast, max(LAST_RAW_COL, LAST_CALC_COL))).AutoFilter()

        # 2) TCD / Bilan factures / Bilan clients : PivotCache redirige vers la vraie plage
        #    large de 'Shipment Detail' (cache du modele fige sur une plage etroite --
        #    'Shipment Detail'!C1:C70 ou C1:C2, meme piege que DPD/Geodis/Mondial Relay/
        #    Chronopost/TNT).
        newRange = ws.Range(ws.Cells(1, 1), ws.Cells(newLast, max(LAST_RAW_COL, LAST_CALC_COL)))
        for sheet_name in ("TCD", "Bilan factures", "Bilan clients"):
            wsPivot = wb.Sheets(sheet_name)
            for i in range(1, wsPivot.PivotTables().Count + 1):
                pt = wsPivot.PivotTables(i)
                src = str(pt.PivotCache().SourceData)
                if "Shipment Detail" not in src:
                    print(f"TCD '{pt.Name}' ({sheet_name}) : source '{src}' n'est pas 'Shipment Detail' -> config préservée.")
                    continue
                newCache = wb.PivotCaches().Create(SourceType=xlDatabase, SourceData=newRange)
                pt.ChangePivotCache(newCache)
                print(f"TCD '{pt.Name}' ({sheet_name}) : PivotCache redirigé vers {newRange.Address}.")
        wb.RefreshAll()
        try:
            xl.CalculateUntilAsyncQueriesDone()
        except Exception:
            pass
        xl.Calculate()

        # 3) "TCD" colonnes manuelles E/F (etendues jusqu'a la vraie derniere ligne native,
        #    colonne A = Shipment Tracking Number).
        wsTcd = wb.Sheets("TCD")
        lastTcd = wsTcd.Cells(wsTcd.Rows.Count, 1).End(xlUp).Row
        if lastTcd >= 2:
            formulas_tcd = {
                5: "=ROUNDUP(D{row}*0.453592,1)",  # E Poids arrondis sup (kg)
                6: "=IF(_xlfn.XLOOKUP(A{row},'Shipment Detail'!N:N,'Shipment Detail'!D:D)=\"entreprise\",\"E\",\"P\")",  # F E/P
            }
            for col_idx, tmpl in formulas_tcd.items():
                retry(lambda c=col_idx, t=tmpl: wsTcd.Range(wsTcd.Cells(2, c), wsTcd.Cells(lastTcd, c))
                      .__setattr__("Formula", [[t.format(row=r)] for r in range(2, lastTcd + 1)]))
            print(f"'TCD' : colonnes E/F étirées jusqu'à la ligne {lastTcd}.")

        # 4) "Bilan factures" colonne D (PDF) : remplie AUTOMATIQUEMENT avec le "Total dû EUR"
        #    extrait de chaque PDF fourni (regle utilisateur 2026-08-20) -- appariee par numero
        #    de facture (colonne A, valeur native du TCD) au lieu d'une saisie manuelle. Colonne
        #    E (Ecart=C-D) etiree jusqu'a la derniere ligne native (colonne A = Invoice Number).
        #    Facture sans PDF correspondant : colonne D laissee vide (saisie manuelle de repli).
        wsBf = wb.Sheets("Bilan factures")
        lastBf = wsBf.Cells(wsBf.Rows.Count, 1).End(xlUp).Row
        if lastBf >= 4:  # en-tete sur la ligne 3 (constate sur le modele de juin)
            n_pdf_ecrits = 0
            for r in range(4, lastBf + 1):
                # Colonne A = valeur NATIVE du TCD (RowField Invoice Number) -- win32com la
                # renvoie en float (ex. 634323394.0), jamais en texte, contrairement au numero
                # de facture extrait du PDF (string via regex) -- BUG TROUVE 2026-08-20 :
                # comparaison str("634323394.0") vs "634323394" echouait systematiquement
                # (0/10 factures appariees). Normalise en entier des deux cotes.
                rawFacture = wsBf.Cells(r, 1).Value
                try:
                    numFacture = str(int(rawFacture))
                except (TypeError, ValueError):
                    numFacture = str(rawFacture or "").strip()
                totalDu = pdf_total_du_par_facture.get(numFacture)
                if totalDu is not None:
                    wsBf.Cells(r, 4).Value = totalDu
                    n_pdf_ecrits += 1
            retry(lambda: wsBf.Range(wsBf.Cells(4, 5), wsBf.Cells(lastBf, 5))
                  .__setattr__("Formula", [[f"=C{r}-D{r}"] for r in range(4, lastBf + 1)]))
            print(f"'Bilan factures' : colonne D (PDF) renseignée automatiquement pour {n_pdf_ecrits}/{lastBf - 3} facture(s) (Total dû EUR extrait du PDF), colonne E (Ecart) étirée jusqu'à la ligne {lastBf}.")

        # 5) "Import ERP" : formules PAR LIGNE (XLOOKUP tracking, PAS de decalage fixe comme
        #    TNT) -- nombre de lignes = nombre de trackings du TCD.
        wsImp = wb.Sheets("Import ERP")
        if wsImp.AutoFilterMode:
            wsImp.AutoFilterMode = False
        oldLastImp = wsImp.Cells(wsImp.Rows.Count, 6).End(xlUp).Row
        LAST_COL_IMPORT = 22
        newLastImp = max(lastTcd, 2)

        formulas_import = {
            1: '=IF(OR(M{row}="FICP",M{row}="IE", M{row}="RE", M{row}="International", M{row}="Europe"),"FEDEX INTERNATIONAL","FEDEX FRANCE")',  # A Transporteur
            6: "=TCD!A{row}",  # F N° Tracking
            8: "=TCD!F{row}",  # H E/P
            9: "=TCD!B{row}",  # I Pays
            10: '=IF(COUNTIF(Zoning!B:B,I{row})=0,"inconnu",_xlfn.XLOOKUP(I{row},Zoning!B:B,Zoning!C:C))',  # J Zone
            11: "=_xlfn.XLOOKUP(F{row},'Shipment Detail'!N:N,'Shipment Detail'!AB:AB)",  # K Nbr Colis
            12: "=TCD!E{row}",  # L Poids
            13: '=IF(J{row}="France","ST",_xlfn.XLOOKUP(F{row},\'Shipment Detail\'!N:N,\'Shipment Detail\'!E:E,""))',  # M mode envoi
            14: '=IF(OR(I{row}="FR",I{row}="DE",I{row}="IT",I{row}="ES",I{row}="BE",I{row}="NL",I{row}="LU",I{row}="PT",I{row}="AT",I{row}="IE",I{row}="FI",I{row}="SE",I{row}="DK",I{row}="PL",I{row}="CZ",I{row}="SK",I{row}="HU",I{row}="RO",I{row}="BG",I{row}="HR",I{row}="SI",I{row}="EE",I{row}="LV",I{row}="LT",I{row}="GR",I{row}="CY",I{row}="MT"),0.2,0)',  # N TVA
            15: "=IF(_xlfn.XLOOKUP(F{row},'Shipment Detail'!N:N,'Shipment Detail'!BM:BM)=0,\"\",_xlfn.XLOOKUP(F{row},'Shipment Detail'!N:N,'Shipment Detail'!BM:BM))",  # O Droits et taxes
            18: "=IF(_xlfn.XLOOKUP(F{row},'Shipment Detail'!N:N,'Shipment Detail'!AV:AV)>60,10,\"\")",  # R Colis volumineux
            20: "=TCD!C{row}",  # T Frêt
        }
        for col_idx, tmpl in formulas_import.items():
            retry(lambda c=col_idx, t=tmpl: wsImp.Range(wsImp.Cells(2, c), wsImp.Cells(newLastImp, c))
                  .__setattr__("Formula", [[t.format(row=r)] for r in range(2, newLastImp + 1)]))

        # DateValidite (colonne B) : valeur ligne 2, puis "=B{n-1}" chaine (modele exact).
        if date_validite_serial is not None:
            wsImp.Cells(2, 2).Value = date_validite_serial
        else:
            print("AVERTISSEMENT: mois de facturation introuvable -> 'Import ERP'!B2 non mise à jour, reste celle du modèle.")
        if newLastImp > 2:
            retry(lambda: wsImp.Range(wsImp.Cells(3, 2), wsImp.Cells(newLastImp, 2))
                  .__setattr__("Formula", [["=B{prev}".format(prev=r - 1)] for r in range(3, newLastImp + 1)]))

        if newLastImp < oldLastImp:
            retry(lambda: wsImp.Range(wsImp.Cells(newLastImp + 1, 1), wsImp.Cells(oldLastImp, LAST_COL_IMPORT)).ClearContents())
        print(f"'Import ERP' : formules reconstruites jusqu'à la ligne {newLastImp}.")

        # 5bis) Supplements "surplus" (Zones eloignees Q/17, Colis volumineux R/18, Plus-value
        #       BtoC U/21) : ECRITS EN VALEUR par-dessus les formules ci-dessus (pas de formule
        #       Excel possible, la donnee n'existe que dans le texte du PDF) -- Colis volumineux
        #       ADDITIONNE le montant PDF a la valeur deja posee par la formule >60cm=10EUR
        #       (2 mecanismes independants, confirme pole transport 2026-08-20, cf. mail "Oubli
        #       facturation surplus Fedex janvier a mai 2026").
        if supplements_par_tracking and newLastImp >= 2:
            n_maj = 0
            for r in range(2, newLastImp + 1):
                tracking = str(wsImp.Cells(r, 6).Value or "").strip()  # F N° Tracking
                if tracking.endswith(".0"):
                    tracking = tracking[:-2]
                supp = supplements_par_tracking.get(tracking)
                if not supp:
                    continue
                if supp["ZonesEloignees"]:
                    wsImp.Cells(r, 17).Value = supp["ZonesEloignees"]
                if supp["ColisVolumineuxPdf"]:
                    existant = wsImp.Cells(r, 18).Value
                    existant = existant if isinstance(existant, (int, float)) else 0
                    wsImp.Cells(r, 18).Value = round(existant + supp["ColisVolumineuxPdf"], 2)
                if supp["PlusValueB2C"]:
                    wsImp.Cells(r, 21).Value = supp["PlusValueB2C"]
                n_maj += 1
            print(f"'Import ERP' : suppléments PDF (Zones éloignées/Colis volumineux/Plus-value BtoC) appliqués sur {n_maj} ligne(s) — cf. mail pôle transport 2026-08-20.")

        wsImp.Range(wsImp.Cells(1, 1), wsImp.Cells(max(newLastImp, 2), LAST_COL_IMPORT)).AutoFilter()
        xl.Calculate()

        # 6) Reconciliation PDF (info console -- colonne D "PDF" de 'Bilan factures' reste a
        #    la saisie manuelle, comme les autres transporteurs sans onglet de controle dedie).
        for p in pdfs:
            if p["total_du"] is not None:
                print(f"PDF {p['file']}{' (n° ' + p['numero_facture'] + ')' if p['numero_facture'] else ''} : "
                      f"Total dû = {p['total_du']:.2f} EUR (à comparer/coller en 'Bilan factures'!D).")
            else:
                print(f"PDF {p['file']} : 'Total dû' introuvable, ignoré.")

        xl.Calculate()
        retry(lambda: wb.Save())
        wb.Close(SaveChanges=True)

        # 7) Export "Import ERP" en VALEURS (pas en JS recalcule a part) : remontee pole
        # transport 2026-08-24 ("le fichier import CSV deconne alors que la feuille Import CSV
        # du fichier de facture est correcte -> copier-coller depuis le fichier de la facture,
        # coller en valeur"). BUG TROUVE 2026-08-24 (cf. finaliser_colissimo.py, meme fix) :
        # lire Range.Value AVANT Save()/Close(), meme juste apres un Calculate(), peut renvoyer
        # des cellules dans un etat transitoire (valeur d'erreur COM brute, ou une ligne vide)
        # qui ne correspond PAS a ce que contient le fichier une fois sauvegarde sur disque.
        # Fix : fermer le classeur (deja fait ci-dessus), le ROUVRIR en lecture, et lire
        # seulement alors -- meme etat que ce que le pole transport voit en ouvrant le fichier.
        wbRead = None
        try:
            # wbRead ferme dans le finally imbrique ci-dessous, quoi qu'il arrive -- cf.
            # finaliser_colissimo.py (meme bug/fix, 2026-08-24) : une exception entre
            # l'ouverture et la fermeture laissait un EXCEL.EXE orphelin, verrouillant le
            # fichier de sortie pour toute regeneration suivante du meme mois.
            wbRead = retry(lambda: xl.Workbooks.Open(os.path.abspath(sortie), UpdateLinks=0, ReadOnly=True))
            wsImpRead = wbRead.Sheets("Import ERP")
            impLast = newLastImp
            if impLast >= 2:
                values = retry(lambda: wsImpRead.Range(wsImpRead.Cells(2, 1), wsImpRead.Cells(impLast, LAST_COL_IMPORT)).Value)
                # BUG TROUVE 2026-08-24 : le TCD (PivotTable, RowField "Shipment Tracking
                # Number") contient un item "(blank)" LEGITIME quand le CSV brut a au moins 1
                # ligne sans tracking (ex. ajustement/note de credit sans envoi associe) -> une
                # ligne "Import ERP" a Tracking vide se glissait dans l'export. Le carrier Node
                # (index.js) exclut deja ce cas explicitement (`if (!tracking) continue;`) --
                # meme filtre applique ici pour rester coherent (pas une vraie ligne facturable).
                # Filet de securite residuel : une cellule encore en erreur COM (marshaling
                # ponctuel, type int au lieu de str/float/datetime/None) est ecrite vide plutot
                # que comme un nombre absurde.
                values = [row for row in values if str(row[5]).strip() and str(row[5]).strip() != "(vide)"]
                n_err = 0
                def clean(v):
                    nonlocal n_err
                    if isinstance(v, int) and not isinstance(v, bool) and v < -1000000:
                        n_err += 1
                        return ""
                    return "" if v is None else v
                export_path = os.path.splitext(sortie)[0] + "_import_valeurs.csv"
                with open(export_path, "w", encoding="utf-8-sig", newline="") as f:
                    w = csv.writer(f, delimiter=";")
                    for row in values:
                        w.writerow([clean(v) for v in row])
                if n_err:
                    print(f"AVERTISSEMENT: {n_err} cellule(s) en erreur COM lors de l'export Import ERP (valeurs) -- laissees vides, a verifier.")
                print(f"EXPORT_IMPORT_VALEURS:{export_path}")
        except Exception as e:
            print("Export Import ERP (valeurs) ignore :", e)
        finally:
            if wbRead is not None:
                try:
                    wbRead.Close(SaveChanges=False)
                except Exception:
                    pass

        print(f"OK -> {sortie}")
    finally:
        try:
            xl.Quit()
        except Exception:
            pass


if __name__ == "__main__":
    main()
