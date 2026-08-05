#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
FINALISEUR GEODIS — produit "AAAA_MM_Facture Geodis.xlsx" A L'IDENTIQUE du
fichier fait a la main (memes feuilles : Factures Geodis, Bilan Factures,
Import CSV, Bilan Client ; memes formules, memes TCD/pivots, meme mise en
forme), en partant du fichier existant comme MODELE et en n'y remplacant que
les DONNEES BRUTES (colonne N "N pi�ce" et suivantes dans "Factures Geodis").

Le fichier d'entree peut etre le CSV/XLSX brut Geodis (entete en ligne 1) OU
un rapport client type "Facturation_client.xlsx" (memes colonnes mais entete
decalee de quelques lignes de titre) -> l'entete est detectee dynamiquement
(recherche de "N recepisse"), comme cote Node (src/carriers/geodis/index.js).

ATTENTION (cf. FACTURATION EXCEL.pdf p.3) : les formules de "Factures Geodis"
referencent des POSITIONS de colonnes fixes (ex. =CS2+CT2+CU2), pas des noms.
Si les colonnes du fichier recu ne sont pas EXACTEMENT dans le meme ordre que
le modele, on arrete avec une erreur explicite plutot que de facturer sur les
mauvaises colonnes en silence.

Necessite : Windows + Excel + pywin32 + openpyxl.
Usage :
  python finaliser_geodis.py "<modele.xlsx>" "<sortie.xlsx>" "<entree1>" [<entree2>...] [--pdf-taxable N] [--frais-gestion N]
"""
import sys, os, shutil, re, csv

FIRST_RAW_COL = 14  # colonne N : debut des donnees brutes Geodis dans "Factures Geodis"
LAST_FORMULA_COL_FACTURES = 13  # colonnes A->M = postes calcules (formules a etendre)
LAST_COL_IMPORT_CSV = 23        # Import CSV : colonnes A->W (Gazole/NbColis compris, meme si vides)


def normalize_header(h):
    return re.sub(r"\s+", " ", str(h or "")).strip()


def compare_key(h):
    """Cle de comparaison ROBUSTE aux variantes cosmetiques vues entre exports
    Geodis (ex. 'RV tel,' vs 'RV tel.', accents) : on ne veut PAS ignorer un
    vrai renommage/decalage de colonne, juste la ponctuation/accentuation."""
    s = normalize_header(h).lower()
    s = s.translate(str.maketrans("éèêëàâäùûüôöîïç", "eeeeaaauuuooiic"))
    s = re.sub(r"[.,]", "", s)
    return re.sub(r"\s+", " ", s).strip()


def find_header_row(rows):
    """Cherche, dans les 15 premieres lignes, celle qui contient une colonne
    ressemblant a 'N recepisse' (insensible aux accents/casse)."""
    for i, row in enumerate(rows[:15]):
        for v in row:
            s = normalize_header(v).lower()
            if "recepiss" in s or "récépiss" in s:
                return i
    return None


def read_csv_rows(path):
    with open(path, encoding="latin-1", newline="") as f:
        return list(csv.reader(f, delimiter=";"))


def is_xlsx(path):
    """Detection par contenu (magic bytes ZIP), PAS par extension : les fichiers
    uploades (multer) arrivent sans extension du tout."""
    import zipfile
    try:
        return zipfile.is_zipfile(path)
    except Exception:
        return False


def read_xlsx_rows(path):
    """Cherche, sur TOUTES les feuilles, celle qui contient l'entete Geodis
    (ex. Facturation_client.xlsx : feuille 'Detail', entete en ligne 4).
    openpyxl refuse d'ouvrir un fichier sans extension .xlsx (verifiee sur le
    NOM, pas le contenu) -> copie temporaire renommee si besoin."""
    import openpyxl, tempfile
    load_path = path
    tmp_path = None
    if not path.lower().endswith((".xlsx", ".xlsm", ".xltx", ".xltm")):
        tmp = tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False)
        tmp.close()
        shutil.copyfile(path, tmp.name)
        load_path = tmp_path = tmp.name
    try:
        wb = openpyxl.load_workbook(load_path, data_only=True)
        for sheet in wb.sheetnames:
            ws = wb[sheet]
            rows = [[c for c in row] for row in ws.iter_rows(values_only=True)]
            hidx = find_header_row(rows)
            if hidx is not None:
                wb.close()
                return rows[hidx:]
        wb.close()
        return None
    finally:
        if tmp_path:
            try:
                os.unlink(tmp_path)
            except OSError:
                pass


def read_input(paths):
    """Concatene toutes les sources -> (header, data_rows). Format detecte par
    CONTENU (is_xlsx), jamais par extension (absente sur les fichiers uploades)."""
    header = None
    all_rows = []
    for p in paths:
        as_xlsx = is_xlsx(p)
        rows = read_xlsx_rows(p) if as_xlsx else read_csv_rows(p)
        if rows is None:
            raise RuntimeError(f"Format non reconnu (aucune feuille avec 'N° récépissé') : {p}")
        hidx = 0  # read_xlsx_rows renvoie deja a partir de l'entete
        if not as_xlsx:
            hidx = find_header_row(rows)
            if hidx is None:
                raise RuntimeError(f"Format non reconnu (aucune ligne d'entete avec 'N° récépissé') : {p}")
            rows = rows[hidx:]
        if header is None:
            header = [normalize_header(h) for h in rows[0]]
        data = [r for r in rows[1:] if any(v not in (None, "") for v in r)]
        all_rows.extend(data)
    return header, all_rows


def coerce(v):
    """Nombre si la valeur est un nombre francais valide (virgule ou point),
    sinon la valeur telle quelle (texte/date/None).
    Exception : un entier avec un zero de tete (ex. '000093630935') est une
    reference/code, pas une quantite -- le convertir perdrait les zeros de
    tete de maniere irreversible (float n'a pas de notion de zero non
    significatif conserve). Reste texte, comme dans le classeur fait a la
    main (colonne Ref.1)."""
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
    return v


def check_columns(header, modele_path):
    """Compare l'entete recue aux colonnes attendues par le modele (a partir
    de FIRST_RAW_COL) : meme nombre, memes noms, MEME ORDRE."""
    import openpyxl
    wb = openpyxl.load_workbook(modele_path, data_only=True)
    ws = wb["Factures Geodis"]
    expected = [normalize_header(ws.cell(row=1, column=c).value) for c in range(FIRST_RAW_COL, ws.max_column + 1)]
    wb.close()
    got = header[:len(expected)]
    if [compare_key(g) for g in got] != [compare_key(e) for e in expected]:
        diffs = [f"  colonne {i + 1} : attendu {e!r}, reçu {g!r}"
                 for i, (e, g) in enumerate(zip(expected, got)) if compare_key(e) != compare_key(g)]
        extra = ""
        if len(header) != len(expected):
            extra = f"\n  (nombre de colonnes : attendu {len(expected)}, reçu {len(header)})"
        raise RuntimeError(
            "Colonnes du fichier d'entrée différentes du modèle (FACTURATION EXCEL.pdf p.3 : "
            "« attention quand copie/colle vérifier même nombre de colonnes et le nom des colonnes ») :\n"
            + "\n".join(diffs[:10]) + extra
        )
    return expected


def extract_pdf_totals(pdf_paths):
    """Facture PDF Geodis : 'Montant Taxable HT' suivi de montants 'X,XX EUR',
    et une ligne dediee 'Frais de gestion de compte X,XX' (cf. adaptateur Node,
    meme logique -> voir src/carriers/geodis/index.js pour le detail du format).
    pypdf garde les espaces INSECABLES (\\xa0) du PDF entre les mots des libelles
    (ex. 'Montant\\xa0Taxable\\xa0HT') alors que pdf-parse (Node) les normalise en
    espaces simples -> on les normalise nous-memes avant toute recherche, sinon
    le libelle ne matche jamais et la reconciliation echoue silencieusement."""
    try:
        import pypdf
    except ImportError:
        return None, None
    taxable, frais = None, None
    for p in pdf_paths:
        try:
            text = "\n".join((page.extract_text() or "") for page in pypdf.PdfReader(p).pages)
        except Exception:
            continue
        text = text.replace("\xa0", " ")
        iTaxable = text.find("Montant Taxable HT")
        if iTaxable != -1:
            amounts = re.findall(r"([\d\s]+,\d{2})\s*EUR", text[iTaxable:])
            if amounts:
                taxable = float(amounts[0].replace(" ", "").replace(",", "."))
        m = re.search(r"Frais de gestion de compte\s*([\d\s]+,\d{2})", text)
        if m:
            frais = float(m.group(1).replace(" ", "").replace(",", "."))
    return taxable, frais


def fill_reconciliation(wb, pdf_taxable, frais_gestion):
    """Onglet Bilan Factures (TCD) : la colonne 'Frais de gestion' (poste sans
    tracking) et 'PDF' (montant taxable HT de la facture) sont saisies a la
    main dans le fichier fait a la main -> on les remplit automatiquement si
    fournies (deja extraites cote Node via pdf-parse).
    Les lignes du pivot demarrent APRES l'entete 'Etiquettes de lignes' (des
    libelles libres au-dessus, ex. 'Frais de tenue de compte 28,5EUR/mois', ne
    font pas partie du tableau -> les compter comme facture fausserait le
    garde-fou 'plusieurs factures' juste en dessous)."""
    ws = wb.Sheets("Bilan Factures")
    last = ws.Cells(ws.Rows.Count, 1).End(-4162).Row
    header_row = None
    for r in range(1, last + 1):
        if str(ws.Cells(r, 1).Value or "").strip() == "Étiquettes de lignes":
            header_row = r
            break
    invoice_rows = []
    blank_row = None
    for r in range((header_row or 0) + 1, last + 1):
        label = str(ws.Cells(r, 1).Value or "").strip()
        if label in ("", "Total général"):
            continue
        if label == "(vide)":
            blank_row = r
        else:
            invoice_rows.append(r)
    if frais_gestion is not None and blank_row:
        ws.Cells(blank_row, 3).Value = frais_gestion  # colonne C = Frais de gestion
    if pdf_taxable is not None:
        if len(invoice_rows) == 1:
            ws.Cells(invoice_rows[0], 4).Value = pdf_taxable  # colonne D = PDF
        else:
            print(f"AVERTISSEMENT: {len(invoice_rows)} ligne(s) facture dans Bilan Factures "
                  f"(pas 1 seule) -> montant PDF ({pdf_taxable}) non affecté automatiquement, à saisir à la main.")


def parse_args(argv):
    """<modele> <sortie> <entree1> [<entree2>...] [--pdf <pdf1> [<pdf2>...]]"""
    modele, sortie = argv[1], argv[2]
    inputs, pdfs, cur = [], [], "in"
    for a in argv[3:]:
        if a == "--pdf":
            cur = "pdf"
        elif cur == "pdf":
            pdfs.append(a)
        else:
            inputs.append(a)
    return modele, sortie, inputs, pdfs


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
    modele, sortie, inputs, pdf_paths = parse_args(sys.argv)
    shutil.copyfile(modele, sortie)  # on ne touche JAMAIS au modele

    header, rows = read_input(inputs)
    check_columns(header, modele)
    n = len(rows)
    ncol = len(header)
    data = [[coerce(v) for v in (list(r) + [None] * ncol)[:ncol]] for r in rows]
    print(f"Entrée : {n} lignes x {ncol} colonnes")

    pdf_taxable, frais_gestion = extract_pdf_totals(pdf_paths) if pdf_paths else (None, None)

    import win32com.client as win32
    xlUp = -4162
    xl = win32.DispatchEx("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    xl.AskToUpdateLinks = False
    try:
        wb = retry(lambda: xl.Workbooks.Open(os.path.abspath(sortie), UpdateLinks=0, ReadOnly=False))
        if wb is None:
            raise RuntimeError("Excel n'a pas pu ouvrir le fichier (déjà ouvert ? verrouillé ?)")

        # 1) Feuille "Factures Geodis" : purge + collage des données brutes (colonne N+)
        ws = wb.Sheets("Factures Geodis")
        lastCol = FIRST_RAW_COL + ncol - 1
        oldLast = ws.Cells(ws.Rows.Count, FIRST_RAW_COL).End(xlUp).Row
        newLast = 1 + n
        maxLast = max(oldLast, newLast)
        retry(lambda: ws.Range(ws.Cells(2, FIRST_RAW_COL), ws.Cells(maxLast, lastCol)).ClearContents())
        retry(lambda: ws.Range(ws.Cells(2, FIRST_RAW_COL), ws.Cells(newLast, lastCol)).__setattr__("Value", data))
        # 2) Étendre les formules des 13 postes calculés (colonnes A->M)
        retry(lambda: ws.Range(ws.Cells(2, 1), ws.Cells(newLast, LAST_FORMULA_COL_FACTURES)).FillDown())
        if newLast < oldLast:
            retry(lambda: ws.Range(ws.Cells(newLast + 1, 1), ws.Cells(oldLast, lastCol)).ClearContents())

        # 3) Onglet "Import CSV" : formules cross-feuille (référencent 'Factures Geodis'
        #    ligne à ligne) -> étendre pareil, sur le même nombre de lignes.
        wsImp = wb.Sheets("Import CSV")
        oldLastImp = wsImp.Cells(wsImp.Rows.Count, 6).End(xlUp).Row  # col F = N Tracking
        if newLast > 2:
            retry(lambda: wsImp.Range(wsImp.Cells(2, 1), wsImp.Cells(newLast, LAST_COL_IMPORT_CSV)).FillDown())
        if newLast < oldLastImp:
            retry(lambda: wsImp.Range(wsImp.Cells(newLast + 1, 1), wsImp.Cells(oldLastImp, LAST_COL_IMPORT_CSV)).ClearContents())

        # 4) Rafraîchir les TCD (Bilan Factures, Bilan Client)
        wb.RefreshAll()
        try:
            xl.CalculateUntilAsyncQueriesDone()
        except Exception:
            pass

        # 5) Réconciliation (Frais de gestion / PDF), si fournis
        if pdf_taxable is not None or frais_gestion is not None:
            try:
                fill_reconciliation(wb, pdf_taxable, frais_gestion)
            except Exception as e:
                print("Réconciliation Bilan Factures ignorée :", e)

        xl.Calculate()
        retry(lambda: wb.Save())
        wb.Close(SaveChanges=True)
        print(f"OK -> {sortie}")
    finally:
        try:
            xl.Quit()
        except Exception:
            pass


if __name__ == "__main__":
    main()
