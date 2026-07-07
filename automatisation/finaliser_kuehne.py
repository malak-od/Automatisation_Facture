#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
FINALISEUR KUEHNE — produit le classeur "AAAA_MM_Facture Kuehne.xlsx" A L'IDENTIQUE
du fichier fait a la main : memes FEUILLES, memes FORMULES, memes TCD (pivots),
meme MISE EN FORME. Principe = exactement le geste manuel :
  1. on part du fichier existant comme MODELE (copie),
  2. on remplace uniquement les DONNEES BRUTES dans l'onglet "Fichier Kuehne",
  3. on etend les formules et on rafraichit les tableaux croises,
  4. Excel recalcule tout.

Necessite : Windows + Excel installe + pywin32.
Usage :
  python finaliser_kuehne.py "<modele.xlsx>" "<sortie.xlsx>" "<dossier_csv>"
  python finaliser_kuehne.py "<modele.xlsx>" "<sortie.xlsx>" "<csv1>" "<csv2>" ...
"""
import sys, os, shutil, csv, glob, re

FIRST_RAW_COL = 13  # colonne M : debut des donnees brutes CSV dans l'onglet "Fichier Kuehne"

def coerce(s):
    """Nombre si la cellule est purement numerique (virgule decimale), sinon texte, sinon vide."""
    if s is None or s == "":
        return None
    if re.fullmatch(r"-?\d+(,\d+)?", s):
        return float(s.replace(",", "."))
    return s

def load_rows(csv_paths):
    hdr, rows = None, []
    for f in csv_paths:
        with open(f, encoding="latin-1", newline="") as fh:
            data = list(csv.reader(fh, delimiter=";"))
        hdr = data[0]
        for r in data[1:]:
            if len(r) < len(hdr):
                r = r + [""] * (len(hdr) - len(r))
            rows.append(r[:len(hdr)])
    return hdr, rows

def gazole_mess_rate(csv_paths):
    """% de la ligne 'TAXE SURCOUT GAZOLE MESS' (ex '... : 11,70 %') -> 0.117. None si absente."""
    for p in csv_paths:
        with open(p, encoding="latin-1", newline="") as fh:
            for row in csv.reader(fh, delimiter=";"):
                for cell in row:
                    if "GAZOLE MESS" in cell.upper():
                        m = re.search(r"(\d+),(\d+)\s*%", cell)
                        if m:
                            return float(m.group(1) + "." + m.group(2)) / 100
    return None

def pdf_taxables(pdf_paths):
    """Montant taxable de chaque PDF (via TTC=taxable*1.2), trie DECROISSANT.
    pdf1 = le plus gros (facture initiale), pdf2 = souffrance, etc."""
    try:
        import pypdf
    except ImportError:
        return []
    out = []
    for p in pdf_paths:
        try:
            txt = pypdf.PdfReader(p).pages[0].extract_text() or ""
        except Exception:
            continue
        nums = [float(x.replace(" ", "").replace("\xa0", "").replace(",", "."))
                for x in re.findall(r"\d[\d \xa0]*,\d{2}", txt)]
        best = None
        for t in nums:
            if t <= 0:
                continue
            if any(abs(v - round(t * 1.2, 2)) < 0.03 for v in nums) and any(abs(v - round(t * 0.2, 2)) < 0.03 for v in nums):
                if best is None or t > best:
                    best = t
        if best:
            out.append(best)
    out.sort(reverse=True)
    return out

def fill_reconciliation(wb, csv_paths, pdf_paths):
    """Onglet TCD : taxe gazole reelle (Z de 'Taxe gazole reelle') + pdf 1..4 (Z de 'pdf n')."""
    tcd = wb.Sheets("TCD")
    yvals = tcd.Range("Y1:Y40").Value  # etiquettes en colonne Y

    def row_of(pred):
        for i, rowv in enumerate(yvals, start=1):
            v = rowv[0] if isinstance(rowv, (tuple, list)) else rowv
            if v and pred(str(v).strip().lower()):
                return i
        return None

    r = row_of(lambda s: s.startswith("taxe gazole"))
    if r:
        tcd.Cells(r, 26).Value = gazole_mess_rate(csv_paths)  # colonne Z = 26

    taxables = pdf_taxables(pdf_paths)
    for k in range(1, 5):
        rr = row_of(lambda s, k=k: s == "pdf %d" % k)
        if rr:
            tcd.Cells(rr, 26).Value = taxables[k - 1] if k - 1 < len(taxables) else None
    print(f"Reconciliation TCD : gazole={gazole_mess_rate(csv_paths)}, pdf taxables={taxables}")

def parse_args(rest):
    """Separe les CSV et PDF (flags --csv / --pdf). 3e arg seul = dossier (glob FcCSV*)."""
    csv_paths, pdf_paths, cur = [], [], None
    for a in rest:
        if a == "--csv":
            cur = csv_paths
        elif a == "--pdf":
            cur = pdf_paths
        elif cur is None:
            csv_paths.append(a)
        else:
            cur.append(a)
    if len(csv_paths) == 1 and os.path.isdir(csv_paths[0]):
        csv_paths = sorted(glob.glob(os.path.join(csv_paths[0], "FcCSV*.csv")))
    return csv_paths, pdf_paths

def main():
    modele, sortie = sys.argv[1], sys.argv[2]
    csv_paths, pdf_paths = parse_args(sys.argv[3:])
    shutil.copyfile(modele, sortie)                 # on ne touche JAMAIS au modele
    hdr, rows = load_rows(csv_paths)
    n, ncol = len(rows), len(hdr)
    data = [[coerce(v) for v in r] for r in rows]   # bloc 2D a coller
    print(f"CSV: {n} lignes x {ncol} colonnes")

    import win32com.client as win32
    xlUp = -4162
    xl = win32.DispatchEx("Excel.Application")  # instance Excel DEDIEE (n'interfere pas avec la tienne)
    xl.Visible = False
    xl.DisplayAlerts = False
    xl.AskToUpdateLinks = False
    try:
        wb = xl.Workbooks.Open(os.path.abspath(sortie), UpdateLinks=0, ReadOnly=False)
        if wb is None:
            raise RuntimeError("Excel n'a pas pu ouvrir le fichier (deja ouvert ? verrouille ?)")
        ws = wb.Sheets("Fichier Kuehne")
        oldLast = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row  # derniere ligne via col B (Tracking)
        newLast = 1 + n
        lastRawCol = FIRST_RAW_COL + ncol - 1
        print(f"Onglet Fichier Kuehne : {oldLast-1} anciennes lignes -> {n} nouvelles")

        # 1) purge des anciennes DONNEES BRUTES (on garde les formules A..L de la ligne 2 comme modele)
        if oldLast >= 2:
            ws.Range(ws.Cells(2, FIRST_RAW_COL), ws.Cells(oldLast, lastRawCol)).ClearContents()
        # 2) coller le bloc de donnees brutes en un seul appel (rapide)
        ws.Range(ws.Cells(2, FIRST_RAW_COL), ws.Cells(newLast, lastRawCol)).Value = data
        # 3) etendre les formules calculees A..L jusqu'a la derniere ligne
        ws.Range(ws.Cells(2, 1), ws.Cells(newLast, 12)).FillDown()
        # 4) si moins de lignes que le mois precedent : effacer le surplus
        if newLast < oldLast:
            ws.Range(ws.Cells(newLast + 1, 1), ws.Cells(oldLast, lastRawCol)).ClearContents()

        # 5) rafraichir les tableaux croises (TCD, Bilan clients)
        wb.RefreshAll()
        try:
            xl.CalculateUntilAsyncQueriesDone()
        except Exception:
            pass
        # 6) reconciliation TCD : pdf 1/2 (taxable) + taxe gazole reelle (% ligne MESS du CSV)
        try:
            fill_reconciliation(wb, csv_paths, pdf_paths)
        except Exception as e:
            print("Reconciliation TCD ignoree :", e)
        xl.Calculate()  # recalcule total pdf / ecart / controle import

        wb.Save()
        wb.Close(SaveChanges=True)
        print("OK -> " + sortie)
    finally:
        xl.Quit()

if __name__ == "__main__":
    main()
