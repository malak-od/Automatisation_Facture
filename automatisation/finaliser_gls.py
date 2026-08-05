#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
FINALISEUR GLS — produit "AAAA_MM_Facture GLS.xlsx" A L'IDENTIQUE du fichier fait
a la main (memes feuilles, memes formules, memes TCD/pivots, meme mise en forme),
en partant du fichier final comme MODELE et en n'y remplacant que les donnees.

Classeur reel (11 feuilles) :
  PA GLS, Zoning, Poids, Comptes GLS, Categories : tables de reference figees, pas
    touchees (Poids/Comptes GLS sont eux-memes des TCD sources sur Facture GLS,
    recalcules par RefreshAll()).
  Facture GLS : 1 ligne = 1 CHARGE brute du CSV BCF (plusieurs lignes par colis :
    Fret, Fret-avise, surcharges...). Colonnes A-E = FORMULES (Categorie, Total HT,
    Gazole, Total HT hors gazole, Poids) ; colonnes F+ = donnees brutes du CSV.
  TCD : tableau croise dynamique (source = Facture GLS!A:BD), 1 ligne par Numero de
    colis, colonnes E-L = les 9 postes GLS internes (Adresse, Colis volumineux,
    Frais de gestion, Fret, Fret-avise, Fret-retour, Zone eloignee, ZZ_Ramasse
    unicolis, (vide)). Colonnes P-S = reconciliation (Total/Gazole/Hors gazole/Ecart).
  Bilan factures / Bilan clients : TCD sources sur Facture GLS / TCD. "Bilan factures"
    = 1 ligne par n. de FACTURE GLS (colonne A du pivot), col B = Somme de Total HT,
    col C = TTC (formule =B*1,2, calcul theorique). Reconciliation PDF (video process,
    verification VISUELLE : TTC calcule compare au "Montant T.T.C." de la facture PDF) :
    on colle le TTC extrait du PDF en colonne E (le n. de facture du nom de fichier
    PDF sert a apparier la bonne ligne) + l'ecart en colonne F.
  Import csv : 1 ligne par colis (=TCD!D{n+1}), formules XLOOKUP vers Facture GLS/
    Zoning/Poids/TCD -- son nombre de lignes depend du TCD (donc du nombre de colis
    UNIQUES), pas du nombre de lignes CSV brutes -> etendu APRES le RefreshAll().
  Avoir : vide (juste l'entete), copie telle quelle.

Necessite : Windows + Excel + pywin32 + pypdf (reconciliation PDF, optionnelle).
Usage : python finaliser_gls.py "<modele.xlsx>" "<sortie.xlsx>" <csv1> [<csv2> ...] [--pdf <pdf1> [...]]
"""
import sys, os, shutil, csv, glob, re

FIRST_RAW_COL = 6   # colonne F : debut des donnees brutes CSV dans "Facture GLS" (A-E = formules)
LAST_FORMULA_COL = 5  # colonne E : derniere formule calculee de "Facture GLS"
LAST_COL_IMPORT_CSV = 22  # colonne V : derniere colonne utile de "Import csv"

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
        with open(f, encoding="utf-8", newline="") as fh:
            data = list(csv.reader(fh, delimiter=";"))
        hdr = data[0]
        for r in data[1:]:
            if len(r) < len(hdr):
                r = r + [""] * (len(hdr) - len(r))
            rows.append(r[:len(hdr)])
    return hdr, rows

def extract_pdf_ttc(pdf_path):
    """Montant T.T.C. de la facture GLS (libelle explicite dans le PDF, cf. video process
    p.ex. '3\xa0683,04Montant T.T.C.:') -- prend la 1ere occurrence trouvee (recap en
    page 0, identique au recap final)."""
    try:
        import pypdf
    except ImportError:
        return None
    try:
        r = pypdf.PdfReader(pdf_path)
    except Exception:
        return None
    for page in r.pages:
        txt = page.extract_text() or ""
        m = re.search(r"(?:^|\n)(\d[\d\s\xa0]*,\d{2})\s*Montant T\.T\.C\.:", txt)
        if m:
            return float(m.group(1).replace("\xa0", "").replace(" ", "").replace(",", "."))
    return None

def pdf_facture_numero(pdf_path):
    """N. de facture GLS (colonne 'Facture' de Facture GLS / colonne A de Bilan
    factures) : 2e groupe de chiffres du nom de fichier (Facture_<payeur>_<facture>_...pdf)."""
    m = re.search(r"^Facture_(\d+)_(\d+)_", os.path.basename(pdf_path))
    return int(m.group(2)) if m else None

def fill_reconciliation(wb, pdf_paths):
    """Onglet 'Bilan factures' : colle le TTC extrait de chaque PDF (colonne E) en
    face de sa ligne (appariee par n. de facture, colonne A), + ecart en colonne F.
    Ecart = TTC THEORIQUE (B*1,2, comme la formule video =Bn*1,2) moins le TTC PDF --
    PAS colonne C, qui dans le fichier de reference actuel contient une valeur figee
    (semble etre une saisie manuelle, egale au HT et non au TTC calcule)."""
    bf = wb.Sheets("Bilan factures")
    bf.Cells(3, 5).Value = "TTC (PDF)"
    bf.Cells(3, 6).Value = "Ecart"
    lastRow = bf.Cells(bf.Rows.Count, 1).End(-4162).Row  # xlUp
    numeros = {}
    for r in range(4, lastRow + 1):
        v = bf.Cells(r, 1).Value
        if isinstance(v, (int, float)):
            numeros[int(v)] = r
    matched = 0
    for p in pdf_paths:
        num = pdf_facture_numero(p)
        ttc = extract_pdf_ttc(p)
        if num is None or ttc is None:
            print(f"Reconciliation GLS : PDF ignore (n. facture ou TTC introuvable) -> {os.path.basename(p)}")
            continue
        row = numeros.get(num)
        if row is None:
            print(f"Reconciliation GLS : facture {num} (PDF {os.path.basename(p)}) absente de 'Bilan factures'")
            continue
        bf.Cells(row, 5).Value = ttc
        bf.Cells(row, 6).Formula = f"=(B{row}*1.2)-E{row}"
        matched += 1
        print(f"Reconciliation GLS : facture {num} -> TTC PDF={ttc}")
    print(f"Reconciliation GLS : {matched}/{len(pdf_paths)} PDF apparies")

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
    """<csv1> [<csv2> ...] [--pdf <pdf1> [...]] (ou un dossier seul en 1er groupe -> glob BCF_*.csv)."""
    csv_paths, pdf_paths, cur = [], [], "csv"
    for a in rest:
        if a == "--pdf":
            cur = "pdf"
        elif cur == "csv":
            csv_paths.append(a)
        else:
            pdf_paths.append(a)
    if len(csv_paths) == 1 and os.path.isdir(csv_paths[0]):
        csv_paths = sorted(glob.glob(os.path.join(csv_paths[0], "BCF_*.csv")))
    return csv_paths, pdf_paths

def main():
    modele, sortie = sys.argv[1], sys.argv[2]
    csv_paths, pdf_paths = parse_args(sys.argv[3:])
    shutil.copyfile(modele, sortie)  # on ne touche JAMAIS au modele
    hdr, rows = load_rows(csv_paths)
    n, ncol = len(rows), len(hdr)
    data = [[coerce(v) for v in r] for r in rows]  # bloc 2D a coller
    print(f"CSV BCF : {n} lignes x {ncol} colonnes")

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

        # ---- 1) Facture GLS : purge + collage des donnees brutes + formules A-E ----
        fg = wb.Sheets("Facture GLS")
        oldLast = fg.Cells(fg.Rows.Count, FIRST_RAW_COL + 21).End(xlUp).Row  # via "Numero de colis" (AB = F+21)
        newLast = 1 + n
        lastRawCol = FIRST_RAW_COL + ncol - 1

        if oldLast >= 2:
            retry(lambda: fg.Range(fg.Cells(2, FIRST_RAW_COL), fg.Cells(oldLast, lastRawCol)).ClearContents())
        retry(lambda: setattr(fg.Range(fg.Cells(2, FIRST_RAW_COL), fg.Cells(newLast, lastRawCol)), "Value", data))
        retry(lambda: fg.Range(fg.Cells(2, 1), fg.Cells(newLast, LAST_FORMULA_COL)).FillDown())
        if newLast < oldLast:
            retry(lambda: fg.Range(fg.Cells(newLast + 1, 1), fg.Cells(oldLast, lastRawCol)).ClearContents())
        print(f"Facture GLS : {oldLast - 1} anciennes lignes -> {n} nouvelles")

        # ---- 2) Rafraichir tous les TCD (Poids, TCD, Comptes GLS, Bilan factures, puis
        #    Bilan clients qui depend de TCD -- RefreshAll() gere l'ordre de dependance) ----
        retry(lambda: wb.RefreshAll())
        try:
            xl.CalculateUntilAsyncQueriesDone()
        except Exception:
            pass
        xl.Calculate()

        # ---- 3) Import csv : TOUT en formules (pas de donnees brutes a coller) -- son nombre
        #    de lignes suit le nombre de colis UNIQUES du TCD recalcule (pas n, le nb de charges
        #    brutes). On ne touche JAMAIS la ligne 2 (modele des formules pour le FillDown).
        tcd = wb.Sheets("TCD")
        tcdLast = tcd.Cells(tcd.Rows.Count, 4).End(xlUp).Row  # colonne D = tracking (categories pivot)
        nbColis = tcdLast - 3  # le TCD demarre a la ligne 4 (lignes 1-3 = entetes/zone de controle)
        impNewLast = 1 + nbColis
        print(f"Import csv : {nbColis} colis uniques (TCD, lignes 4..{tcdLast})")

        imp = wb.Sheets("Import csv")
        impOldLast = imp.Cells(imp.Rows.Count, 7).End(xlUp).Row  # colonne G = Tracking
        if impNewLast > impOldLast:
            retry(lambda: imp.Range(imp.Cells(2, 1), imp.Cells(impNewLast, LAST_COL_IMPORT_CSV)).FillDown())
        elif impNewLast < impOldLast:
            retry(lambda: imp.Range(imp.Cells(impNewLast + 1, 1), imp.Cells(impOldLast, LAST_COL_IMPORT_CSV)).ClearContents())
        xl.Calculate()  # recalcule Import csv une fois etendu/reduit

        # ---- 4) Reconciliation PDF : TTC facture GLS vs TTC calcule (Bilan factures) ----
        if pdf_paths:
            try:
                fill_reconciliation(wb, pdf_paths)
            except Exception as e:
                print("Reconciliation GLS ignoree :", e)
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
