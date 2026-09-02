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
        # BUG TROUVE 2026-08-27 : echec silencieux (juste une liste vide, TCD reste vide sans
        # aucune explication) si pypdf n'est pas installe -- constate sur la VM prod, ou
        # requirements.txt ne listait pas cette dependance (utilisee par 11/12 finaliseurs).
        print("AVERTISSEMENT: module 'pypdf' introuvable -- reconciliation PDF (TCD) impossible. Installer via 'pip install -r Automatisation/requirements.txt'.")
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

def _est_valeur_com_vide_ou_erreur(v):
    """True si v est vide/absent, "(vide)" (etiquette PivotTable native pour une ligne sans
    donnee reelle, confirme sur le TCD Kuehne), ou une VALEUR D'ERREUR EXCEL (#N/A...).
    Meme piege deja identifie/corrige cote UPS (finaliser_ups.py) : Excel via win32com NE
    retourne PAS la chaine "#N/A" pour une cellule en erreur -- un ENTIER NEGATIF (code
    d'erreur OLE, ex. NA()=-2146826246), confirme empiriquement. Les codes d'erreur sont tous
    des entiers tres negatifs (< -2000000000) ; un tracking/numero de compte reel ne descend
    jamais a ce niveau -- distinction fiable."""
    if v is None:
        return True
    if isinstance(v, int) and v < -2000000000:
        return True
    s = str(v).strip()
    return (not s) or s.startswith("#") or s == "(vide)"


def derniere_ligne_reelle(ws, last_col_check, last_row_end_xlup):
    """Rogne 'last_row_end_xlup' (resultat de Cells(Rows.Count,c).End(xlUp).Row) en remontant
    tant que la cellule de la colonne 'last_col_check' est vide/"(vide)"/en erreur (cf.
    _est_valeur_com_vide_ou_erreur). BUG TROUVE 2026-09-01 : le TCD Kuehne a un PivotCache
    fige sur une plage plus large que les vraies donnees (cf. redirection PivotCache
    ci-dessus) -- meme apres redirection, End(xlUp) peut encore s'arreter sur une ligne
    residuelle si le TCD n'a pas immediatement rafraichi sa taille de sortie. Lecture en BLOC
    (1 aller-retour COM), pas cellule par cellule."""
    if last_row_end_xlup < 2:
        return last_row_end_xlup
    values = ws.Range(ws.Cells(2, last_col_check), ws.Cells(last_row_end_xlup, last_col_check)).Value
    if not isinstance(values, tuple):
        values = ((values,),)
    row = last_row_end_xlup
    while row >= 2:
        if not _est_valeur_com_vide_ou_erreur(values[row - 2][0]):
            break
        row -= 1
    return row


def mois_cible_serial(csv_paths):
    """Serie Excel (jours depuis 1899-12-30) du 1er du mois majoritaire de la colonne
    "Date facture" des CSV bruts Kuehne (format JJ/MM/AAAA, confirme sur un vrai CSV
    d'entree) -- BUG TROUVE 2026-09-01 : la Date validite tarif de "Kuehne_Import" restait
    figee a la date du modele (01/06/2026), jamais mise a jour au mois reellement traite
    (meme famille de bug deja corrigee cote UPS/Delivengo). None si aucune date exploitable."""
    from collections import Counter
    import datetime as _dt
    compte_mois = Counter()
    for p in csv_paths:
        with open(p, encoding="latin-1", newline="") as fh:
            rows = list(csv.reader(fh, delimiter=";"))
        if not rows:
            continue
        header = rows[0]
        try:
            i_date = header.index("Date facture")
        except ValueError:
            continue
        for r in rows[1:]:
            if len(r) <= i_date:
                continue
            m = re.fullmatch(r"(\d{2})/(\d{2})/(\d{4})", str(r[i_date]).strip())
            if m:
                compte_mois[f"{m.group(3)}{m.group(2)}"] += 1
    if not compte_mois:
        return None
    aaaamm = compte_mois.most_common(1)[0][0]
    annee, mois = int(aaaamm[:4]), int(aaaamm[4:6])
    epoch = _dt.datetime(1899, 12, 30)
    return (_dt.datetime(annee, mois, 1) - epoch).days


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
    date_validite_serial = mois_cible_serial(csv_paths)
    if date_validite_serial is None:
        print("AVERTISSEMENT: mois de facturation introuvable dans les CSV -- 'Kuehne_Import'!B2 non mise a jour, reste celle du modele.")

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

        # BUG TROUVE 2026-09-01 : le TCD source de "Kuehne_Import" a son PivotCache fige sur
        # 'Fichier Kuehne'!A1:GM220 (plage FIXE a 220 lignes, jamais resynchronisee -- confirme
        # par analyse XML directe du modele). Un mois avec PLUS de 220 lignes perdrait des
        # trackings reels silencieusement dans le TCD (donc dans Kuehne_Import) ; un mois avec
        # MOINS laisse des lignes "(vide)"/0 residuelles jusqu'a 220. Fix : rediriger le
        # PivotCache vers la vraie plage de donnees a CHAQUE run, meme principe deja applique
        # cote UPS (finaliser_ups.py, redirection newRangeFacture).
        xlDatabase = 1
        newRangeKuehne = ws.Range(ws.Cells(1, 1), ws.Cells(newLast, lastRawCol))
        wsTcd = wb.Sheets("TCD")
        for i in range(1, wsTcd.PivotTables().Count + 1):
            pt = wsTcd.PivotTables(i)
            newCache = wb.PivotCaches().Create(SourceType=xlDatabase, SourceData=newRangeKuehne)
            pt.ChangePivotCache(newCache)
        print(f"'TCD' : PivotCache redirige vers {newRangeKuehne.Address}.")

        # 5) rafraichir les tableaux croises (TCD, Bilan clients)
        wb.RefreshAll()
        try:
            xl.CalculateUntilAsyncQueriesDone()
        except Exception:
            pass

        # BUG TROUVE 2026-09-01 : "Kuehne_Import" (onglet livre a l'ERP) n'etait JAMAIS
        # redimensionne ni date -- reste fige a l'etat du modele (190 lignes, Date validite
        # tarif=01/06/2026) quel que soit le mois reellement traite. Colonne C = "Tracking"
        # (jamais None, "(vide)" sur une ligne residuelle -- signal fiable confirme).
        lastTcdBrut = wsTcd.Cells(wsTcd.Rows.Count, 3).End(xlUp).Row
        lastTcd = derniere_ligne_reelle(wsTcd, 3, lastTcdBrut)
        print(f"'TCD' : {lastTcd - 1} tracking(s) reel(s) (End(xlUp) brut={lastTcdBrut}).")

        wsImp = wb.Sheets("Kuehne_Import")
        if date_validite_serial is not None:
            wsImp.Cells(2, 2).Value = date_validite_serial  # B2 Date validite tarif (=B{prev} propage aux lignes suivantes)

        # Redimensionne "Kuehne_Import" au vrai nombre de lignes. 1 seule ligne d'en-tete
        # cote TCD ET Kuehne_Import, MEME indice de ligne des deux cotes (confirme par lecture
        # directe, pas de decalage tcdrow contrairement a UPS -- TCD a 2 lignes d'en-tete la-bas)
        # -- newLastImp = lastTcd, PAS lastTcd-1. BUG TROUVE 2026-09-02 (run reel aout 2026) :
        # le "-1" etait copie a tort du raisonnement UPS et faisait perdre la toute derniere
        # ligne reelle du TCD a chaque run (ex. tracking EXP20260817-2907610, poste Zones
        # eloignees=25,70EUR issu de la facture "evenements" -- absent de Kuehne_Import alors
        # que present dans TCD), cassant le controle TCD!Y18 ("Import en erreur").
        # Formules =TCD!<col><row>/=B{prev} deja dans le modele, INCHANGEES (contrairement a
        # "Fichier import" UPS qui reecrit ses formules a chaque run) -- on ne fait qu'etendre
        # ou purger la plage deja peuplee.
        newLastImp = max(lastTcd, 2)
        oldLastImp = wsImp.Cells(wsImp.Rows.Count, 6).End(xlUp).Row  # colonne F = "N Tracking"
        LAST_COL_IMPORT = 23  # A..W (23 colonnes standard ERP)
        if newLastImp > oldLastImp:
            # Mois avec PLUS de lignes que l'etat precedent : FillDown() DEPUIS la derniere
            # ligne DEJA peuplee (oldLastImp), PAS depuis la ligne 2 -- sinon la formule de la
            # ligne 2 (souvent litterale/differente, cf. B2 date en dur vs B3+=B{prev}) serait
            # copiee sur TOUTES les lignes intermediaires deja correctement peuplees,
            # ecrasant leurs propres formules relatives.
            wsImp.Range(wsImp.Cells(max(oldLastImp, 2), 1), wsImp.Cells(newLastImp, LAST_COL_IMPORT)).FillDown()
            print(f"'Kuehne_Import' : formules etendues jusqu'a la ligne {newLastImp}.")
        elif newLastImp < oldLastImp:
            # Mois avec MOINS de lignes : purge le surplus residuel (ClearContents, PAS
            # EntireRow.Delete -- les formules du modele au-dela doivent rester intactes pour
            # un futur mois plus long).
            wsImp.Range(wsImp.Cells(newLastImp + 1, 1), wsImp.Cells(oldLastImp, LAST_COL_IMPORT)).ClearContents()
            print(f"'Kuehne_Import' : lignes residuelles au-dela de {newLastImp} effacees.")
        xl.Calculate()

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
