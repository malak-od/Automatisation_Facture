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
    = 1 ligne par n. de FACTURE GLS (colonne A du pivot), col B = Somme de Total HT.
    Sans PDF fourni, col C/D restent vides (plus de calcul theorique B*1,2, decision
    utilisateur 2026-09-07). Reconciliation PDF (fill_reconciliation) : colle le HT
    extrait du "Montant H.T." en colonne C, un message "ok"/"pas bien" en colonne D
    (B=C a l'arrondi au centime pres), le TTC extrait du "Montant T.T.C." en colonne
    E -- appariement par n. de facture lu dans le CONTENU du PDF (motif '<10
    chiffres>Document'), pas le nom de fichier (perdu par multer via l'app web).
    Colonne F (Ecart) supprimee (decision utilisateur 2026-09-08).
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

def first_day_of_month_serial(header, rows, date_col_name="Date jour"):
    """1er jour du mois de la premiere 'Date jour' trouvee (format brut GLS,
    JJ.MM.AAAA), en NOMBRE SERIEL Excel -- utilise pour 'Import csv'!C2 ('Date
    validite tarif'), une valeur FIXE saisie a la main dans le modele (=C2
    recopie sur les lignes suivantes) qui reste sinon bloquee sur le mois du
    modele (juin) meme en traitant un autre mois (constate sur aout 2026,
    meme piege deja corrige sur DPD/Mondial Relay)."""
    import datetime as _dt
    i = next((idx for idx, name in enumerate(header) if name.strip().lower() == date_col_name.strip().lower()), None)
    if i is None:
        return None
    EXCEL_EPOCH = _dt.datetime(1899, 12, 30)
    for r in rows:
        v = r[i] if i < len(r) else None
        if v not in (None, ""):
            m = re.match(r"^(\d{2})\.(\d{2})\.(\d{4})$", str(v).strip())
            if m:
                first = _dt.datetime(int(m.group(3)), int(m.group(2)), 1)
                return (first - EXCEL_EPOCH).days
    return None

def _extract_pdf_montant(pdf_path, label_regex):
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
        m = re.search(r"(?:^|\n)(\d[\d\s\xa0]*,\d{2})\s*" + label_regex, txt)
        if m:
            return float(m.group(1).replace("\xa0", "").replace(" ", "").replace(",", "."))
    return None

def extract_pdf_ttc(pdf_path):
    """Montant T.T.C. de la facture GLS (libelle explicite dans le PDF, cf. video process
    p.ex. '3\xa0683,04Montant T.T.C.:') -- prend la 1ere occurrence trouvee (recap en
    page 0, identique au recap final)."""
    return _extract_pdf_montant(pdf_path, r"Montant T\.T\.C\.:")

def extract_pdf_ht(pdf_path):
    """Montant H.T. de la facture GLS (libelle explicite, ex. '3\xa0069,20 Montant
    H.T.: ' -- constate en derniere page, recapitulatif final, juste avant le TTC)."""
    return _extract_pdf_montant(pdf_path, r"Montant H\.T\.:")

def pdf_facture_numero(pdf_path):
    """N. de facture GLS (colonne 'Facture' de Facture GLS / colonne A de Bilan
    factures). BUG TROUVE 2026-09-08 : lu depuis le NOM DE FICHIER (Facture_<payeur>_
    <facture>_...pdf) -- fonctionne seulement si ce nom original survit jusqu'au
    script (teste en ligne de commande, chemins nommes explicitement). Mais l'app web
    (server.js, carriers/gls/index.js!buildArgs) passe les chemins UPLOADS bruts,
    renommes en hex par multer SANS le nom d'origine -> le numero de facture n'etait
    JAMAIS trouve en pratique via l'interface, PDF toujours ignore silencieusement
    (colonnes HT/TTC vides), meme regle que DPD/Mondial Relay qui lisent deja depuis
    le CONTENU. Lu desormais depuis le texte du PDF lui-meme (motif '<10 chiffres>
    Document', constate sur les factures de juin et aout 2026, absent de la page 0
    -- recap client -- present a partir de la page 1)."""
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
        m = re.search(r"(\d{10})Document", txt)
        if m:
            return int(m.group(1))
    return None

def fill_reconciliation(wb, pdf_paths):
    """Onglet 'Bilan factures' : colle le HT extrait de chaque PDF (colonne C) et le
    TTC (colonne E) en face de sa ligne (appariee par n. de facture, colonne A).
    Colonne D : message 'ok' si le Total HT calcule (B) egale le HT du PDF (C) a
    l'arrondi au centime pres, 'pas bien' sinon -- decision utilisateur 2026-09-08.
    Colonne F (Ecart) supprimee -- decision utilisateur 2026-09-08."""
    bf = wb.Sheets("Bilan factures")
    bf.Cells(3, 3).Value = "HT"
    bf.Cells(3, 4).Value = "Controle"
    bf.Cells(3, 5).Value = "TTC (PDF)"
    bf.Cells(3, 6).ClearContents()
    lastRow = bf.Cells(bf.Rows.Count, 1).End(-4162).Row  # xlUp
    numeros = {}
    for r in range(4, lastRow + 1):
        v = bf.Cells(r, 1).Value
        if isinstance(v, (int, float)):
            numeros[int(v)] = r
    matched = 0
    for p in pdf_paths:
        num = pdf_facture_numero(p)
        ht = extract_pdf_ht(p)
        ttc = extract_pdf_ttc(p)
        if num is None or (ht is None and ttc is None):
            print(f"Reconciliation GLS : PDF ignore (n. facture, HT et TTC introuvables) -> {os.path.basename(p)}")
            continue
        row = numeros.get(num)
        if row is None:
            print(f"Reconciliation GLS : facture {num} (PDF {os.path.basename(p)}) absente de 'Bilan factures'")
            continue
        if ht is not None:
            bf.Cells(row, 3).Value = ht
            bf.Cells(row, 4).Formula = f'=IF(ROUND(B{row},2)=ROUND(C{row},2),"ok","pas bien")'
        if ttc is not None:
            bf.Cells(row, 5).Value = ttc
        bf.Cells(row, 6).ClearContents()
        matched += 1
        print(f"Reconciliation GLS : facture {num} -> HT PDF={ht}, TTC PDF={ttc}")
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

        # ---- 2bis) 'Bilan factures' colonnes C/D/F : BUG TROUVE 2026-09-03 -- ces
        #    cellules NE SONT NI un TCD (ne suivent pas RefreshAll) NI des formules
        #    dans le modele clone, juste des VALEURS FIGEES heritees du modele de
        #    reference (juin 2026, facture 2501382993) -- jamais recalculees pour
        #    le mois traite. Purgees ici (avant tout) pour repartir propre ; si un
        #    PDF est fourni, fill_reconciliation() repeuple ensuite C (HT du PDF)
        #    et D (message ok/pas bien) -- decision utilisateur 2026-09-08. F
        #    (Ecart) supprimee (decision utilisateur 2026-09-08), toujours vide.
        bf = wb.Sheets("Bilan factures")
        bfLast = bf.Cells(bf.Rows.Count, 1).End(xlUp).Row
        for r in range(4, bfLast + 1):
            v = bf.Cells(r, 1).Value
            if isinstance(v, (int, float)):
                bf.Cells(r, 3).ClearContents()
                bf.Cells(r, 4).ClearContents()
                bf.Cells(r, 6).ClearContents()
        xl.Calculate()
        print("Bilan factures : colonnes C/D/F (valeurs figees du modele / colonne supprimee) purgees.")

        # ---- 2ter) 'Bilan clients' colonnes E/F ("Facture pdf"/"Montant") : meme
        #    piege que 2bis, en pire -- pas une valeur figee du mois precedent mais
        #    un CONTROLE PONCTUEL fait a la main en NOVEMBRE 2022 (note "novembre
        #    2022 -> ok" en G15 du modele), donc totalement sans rapport avec le
        #    mois traite. Contrairement a "Bilan factures" (1 ligne = 1 facture),
        #    la colonne A ici est un NUMERO DE COMPTE GLS (pas une facture) -> pas
        #    d'appariement facture<->PDF possible sur cette feuille. Purge demandee
        #    par l'utilisateur plutot que d'inventer une reconciliation non
        #    documentee (aucune video/doc ne couvre ce bloc).
        #    BUG TROUVE 2026-09-07 : bcLast etait calcule sur la colonne A (qui n'a
        #    que 2 lignes, "(vide)" en ligne 2, le TCD n'ayant qu'une seule
        #    categorie ce mois-ci) -> la purge s'arretait a la ligne 2, laissant
        #    intact tout le reliquat E3:G15 (visible sur aout 2026 : "2,501E+09"/
        #    "17,73" en E3/F3 + la note "novembre 2022 -> ok" en E15/F15/G15).
        #    On calcule desormais la derniere ligne sur la colonne E elle-meme
        #    (ou F/G), la ou vit reellement le reliquat a purger.
        bc = wb.Sheets("Bilan clients")
        bcLastCol = bc.Cells(bc.Rows.Count, 5).End(xlUp).Row
        for col in (6, 7):
            bcLastCol = max(bcLastCol, bc.Cells(bc.Rows.Count, col).End(xlUp).Row)
        retry(lambda: bc.Range(bc.Cells(2, 5), bc.Cells(max(bcLastCol, 2), 7)).ClearContents())
        print("Bilan clients : colonnes E-G (reliquat de controle manuel novembre 2022) purgees.")

        # ---- 3) Import csv : TOUT en formules (pas de donnees brutes a coller) -- son nombre
        #    de lignes suit le nombre de colis UNIQUES du TCD recalcule (pas n, le nb de charges
        #    brutes). On ne touche JAMAIS la ligne 2 (modele des formules pour le FillDown).
        tcd = wb.Sheets("TCD")
        tcdLast = tcd.Cells(tcd.Rows.Count, 4).End(xlUp).Row  # colonne D = tracking (categories pivot)
        nbColis = tcdLast - 3  # le TCD demarre a la ligne 4 (lignes 1-3 = entetes/zone de controle)
        impNewLast = 1 + nbColis
        print(f"Import csv : {nbColis} colis uniques (TCD, lignes 4..{tcdLast})")

        imp = wb.Sheets("Import csv")
        # C2 ('Date validite tarif') : VALEUR FIXE saisie a la main dans le
        # modele (C3+ = "=C2" recopie par le FillDown ci-dessous) -> mise a
        # jour vers le 1er du mois traite AVANT le FillDown, sinon elle reste
        # bloquee sur le mois du modele (juin) quel que soit le mois traite
        # (constate sur aout 2026, meme piege deja corrige sur DPD/Mondial
        # Relay). Source : colonne 'Date jour' du CSV BCF brut.
        date_validite_serial = first_day_of_month_serial(hdr, rows)
        if date_validite_serial is not None:
            imp.Cells(2, 3).Value = date_validite_serial
        else:
            print("AVERTISSEMENT: 'Date jour' introuvable dans le CSV BCF -> "
                  "'Date validité tarif' (Import csv!C2) non mise à jour, reste celle du modèle.")
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
