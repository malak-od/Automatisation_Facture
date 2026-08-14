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
    Taxe gazole, Zones eloignees). Colonne A "ID Client" : AUCUNE formule ni
    source automatisable identifiee dans le classeur (valeurs figees,
    probablement saisies/collees a la main) -- NON reconstruite ici, comme
    deja decide pour Chronopost (meme cas).
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
    """'TOTAL HT' de la facture Colissimo (libelle explicite en clair sur le
    recapitulatif, ex. 'TOTAL HT 23 220,68 �')."""
    try:
        import pdfplumber
    except ImportError:
        return None, None
    try:
        with pdfplumber.open(pdf_path) as pdf:
            text = pdf.pages[0].extract_text() or ""
    except Exception:
        return None, None
    m_num = re.search(r"FACTURE\s*N[°\s]*([A-Z0-9]+)", text)
    # "€" du PDF Colissimo parfois mal encode (extrait comme "�" par pdfplumber selon
    # la police embarquee) -- "." accepte n'importe quel caractere de fin, pas seulement "€".
    m_total = re.search(r"TOTAL\s*HT\s*([\d\s\xa0]+,\d{2})\s*.", text)
    numero = m_num.group(1) if m_num else None
    total = float(m_total.group(1).replace("\xa0", "").replace(" ", "").replace(",", ".")) if m_total else None
    return numero, total


def fill_reconciliation(wb, pdf_paths):
    """Onglet 'Bilan Factures' : colle le 'TOTAL HT' extrait de chaque PDF (colonne D,
    'PDF HT') en face de sa ligne (appariee par N facture, colonne A du TCD).
    Indemnisation/Avoir (colonnes E/F) restent a la main -- non extractibles
    automatiquement d'un PDF simple (pas de libelle structure pour ces 2 montants)."""
    bf = wb.Sheets("Bilan Factures")
    xlUp = -4162
    lastRow = bf.Cells(bf.Rows.Count, 1).End(xlUp).Row
    numeros = {}
    for r in range(4, lastRow + 1):
        v = bf.Cells(r, 1).Value
        if v:
            numeros[str(v).strip()] = r
    matched = 0
    for p in pdf_paths:
        numero, total = extract_pdf_total_ht(p)
        if numero is None or total is None:
            print(f"Reconciliation Colissimo : PDF ignore (numero de facture ou TOTAL HT introuvable) -> {os.path.basename(p)}")
            continue
        row = numeros.get(numero)
        if row is None:
            print(f"Reconciliation Colissimo : facture {numero} (PDF {os.path.basename(p)}) absente de 'Bilan Factures'")
            continue
        bf.Cells(row, 4).Value = total
        matched += 1
        print(f"Reconciliation Colissimo : facture {numero} -> TOTAL HT PDF={total}")
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
    """--presta <csv...> [--douane <csv...>] [--pdf <pdf...>]"""
    presta, douane, pdf_paths = [], [], []
    cur = None
    for a in rest:
        if a == "--presta":
            cur = "presta"
        elif a == "--douane":
            cur = "douane"
        elif a == "--pdf":
            cur = "pdf"
        elif cur == "presta":
            presta.append(a)
        elif cur == "douane":
            douane.append(a)
        elif cur == "pdf":
            pdf_paths.append(a)
    return presta, douane, pdf_paths


def main():
    modele, sortie = sys.argv[1], sys.argv[2]
    presta_paths, douane_paths, pdf_paths = parse_args(sys.argv[3:])
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

        # ---- 2) Rafraichir tous les TCD (Poids, TCD, Bilan Factures, puis Bilan Client qui
        #    depend de TCD -- RefreshAll() gere l'ordre de dependance) ----
        retry(lambda: wb.RefreshAll())
        try:
            xl.CalculateUntilAsyncQueriesDone()
        except Exception:
            pass
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
