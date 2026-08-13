#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
FINALISEUR CHRONOPOST — produit "AAAA_MM_Facture Chronopost.xlsx" A
L'IDENTIQUE du fichier fait a la main (14 feuilles), en partant du fichier
existant comme MODELE. Le transporteur le plus complexe du projet : 2
fichiers Excel bruts (1 par sous-compte/contrat) consolides, un POOL gazole
par facture redistribue au prorata du fret, un TCD en cascade (TCD poids +
TCD alimentent "Fichier import" par formule), plusieurs tables de lookup
(Bibliotheque transporteurs, Categories, Pays TVA, Sous-comptes, Zoning
2shop).

Feuilles STATIQUES (jamais touchees, copiees telles quelles avec le modele) :
  Pays TVA, Sous-comptes, Bibliothèque transporteurs, Catégories, TARIFS,
  cap à 5%, Zoning 2shop, Bilan clients, Avoir.

"Facture Chronopost" (A1:AJ) : colonne A ("ID Clients") est une colonne
CALCULEE (PAS dans le brut -- confirme : le brut recu commence directement
par "No Facture" = colonne B du modele, pas colonne A). Donnees brutes des 2
fichiers sources collees a partir de la colonne B (en-tete des 2 fichiers
Chronopost STABLE, confirme identique sur les 2 sous-comptes de juin 2026 --
pas de decalage mois a mois comme Mondial Relay/DPD, donc pas de resolution
par NOM colonne par colonne necessaire pour le COLLAGE, seulement pour les
formules qui les referencent), puis colonnes calculees W->AE reconstruites
par formule :
  W (Code produit modifie 2shop-Kersun) : resynchronisee sur le modele 2026_07
  (a jour, 2026-08-12) -- =IF(S deja code standard connu,S,IF(sous-compte 0/1
  + produit 5X/5Y/6B/6C,S,IF(C<>2,"",IF(S="5X","5XK",...)))). La version de
  juin ne gerait que le dernier cas (Kersun) -- colonne d'affichage/controle
  interne uniquement, n'alimente aucun calcul en aval de ce script.
  X (Zoning 2shop) : =IF(S="6B",XLOOKUP(H,'Zoning 2shop'!A:A,B:B),
     IF(S="6C",XLOOKUP(G,'Zoning 2shop'!C:C,D:D),""))
  Y (Gazole %) : =AB/Z   Z (fret) : =T   AA (surete+eco taux) : =IF(...,0.08,0.5)
  AB (gazole) : =$AG$5/$AG$2*Z  (pool gazole / total fret * fret ligne)
  AC (hors gazole) : =Z+AA   AD (Categories) : =IF(COUNTIF(...)=0,"catégorie
  inconnue",LOOKUP(N,Catégories!A:A,Catégories!B:B))
  AE (Total avec GO) : =SUM(Z:AB)
Zone recap AF1:AI9 (labels/sous-totaux Frêt/eco/sureté/Gazole/Gestion, taux
gasoil routier/aerien) : AG2/AG3/AG4/AG5 (Frêt/eco/sûreté/Gazole) calcules en
PYTHON (somme par prefixe Numero LT : CAP*/ECO*/SUR*/reste) et ecrits EN
VALEURS ici -- PAS par formule Excel (positions/plages variables mois a
mois selon le tri manuel, cf. transcription video, non fiable a automatiser
par formule). Le reste de la zone recap (AF6:AI9, "Gestion"/"Gazole
reel"/taux gasoil routier-aerien) reste fige au modele, jamais recalcule.

"TCD poids"/"TCD"/"Contrôle pdf" : VRAIS TCD Excel, PivotCache redirige vers
la nouvelle plage de "Facture Chronopost" (meme piege que DPD/Geodis/Mondial
Relay : le cache du modele pointe une plage figee/etroite).

"TCD" colonnes de calcul manuel (S/T/U/V, puis A/B/C/W/Y/Z a partir de la
ligne ou "ID client" natif du TCD existe) reconstruites par formule EXACTE
du modele (RECHERCHEX vers Bibliothèque transporteurs/Facture Chronopost).

"Fichier import" : formules FIXES referencant TCD/TCD poids par POSITION
(stables une fois TCD/TCD poids reconstruits) -- MAIS le carrier Node a DEJA
identifie/supprime les lignes forfaitaires CAP/ECO/SUR de l'import (jamais
dans le TCD natif ventile par Numero LT de facon utile pour l'import ERP) --
donc ICI on ecrit "Fichier import" en VALEURS deja calculees par le carrier
Node (fiable, deja valide sur les 7 PDF de juin 2026 a 0,00€ d'ecart), pas
par re-derivation de formules TCD -- meme principe que "Import CSV" de BLS.

Reconciliation PDF (onglet "Contrôle pdf") : Total HT extrait du bloc "TOTAL
FACTURE<HT><TVA><TTC>" de chaque PDF (format texte structure, confirme sur
les 7 PDF de juin 2026), colle en face du No Facture correspondant.

Necessite : Windows + Excel + pywin32 + pypdf.
Usage :
  python finaliser_chronopost.py "<modele.xlsx>" "<sortie.xlsx>" "<facture1.xlsx>" ["<facture2.xlsx>"...] [--pdf <pdf1> [<pdf2>...]]
"""
import sys, os, shutil, re


def normalize_header(h):
    return re.sub(r"\s+", " ", str(h or "")).strip()


def compare_key(h):
    s = normalize_header(h).lower()
    s = s.translate(str.maketrans("éèêëàâäùûüôöîïç", "eeeeaaauuuooiic"))
    s = re.sub(r"[.,]", "", s)
    return re.sub(r"\s+", " ", s).strip()


def is_xlsx(path):
    import zipfile
    try:
        return zipfile.is_zipfile(path)
    except Exception:
        return False


def col_letter(idx0):
    import openpyxl
    return openpyxl.utils.get_column_letter(idx0 + 1)


def col_index(header, name):
    target = compare_key(name)
    for i, h in enumerate(header):
        if compare_key(h) == target:
            return i
    return None


def to_num(v):
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v or "").strip().replace(" ", "").replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return 0.0


def read_facture_brute(path):
    """Facture Chronopost brute : feuille unique, en-tete EXACT en ligne 4 (index 3),
    donnees a partir de la ligne 5 -- meme structure confirmee sur les 2 sous-comptes."""
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
        ws = wb[wb.sheetnames[0]]
        rows = [[c for c in row] for row in ws.iter_rows(values_only=True)]
        wb.close()
    finally:
        if tmp_path:
            try:
                os.unlink(tmp_path)
            except OSError:
                pass
    header = [normalize_header(h) for h in rows[3]]
    data = [r for r in rows[4:] if any(v not in (None, "") for v in r)]
    return {"file": os.path.basename(path), "header": header, "rows": data}


def excel_serial_from_date(v):
    """v = 'DD/MM/AAAA' str, serial Excel deja numerique, ou datetime.datetime."""
    import datetime as _dt
    EXCEL_EPOCH = _dt.datetime(1899, 12, 30)
    if isinstance(v, _dt.datetime):
        return (v - EXCEL_EPOCH).days
    if isinstance(v, (int, float)):
        return int(v)
    m = re.match(r"^(\d{2})/(\d{2})/(\d{4})$", str(v or "").strip())
    if not m:
        return None
    d, mo, y = int(m.group(1)), int(m.group(2)), int(m.group(3))
    return (_dt.datetime(y, mo, d) - EXCEL_EPOCH).days


def extract_pdf_total(pdf_path):
    """PDF Chronopost : bloc 'Compte<n>Facture<n>...' + 'TOTAL FACTURE<HT><TVA><TTC>'
    colle en fin de document (format texte structure, confirme sur les 7 PDF de juin
    2026) -- meme logique que src/carriers/chronopost/index.js, a garder synchronisee."""
    try:
        import pypdf
    except ImportError:
        return None
    try:
        text = "\n".join((page.extract_text() or "") for page in pypdf.PdfReader(pdf_path).pages)
    except Exception:
        return None
    m_compte = re.search(r"Compte\s*(\d+)\s*Facture\s*(\d+)", text)
    m_total = re.search(r"TOTAL\s*FACTURE\s*([\d\s]+[.,]\d{2})\s*([\d\s]+[.,]\d{2})\s*([\d\s]+[.,]\d{2})", text)
    if not m_compte or not m_total:
        return None
    return {
        "file": os.path.basename(pdf_path),
        "compte": m_compte.group(1),
        "facture": m_compte.group(2),
        "total_ht": round(to_num(m_total.group(1).replace(" ", "")), 2),
    }


def parse_args(argv):
    modele, sortie = argv[1], argv[2]
    factures, pdfs, cur = [], [], "in"
    for a in argv[3:]:
        if a == "--pdf":
            cur = "pdf"
        elif cur == "pdf":
            pdfs.append(a)
        else:
            factures.append(a)
    return modele, sortie, factures, pdfs


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


def redirect_pivot_caches(wb, source_ws, sheet_names, source_sheet, last_row, first_col, last_col):
    """Redirige le PivotCache de tous les TCD des feuilles donnees vers la vraie plage de
    donnees (colonnes first_col->last_col de source_ws), SEULEMENT si leur source actuelle
    reference deja source_sheet (evite de casser un TCD imbrique sur un autre TCD, meme piege
    que Mondial Relay/DPD).

    IMPORTANT : la SourceData du MODELE pointe vers une plage MINUSCULE et locale (constate
    via COM sur les 3 TCD Chronopost : 'Facture Chronopost'!C2:C18/C2:C30/C2:C20, PAS
    B1:AD1048576 comme une premiere lecture openpyxl l'avait suggere -- openpyxl et COM
    peuvent lire des metadonnees differentes/desynchronisees) -- meme "vestige de plage
    figee" deja rencontre sur DPD/Geodis/Mondial Relay. On ignore donc DELIBEREMENT cette
    plage figee et on redirige vers TOUTE la largeur utile (first_col->last_col, fixee par
    l'appelant selon les colonnes reelles de 'Facture Chronopost'), PAS une colonne derivee
    de l'ancienne SourceData."""
    xlDatabase = 1
    new_range = source_ws.Range(source_ws.Cells(1, first_col), source_ws.Cells(last_row, last_col))
    for sheet_name in sheet_names:
        ws = wb.Sheets(sheet_name)
        count = ws.PivotTables().Count
        for i in range(1, count + 1):
            pt = ws.PivotTables(i)
            src = str(pt.PivotCache().SourceData)
            if source_sheet not in src:
                print(f"TCD '{pt.Name}' ({sheet_name}) : source '{src}' n'est pas '{source_sheet}' "
                      f"(TCD imbrique ?) -> non redirige, config preservee.")
                continue
            new_cache = wb.PivotCaches().Create(SourceType=xlDatabase, SourceData=new_range)
            pt.ChangePivotCache(new_cache)
            print(f"TCD '{pt.Name}' ({sheet_name}) : PivotCache redirige vers {new_range.Address} "
                  f"(ancienne source figee '{src}' ignoree).")


def main():
    modele, sortie, facture_paths, pdf_paths = parse_args(sys.argv)
    shutil.copyfile(modele, sortie)  # on ne touche JAMAIS au modele

    factures_brutes = [read_facture_brute(p) for p in facture_paths if is_xlsx(p)]
    if not factures_brutes:
        raise RuntimeError("Aucune facture Chronopost fournie (xlsx, 1 par sous-compte).")

    header = factures_brutes[0]["header"]
    all_rows = []
    for f in factures_brutes:
        i_map = {name: col_index(f["header"], name) for name in header}
        for r in f["rows"]:
            all_rows.append([r[i_map[name]] if i_map[name] is not None and i_map[name] < len(r) else None for name in header])
    ncol = len(header)
    n = len(all_rows)
    print(f"Entrée : {len(factures_brutes)} facture(s), {n} ligne(s)")

    i_facture = col_index(header, "No Facture")
    i_numero_lt = col_index(header, "Numero LT")
    i_produit = col_index(header, "Produit")
    i_date_lt = col_index(header, "Date LT")
    i_sous_compte = col_index(header, "Sous-compte")
    i_pays_arrivee = col_index(header, "Pays arrivee")
    i_pays_depart = col_index(header, "Pays depart")

    # TRI DES LIGNES (2026-08-13) : le fichier fait a la main regroupe TOUTES les lignes
    # forfaitaires (CAP*/ECO*/SUR*) EN TETE de "Facture Chronopost", triees CAP puis ECO puis
    # SUR (confirme cellule par cellule sur 2026_07_Facture Chronopost.xlsx, lignes 2-33 =
    # exactement les 32 lignes forfaitaires dans cet ordre, les lignes "Transport" ne
    # commencent qu'a partir de la ligne 34) -- pas un artefact, c'est CE TRI qui permet les
    # formules modele AF2:AF5 en plages contiguees (=SOMME(T<debut>:T<fin>), cf. transcription
    # video). Avant ce fix, les lignes forfaitaires restaient dispersees a leur position
    # naturelle de lecture du brut (melangees aux lignes normales), ce qui ne reproduisait pas
    # la presentation du fichier reel (capture ecran utilisateur "reproduire la meme DA").
    def _cle_tri_forfait(row):
        lt = str(row[i_numero_lt] or "").strip().upper()
        if lt.startswith("CAP"):
            return (0, lt)
        if lt.startswith("ECO"):
            return (1, lt)
        if lt.startswith("SUR"):
            return (2, lt)
        return (3, "")
    all_rows.sort(key=_cle_tri_forfait)

    # Date validite = mois MAJORITAIRE sur l'ensemble des lignes (PAS la 1re ligne trouvee) --
    # BUG TROUVE 2026-08-13 (meme fix cote Node, chronopost/index.js) : le brut Chronopost
    # recu pour un mois donne contient parfois quelques lignes residuelles du mois precedent
    # EN TETE de fichier (livraisons tardives facturees avec la coupure du mois suivant,
    # confirme sur facture_chronopost_51291303_202607.xlsx : 3 lignes datees du 30/06/2026 en
    # tete de fichier, sur 1717 lignes totales dont 1714 de juillet). Prendre la 1re ligne
    # donnait "juin" pour un fichier envoye et rempli a 99,8% de juillet.
    import datetime as _dt
    EXCEL_EPOCH = _dt.datetime(1899, 12, 30)
    comptage_mois = {}
    for r in all_rows:
        s = excel_serial_from_date(r[i_date_lt]) if i_date_lt is not None else None
        if s is not None:
            d = EXCEL_EPOCH + _dt.timedelta(days=s)
            cle = (d.year, d.month)
            comptage_mois[cle] = comptage_mois.get(cle, 0) + 1
    date_validite_serial = None
    if comptage_mois:
        annee, mois = max(comptage_mois.items(), key=lambda kv: kv[1])[0]
        date_validite_serial = (_dt.datetime(annee, mois, 1) - EXCEL_EPOCH).days
        if len(comptage_mois) > 1:
            detail = ", ".join(f"{a}-{m:02d}: {n} ligne(s)" for (a, m), n in sorted(comptage_mois.items(), key=lambda kv: -kv[1]))
            print(f"INFO: plusieurs mois detectes dans les fichiers recus ({detail}) — mois retenu : {mois:02d}/{annee} (majoritaire).")

    pdfs = [r for r in (extract_pdf_total(p) for p in pdf_paths) if r]

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

        # 1) "Facture Chronopost" : purge + collage donnees brutes a partir de la colonne B
        #    (colonne A = "ID Clients", CALCULEE, PAS dans le brut recu -- confirme : le brut
        #    commence directement par "No Facture" = colonne B du modele). En-tete des 2
        #    fichiers Chronopost STABLE (confirme identique sur les 2 sous-comptes de juin
        #    2026) -- pas de decalage mois a mois comme Mondial Relay/DPD, mais l'OFFSET fixe
        #    de +1 colonne (a cause de "ID Clients") doit etre applique partout.
        ws = wb.Sheets("Facture Chronopost")
        # PIEGE MAJEUR (decouvert en debug) : le modele a un AutoFilter actif sur cette feuille
        # -> ClearContents()/Value= sur une plage de lignes echouent SILENCIEUSEMENT (aucune
        # exception levee, aucun avertissement) des que le filtre masque des lignes -- Excel
        # restreint alors les operations de plage aux lignes VISIBLES uniquement, ce qui laisse
        # les anciennes valeurs du modele intactes tout en semblant reussir. DOIT etre desactive
        # AVANT toute manipulation de plage sur cette feuille.
        if ws.AutoFilterMode:
            ws.AutoFilterMode = False
        FIRST_RAW_COL = 2  # colonne B
        LAST_RAW_COL = FIRST_RAW_COL - 1 + ncol  # B -> U (20 colonnes brutes)
        oldLast = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row
        newLast = 1 + n
        maxLast = max(oldLast, newLast, 2)
        LAST_CALC_COL = 31  # AE (colonnes calculees W->AE = 23->31 ; AF/32 = 'Total', reste une formule fixe du modele, jamais reecrite ici)

        retry(lambda: ws.Range(ws.Cells(2, 1), ws.Cells(maxLast, LAST_CALC_COL)).ClearContents())
        # "No Facture" (et "Numero LT") sont du TEXTE dans le modele ET dans le brut recu
        # ("13655988", pas 13655988) -- ecrire via COM .Value sur une plage dont le format de
        # cellule n'est pas force en Texte laisse Excel AUTO-CONVERTIR certaines valeurs en
        # nombre selon le contexte, ce qui casse le TCD natif "Contrôle pdf"/"TCD" (champ
        # rowField mixte texte/nombre -> agregation aberrante, constate : une concatenation de
        # plusieurs "13655988" lue comme un seul grand nombre). Force NumberFormat="@" (Texte)
        # sur ces 2 colonnes AVANT d'ecrire les valeurs pour eviter toute auto-conversion.
        col_no_facture_letter = col_letter(FIRST_RAW_COL - 1 + i_facture) if i_facture is not None else None
        col_numero_lt_letter = col_letter(FIRST_RAW_COL - 1 + i_numero_lt) if i_numero_lt is not None else None
        for letter in (col_no_facture_letter, col_numero_lt_letter):
            if letter:
                retry(lambda l=letter: ws.Range(f"{l}2:{l}{maxLast}").__setattr__("NumberFormat", "@"))
        retry(lambda: ws.Range(ws.Cells(2, FIRST_RAW_COL), ws.Cells(newLast, LAST_RAW_COL)).__setattr__("Value", all_rows))
        if newLast < oldLast:
            retry(lambda: ws.Range(ws.Cells(newLast + 1, 1), ws.Cells(oldLast, LAST_CALC_COL)).ClearContents())

        # Colonnes calculees X->AF (formules du modele, colonnes source resolues par nom + offset).
        col_c = col_letter(FIRST_RAW_COL - 1 + i_sous_compte) if i_sous_compte is not None else "D"
        col_s = col_letter(FIRST_RAW_COL - 1 + i_produit) if i_produit is not None else "T"
        col_g = col_letter(FIRST_RAW_COL - 1 + i_pays_depart) if i_pays_depart is not None else "H"
        col_h = col_letter(FIRST_RAW_COL - 1 + i_pays_arrivee) if i_pays_arrivee is not None else "I"
        col_t = col_letter(FIRST_RAW_COL - 1 + col_index(header, "Montant HT")) if col_index(header, "Montant HT") is not None else "U"
        col_n = col_letter(FIRST_RAW_COL - 1 + col_index(header, "Type prestation")) if col_index(header, "Type prestation") is not None else "O"

        # Colonnes/references VERIFIEES cellule par cellule contre le vrai modele (row 2,
        # 2026_06_Facture Chronopost.xlsx) -- un decalage d'1 colonne ici (cle ET references
        # internes AA/AB/AC/Z) avait ete introduit par erreur, laissant W (Code produit
        # modifie 2shop-Kersun) TOUJOURS VIDE malgre la bibliotheque a jour (bug trouve et
        # corrige 2026-08-12, cf. capture ecran utilisateur "Facture Chronopost").
        formulas_w_ae = {
            # W : formule resynchronisee sur le classeur 2026_07 (a jour, verifie 2026-08-12
            # -- celle de juin ne gerait QUE le cas Kersun, juillet ajoute 2 cas manquants :
            # (a) produit deja un code standard connu -> recopie tel quel, (b) sous-compte 0/1
            # (pas Kersun) avec produit 5X/5Y/6B/6C -> recopie tel quel sans le suffixe K).
            23: f'=IF(OR({{col_s}}{{{{row}}}}=2,{{col_s}}{{{{row}}}}=16,{{col_s}}{{{{row}}}}=17,{{col_s}}{{{{row}}}}=44,{{col_s}}{{{{row}}}}=86,{{col_s}}{{{{row}}}}=1,{{col_s}}{{{{row}}}}="1S",{{col_s}}{{{{row}}}}="3Z",{{col_s}}{{{{row}}}}="X"),{{col_s}}{{{{row}}}},IF(AND(OR({{col_c}}{{{{row}}}}=0,{{col_c}}{{{{row}}}}=1),OR({{col_s}}{{{{row}}}}="5X",{{col_s}}{{{{row}}}}="5Y",{{col_s}}{{{{row}}}}="6B",{{col_s}}{{{{row}}}}="6C")),{{col_s}}{{{{row}}}},IF({{col_c}}{{{{row}}}}<>2,"",IF({{col_s}}{{{{row}}}}="5X","5XK",IF({{col_s}}{{{{row}}}}="5Y","5YK",IF({{col_s}}{{{{row}}}}="6B","6BK",IF({{col_s}}{{{{row}}}}="6C","6CK","")))))))'.format(col_c=col_c, col_s=col_s),  # W
            24: f'=IF({{col_s}}{{{{row}}}}="6B",_xlfn.XLOOKUP({{col_h}}{{{{row}}}},\'Zoning 2shop\'!A:A,\'Zoning 2shop\'!B:B),IF({{col_s}}{{{{row}}}}="6C",_xlfn.XLOOKUP({{col_g}}{{{{row}}}},\'Zoning 2shop\'!C:C,\'Zoning 2shop\'!D:D),""))'.format(col_s=col_s, col_g=col_g, col_h=col_h),  # X
            25: "=AB{row}/Z{row}",  # Y (Gazole %) = gazole / fret
            26: f"={{col_t}}{{{{row}}}}".format(col_t=col_t),  # Z (fret)
            27: f'=IF(OR({{col_s}}{{{{row}}}}="6C",{{col_s}}{{{{row}}}}="6B",{{col_s}}{{{{row}}}}="5X",{{col_s}}{{{{row}}}}="5Y"),0.08,0.5)'.format(col_s=col_s),  # AA (taux surete+eco)
            28: "=$AG$5/$AG$2*Z{row}",  # AB (gazole reparti -- pool/total fret * fret ligne)
            29: "=Z{row}+AA{row}",  # AC (hors gazole)
            30: f'=IF(COUNTIF(Catégories!A:A,{{col_n}}{{{{row}}}})=0,"catégorie inconnue",LOOKUP({{col_n}}{{{{row}}}},Catégories!A:A,Catégories!B:B))'.format(col_n=col_n),  # AD (Categories)
            31: "=SUM(Z{row}:AB{row})",  # AE (Total avec GO)
        }
        for col_idx, tmpl in formulas_w_ae.items():
            retry(lambda c=col_idx, t=tmpl: ws.Range(ws.Cells(2, c), ws.Cells(newLast, c))
                  .__setattr__("Formula", [[t.format(row=r)] for r in range(2, newLast + 1)]))

        # Zone recap AG2:AG5 : VRAIES FORMULES Excel SUMIF sur "Type prestation" (colonne N),
        # PAS des valeurs Python (decision utilisateur 2026-08-13 : "pour les colonnes AF et
        # AG faut laisser les formules dans les cellules"). SUMIF sur le LIBELLE plutot que
        # SUM sur une plage fixe recalee a la main (methode du fichier reel, cf. transcription
        # video "AF2:AF5 = SOMME(T<debut>:T<fin>) sur plage FIXE... recalee a la main apres
        # tri") : verifie sur 2026_07_Facture Chronopost.xlsx que la plage fixe AG2=SUM(T31:
        # T5000) EMPIETE en fait sur les lignes SURT* (31-33 sont encore forfaitaires a ce
        # moment-la, pas des lignes Transport) -- erreur de calage manuel qui gonfle le Fret
        # de ~130€. SUMIF(Type prestation) est ROBUSTE a ce risque (independant du tri/de la
        # position des lignes) : AG2=SUMIF(N,"Transport",T), AG3=SUMIF(N,"Participation
        # Eco-Responsable",T), AG4=SUMIF(N,"Sûreté colis",T) (libelles confirmes via la table
        # 'categories' de config.json). AG5 (Gazole, pool CAP*) reste par PREFIXE Numero LT
        # (2 libelles distincts "Surcharge Carburant Aérien"/"Routier" regroupes sous un seul
        # poste "Gazole") -- SUMIF sur les 2 libelles cumules.
        # CONSEQUENCE ATTENDUE (2026-08-13, confirme utilisateur) : la colonne AB "gazole"
        # (=$AG$5/$AG$2*Z{row}) depend directement de ce AG2 -- comme notre AG2 est maintenant
        # STRICTEMENT "Transport" (vs le fichier reel dont l'AG2 mal cale inclut ~130€ de
        # lignes SURT*), le taux AG5/AG2 differe legerement du fichier reel (0,109 ici vs
        # 0,101 dans '2026_07_Facture Chronopost.xlsx'), ce qui cree un ecart cumule de
        # plusieurs centaines d'euros sur la colonne AB par rapport A CE FICHIER REEL
        # SPECIFIQUE -- CE N'EST PAS UN BUG : c'est le fichier reel qui a l'erreur de calage,
        # notre AG2 est correct par construction (SUMIF, insensible au tri/position).
        ws.Cells(2, 33).Formula = f'=SUMIF({col_n}:{col_n},"Transport",{col_t}:{col_t})'  # AG2 Frêt
        ws.Cells(3, 33).Formula = f'=SUMIF({col_n}:{col_n},"Participation Eco-Responsable",{col_t}:{col_t})'  # AG3 eco
        ws.Cells(4, 33).Formula = f'=SUMIF({col_n}:{col_n},"Sûreté colis",{col_t}:{col_t})'  # AG4 sûreté
        ws.Cells(5, 33).Formula = (
            f'=SUMIF({col_n}:{col_n},"Surcharge Carburant Aérien",{col_t}:{col_t})'
            f'+SUMIF({col_n}:{col_n},"Surcharge Carburant Routier",{col_t}:{col_t})'
        )  # AG5 Gazole
        print("Zone récap : formules SUMIF ecrites en AG2 (Frêt) / AG3 (eco) / AG4 (sûreté) / AG5 (Gazole).")

        # 2) "TCD poids"/"Contrôle pdf" : PivotCache redirige vers TOUTE la largeur utile de
        #    "Facture Chronopost" (B->AF), PAS la plage figee/etroite du modele (cf. docstring
        #    de redirect_pivot_caches -- meme piege que DPD/Geodis/Mondial Relay). Ces 2 TCD
        #    fonctionnent correctement avec cette redirection (donnees + reconciliation PDF
        #    verifiees exactes).
        #
        #    "TCD" est VOLONTAIREMENT EXCLU de cette redirection (limite connue, acceptee par
        #    l'utilisateur 2026-08-12) : la redirection vers la plage large produit des
        #    #VALUE! dans tout le TCD (constate en test), alors que les colonnes source
        #    ('Facture Chronopost') sont elles-memes correctes -- signe d'un TCD dont la
        #    configuration RowField/ColField/DataField (4 niveaux : Numero LT>Produit>TVA>Zone
        #    Tarifaire en lignes, Categories en colonnes) est FRAGILE a un changement de
        #    largeur de plage, meme piege deja documente pour le TCD imbrique 'Bilan clients'
        #    de Mondial Relay. Le TCD 'TCD' reste donc sur sa plage figee d'origine du modele
        #    (affiche des valeurs OBSOLETES du mois modele, PAS le mois traite) -- accepte
        #    car aucun controle utilise ce TCD precisement (reconciliation PDF passe par
        #    'Contrôle pdf', poids par 'TCD poids', tous deux fonctionnels).
        redirect_pivot_caches(wb, ws, ["TCD poids", "Contrôle pdf"], "Facture Chronopost",
                               max(newLast, 2), FIRST_RAW_COL, LAST_CALC_COL)
        wb.RefreshAll()
        try:
            xl.CalculateUntilAsyncQueriesDone()
        except Exception:
            pass
        xl.Calculate()

        # 3) "Fichier import" : le carrier Node a DEJA reclasse/calcule les lignes (postes ERP,
        #    gazole reparti, mapping mode envoi/zone) -- valide sur les 7 PDF de juin 2026 a
        #    0,00€ d'ecart. On lit son resultat depuis un fichier CSV intermediaire (colle par
        #    server.js AVANT d'appeler ce finaliseur -- cf. buildArgs) : NON, en realite ce
        #    script est appele avec les fichiers BRUTS, pas les resultats deja calcules --
        #    on reproduit ICI la meme logique de reclassement en Python (dupliquee depuis
        #    src/carriers/chronopost/index.js, a garder synchronisee) pour ecrire "Fichier
        #    import" en VALEURS (comme "Import CSV" de BLS), la Date Validite en premiere
        #    colonne mise a jour dynamiquement (meme piege que DPD/Mondial Relay/BLS).
        wsImp = wb.Sheets("Fichier import")
        if wsImp.AutoFilterMode:  # meme piege que "Facture Chronopost" ci-dessus
            wsImp.AutoFilterMode = False
        oldLastImp = wsImp.Cells(wsImp.Rows.Count, 6).End(xlUp).Row
        LAST_COL_IMPORT = 23

        import json
        biblio_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "facturation-app", "src", "carriers", "chronopost", "config.json")
        with open(biblio_path, encoding="utf-8") as f:
            cfg = json.load(f)

        def is_forfaitaire(numero_lt):
            s = str(numero_lt or "").strip().upper()
            return bool(re.match(r"^(CAP|ECO|SUR)", s))

        def type_forfaitaire(numero_lt):
            s = str(numero_lt or "").strip().upper()
            if s.startswith("CAP"):
                return "CAP"
            if s.startswith("ECO"):
                return "ECO"
            if s.startswith("SUR"):
                return "SUR"
            return None

        i_type_prestation = col_index(header, "Type prestation")
        i_tva_col = col_index(header, "TVA")
        i_zone_tarifaire = col_index(header, "Zone Tarifaire")
        i_poids = col_index(header, "Poids")
        i_montant_ht = col_index(header, "Montant HT")

        lignes_normales, lignes_forfaitaires = [], []
        for r in all_rows:
            facture = str(r[i_facture] or "").strip()
            numero_lt = str(r[i_numero_lt] or "").strip()
            montant_ht = to_num(r[i_montant_ht])
            if not facture or not numero_lt:
                continue
            if is_forfaitaire(numero_lt):
                if type_forfaitaire(numero_lt) == "CAP":
                    lignes_forfaitaires.append({"facture": facture, "montant_ht": montant_ht})
                continue
            lignes_normales.append({
                "facture": facture, "sous_compte": str(r[i_sous_compte] or "").strip(),
                "numero_lt": numero_lt,
                "type_prestation": str(r[i_type_prestation] or "").strip() if i_type_prestation is not None else "",
                "zone_tarifaire": str(r[i_zone_tarifaire] or "").strip() if i_zone_tarifaire is not None else "",
                "poids": to_num(r[i_poids]) if i_poids is not None else 0,
                "produit": str(r[i_produit] or "").strip() if i_produit is not None else "",
                "montant_ht": montant_ht,
                "pays_depart": str(r[i_pays_depart] or "").strip() if i_pays_depart is not None else "",
                "pays_arrivee": str(r[i_pays_arrivee] or "").strip() if i_pays_arrivee is not None else "",
            })

        pool_gazole = {}
        for l in lignes_forfaitaires:
            pool_gazole[l["facture"]] = pool_gazole.get(l["facture"], 0.0) + l["montant_ht"]
        total_fret = {}
        for l in lignes_normales:
            total_fret[l["facture"]] = total_fret.get(l["facture"], 0.0) + l["montant_ht"]

        POSTE_KEYS = ["Adresse", "Assurance", "Colis volumineux", "Corse", "Droits et taxes", "Frais facturation", "Frêt", "Gazole", "Zones éloignées"]

        def mapping_erp(produit, zone_tarifaire, pays_arrivee, pays_depart):
            entry = cfg["bibliotheque"].get(produit)
            if not entry:
                return {"modeEnvoi": "inconnu", "zone": "inconnu", "transporteur": "inconnu"}
            zone = entry["zone"]
            if produit == "6B":
                z = cfg["zoning_2shop"].get(pays_arrivee.upper())
                zone = str(z) if z is not None else "inconnu"
            elif produit == "6C":
                z = cfg["zoning_2shop_6c"].get(pays_depart.upper())
                zone = str(z) if z is not None else "inconnu"
            elif produit in ("17", "44"):
                zone = f"{produit}_{zone_tarifaire}"
            return {"modeEnvoi": entry["modeEnvoi"], "zone": zone, "transporteur": entry["transporteur"]}

        def ep_pour_mode(mode):
            # SEUL le mode envoi '1S' a une grille tarifaire Entreprise -- corrige par le pole
            # transport (2026-08-12) apres 64 avaries en juillet 2026 (E fixe errone avant).
            return "E" if str(mode or "").strip() == "1S" else "P"

        def forfait_zones_eloignees(montant_brut):
            # "Zones eloignees" export = forfait binaire (Corse + Zones eloignees brutes
            # fusionnees), PAS le montant brut -- verifie empiriquement contre le fichier
            # fait a la main de juillet 2026 (2026-08-12).
            if montant_brut <= 0:
                return None
            return 29 if montant_brut > 15 else 9.5

        def forfait_adresses(montant_brut):
            if montant_brut <= 0:
                return None
            return 17 if montant_brut > 8.5 else 8.5

        def round_up_1(x):
            # Poids arrondi au dixieme de kg SUPERIEUR (meme convention que les autres
            # transporteurs du projet -- roundUp1 cote Node) -- confirme empiriquement
            # (2026-08-13) contre le CSV fait a la main de juillet 2026 : match exact sur
            # 3917/3917 trackings (poids brut Chronopost a 3 decimales, ex. 4,638kg -> 4,7kg).
            import math
            return math.ceil(x * 10 - 1e-9) / 10

        produits_hors_contrat = set(cfg.get("produits_hors_contrat", []))
        n_hors_contrat = 0
        n_produit_inconnu = 0

        # Calcul par ligne brute (comme avant), PUIS regroupement par tracking (Numero LT)
        # avant d'ecrire import_data -- cf. meme fix cote Node (chronopost/index.js) : un
        # meme tracking a souvent plusieurs lignes dans le brut (1 "Transport" + N
        # "supplements", ex. "Zones Difficiles d'acces", poids=0 car pas un colis separe).
        # Le fichier fait a la main n'a QU'UNE SEULE ligne par tracking (confirme : 3917
        # trackings uniques = 3917 lignes du CSV reel de juillet 2026). BUG TROUVE 2026-08-13
        # (7e bug) : avant ce fix, chaque ligne du brut devenait une ligne d'import separee
        # (fausses alertes POIDS=0 + double comptage NbrColis potentiel cote ERP).
        lignes_calculees = []  # { numero_lt, m, postes, poids, produit }
        for l in lignes_normales:
            categorie = cfg["categories"].get(l["type_prestation"])
            tf = total_fret.get(l["facture"], 0.0)
            pg = pool_gazole.get(l["facture"], 0.0)
            gazole_ligne = (pg / tf) * l["montant_ht"] if tf > 0 else 0.0

            m = mapping_erp(l["produit"], l["zone_tarifaire"], l["pays_arrivee"], l["pays_depart"])
            if m["modeEnvoi"] == "inconnu":
                # Produit Chrono non reconnu dans "Bibliotheque transporteurs" (ex. produit "0"
                # vu en juillet 2026, EXP20260727-2876730, 11 trackings) -- Transporteur/mode
                # envoi/Zone tous "inconnu", inexploitable pour l'ERP. Confirme empiriquement
                # (2026-08-12) face a la reference faite a la main de juillet 2026 (3917 lignes
                # exactement, notre sortie en avait 3928 -- l'ecart = ces 11 trackings). Meme
                # traitement que les lignes forfaitaires CAP/ECO/SUR : exclu de l'import, pas
                # de tarif invente.
                n_produit_inconnu += 1
                continue

            # "sureté + eco" (colonne AA du modele) : forfait FIXE ajoute a TOUTE ligne normale
            # -- 0,08€ pour les modes 2SHOP (6B/6C/5X/5Y), 0,5€ sinon. Formule modele EXACTE :
            # AA=IF(mode∈{6C,6B,5X,5Y},0.08,0.5), AC(hors gazole)=Z+AA -- s'ajoute au Frêt de
            # CHAQUE ligne du CLASSEUR (fidele au modele Excel). BUG TROUVE 2026-08-12 (5e bug) :
            # ce forfait n'etait pas applique du tout, cause principale de l'ecart residuel de
            # ~1040€ sur juillet 2026.
            surete_eco_ligne = 0.08 if m["modeEnvoi"] in ("6B", "6C", "5X", "5Y") else 0.5

            postes = {k: 0.0 for k in POSTE_KEYS}
            if categorie in POSTE_KEYS:
                postes[categorie] = round(postes[categorie] + l["montant_ht"], 2)
            # "Frais facturation" fusionne dans "Frêt" (pas de colonne ERP dediee, meme trou
            # que "Corse" -- decision utilisateur 2026-08-12, precaution sans donnee reelle).
            if postes["Frais facturation"] > 0:
                postes["Frêt"] = round(postes["Frêt"] + postes["Frais facturation"], 2)
                postes["Frais facturation"] = 0.0
            postes["Frêt"] = round(postes["Frêt"] + surete_eco_ligne, 2)
            postes["Gazole"] = round(postes["Gazole"] + gazole_ligne, 2)

            if l["produit"] in produits_hors_contrat:
                n_hors_contrat += 1

            # "Pays arrivee" brut VIDE = envoi NATIONAL (Chronopost ne renseigne ce champ que
            # pour l'international, confirme sur ~2440 lignes/4148 de juillet 2026) -- PAS une
            # donnee manquante, meme principe que pays_tva du modele ("vide/non liste = 20%
            # par defaut"). Defaut "FR" applique ICI pour l'export (Pays/TVA), PAS sur
            # l["pays_arrivee"] passe a mapping_erp() plus haut (reste tel quel pour le
            # lookup zoning_2shop 6B/6C, qui n'a de toute facon aucune cle "FR").
            pays_arrivee_export = l["pays_arrivee"] or "FR"
            tva = cfg["pays_tva"].get(pays_arrivee_export.upper(), 0.2)

            lignes_calculees.append({
                "numero_lt": l["numero_lt"], "transporteur": m["transporteur"],
                "mode": m["modeEnvoi"], "zone": m["zone"], "pays": pays_arrivee_export,
                "tva": tva, "poids": l["poids"], "postes": postes,
            })

        # Regroupement par tracking : Poids/Zone/Mode/Transporteur/Pays/TVA pris sur la
        # premiere ligne avec poids>0 (la ligne "Transport"), postes CUMULES sur le groupe.
        groupes = {}
        ordre = []
        for lc in lignes_calculees:
            if lc["numero_lt"] not in groupes:
                groupes[lc["numero_lt"]] = []
                ordre.append(lc["numero_lt"])
            groupes[lc["numero_lt"]].append(lc)

        import_data = []
        for numero_lt in ordre:
            groupe = groupes[numero_lt]
            base = next((g for g in groupe if g["poids"] > 0), groupe[0])
            # Poids = SOMME de toutes les lignes du groupe (pas juste celui de "base") --
            # reproduit exactement le mecanisme du TCD Excel reel ("TCD poids", dataField=
            # Somme de Poids, verifie dans la video process 2026-08-13) : le TCD ne suppose
            # JAMAIS que les lignes supplement ont poids=0, il somme mecaniquement. Le
            # "base['poids']" seul donnait le meme resultat SEULEMENT parce que les lignes
            # supplement observees ont toujours eu poids=0 (verifie sur 3917 trackings de
            # juillet 2026, aucune exception) -- mais c'etait une supposition fragile, pas une
            # garantie structurelle (meme fix que cote Node). PAS de round(x,2) ici :
            # round_up_1 a besoin de la precision brute (poids Chronopost a 3 decimales) pour
            # arrondir correctement au dixieme superieur -- round(4.403,2)=4.4 puis
            # round_up_1(4.4)=4.4 est FAUX vs round_up_1(4.403)=4.5 direct.
            poids_somme = sum(g["poids"] for g in groupe)
            postes = {k: round(sum(g["postes"][k] for g in groupe), 2) for k in POSTE_KEYS}
            # Le forfait sureté+eco est ajoute A CHAQUE LIGNE BRUTE ci-dessus (fidele au
            # classeur/modele Excel), mais au niveau TRACKING il ne doit s'appliquer qu'UNE
            # SEULE FOIS -- sinon un tracking a N lignes se voit facturer N x le forfait.
            # Retire les (N-1) forfaits en trop du Frêt cumule (meme fix que cote Node).
            surete_eco_forfait = 0.08 if base["mode"] in ("6B", "6C", "5X", "5Y") else 0.5
            postes["Frêt"] = round(postes["Frêt"] - surete_eco_forfait * (len(groupe) - 1), 2)

            # Le gazole (pool CAP* reparti au prorata) N'EST PAS repercute au client dans le
            # fichier livre -- confirme empiriquement (2026-08-12) : la colonne "gazole" du
            # CSV reel de juillet 2026 est VIDE sur les 3917 lignes, et "Frêt" reel = notre
            # Frêt SANS le gazole (matches exacts). postes["Gazole"] reste calcule (utilise
            # ailleurs pour la reconciliation PDF sur le total BRUT) mais n'est jamais ecrit
            # dans le fichier d'import.
            import_data.append([
                base["transporteur"], date_validite_serial, "", "", "",
                numero_lt, "", ep_pour_mode(base["mode"]), base["pays"], base["zone"],
                1, round_up_1(poids_somme), base["mode"], base["tva"],
                postes["Droits et taxes"] or None, postes["Assurance"] or None,
                forfait_zones_eloignees(postes["Zones éloignées"] + postes["Corse"]),
                postes["Colis volumineux"] or None,
                forfait_adresses(postes["Adresse"]), postes["Frêt"] or None, None,
                None, None,
            ])
        if n_hors_contrat:
            print(f"AVERTISSEMENT: {n_hors_contrat} ligne(s) avec produit Chrono HORS CONTRAT "
                  f"({', '.join(sorted(produits_hors_contrat))}) — probable erreur de facturation Chronopost à réclamer.")
        if n_produit_inconnu:
            print(f"AVERTISSEMENT: {n_produit_inconnu} ligne(s) avec produit Chrono absent de "
                  f"'Bibliothèque transporteurs' (mode envoi/zone 'inconnu') — exclue(s) de l'import.")

        newLastImp = 1 + len(import_data)
        if import_data:
            retry(lambda: wsImp.Range(wsImp.Cells(2, 1), wsImp.Cells(newLastImp, LAST_COL_IMPORT)).__setattr__("Value", import_data))
        if newLastImp < oldLastImp:
            retry(lambda: wsImp.Range(wsImp.Cells(newLastImp + 1, 1), wsImp.Cells(oldLastImp, LAST_COL_IMPORT)).ClearContents())

        xl.Calculate()

        # 4) Réconciliation PDF (onglet "Contrôle pdf", colonne C = 'pdf' saisie a la main
        #    dans le modele, en face de No Facture en colonne A du TCD).
        try:
            ws_pdf = wb.Sheets("Contrôle pdf")
            lastA = ws_pdf.Cells(ws_pdf.Rows.Count, 1).End(xlUp).Row
            facture_to_row = {}
            for r in range(1, lastA + 1):
                v = str(ws_pdf.Cells(r, 1).Value or "").strip()
                if v and v not in ("Étiquettes de lignes", "(vide)", "Total général"):
                    facture_to_row[v] = r
            lastC = ws_pdf.Cells(ws_pdf.Rows.Count, 3).End(xlUp).Row
            if lastC >= 4:
                ws_pdf.Range(ws_pdf.Cells(4, 3), ws_pdf.Cells(lastC, 3)).ClearContents()
            matched = 0
            for p in pdfs:
                row = facture_to_row.get(p["facture"])
                if row is None:
                    print(f"AVERTISSEMENT: facture PDF {p['file']} ({p['facture']}) -> introuvable dans Contrôle pdf.")
                    continue
                ws_pdf.Cells(row, 3).Value = p["total_ht"]
                matched += 1
            print(f"Réconciliation PDF : {matched}/{len(pdfs)} facture(s) rapprochée(s) dans Contrôle pdf.")
        except Exception as e:
            print("Réconciliation Contrôle pdf ignorée :", e)

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
