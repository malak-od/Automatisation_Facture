#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
FINALISEUR UPS — produit "AAAA_MM_Facture UPS.xlsx" A L'IDENTIQUE du fichier
fait a la main (16 feuilles : Gazole, CODES SVCE LEVEL, CODIFICATION CODE
EXCEPTION, Charge.CHG_CODE, ST SV, Zone, Clients log, zone colis poids
assurance, Bilan factures, Facture UPS, Comptes UPS, TCD, Fichier import,
Demande avoir, Bilan clients, Adresse), en partant du fichier existant
comme MODELE.

Classeur reel :
  Facture UPS : 1 ligne = 1 evenement de facturation UPS Billing (PAS 1
    colis), donnees brutes CSV UPS (SANS EN-TETE, 250 colonnes, colonnes
    resolues PAR POSITION -- decalage CONFIRME -4 : position CSV brut N =
    colonne Facture UPS N+4), colonnes A->D CALCULEES : A="Clients"=vide
    (manuel), B="Montant assurance"=IF(codeDescription="EVS",valeurBase,0),
    C="Mode envoi"=IF(codeClasse="FRT",XLOOKUP(description,'ST SV'!A:A,B:B),0),
    D="Categorie"=cascade NB.SI('ST SV'!Q:Q/D:D)->Adresse/plus-value BtoC,
    puis codeClasse="FRT"->Fret, "TAX"->TVA, sinon XLOOKUP(codeDescription,
    'Charge.CHG_CODE'!A:C) -- "CODE INCONNU" est une VRAIE categorie du
    modele reel (239/472 codes), jamais reclassee.
  5 TCD (tous PivotCache source ETROITE/FIGEE sur le modele, meme piege
    que DPD/Geodis/Mondial Relay/Chronopost/TNT/FedEx) :
    - "Bilan factures" (2 TCD) : par compte, et par date+numero facture.
    - "zone colis poids assurance" : Max Nombre de colis/Zone/Assurance/
      Poids par tracking.
    - "ST SV" : Nombre de Categorie par tracking x Mode envoi (ST/SV/vide).
    - "TCD" : Somme Montant net par tracking x Categorie.
    - "Bilan clients" (source='TCD', PAS 'Facture UPS') : par client.
  Colonnes MANUELLES juxtaposees (etirees apres redirection PivotCache) :
    - "zone colis poids assurance"!B (Logistique)=COUNTIF('Clients log'!
      A:A,A), C (Colis)=IF(E=0,1,E), I (Poids UPS)=ROUNDUP(H,0), J (Poids
      UPS_COD)=IF(F>3,ROUNDUP(H,0),IF(H<10,ROUNDUP(H,1),ROUNDUP(H,0))).
    - "ST SV"!N = IF(K<>"","SV","ST").
    - "TCD"!B (Cout)=SUM(F:K)+O+M (controle), D (Poids)=XLOOKUP vers
      'zone colis poids assurance'!I. A (Logistique)/C (Client) manuels,
      jamais reconstruits (Client = regle transversale ID client, saisie
      humaine post-generation).
    - "Fichier import" : formules PAR LIGNE completes (cf. carrier Node
      index.js pour le detail complet des 24 colonnes, notamment
      Transporteur=compte extrait du tracking via 'Comptes UPS', Zone=
      FOURNIE par UPS (native, jamais calculee), E/P=cascade ERP/plus-
      value BtoC, TVA=IF(TCD!N="",0,0.2) (poste TVA reel, PAS liste de
      pays), Colis volumineux=bareme par palier sur TCD!H (poste ERP
      MONTANT, PAS le poids -- piege deja documente).

Colis 1Z79 (regle FACTURATION EXCEL.docx) : colis viticulteur retourne
  chez La Ruche -- EXCLUS de Facture UPS (jamais colles), a signaler pour
  demande d'avoir (infos console).

Trackings SANS AUCUNE charge facturable (tous postes ERP a 0, ligne
  "INF"/retour indelivrable isolee sans ligne FRT) : EXCLUS de "Fichier
  import" (confirme par comparaison au fichier reel de juin 2026 :
  8736/8736 lignes ont au moins un poste non nul, 0 ligne totalement
  vide) -- mais RESTENT dans "Facture UPS" (donnees brutes completes,
  la colonne "Facture UPS" du classeur modele n'est jamais filtree).

Necessite : Windows + Excel + pywin32.
Usage :
  python finaliser_ups.py "<modele.xlsx>" "<sortie.xlsx>" --csv <csv1> [<csv2>...] [--brut <brutM.xlsx> [<brutM-1.xlsx>]]
"""
import sys, os, shutil, re, csv


def normalize_header(h):
    return re.sub(r"\s+", " ", str(h or "")).strip()


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
    return v


def num(x):
    try:
        return float(str(x).replace(",", ".").strip())
    except (ValueError, AttributeError):
        return 0.0


# Positions CSV brut UPS (0-based) -- decalage CONFIRME -4 vs colonnes "Facture UPS" (1-based) :
# position CSV brut N (1-based) = colonne Facture UPS (N+4). Formule uniforme : <col FU> - 4 - 1.
COL = {
    "numero_compte": 7 - 4 - 1,
    "date_facture": 9 - 4 - 1,
    "numero_facture": 10 - 4 - 1,
    "ref1": 20 - 4 - 1,
    "numero_suivi": 25 - 4 - 1,
    "nombre_colis": 23 - 4 - 1,
    "poids_facture": 33 - 4 - 1,
    "zone": 38 - 4 - 1,
    "code_classe": 48 - 4 - 1,
    "code_description": 49 - 4 - 1,
    "description": 50 - 4 - 1,
    "valeur_base": 53 - 4 - 1,
    "montant_net": 57 - 4 - 1,
    "pays": 86 - 4 - 1,
}


def read_ups_csv(path):
    """CSV UPS Billing brut, SANS EN-TETE, separateur virgule, latin-1."""
    with open(path, encoding="latin-1", newline="") as f:
        rows = list(csv.reader(f))
    return [r for r in rows if any(v != "" for v in r)]


def load_config(carrier_dir):
    import json
    with open(os.path.join(carrier_dir, "config.json"), encoding="utf-8") as f:
        return json.load(f)


def poids_arrondi(transporteur, nb_colis, poids):
    import math
    if transporteur == "UPS_COD":
        if nb_colis > 3:
            return math.ceil(poids)
        if poids < 10:
            return math.ceil(poids * 10) / 10
        return math.ceil(poids)
    return math.ceil(poids)


def colis_volumineux_montant(montant_reel):
    import math
    if not montant_reel:
        return 0
    if montant_reel < 3:
        return 3
    if montant_reel < 15:
        return 15
    if montant_reel < 50:
        return 35
    if montant_reel < 100:
        return 59
    if montant_reel < 150:
        return 177
    return math.ceil(montant_reel / 59) * 59


def parse_args(argv):
    """--csv <csv...> [--brut <xlsx...>] [--period AAAA_MM]"""
    modele, sortie = argv[1], argv[2]
    csvs, brut, cur = [], [], None
    period = None
    for a in argv[3:]:
        if a == "--csv":
            cur = "c"
        elif a == "--brut":
            cur = "b"
        elif a == "--period":
            cur = "p"
        elif cur == "c":
            csvs.append(a)
        elif cur == "b":
            brut.append(a)
        elif cur == "p":
            period = a
            cur = None
    return modele, sortie, csvs, brut, period


def load_brut_ep(brut_paths):
    """Map {tracking -> 'entreprise'/'particulier'} depuis l'export WMS partage (meme
    mecanisme que Delivengo/Geodis/DPD/FedEx) -- colonnes AP=PRO_TRACKING(41)/Q=
    DES_PARTICULIER(16), 0-based. Point relais -> entreprise."""
    import io
    import openpyxl
    m = {}
    for p in [bp for bp in brut_paths if bp]:
        # BUG TROUVE 2026-08-26 : openpyxl refuse un chemin SANS EXTENSION (InvalidFileException
        # "openpyxl does not support  file format") -- les fichiers uploades via multer sont
        # renommes en identifiant hexadecimal sans extension (piege deja connu, cf. BLS). Fix :
        # ouvrir en binaire et passer un buffer memoire (openpyxl accepte un objet fichier-like,
        # detecte le format par SIGNATURE ZIP interne, pas par le nom de fichier).
        with open(p, "rb") as f:
            buf = io.BytesIO(f.read())
        wb = openpyxl.load_workbook(buf, read_only=True)
        ws = wb[wb.sheetnames[0]]
        for row in ws.iter_rows(min_row=2, values_only=True):
            if len(row) > 41:
                t = str(row[41] or "").strip()
                v = str(row[16] or "").strip().lower()
                ep = "particulier" if v == "particulier" else ("entreprise" if v in ("entreprise", "point relais") else None)
                if t and ep and t not in m:
                    m[t] = ep
        wb.close()
    return m


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
    modele, sortie, csv_paths, brut_paths, period = parse_args(sys.argv)
    if not csv_paths:
        raise RuntimeError("Aucun CSV fourni (--csv <facture1.csv> [...]).")
    shutil.copyfile(modele, sortie)  # on ne touche JAMAIS au modele

    carrier_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "facturation-app", "src", "carriers", "ups")
    cfg = load_config(carrier_dir)

    all_rows = []
    for p in csv_paths:
        all_rows.extend(read_ups_csv(p))
    if not all_rows:
        raise RuntimeError("Fichier(s) UPS vide(s) ou illisible(s).")
    print(f"Entrée : {len(all_rows)} ligne(s) brute(s), {len(csv_paths)} fichier(s).")

    # Mois cible = choix utilisateur (--period, source de verite, decision 2026-08-20 --
    # BUG TROUVE 2026-08-26 : --period n'etait jamais transmis au finaliseur, "Date validite
    # tarif" restait calculee par majorite sur "Date de la facture" du CSV brut, peu fiable,
    # cf. carrier Node index.js meme fix) -- repli sur le mois majoritaire si absent.
    from collections import Counter
    import datetime as _dt
    EXCEL_EPOCH = _dt.datetime(1899, 12, 30)
    date_validite_serial = None
    mois_cible = None
    if period:
        m = re.fullmatch(r"(\d{4})_(\d{2})", period)
        if m:
            mois_cible = f"{m.group(1)}{m.group(2)}"
            date_validite_serial = (_dt.datetime(int(m.group(1)), int(m.group(2)), 1) - EXCEL_EPOCH).days
    if mois_cible is None:
        compte_mois = Counter()
        for r in all_rows:
            d = str(r[COL["date_facture"]] if len(r) > COL["date_facture"] else "").strip()
            m = re.match(r"^(\d{4})-(\d{2})-\d{2}$", d)
            if m:
                compte_mois[f"{m.group(1)}{m.group(2)}"] += 1
        if compte_mois:
            mois_cible = compte_mois.most_common(1)[0][0]
            annee, mois = int(mois_cible[:4]), int(mois_cible[4:6])
            date_validite_serial = (_dt.datetime(annee, mois, 1) - EXCEL_EPOCH).days

    # Colis 1Z79 (regle FACTURATION EXCEL.docx) : colis viticulteur retourne chez La Ruche --
    # reste dans "Facture UPS" (donnees brutes, decision utilisateur 2026-08-25 -- confirme
    # contre le fichier reel fait-main de juillet 2026), EXCLU uniquement du fichier import,
    # et agrege dans l'onglet "Demande avoir" (Tracking/Nb colis/Montant/Cause). BUG TROUVE
    # 2026-08-25 (test sur 14 trackings reels de juillet 2026, comparaison au fichier
    # fait-main) : "code_classe=FRT strict" etait FAUX -- Montant = somme de TOUTES les
    # lignes SAUF TVA (code_classe=TAX) et Taxe gazole (code_classe=FSC), confirme exact sur
    # 2/3 cas d'ecart reexamines (le 3e reste un ecart residuel inexplique, probable
    # ajustement manuel isole). Nb colis = FIGE A 1 par tracking (valeur du fait-main sur
    # 13/14 trackings testes, aucune formule de somme sur les colonnes brutes ne colle).
    # Lignes SANS Numero de suivi NI Numero de reference 1 (Ref1, colonne T) : demande
    # utilisateur 2026-08-25 -- supprimees de "Facture UPS" (jamais collees, meme regle que
    # le carrier Node index.js), aucune cle d'identification exploitable. Ref1 vide, "." OU
    # "null" en texte (valeurs non-renseignees frequentes cote UPS) compte comme "pas de reference".
    def ref_vide(v):
        v = (v or "").strip()
        return (not v) or set(v) == {"."} or v.lower() == "null"

    CAUSE_1Z79 = "Nouveau compte pour le renvoi des colis depuis LR  l'avoir est censé être remis automatiquement sans avoir besoin de le demander"
    demandes_avoir_1z79 = {}  # tracking -> montant
    n_sans_identification = 0
    n_montant_net_zero = 0
    lignes_retenues = []
    for r in all_rows:
        t = str(r[COL["numero_suivi"]] if len(r) > COL["numero_suivi"] else "").strip()
        montant_net = coerce(r[COL["montant_net"]] if len(r) > COL["montant_net"] else "")
        montant_net = montant_net if isinstance(montant_net, (int, float)) else 0
        if t.upper().startswith("1Z79"):
            code_classe = str(r[COL["code_classe"]] if len(r) > COL["code_classe"] else "").strip().upper()
            if code_classe not in ("TAX", "FSC"):
                demandes_avoir_1z79[t] = round(demandes_avoir_1z79.get(t, 0.0) + montant_net, 2)
            else:
                demandes_avoir_1z79.setdefault(t, 0.0)
            # Lignes a Montant net = 0 : demande utilisateur 2026-08-25 -- supprimees de
            # "Facture UPS" MEME pour les 1Z79 (BUG TROUVE 2026-08-25, cf. carrier Node
            # index.js pour le detail complet : le fait-main a 0/49 lignes 1Z79 a montant=0,
            # alors que le CSV brut en a 9/58 -- toutes supprimees). L'agregation "Demande
            # avoir" ci-dessus reste faite AVANT ce filtre.
            if montant_net == 0:
                n_montant_net_zero += 1
                continue
            lignes_retenues.append(r)
            continue
        ref1 = str(r[COL["ref1"]] if len(r) > COL["ref1"] else "").strip()
        if not t and ref_vide(ref1):
            n_sans_identification += 1
            continue
        # Lignes a Montant net = 0 (hors 1Z79, deja traite ci-dessus) : demande utilisateur
        # 2026-08-25 -- supprimees de "Facture UPS" (pas seulement du fichier import),
        # confirme sur reel (facture 202600782885 : 13513 lignes a montant=0 AVEC tracking
        # rempli, distinct du filtre "sans identification" ci-dessus).
        if montant_net == 0:
            n_montant_net_zero += 1
            continue
        lignes_retenues.append(r)
    if n_montant_net_zero:
        print(f"{n_montant_net_zero} ligne(s) à Montant net = 0 supprimée(s) de 'Facture UPS'.")
    if demandes_avoir_1z79:
        print(f"{len(demandes_avoir_1z79)} colis en 1Z79 (retour viticulteur chez La Ruche) exclus de l'import — reportés dans l'onglet 'Demande avoir' : {sorted(demandes_avoir_1z79.keys())}.")
    if n_sans_identification:
        print(f"{n_sans_identification} ligne(s) sans Numéro de suivi ni Numéro de référence 1 supprimée(s) (aucune clé d'identification exploitable).")

    ncol = max((len(r) for r in lignes_retenues), default=0)
    data_brut = [[coerce(v) for v in (r + [None] * ncol)[:ncol]] for r in lignes_retenues]

    ep_map = load_brut_ep(brut_paths) if brut_paths else {}
    if not brut_paths:
        print("AVERTISSEMENT: aucun export 'expéditions_brut' fourni -> E/P classé 'P' par défaut sauf plus-value BtoC détectée.")

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

        # 1) "Facture UPS" : purge + collage donnees brutes a partir de la colonne E (A-D =
        #    calculees : Clients/Montant assurance/Mode envoi/Categorie).
        ws = wb.Sheets("Facture UPS")
        if ws.AutoFilterMode:
            ws.AutoFilterMode = False
        FIRST_RAW_COL = 5  # colonne E
        LAST_RAW_COL = FIRST_RAW_COL - 1 + ncol
        LAST_CALC_COL = 4  # A->D
        oldLast = ws.Cells(ws.Rows.Count, FIRST_RAW_COL).End(xlUp).Row
        n = len(data_brut)
        newLast = 1 + n
        maxLast = max(oldLast, newLast, 2)
        purgeUntil = maxLast + 200
        retry(lambda: ws.Range(ws.Cells(2, 1), ws.Cells(purgeUntil, max(LAST_RAW_COL, LAST_CALC_COL))).ClearContents())

        # "Numero de suivi" est TOUJOURS numerique/alphanumerique sans zero de tete dans le
        # brut reel UPS -- pas de force NumberFormat texte necessaire (contrairement a TNT).
        retry(lambda: ws.Range(ws.Cells(2, FIRST_RAW_COL), ws.Cells(newLast, LAST_RAW_COL)).__setattr__("Value", data_brut))
        if newLast < oldLast:
            retry(lambda: ws.Range(ws.Cells(newLast + 1, 1), ws.Cells(oldLast, max(LAST_RAW_COL, LAST_CALC_COL))).ClearContents())

        # Colonnes A-D : formules du modele, resolues par POSITION FIXE (decalage +4 confirme).
        def col_letter(idx0):
            import openpyxl
            return openpyxl.utils.get_column_letter(idx0 + 1)

        col_code_classe = col_letter(FIRST_RAW_COL - 1 + COL["code_classe"])
        col_code_description = col_letter(FIRST_RAW_COL - 1 + COL["code_description"])
        col_description = col_letter(FIRST_RAW_COL - 1 + COL["description"])
        col_valeur_base = col_letter(FIRST_RAW_COL - 1 + COL["valeur_base"])

        # B (Montant assurance) : =IF(codeDescription="EVS",valeurBase,0)
        formula_b = f'=IF({col_code_description}{{row}}="EVS",{col_valeur_base}{{row}},0)'
        retry(lambda t=formula_b: ws.Range(ws.Cells(2, 2), ws.Cells(newLast, 2)).__setattr__("Formula", [[t.format(row=r)] for r in range(2, newLast + 1)]))

        # C (Mode envoi) : =IF(codeClasse="FRT",XLOOKUP(description,'ST SV'!A:A,B:B),0)
        formula_c = f'=IF({col_code_classe}{{row}}="FRT",_xlfn.XLOOKUP({col_description}{{row}},\'ST SV\'!A:A,\'ST SV\'!B:B),0)'
        retry(lambda t=formula_c: ws.Range(ws.Cells(2, 3), ws.Cells(newLast, 3)).__setattr__("Formula", [[t.format(row=r)] for r in range(2, newLast + 1)]))

        # D (Categorie) : cascade exacte du modele.
        formula_d = (
            f'=IF(COUNTIF(\'ST SV\'!Q:Q,{col_description}{{row}})<>0,"Adresse",'
            f'IF(COUNTIF(\'ST SV\'!D:D,{col_description}{{row}})<>0,"plus-value BtoC",'
            f'IF({col_code_classe}{{row}}="FRT","Frêt",'
            f'IF({col_code_classe}{{row}}="TAX","TVA",'
            f'IF(COUNTIF(\'Charge.CHG_CODE\'!A:A,{col_code_description}{{row}})=0,"code inconnu",'
            f'_xlfn.XLOOKUP({col_code_description}{{row}},\'Charge.CHG_CODE\'!A:A,\'Charge.CHG_CODE\'!C:C))))))'
        )
        retry(lambda t=formula_d: ws.Range(ws.Cells(2, 4), ws.Cells(newLast, 4)).__setattr__("Formula", [[t.format(row=r)] for r in range(2, newLast + 1)]))

        ws.Range(ws.Cells(1, 1), ws.Cells(newLast, max(LAST_RAW_COL, LAST_CALC_COL))).AutoFilter()

        # 2) 5 TCD : PivotCache redirige vers la vraie plage large de 'Facture UPS' (cache du
        #    modele fige sur une plage etroite -- meme piege que DPD/Geodis/Mondial Relay/
        #    Chronopost/TNT/FedEx). "Bilan clients" pointe sur 'TCD' (pas 'Facture UPS'),
        #    redirige separement APRES redimensionnement du TCD "TCD" (etape 3).
        newRangeFacture = ws.Range(ws.Cells(1, 1), ws.Cells(newLast, max(LAST_RAW_COL, LAST_CALC_COL)))
        for sheet_name in ("Bilan factures", "zone colis poids assurance", "ST SV", "TCD"):
            wsPivot = wb.Sheets(sheet_name)
            for i in range(1, wsPivot.PivotTables().Count + 1):
                pt = wsPivot.PivotTables(i)
                src = str(pt.PivotCache().SourceData)
                if "Facture UPS" not in src:
                    print(f"TCD '{pt.Name}' ({sheet_name}) : source '{src}' n'est pas 'Facture UPS' -> config préservée.")
                    continue
                newCache = wb.PivotCaches().Create(SourceType=xlDatabase, SourceData=newRangeFacture)
                pt.ChangePivotCache(newCache)
                print(f"TCD '{pt.Name}' ({sheet_name}) : PivotCache redirigé vers {newRangeFacture.Address}.")
        wb.RefreshAll()
        try:
            xl.CalculateUntilAsyncQueriesDone()
        except Exception:
            pass
        xl.Calculate()

        # 3) "zone colis poids assurance" colonnes manuelles B/C/I/J (etendues jusqu'a la
        #    vraie derniere ligne native, colonne D = Numero de suivi).
        wsZcp = wb.Sheets("zone colis poids assurance")
        lastZcp = wsZcp.Cells(wsZcp.Rows.Count, 4).End(xlUp).Row
        if lastZcp >= 2:
            formulas_zcp = {
                2: "=COUNTIF('Clients log'!A:A,A{row})",          # B Logistique
                3: "=IF(E{row}=0,1,E{row})",                        # C Colis
                9: "=ROUNDUP(H{row},0)",                            # I Poids UPS
                10: '=IF(F{row}>3,ROUNDUP(H{row},0),IF(H{row}<10,ROUNDUP(H{row},1),ROUNDUP(H{row},0)))',  # J Poids UPS_COD
            }
            for col_idx, tmpl in formulas_zcp.items():
                retry(lambda c=col_idx, t=tmpl: wsZcp.Range(wsZcp.Cells(2, c), wsZcp.Cells(lastZcp, c))
                      .__setattr__("Formula", [[t.format(row=r)] for r in range(2, lastZcp + 1)]))
            print(f"'zone colis poids assurance' : colonnes B/C/I/J étirées jusqu'à la ligne {lastZcp}.")

        # 4) "ST SV" colonne manuelle N (etendue jusqu'a la derniere ligne native, colonne H).
        wsStSv = wb.Sheets("ST SV")
        lastStSv = wsStSv.Cells(wsStSv.Rows.Count, 8).End(xlUp).Row
        if lastStSv >= 3:  # entete sur 2 lignes (constate sur le modele)
            retry(lambda: wsStSv.Range(wsStSv.Cells(3, 14), wsStSv.Cells(lastStSv, 14))
                  .__setattr__("Formula", [[f'=IF(K{r}<>"","SV","ST")'] for r in range(3, lastStSv + 1)]))
            print(f"'ST SV' : colonne N étirée jusqu'à la ligne {lastStSv}.")

        # 5) "TCD" colonne manuelle B/D (etendues jusqu'a la vraie derniere ligne native,
        #    colonne E = Numero de suivi). Colonnes A (Logistique)/C (Client) NON reconstruites
        #    (manuelles -- Client = regle transversale ID client, saisie humaine).
        wsTcd = wb.Sheets("TCD")
        lastTcd = wsTcd.Cells(wsTcd.Rows.Count, 5).End(xlUp).Row
        if lastTcd >= 3:  # entete sur 2 lignes
            formulas_tcd = {
                2: "=SUM(F{row}:K{row})+O{row}+M{row}",  # B Cout (controle)
                4: "=_xlfn.XLOOKUP(E{row},'zone colis poids assurance'!D:D,'zone colis poids assurance'!I:I)",  # D Poids
            }
            for col_idx, tmpl in formulas_tcd.items():
                retry(lambda c=col_idx, t=tmpl: wsTcd.Range(wsTcd.Cells(3, c), wsTcd.Cells(lastTcd, c))
                      .__setattr__("Formula", [[t.format(row=r)] for r in range(3, lastTcd + 1)]))
            print(f"'TCD' : colonnes B/D étirées jusqu'à la ligne {lastTcd}.")

        # 1Z79 EXCLUS du TCD "TCD" (donc de "Fichier import", qui s'etend sur lastTcd) --
        # decision utilisateur 2026-08-25 : reste dans "Facture UPS" (donnees brutes) mais pas
        # dans le fichier import. Le TCD etant un vrai PivotTable (source='Facture UPS'
        # redirigee ci-dessus), les 1Z79 y remonteraient automatiquement -- masques ici via
        # PivotItems du champ "Numéro de suivi" (RowField), pas via suppression de lignes
        # (casserait la structure du pivot). Volume negligeable (15 trackings distincts/mois
        # observe sur juin 2026), boucle simple sans impact perf.
        try:
            tcd_pt = wsTcd.PivotTables(1)
            for pf in tcd_pt.PivotFields():
                if pf.Orientation == 1 and "suivi" in str(pf.Name).lower():  # xlRowField
                    n_masques = 0
                    for pi in pf.PivotItems():
                        if str(pi.Name).upper().startswith("1Z79"):
                            try:
                                pi.Visible = False
                                n_masques += 1
                            except Exception as e:
                                print(f"TCD : impossible de masquer le tracking 1Z79 {pi.Name!r} :", e)
                    if n_masques:
                        print(f"'TCD' : {n_masques} tracking(s) 1Z79 masqué(s) (exclus de Fichier import).")
                    break
            wb.RefreshAll()
            xl.Calculate()
            lastTcd = wsTcd.Cells(wsTcd.Rows.Count, 5).End(xlUp).Row
        except Exception as e:
            print("Avertissement : masquage 1Z79 sur le TCD a échoué :", e)

        # "Bilan clients" (source='TCD', PAS 'Facture UPS') : redirige APRES redimensionnement
        # de "TCD" ci-dessus, sur toute la hauteur POSSIBLE (pas juste celle deja peuplee).
        for i in range(1, wb.Sheets("Bilan clients").PivotTables().Count + 1):
            pt = wb.Sheets("Bilan clients").PivotTables(i)
            src = str(pt.PivotCache().SourceData)
            if "TCD" not in src:
                continue
            newCache = wb.PivotCaches().Create(SourceType=xlDatabase, SourceData=wsTcd.Range(wsTcd.Cells(1, 1), wsTcd.Cells(max(lastTcd, 2), 16)))
            pt.ChangePivotCache(newCache)
        wb.RefreshAll()
        xl.Calculate()

        # 6) "Fichier import" : formules PAR LIGNE completes (24 colonnes, cf. docstring en
        #    tete + carrier Node index.js pour le detail complet). Nombre de lignes = nombre
        #    de trackings du TCD (lastTcd-1, entete 2 lignes).
        #
        # BUG TROUVE 2026-08-20 (meme piege deja documente sur TNT/Chronopost) : le ColField
        # natif du TCD ("Categorie") n'affiche QUE les postes REELLEMENT presents dans le mois
        # traite -- son ORDRE/POSITION (colonnes F/G/H...) depend des donnees, PAS fixe comme
        # dans le modele de juin. Un mois avec un poste supplementaire (ex. "code inconnu",
        # absent du modele de juin) DECALE tout ce qui suit -- coder TCD!I/N/O/H/F/J/K en dur
        # (comme le finaliseur de premier jet) casse silencieusement le mapping poste->colonne.
        # Fix : resout chaque lettre par NOM d'en-tete (ligne 2 du TCD natif), APRES le
        # RefreshAll deja fait plus haut.
        headerRowTcd = [normalize_header(wsTcd.Cells(2, c).Value) for c in range(6, 17)]  # F..P

        def find_tcd_col(nom):
            for i, h in enumerate(headerRowTcd):
                if h.lower() == nom.lower():
                    return col_letter(5 + i)  # colonne 6 = F
            return None

        col_adresse = find_tcd_col("Adresse")
        col_assurance = find_tcd_col("Assurance")
        col_colis_vol = find_tcd_col("Colis volumineux")
        col_droits_taxes = find_tcd_col("Droits et taxes")
        col_fret = find_tcd_col("Frêt")
        col_plus_value = find_tcd_col("plus-value BtoC")
        col_tva = find_tcd_col("TVA")
        col_zones_eloignees = find_tcd_col("Zones éloignées")
        missing = [n for n, c in (("Adresse", col_adresse), ("Assurance", col_assurance), ("Colis volumineux", col_colis_vol),
                                    ("Droits et taxes", col_droits_taxes), ("Frêt", col_fret), ("plus-value BtoC", col_plus_value),
                                    ("TVA", col_tva), ("Zones éloignées", col_zones_eloignees)) if c is None]
        if missing:
            print(f"INFO: poste(s) ERP absent(s) du TCD ce mois-ci (aucune ligne) : {', '.join(missing)} -- colonne(s) correspondante(s) non trouvée(s), formules correspondantes videes.")

        wsImp = wb.Sheets("Fichier import")
        if wsImp.AutoFilterMode:
            wsImp.AutoFilterMode = False
        oldLastImp = wsImp.Cells(wsImp.Rows.Count, 9).End(xlUp).Row
        LAST_COL_IMPORT = 24
        newLastImp = max(lastTcd - 1, 2)

        formulas_import = {
            2: '=IF(COUNTIF(\'zone colis poids assurance\'!D:D,I{row})=0,"inconnu",_xlfn.XLOOKUP(I{row},\'zone colis poids assurance\'!D:D,\'zone colis poids assurance\'!F:F))',  # C Zone (colonne interne de controle)
            4: '=IF(COUNTIF(\'Comptes UPS\'!A:A,RIGHT(LEFT(I{row},8),6))=0,"inconnu",_xlfn.XLOOKUP(RIGHT(LEFT(I{row},8),6),\'Comptes UPS\'!A:A,\'Comptes UPS\'!B:B))',  # D Transport
            6: "=_xlfn.XLOOKUP(I{row},'Facture UPS'!Y:Y,'Facture UPS'!T:T)",  # F Ref.1 (simplifie : pas de IF("","") ici, valeur brute)
            7: "=_xlfn.XLOOKUP(I{row},'Facture UPS'!Y:Y,'Facture UPS'!U:U)",  # G Ref.2
            9: "=TCD!E{tcdrow}",  # I N° Tracking
            11: '=IF(B{row}="particulier","P",IF(X{row}="","E","P"))',  # K E/P
            12: "=_xlfn.XLOOKUP(I{row},'Facture UPS'!Y:Y,'Facture UPS'!CH:CH)",  # L Pays
            13: '=IF(LEN(C{row})>2,C{row},IF(L{row}="FR","France",_xlfn.XLOOKUP(I{row},\'zone colis poids assurance\'!D:D,\'zone colis poids assurance\'!F:F)))',  # M Zone
            14: "=MAX(_xlfn.XLOOKUP(I{row},'zone colis poids assurance'!D:D,'zone colis poids assurance'!E:E),A{row})",  # N Nbr Colis
            15: '=IF(D{row}="UPS_COD",_xlfn.XLOOKUP(I{row},\'zone colis poids assurance\'!D:D,\'zone colis poids assurance\'!J:J),_xlfn.XLOOKUP(I{row},\'zone colis poids assurance\'!D:D,\'zone colis poids assurance\'!I:I))',  # O Poids
            16: '=IF(COUNTIF(\'ST SV\'!H:H,I{row})=0,"inconnu",_xlfn.XLOOKUP(I{row},\'ST SV\'!H:H,\'ST SV\'!N:N))',  # P mode envoi
        }
        # BUG TROUVE 2026-08-25 (signale par l'utilisateur : "Fichier import" n'a pas de
        # formules comme le fichier fait-main) : les formules ci-dessous referencaient le TCD
        # par LIGNE FIXE (TCD!col{tcdrow}) -- fige en valeurs plus bas AVANT suppression des
        # lignes vides pour eviter un decalage (une ligne supprimee aurait laisse les lignes
        # suivantes pointer vers le mauvais tracking du TCD). Fix : remplace par un XLOOKUP
        # dynamique sur le tracking (colonne I, deja utilise par les autres formules
        # ci-dessus) -- exactement le meme principe que les formules "zone colis poids
        # assurance"/"Facture UPS" deja en XLOOKUP par tracking, jamais par position. Une
        # formule qui cherche par tracking reste valide meme apres suppression de lignes ->
        # permet de GARDER des formules vivantes dans "Fichier import" (fidele au fait-main)
        # au lieu de figer en valeurs.
        def tcd_lookup(col):
            return f"_xlfn.XLOOKUP(I{{row}},TCD!E:E,TCD!{col}:{col})"

        if col_tva:
            formulas_import[17] = f'=IF({tcd_lookup(col_tva)}="",0,0.2)'  # Q TVA
        else:
            retry(lambda: wsImp.Range(wsImp.Cells(2, 17), wsImp.Cells(newLastImp, 17)).__setattr__("Value", 0))
        if col_droits_taxes:
            formulas_import[18] = f'=IF({tcd_lookup(col_droits_taxes)}=0,"",{tcd_lookup(col_droits_taxes)})'  # R Droits et taxes
        else:
            retry(lambda: wsImp.Range(wsImp.Cells(2, 18), wsImp.Cells(newLastImp, 18)).ClearContents())
        formulas_import[19] = "=IF(_xlfn.XLOOKUP(I{row},'zone colis poids assurance'!D:D,'zone colis poids assurance'!G:G)=0,\"\",MAX(10,ROUNDUP(0.02*_xlfn.XLOOKUP(I{row},'zone colis poids assurance'!D:D,'zone colis poids assurance'!G:G),2)))"  # S Assurance
        if col_zones_eloignees:
            # BUG TROUVE 2026-08-20 : =IF(TCD!col="","",40) (formule modele litterale) declenche
            # a tort le forfait 40EUR quand la seule ligne "Zones eloignees" du tracking a un
            # montant reel de 0EUR (ex. code ESP "Supplt zone enlev. etendue" facture a 0,00EUR
            # -- observe sur 763/8719 trackings reels de juin 2026) : le TCD affiche alors 0
            # (pas vide), donc TCD!col<>"" est vrai -> formule litterale donne 40 alors que le
            # fichier reel livre montre 0. Fix : traiter aussi 0 comme "absence de charge",
            # coherent avec le fait qu'un forfait de 40EUR n'a pas de sens pour un poste facture
            # 0EUR.
            formulas_import[20] = f'=IF(OR({tcd_lookup(col_zones_eloignees)}="",{tcd_lookup(col_zones_eloignees)}=0),"",40)'  # T Zones éloignées
        else:
            retry(lambda: wsImp.Range(wsImp.Cells(2, 20), wsImp.Cells(newLastImp, 20)).ClearContents())
        if col_colis_vol:
            cv = tcd_lookup(col_colis_vol)
            formulas_import[21] = (
                f'=IF({cv}=0,"",IF({cv}<3,3,'
                f'IF({cv}<15,15,IF({cv}<50,35,'
                f'IF({cv}<100,59,IF({cv}<150,177,'
                f'ROUNDUP({cv}/59,0)*59))))))'
            )  # U Colis volumineux
        else:
            retry(lambda: wsImp.Range(wsImp.Cells(2, 21), wsImp.Cells(newLastImp, 21)).ClearContents())
        if col_adresse:
            formulas_import[22] = f'=IF({tcd_lookup(col_adresse)}="","",11.5)'  # V Adresses
        else:
            retry(lambda: wsImp.Range(wsImp.Cells(2, 22), wsImp.Cells(newLastImp, 22)).ClearContents())
        if col_fret:
            formulas_import[23] = f'=IF({tcd_lookup(col_fret)}="","",ROUND({tcd_lookup(col_fret)},2))'  # W Frêt
        else:
            retry(lambda: wsImp.Range(wsImp.Cells(2, 23), wsImp.Cells(newLastImp, 23)).ClearContents())
        if col_plus_value:
            formulas_import[24] = f'=IF({tcd_lookup(col_plus_value)}="","",{tcd_lookup(col_plus_value)})'  # X plus-value BtoC
        else:
            retry(lambda: wsImp.Range(wsImp.Cells(2, 24), wsImp.Cells(newLastImp, 24)).ClearContents())

        for col_idx, tmpl in formulas_import.items():
            retry(lambda c=col_idx, t=tmpl: wsImp.Range(wsImp.Cells(2, c), wsImp.Cells(newLastImp, c))
                  .__setattr__("Formula", [[t.format(row=r, tcdrow=r + 2) for _ in [0]][0] for r in range(2, newLastImp + 1)]))

        if date_validite_serial is not None:
            wsImp.Cells(2, 5).Value = date_validite_serial
        else:
            print("AVERTISSEMENT: mois de facturation introuvable -> 'Fichier import'!E2 non mise à jour, reste celle du modèle.")
        if newLastImp > 2:
            retry(lambda: wsImp.Range(wsImp.Cells(3, 5), wsImp.Cells(newLastImp, 5))
                  .__setattr__("Formula", [["=E{prev}".format(prev=r - 1)] for r in range(3, newLastImp + 1)]))

        if newLastImp < oldLastImp:
            retry(lambda: wsImp.Range(wsImp.Cells(newLastImp + 1, 1), wsImp.Cells(oldLastImp, LAST_COL_IMPORT)).ClearContents())
        print(f"'Fichier import' : formules reconstruites jusqu'à la ligne {newLastImp}.")

        # Colonne B "E/P ERP" : ECRITE EN VALEUR (declaration ERP, export WMS m/m-1, meme
        # mecanisme que Delivengo/Geodis/DPD/FedEx -- PAS une formule Excel, la donnee vient
        # d'un fichier externe). Lue via la colonne I (N° Tracking, deja calculee ci-dessus)
        # apres Calculate() pour resoudre les formules TCD!E.
        xl.Calculate()
        if newLastImp >= 2 and ep_map:
            ep_values = []
            for r in range(2, newLastImp + 1):
                tracking = str(wsImp.Cells(r, 9).Value or "").strip()
                ep_values.append([ep_map.get(tracking, "")])
            retry(lambda: wsImp.Range(wsImp.Cells(2, 2), wsImp.Cells(newLastImp, 2)).__setattr__("Value", ep_values))
            print(f"'Fichier import' : colonne B (E/P ERP) renseignée depuis l'export WMS pour {sum(1 for v in ep_values if v[0])}/{newLastImp - 1} ligne(s).")
        else:
            print("AVERTISSEMENT: pas d'export WMS fourni -> colonne B (E/P ERP) laissée vide, E/P final se rabattra sur 'P' par défaut (sauf plus-value BtoC détectée).")

        xl.Calculate()

        # Trackings SANS AUCUNE charge facturable (tous postes ERP a 0 -- ligne "INF"/retour
        # indelivrable isolee sans ligne FRT) : EXCLUS de "Fichier import" (confirme par
        # comparaison au fichier reel de juin 2026, 0 ligne totalement vide sur 8736 -- decision
        # utilisateur 2026-08-20 : garder TOUTES les lignes dans "Facture UPS", filtrer
        # UNIQUEMENT "Fichier import"). BUG TROUVE 2026-08-25 (signale par l'utilisateur :
        # "Fichier import" n'a pas de formules comme le fichier fait-main) -- les formules des
        # postes ERP ci-dessus etaient figees en VALEURS ici avant suppression des lignes vides
        # (necessaire tant qu'elles referencaient le TCD par ligne fixe). Elles referencent
        # desormais le tracking par XLOOKUP (cf. plus haut) -- restent valides meme apres
        # suppression de lignes, donc PLUS BESOIN de figer : "Fichier import" garde des
        # formules vivantes, fidele au fichier fait-main.
        if newLastImp >= 2:
            # Lecture en BLOC (1 seul aller-retour COM au lieu de milliers) -- la boucle
            # cellule-par-cellule/Delete()-par-ligne precedente timeoutait (>590s sur 8700+
            # lignes, BUG PERF TROUVE 2026-08-20).
            rangeImp = wsImp.Range(wsImp.Cells(2, 1), wsImp.Cells(newLastImp, LAST_COL_IMPORT))
            valuesImp = rangeImp.Value  # tuple de tuples (lecture seule, formules NON figees)
            # Colonnes montant (postes ERP, 1-based -> index 0-based dans valuesImp) : R Droits
            # et taxes(18), S Assurance(19), T Zones eloignees(20), U Colis volumineux(21),
            # V Adresses(22), W Fret(23), X plus-value BtoC(24).
            POSTE_COLS_0BASED = [17, 18, 19, 20, 21, 22, 23]
            lignes_vides = [i for i, row in enumerate(valuesImp) if all((row[c] in (None, "", 0)) for c in POSTE_COLS_0BASED)]
            if lignes_vides:
                # Regroupe les indices contigus en plages pour minimiser le nombre d'appels
                # Delete() (une plage multi-lignes se supprime en 1 seul appel COM).
                plages = []
                debut = lignes_vides[0]
                prec = lignes_vides[0]
                for idx in lignes_vides[1:]:
                    if idx != prec + 1:
                        plages.append((debut, prec))
                        debut = idx
                    prec = idx
                plages.append((debut, prec))
                # Suppression de la DERNIERE plage vers la PREMIERE (evite tout decalage des
                # indices de lignes non encore traitees).
                for i0, i1 in reversed(plages):
                    r0, r1 = i0 + 2, i1 + 2  # index 0-based -> ligne Excel 1-based (+entete)
                    retry(lambda a=r0, b=r1: wsImp.Range(wsImp.Cells(a, 1), wsImp.Cells(b, LAST_COL_IMPORT)).EntireRow.Delete())
                newLastImp -= len(lignes_vides)
                print(f"'Fichier import' : {len(lignes_vides)} ligne(s) sans aucune charge facturable supprimée(s) (tous postes ERP à 0).")

        wsImp.Range(wsImp.Cells(1, 1), wsImp.Cells(max(newLastImp, 2), LAST_COL_IMPORT)).AutoFilter()
        xl.Calculate()

        # BUG TROUVE 2026-08-26 (signale par l'utilisateur : "la feuille ne doit pas trainer
        # apres la derniere ligne") : ClearContents() (etape precedente) vide le CONTENU des
        # lignes residuelles du modele clone (mise en forme/bordures gardees), donc Excel
        # continue de considerer ces lignes comme faisant partie de la feuille (UsedRange
        # gonfle bien au-dela de newLastImp -- constate : max_row=7261 pour 6914 lignes de
        # donnees reelles, 347 lignes de residu "vide" mais toujours dans la zone utilisee).
        # Fix : supprimer PHYSIQUEMENT les lignes au-dela de newLastImp (EntireRow.Delete(),
        # pas juste ClearContents) -- force Excel a reduire UsedRange a la vraie derniere ligne.
        wsImpUsedRange = wsImp.UsedRange
        wsImpUsedLastRow = wsImpUsedRange.Row + wsImpUsedRange.Rows.Count - 1
        if wsImpUsedLastRow > newLastImp:
            retry(lambda: wsImp.Range(
                wsImp.Cells(newLastImp + 1, 1), wsImp.Cells(wsImpUsedLastRow, LAST_COL_IMPORT)
            ).EntireRow.Delete())
            print(f"'Fichier import' : {wsImpUsedLastRow - newLastImp} ligne(s) residuelle(s) (mise en forme sans donnees) supprimee(s) au-dela de la ligne {newLastImp}.")
        xl.Calculate()

        # 7) "Demande avoir" (1Z79) : Tracking/Nb colis/Montant/Cause -- Factures/Poids/Mode
        #    livraison laisses vides (saisie manuelle du pole transport, jamais remplis meme
        #    dans le fichier fait-main). Purge d'abord les lignes du modele clone (mois
        #    precedent), puis ecrit les trackings 1Z79 du mois traite (demandes_avoir_1z79,
        #    calcule plus haut, filtre code_classe=FRT uniquement).
        try:
            wsAvoir = wb.Sheets("Demande avoir")
            oldLastAvoir = wsAvoir.Cells(wsAvoir.Rows.Count, 1).End(xlUp).Row
            if oldLastAvoir >= 2:
                retry(lambda: wsAvoir.Range(wsAvoir.Cells(2, 1), wsAvoir.Cells(oldLastAvoir, 7)).ClearContents())
            if demandes_avoir_1z79:
                rows_avoir = [
                    [tracking, None, 1, None, None, montant, CAUSE_1Z79]
                    for tracking, montant in sorted(demandes_avoir_1z79.items())
                ]
                retry(lambda: wsAvoir.Range(wsAvoir.Cells(2, 1), wsAvoir.Cells(1 + len(rows_avoir), 7))
                      .__setattr__("Value", rows_avoir))
                print(f"'Demande avoir' : {len(rows_avoir)} colis 1Z79 reportés.")
        except Exception as e:
            print("Avertissement : remplissage 'Demande avoir' a échoué :", e)

        xl.Calculate()
        retry(lambda: wb.Save())
        wb.Close(SaveChanges=True)

        # 8) Export "Fichier import" en VALEURS (pas en JS recalcule a part) : remontee
        # utilisateur 2026-08-26 -- les alertes affichees (POIDS=0 etc.) et le CSV livre
        # doivent refleter le VRAI classeur (formules XLOOKUP resolues), pas le calcul JS
        # separe du carrier Node (qui peut diverger, ex. Réf.1/Réf.2 jamais remplis cote JS).
        # Meme mecanisme que finaliser_colissimo.py/finaliser_fedex.py (2026-08-24/25) :
        # rouvrir le classeur SAUVEGARDE en lecture (etat stable, pas de valeur d'erreur COM
        # transitoire), lire les colonnes D->X (4->24 = les 21 colonnes standard ERP presentes
        # dans "Fichier import", en commencant par "Transporteur" -- BUG TROUVE 2026-08-26 :
        # demarrer a la colonne E (DateValidite) au lieu de D decalait TOUT l'export d'1
        # colonne vers la gauche par rapport a IMPORT_COLUMNS. TaxeGasoil/NbColis n'existent
        # pas dans cet onglet, resteront vides cote CSV final -- cf. IMPORT_COLUMNS, meme
        # convention que Colissimo).
        wbRead = None
        try:
            wbRead = retry(lambda: xl.Workbooks.Open(os.path.abspath(sortie), UpdateLinks=0, ReadOnly=True))
            wsImpRead = wbRead.Sheets("Fichier import")
            impLast = wsImpRead.Cells(wsImpRead.Rows.Count, 9).End(xlUp).Row  # colonne I = N Tracking
            if impLast >= 2:
                values = retry(lambda: wsImpRead.Range(wsImpRead.Cells(2, 4), wsImpRead.Cells(impLast, 24)).Value)
                # BUG TROUVE 2026-08-26 (constate en test reel) : une ligne fantome (Transporteur
                # ="inconnu", Tracking="0.0") apparaissait au tout debut de l'export -- meme
                # famille de bug que Colissimo/FedEx (cellule encore en etat transitoire au
                # moment de la lecture COM, malgre la reouverture apres sauvegarde). Filtre sur
                # le tracking (index 5 dans row = colonne I=9 moins le decalage colonne D=4).
                values = [row for row in values if str(row[5] or "").strip() not in ("", "0.0", "(vide)", "inconnu")]
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
                    print(f"AVERTISSEMENT: {n_err} cellule(s) en erreur COM lors de l'export Fichier import (valeurs) -- laissees vides, a verifier.")
                print(f"EXPORT_IMPORT_VALEURS:{export_path}")
        except Exception as e:
            print("Export Fichier import (valeurs) ignore :", e)
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
