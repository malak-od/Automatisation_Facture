#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
FINALISEUR BLS — produit "AAAA_MM_Facture BLS.xlsx" A L'IDENTIQUE du fichier
fait a la main (5 feuilles : Factures BLS, Bilan PDF, Import CSV, Bilan
client, Avoir ; memes formules, memes TCD, meme mise en forme), en partant du
fichier existant comme MODELE.

Source UNIQUE = la facture PDF recue (decision utilisateur 2026-08-11) -- le
xlsx brut equivalent contient la MEME donnee mais en mise en page facture
avec cellules fusionnees non alignees entre l'en-tete et les lignes de
donnees (source d'erreurs de parsing constatee en pratique, cf. historique
de ce fichier). Le texte PDF (extrait via pypdf) est un format lineaire plus
simple et plus fiable a parser : chaque ligne "Dossier" est identifiee par le
motif "DD/MM/AAAA NNNNNNNN ...", le libelle/montant pouvant etre sur cette
meme ligne texte ou sur la ligne suivante (retour a la ligne PDF, ex.
"V/Réf ..." ou juste vide) -- fusionnes en un seul item. Une ligne "Dossier"
SANS AUCUN montant trouvable est un montant reellement NUL (verifie en
comparant la somme au Total HT officiel de la facture : la somme des autres
lignes egale deja EXACTEMENT le total), PAS une ligne a ignorer -- confirme
sur juillet 2026 (Dossier 13301 "Vannes", montant 0).

Contrairement a DPD/Geodis/Mondial Relay, "Factures BLS" n'a AUCUNE formule
calculee -- ce sont des valeurs litterales (copiees-collees a la main depuis
la facture recue dans le process reel, ici deja extraites/reclassees en amont
par ce script/le carrier Node -- src/carriers/bls/index.js, meme logique de
parsing PDF, a garder synchronisee). "ID Client"/"Impact CO2"/"Prix Unitaire"
restent VIDES (jamais remplis, ni dans le modele ni dans les 2 videos process
-- aucun mecanisme de determination identifie pour "ID Client", rempli a la
main par le pole transport apres coup dans le classeur reel).

"Import CSV" : formules FIXES referencant "Factures BLS" ligne a ligne
(F=Dossier, G=Libelle, T=Montant H,T,), + valeurs fixes (Transporteur/E-P/
Pays/Zone/mode envoi), + B ("Date validite tarif") et N ("TVA") en CHAINE de
formules "=B{n-1}"/"=N{n-1}" a partir d'une valeur FIXE en ligne 2 (meme piege
que "Date validite tarif" sur DPD/Mondial Relay) -- B2 mis a jour vers le 1er
jour du mois traite avant le FillDown. K/L (Nbr Colis/Poids) : VALEURS SAISIES
A LA MAIN (RECHERCHEX cote pole transport vers l'export "AffreTrans", non
disponible ici) -- laissees VIDES, a completer manuellement comme le fait deja
le pole transport pour les clients/dossiers inconnus.

"Bilan PDF" (1 TCD, source 'Factures BLS'!B1:K1048576 -- colonne OUVERTE,
s'etend automatiquement, PAS le piege de plage figee/etroite deja rencontre
sur DPD/Geodis/Mondial Relay) : RefreshAll() suffit, pas besoin de rediriger
le PivotCache. Reconciliation PDF manuelle (D4 = total PDF colle, E4 = ecart
formule deja presente dans le modele).

"Bilan client" (1 TCD, source 'Factures BLS'!A1:K1048576, meme motif "colonne
ouverte") : idem, RefreshAll() suffit. RowField = "ID Client", qui reste vide
dans "Factures BLS" -> ce TCD affichera un seul groupe "(vide)" tant que "ID
Client" n'est pas rempli manuellement (fidele au process reel, PAS une
regression -- confirme par les 2 videos process : ID Client vide de bout en
bout, y compris cote portail AffreTrans).

Necessite : Windows + Excel + pywin32 + pypdf.
Usage :
  python finaliser_bls.py "<modele.xlsx>" "<sortie.xlsx>" --pdf <pdf1> [<pdf2>...]
"""
import sys, os, shutil, re


def to_num(v):
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip().replace(" ", "").replace("\xa0", "").replace(",", ".")
    return float(s)


def to_excel_serial_from_ddmmyyyy(date_str):
    """'DD/MM/AAAA' -> nombre serial Excel (jours depuis 1899-12-30)."""
    import datetime as _dt
    m = re.match(r"^(\d{2})/(\d{2})/(\d{4})$", date_str)
    if not m:
        return None
    d, mo, y = int(m.group(1)), int(m.group(2)), int(m.group(3))
    EXCEL_EPOCH = _dt.datetime(1899, 12, 30)
    return (_dt.datetime(y, mo, d) - EXCEL_EPOCH).days


def parse_montant(s):
    """Nombre 'français' toléré (espaces normaux/insécables comme séparateur de
    milliers, virgule OU point comme décimale) -> float, ou None si pas un montant."""
    t = str(s or "").strip().replace(" ", "").replace("\xa0", "")
    if not re.fullmatch(r"\d+([.,]\d{1,2})?", t):
        return None
    return to_num(t.replace(",", "."))


def match_date_dossier(line):
    """Découpe une ligne de données PDF en (date, dossier, reste) si elle commence
    par 'DD/MM/AAAA NNNNNNNN' (dossier sur 8 chiffres tel qu'imprimé dans le PDF)."""
    m = re.match(r"^(\d{2}/\d{2}/\d{4})\s+0*(\d+)\s*(.*)$", line.strip())
    if not m:
        return None
    return m.group(1), m.group(2), m.group(3)


def parse_reste_ligne(reste):
    """Extrait libelle/unite/quantite/montant/code_tva depuis la fin d'une ligne de
    type 'De X A Y [M QQ.QQ] MMM.MM C' -- unite+quantite optionnelles (absentes sur
    certaines lignes), montant+code_tva optionnels (ligne 'Dossier' sans montant ->
    montant=None)."""
    t = reste.strip()
    libelle, montant, code_tva = t, None, ""
    m_end = re.match(r"^(.*?)(?:\s+([\d ]+[.,]\d{1,2})\s+(\d+))?$", t)
    if m_end and m_end.group(2) is not None:
        libelle = m_end.group(1).strip()
        montant = parse_montant(m_end.group(2))
        code_tva = m_end.group(3)
    unite, quantite = "", 0
    m_unite = re.match(r"^(.*?)\s+([A-Z])\s+([\d ]+[.,]\d{1,2})$", libelle)
    if m_unite:
        libelle = m_unite.group(1).strip()
        unite = m_unite.group(2)
        quantite = parse_montant(m_unite.group(3)) or 0
    return {"libelle": libelle, "unite": unite, "quantite": quantite, "montant": montant, "code_tva": code_tva}


def extract_n_facture(text):
    """Numero de facture BLS : bloc 'N Client'+7 chiffres+date+folio colles/espaces
    ensemble dans l'en-tete PDF (ex. '01LARUCH2601856 31/07/2026 1' -> '2601856')."""
    m = re.search(r"[A-Z0-9]{5,}(\d{7})\s*\d{2}/\d{2}/\d{4}\s*\d+", text)
    return m.group(1) if m else None


def extract_total_ht(text):
    """Total HT officiel : 1re des 3 lignes-nombres isolees juste avant '...EUR' en
    fin de facture (Total HT / Montant TVA / Net a payer)."""
    m = re.search(r"^([\d ]+[.,]\d{1,2})\s*$\n^([\d ]+[.,]\d{1,2})\s*$\n^([\d ]+[.,]\d{1,2})\s*EUR\s*$", text, re.MULTILINE)
    return parse_montant(m.group(1)) if m else None


def parse_facture_pdf(pdf_path):
    """Parse le tableau de lignes 'Dossier' d'une facture PDF BLS -- source UNIQUE.
    Chaque ligne de données commence par 'DD/MM/AAAA NNNNNNNN' ; le reste (libellé +
    montant) peut être sur cette même ligne texte OU splitté sur la ligne suivante
    (retour à la ligne PDF, motif 'V/Réf ...' ou libellé vide) -- fusionné en un seul
    item. Une ligne 'Dossier' sans AUCUN montant trouvable est conservée avec
    montant=0 (pas ignorée)."""
    import pypdf
    text = "\n".join((page.extract_text() or "") for page in pypdf.PdfReader(pdf_path).pages)

    n_facture = extract_n_facture(text)
    total_ht = extract_total_ht(text)

    lines = text.split("\n")
    try:
        start_idx = next(i for i, l in enumerate(lines) if l.strip() == "FACTURE") + 1
    except StopIteration:
        start_idx = 0
    try:
        end_idx = next(i for i, l in enumerate(lines) if re.match(r"^Escompte pour r", l.strip(), re.I))
    except StopIteration:
        end_idx = len(lines)
    data_lines = lines[start_idx:end_idx]

    items = []
    pending = None  # {"date","dossier","libelle_prefix"} en attente d'un montant sur la ligne suivante
    for raw_line in data_lines:
        line = raw_line.strip()
        if not line:
            continue
        head = match_date_dossier(line)

        if head:
            if pending:
                items.append({"date": pending["date"], "dossier": pending["dossier"], "libelle": pending["libelle_prefix"],
                              "unite": "", "quantite": 0, "montant": 0.0, "code_tva": ""})
            date, dossier, reste = head
            parsed = parse_reste_ligne(reste)
            if parsed["montant"] is not None:
                items.append({"date": date, "dossier": dossier, "libelle": parsed["libelle"],
                              "unite": parsed["unite"], "quantite": parsed["quantite"],
                              "montant": parsed["montant"], "code_tva": parsed["code_tva"]})
                pending = None
            else:
                pending = {"date": date, "dossier": dossier, "libelle_prefix": parsed["libelle"]}
            continue

        if pending:
            parsed = parse_reste_ligne(line)
            libelle = f"{pending['libelle_prefix']}\n{parsed['libelle']}" if pending["libelle_prefix"] else parsed["libelle"]
            items.append({"date": pending["date"], "dossier": pending["dossier"], "libelle": libelle,
                          "unite": parsed["unite"], "quantite": parsed["quantite"],
                          "montant": parsed["montant"] if parsed["montant"] is not None else 0.0,
                          "code_tva": parsed["code_tva"]})
            pending = None
            continue
        # Ligne non rattachable -> ignorée.

    if pending:
        items.append({"date": pending["date"], "dossier": pending["dossier"], "libelle": pending["libelle_prefix"],
                      "unite": "", "quantite": 0, "montant": 0.0, "code_tva": ""})

    return {"file": os.path.basename(pdf_path), "n_facture": n_facture, "total_ht": total_ht, "items": items}


def parse_affretement_csv(csv_path):
    """Export CSV du portail 'AffreTrans' (';' separe, UTF-8 avec BOM). Cle de jointure :
    colonne 'Récépissé' = 'Dossier' BLS (confirme sur juin 2026 : 13/15 lignes retrouvees
    directement). Fournit 'ID Client' (deja le bon format), 'Poids (kg)', 'Nb palettes'.
    Filtre implicitement sur Transporteur='BLS'. Meme logique que
    src/carriers/bls/index.js/parseAffretementCsv, a garder synchronisee."""
    with open(csv_path, encoding="utf-8-sig") as f:
        lines = [l for l in f.read().split("\n") if l.strip("\r\n")]
    if not lines:
        return {}
    header = [h.strip() for h in lines[0].split(";")]
    try:
        i_recepisse = header.index("Récépissé")
    except ValueError:
        return {}
    i_id_client = header.index("ID Client") if "ID Client" in header else -1
    i_poids = header.index("Poids (kg)") if "Poids (kg)" in header else -1
    i_nb_palettes = header.index("Nb palettes") if "Nb palettes" in header else -1
    i_transporteur = header.index("Transporteur") if "Transporteur" in header else -1

    result = {}
    for line in lines[1:]:
        cols = line.rstrip("\r\n").split(";")
        if i_transporteur >= 0 and (cols[i_transporteur] if i_transporteur < len(cols) else "").strip().upper() != "BLS":
            continue
        recepisse = (cols[i_recepisse] if i_recepisse < len(cols) else "").strip()
        if not recepisse or recepisse in result:
            continue
        result[recepisse] = {
            "id_client": (cols[i_id_client] if 0 <= i_id_client < len(cols) else "").strip(),
            "poids": to_num(cols[i_poids]) if 0 <= i_poids < len(cols) and cols[i_poids].strip() else 0,
            "nb_palettes": to_num(cols[i_nb_palettes]) if 0 <= i_nb_palettes < len(cols) and cols[i_nb_palettes].strip() else 0,
        }
    return result


def parse_args(argv):
    modele, sortie = argv[1], argv[2]
    pdfs, affretements, cur = [], [], None
    for a in argv[3:]:
        if a == "--pdf":
            cur = "pdf"
        elif a == "--affretement":
            cur = "affretement"
        elif cur == "pdf":
            pdfs.append(a)
        elif cur == "affretement":
            affretements.append(a)
    return modele, sortie, pdfs, affretements


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
    modele, sortie, pdf_paths, affretement_paths = parse_args(sys.argv)
    shutil.copyfile(modele, sortie)  # on ne touche JAMAIS au modele

    if not pdf_paths:
        raise RuntimeError("Aucune facture PDF fournie (--pdf <pdf1> [<pdf2>...]).")
    if not affretement_paths:
        raise RuntimeError('Export "AffreTrans" manquant (obligatoire, --affretement <csv>) : sans lui, "Id client"/"Nbr Colis"/"Poids" ne peuvent pas être complétés.')
    factures = [parse_facture_pdf(p) for p in pdf_paths]
    all_items = []
    for f in factures:
        for it in f["items"]:
            it["n_facture"] = f["n_facture"] or ""
            # Navette = trajet INTERNE entre les 2 sites "21 Longvic" <-> "21 Créancey" (pas
            # seulement les lignes libellees "navette" -- meme regle que le carrier Node,
            # src/carriers/bls/index.js, a garder synchronisee -- decision utilisateur
            # 2026-08-12). Cout REEL paye a BLS mais PAS refacturable au client.
            libelle_low = (it["libelle"] or "").lower()
            it["is_navette"] = ("21 longvic" in libelle_low) and ("21 cr" in libelle_low and "ancey" in libelle_low)
            all_items.append(it)
    if not all_items:
        raise RuntimeError("Aucune ligne 'Dossier' trouvée dans les facture(s) fournies.")
    print(f"Entrée : {len(factures)} facture(s), {len(all_items)} ligne(s) 'Dossier'")

    # "Factures BLS"/"Import CSV" : lignes a montant 0 retirees (rien a facturer nulle part,
    # confirme = montant reellement nul -- cf. docstring), navette CONSERVEE et REFACTURABLE
    # au client (montant reel paye a BLS -- decision utilisateur 2026-08-24, remplace la regle
    # "Fret=0 pour navette" du 2026-08-12).
    items = [it for it in all_items if it["montant"] != 0]
    n_montant_nul = len(all_items) - len(items)
    if n_montant_nul:
        print(f"{n_montant_nul} ligne(s) à montant nul supprimée(s) (rien à facturer).")
    n_navette = sum(1 for it in items if it["is_navette"])
    if n_navette:
        montant_navette = round(sum(it["montant"] for it in items if it["is_navette"]), 2)
        print(f"{n_navette} ligne(s) navette ({montant_navette} EUR payés à BLS) : Frêt refacturé au client (montant réel), Poids=13200 (forfaitaire).")

    affretement_map = {}
    for p in affretement_paths:
        try:
            affretement_map.update(parse_affretement_csv(p))
        except Exception as e:
            print(f"AVERTISSEMENT: export affrètement {os.path.basename(p)} illisible ({e}).")
    if affretement_paths:
        items_hors_navette = [it for it in items if not it["is_navette"]]
        n_trouve = sum(1 for it in items_hors_navette if it["dossier"] in affretement_map)
        print(f"Export affrètement : {n_trouve}/{len(items_hors_navette)} ligne(s) complétée(s) (Id client/Poids/Nb palettes, hors navette).")

    # Réconciliation : Total HT officiel (extrait du PDF) vs somme des montants extraits --
    # même fichier, même source, donc un écart ici signale un bug de parsing.
    total_par_facture = {}
    for it in all_items:
        total_par_facture[it["n_facture"]] = round(total_par_facture.get(it["n_facture"], 0.0) + it["montant"], 2)
    for f in factures:
        if f["total_ht"] is None:
            print(f"AVERTISSEMENT: Total HT introuvable dans {f['file']}, réconciliation ignorée.")
            continue
        calc = total_par_facture.get(f["n_facture"])
        ecart = round((f["total_ht"] or 0) - (calc or 0), 2)
        print(f"Réconciliation {f['file']} ({f['n_facture']}) : PDF={f['total_ht']:.2f} calculé={calc:.2f} écart={ecart:+.2f}")

    date_validite_serial = None
    for it in all_items:
        s = to_excel_serial_from_ddmmyyyy(it["date"])
        if s is not None:
            import datetime as _dt
            EXCEL_EPOCH = _dt.datetime(1899, 12, 30)
            d = EXCEL_EPOCH + _dt.timedelta(days=s)
            date_validite_serial = (_dt.datetime(d.year, d.month, 1) - EXCEL_EPOCH).days
            break

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

        # 1) "Factures BLS" : purge + collage des valeurs (PAS de formules sur cette feuille
        #    dans le modele -- copie-colle a la main dans le process reel, cf. docstring).
        #    Colonnes : A=ID Client(vide) B=n facture C=Date Prestation D=Dossier E=Libelle
        #    F=Impact CO2(vide) G=Unite H=Quantite I=Prix Unitaire(vide) J=Montant H,T, K=Code Tva
        ws = wb.Sheets("Factures BLS")
        oldLast = ws.Cells(ws.Rows.Count, 4).End(xlUp).Row  # D = Dossier
        newLast = 1 + len(items)
        maxLast = max(oldLast, newLast, 2)
        retry(lambda: ws.Range(ws.Cells(2, 1), ws.Cells(maxLast, 11)).ClearContents())

        data = []
        for it in items:
            date_val = to_excel_serial_from_ddmmyyyy(it["date"])
            aff = affretement_map.get(it["dossier"])
            data.append([
                # A ID Client : TOUJOURS vide en sortie (decision utilisateur 2026-08-14) --
                # meme quand l'export AffreTrans fournit une valeur, ce n'est pas le vrai ID
                # client ERP attendu ici. Reste a saisir manuellement par le pole transport
                # apres coup (le TCD "Bilan client", RowField="ID Client", se reactualise de
                # lui-meme par simple RefreshAll une fois la colonne renseignee a la main).
                None,
                it["n_facture"],                    # B n facture
                date_val,                           # C Date Prestation
                it["dossier"],                      # D Dossier
                it["libelle"],                      # E Libelle
                None,                                # F Impact CO2 (vide)
                it["unite"] or None, it["quantite"] or None,  # G Unite, H Quantite
                None,                                # I Prix Unitaire (vide)
                round(it["montant"], 2),            # J Montant H,T,
                int(it["code_tva"]) if str(it["code_tva"]).isdigit() else 1,  # K Code Tva
            ])
        retry(lambda: ws.Range(ws.Cells(2, 1), ws.Cells(1 + len(data), 11)).__setattr__("Value", data))
        if newLast < oldLast:
            retry(lambda: ws.Range(ws.Cells(newLast + 1, 1), ws.Cells(oldLast, 11)).ClearContents())

        # Mise en forme : la ligne 2 du modele porte bordures/alignement/wrap text (ex. E =
        # Libelle en wrap text pour le motif "V/Ref...\n..." sur 2 lignes visuelles) -> copiee
        # sur TOUTES les lignes de donnees ecrites (le ClearContents ci-dessus efface le
        # contenu mais PAS la mise en forme d'origine, qui ne s'etend donc jamais au-dela de
        # la hauteur du modele clone -- cf. capture d'ecran pole transport, juillet 2026 :
        # tableau sans bordures/wrap au-dela de la ligne 16 d'origine).
        if len(data) > 1:
            retry(lambda: ws.Range(ws.Cells(2, 1), ws.Cells(2, 11)).Copy())
            retry(lambda: ws.Range(ws.Cells(3, 1), ws.Cells(1 + len(data), 11))
                  .PasteSpecial(Paste=-4122))  # xlPasteFormats
            xl.CutCopyMode = False
        # Hauteur de ligne auto (le wrap text de Libelle peut faire 2+ lignes visuelles selon
        # le motif "V/Ref...\n...").
        retry(lambda: ws.Range(ws.Cells(2, 1), ws.Cells(1 + len(data), 11)).EntireRow.AutoFit())

        # 2) "Import CSV" : formules FIXES (F/G/T referencent 'Factures BLS' ligne a ligne),
        #    valeurs fixes deja dans le modele (A/H/I/J/M), B2/N2 (Date validite/TVA) sont
        #    des valeurs FIXES chainees (B3="=B2", N3="=N2"...) -- B2 mis a jour vers le 1er
        #    du mois traite AVANT le FillDown (meme piege que DPD/Mondial Relay).
        wsImp = wb.Sheets("Import CSV")
        oldLastImp = wsImp.Cells(wsImp.Rows.Count, 6).End(xlUp).Row  # F = Tracking/Dossier
        if date_validite_serial is not None:
            wsImp.Cells(2, 2).Value = date_validite_serial
        else:
            print("AVERTISSEMENT: 'Date Prestation' introuvable -> 'Date validité tarif' (Import CSV!B2) non mise à jour, reste celle du modèle.")
        newLastImp = 1 + len(items)
        LAST_COL_IMPORT = 23  # A -> W (23 colonnes standard ERP)
        if newLastImp > 2:
            retry(lambda: wsImp.Range(wsImp.Cells(2, 1), wsImp.Cells(newLastImp, LAST_COL_IMPORT)).FillDown())
        if newLastImp < oldLastImp:
            retry(lambda: wsImp.Range(wsImp.Cells(newLastImp + 1, 1), wsImp.Cells(oldLastImp, LAST_COL_IMPORT)).ClearContents())
        # K/L (Nbr Colis/Poids) : VALEURS SAISIES A LA MAIN dans le modele (RECHERCHEX cote
        # pole transport vers l'export "AffreTrans") -- le FillDown ci-dessus les aurait
        # recopiees depuis la ligne 2 du modele (valeurs de juin, sans rapport avec le mois
        # traite) -> purgees d'abord, puis remplies via l'export affretement fourni (jointure
        # sur Dossier=Récépissé, meme logique que le carrier Node). Navette : Nbr Colis/Poids
        # FIXES forfaitaires (pas de jointure AffreTrans attendue, meme principe pour les
        # 2 champs -- decision utilisateur 2026-08-12). Les lignes hors-navette sans
        # correspondance restent vides, a completer manuellement comme avant.
        POIDS_NAVETTE_DEFAUT = 13200
        NBRCOLIS_NAVETTE_DEFAUT = 33
        if newLastImp >= 2:
            retry(lambda: wsImp.Range(wsImp.Cells(2, 11), wsImp.Cells(newLastImp, 12)).ClearContents())
            kl_data = []
            for it in items:
                aff = affretement_map.get(it["dossier"])
                if it["is_navette"]:
                    nb_colis, poids = NBRCOLIS_NAVETTE_DEFAUT, POIDS_NAVETTE_DEFAUT
                else:
                    nb_colis = round(aff["nb_palettes"]) if aff else None
                    poids = aff["poids"] if aff else None
                kl_data.append([nb_colis, poids])
            retry(lambda: wsImp.Range(wsImp.Cells(2, 11), wsImp.Cells(1 + len(kl_data), 12)).__setattr__("Value", kl_data))
        # T (Frêt) : formule ='Factures BLS'!J{n} (montant REEL, y compris navette) -> navette
        # REFACTURABLE au client depuis le 2026-08-24 (decision utilisateur, remplace la regle
        # "Fret=0 pour navette" du 2026-08-12, remontee pole transport : "etendre formule Fret,
        # ne pas supprimer pour navette") -- la formule FillDown suffit, plus de gel a 0.

        # 3) "Bilan PDF"/"Bilan client" : TCD dont le PivotCache pointe deja une plage OUVERTE
        #    ('Factures BLS'!B1:K1048576 / A1:K1048576 -- confirme sur le modele, contrairement
        #    au piege de plage figee/etroite deja rencontre sur DPD/Geodis/Mondial Relay) ->
        #    RefreshAll() suffit, pas besoin de rediriger le PivotCache.
        wb.RefreshAll()
        try:
            xl.CalculateUntilAsyncQueriesDone()
        except Exception:
            pass
        xl.Calculate()

        # 4) Réconciliation PDF (onglet "Bilan PDF", cellule D4 = total PDF colle a la main
        #    dans le modele, E4 = formule d'ecart deja presente, pas touchee).
        wsPdf = wb.Sheets("Bilan PDF")
        wsPdf.Cells(4, 4).ClearContents()  # D4
        valid_totals = [f for f in factures if f["total_ht"] is not None]
        if valid_totals:
            total = round(sum(f["total_ht"] for f in valid_totals), 2)
            wsPdf.Cells(4, 4).Value = total
            print(f"Réconciliation PDF -> Bilan PDF!D4 = {total}")
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
