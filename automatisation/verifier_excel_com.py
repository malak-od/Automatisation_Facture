# -*- coding: utf-8 -*-
"""Verifie que l'automatisation Excel COM fonctionne sur cette machine.

A LANCER DANS LA SESSION DU COMPTE DE SERVICE (celui qui executera le service
Windows `facturation`), et non sous le compte d'administration : les reglages
Excel qui conditionnent COM sont stockes par profil utilisateur (HKCU).

  python automatisation/verifier_excel_com.py                  # modele UPS (pire cas)
  python automatisation/verifier_excel_com.py "<autre.xlsx>"

Le modele UPS est le defaut volontairement : 47 Mo et ~200 000 formules
XLOOKUP, c'est le pire cas du parc. S'il passe, les onze autres passent.

Chaque etape est annoncee AVANT d'etre tentee. Si le script se fige sans rien
afficher de plus, la derniere ligne affichee designe l'etape bloquante -- c'est
le symptome d'une boite de dialogue Excel invisible (mode protege, ecran de
premier lancement non purge), qui n'emet ni exception ni timeout.

Voir deploy/DEPLOIEMENT.md 1.d pour la configuration prealable.
"""
import os
import sys
import time

RACINE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
MODELE_DEFAUT = os.path.join(RACINE, "Transporteurs", "UPS", "2026_06_Facture UPS.xlsx")


def etape(msg):
    """Annonce une etape et vide le tampon : indispensable pour que la derniere
    ligne visible designe l'etape ou le script s'est fige."""
    sys.stdout.write("... %s\n" % msg)
    sys.stdout.flush()


def echec(titre, *pistes):
    sys.stdout.write("\nECHEC : %s\n" % titre)
    for p in pistes:
        sys.stdout.write("  -> %s\n" % p)
    sys.stdout.flush()
    sys.exit(1)


def main():
    chemin = os.path.abspath(sys.argv[1]) if len(sys.argv) > 1 else MODELE_DEFAUT

    sys.stdout.write("Modele teste : %s\n" % chemin)
    if not os.path.exists(chemin):
        echec(
            "modele introuvable",
            "Le depot doit etre clone dans C:\\Projets\\automatisation-facturation",
            "Les 12 transporteurs codent leur modele en dur en ../Transporteurs/",
        )
    sys.stdout.write("Taille       : %.1f Mo\n\n" % (os.path.getsize(chemin) / 1048576.0))

    etape("import de win32com (pywin32)")
    try:
        import win32com.client as win32
    except ImportError:
        echec(
            "pywin32 absent",
            "pip install -r automatisation/requirements.txt",
        )

    t0 = time.time()
    etape("demarrage d'Excel (DispatchEx)")
    try:
        xl = win32.DispatchEx("Excel.Application")
    except Exception as e:
        echec(
            "Excel n'a pas demarre : %s" % e,
            "Excel est-il installe ET active ? (ospp.vbs /dstatus)",
            "Es-tu bien dans une session interactive ? COM ne marche pas en session 0",
        )
    xl.Visible = False
    xl.DisplayAlerts = False
    xl.AskToUpdateLinks = False
    t_start = time.time() - t0

    code = 0
    wb = None
    try:
        t0 = time.time()
        etape("ouverture du classeur (blocage ici = mode protege)")
        wb = xl.Workbooks.Open(chemin, UpdateLinks=0, ReadOnly=True)
        t_open = time.time() - t0

        # Le mode protege n'ouvre pas un Workbook mais une ProtectedViewWindow :
        # le code des finaliseurs echoue alors plus loin, sur un objet COM qui
        # n'a pas les memes membres. On le detecte ici, tant que c'est lisible.
        try:
            pv = xl.ProtectedViewWindows.Count
        except Exception:
            pv = 0
        if pv:
            echec(
                "classeur ouvert en MODE PROTEGE (%d fenetre(s))" % pv,
                "Excel > Fichier > Options > Centre de gestion de la confidentialite",
                "Emplacements approuves : ajouter C:\\Projets\\automatisation-facturation",
                "en cochant 'Les sous-dossiers de cet emplacement sont egalement approuves'",
            )

        etape("%d feuille(s) -- RefreshAll (TCD)" % wb.Sheets.Count)
        t0 = time.time()
        wb.RefreshAll()
        t_refresh = time.time() - t0

        etape("Calculate (recalcul des formules)")
        t0 = time.time()
        xl.Calculate()
        t_calc = time.time() - t0

        sys.stdout.write("\n--- Durees ---\n")
        sys.stdout.write("  demarrage Excel : %6.1f s\n" % t_start)
        sys.stdout.write("  ouverture       : %6.1f s\n" % t_open)
        sys.stdout.write("  RefreshAll      : %6.1f s\n" % t_refresh)
        sys.stdout.write("  Calculate       : %6.1f s\n" % t_calc)
        sys.stdout.write("  TOTAL           : %6.1f s\n" % (t_start + t_open + t_refresh + t_calc))
        sys.stdout.write("\nTOUT EST OK -- l'automatisation Excel fonctionne sur cette machine.\n")
        sys.stdout.write("Ce total est le pire cas du parc : il sert de reference pour les\n")
        sys.stdout.write("timeouts (900 s cote nginx ET cote NPM) et pour dimensionner les vCPU.\n")
    except SystemExit:
        raise
    except Exception as e:
        sys.stdout.write("\nECHEC pendant le traitement : %s\n" % e)
        code = 1
    finally:
        # Sans ce nettoyage, un EXCEL.EXE orphelin survit et fait echouer les
        # generations suivantes ("fichier deja ouvert dans Excel").
        try:
            if wb is not None:
                wb.Close(False)
        except Exception:
            pass
        try:
            xl.Quit()
        except Exception:
            pass

    sys.exit(code)


if __name__ == "__main__":
    main()
