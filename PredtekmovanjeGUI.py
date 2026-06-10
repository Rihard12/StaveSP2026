import os
import sys
import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext
import threading
import logging
import openpyxl
import xlwings as xw
from rapidfuzz import process
from openpyxl.utils import column_index_from_string, get_column_letter
import numpy as np
from datetime import datetime

def get_base_dir():
    """Vrne mapo kjer je .exe (pri buildu) oz. kjer je gui.py (pri razvoju)."""
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))

BASE_DIR = get_base_dir()

# ──────────────────────────────────────────────
#  GLOBALNE SPREMENLJIVKE (nastavljive v GUI)
# ──────────────────────────────────────────────
NAPOVEDI_MAPA     = os.path.join(BASE_DIR, "TESTNapovediPredtekmovanje")
REZULTATI_FILE    = os.path.join(BASE_DIR, "TEST - SP 2026 - predtekmovanje - rezultati.xlsx")
LOG_FILE          = os.path.join(BASE_DIR, "log.txt")

IME_CELL          = "I3"
DOMACI_COL        = "I"
GOST_COL          = "K"
TOCKE_COL         = "L"
IME_COLUMN        = "B"

TEKME_FILE        = os.path.join(BASE_DIR, "SP 2026 - tekme.xlsx")
TEKME_SHEET_IDX   = 0    # Indeks lista za tekme
SKUPINE_SHEET_IDX = 1    # Indeks lista za skupine

IMENA_VRSTICE     = {}

logging.basicConfig(
    filename=LOG_FILE,
    level=logging.INFO,
    format="%(asctime)s - %(message)s"
)


CONFIG_FILE = os.path.join(BASE_DIR, "config.json")

def nalozi_config():
    """Naloži nastavitve iz config.json, če obstaja."""
    global NAPOVEDI_MAPA, REZULTATI_FILE, TEKME_FILE, LOG_FILE
    global IME_CELL, DOMACI_COL, GOST_COL, TOCKE_COL, IME_COLUMN
    global TEKME_SHEET_IDX, SKUPINE_SHEET_IDX
    if not os.path.exists(CONFIG_FILE):
        return
    try:
        import json
        with open(CONFIG_FILE, "r", encoding="utf-8") as f:
            cfg = json.load(f)
        NAPOVEDI_MAPA     = cfg.get("NAPOVEDI_MAPA",     NAPOVEDI_MAPA)
        REZULTATI_FILE    = cfg.get("REZULTATI_FILE",    REZULTATI_FILE)
        TEKME_FILE        = cfg.get("TEKME_FILE",        TEKME_FILE)
        IME_CELL          = cfg.get("IME_CELL",          IME_CELL)
        DOMACI_COL        = cfg.get("DOMACI_COL",        DOMACI_COL)
        GOST_COL          = cfg.get("GOST_COL",          GOST_COL)
        TOCKE_COL         = cfg.get("TOCKE_COL",         TOCKE_COL)
        IME_COLUMN        = cfg.get("IME_COLUMN",        IME_COLUMN)
        TEKME_SHEET_IDX   = cfg.get("TEKME_SHEET_IDX",   TEKME_SHEET_IDX)
        SKUPINE_SHEET_IDX = cfg.get("SKUPINE_SHEET_IDX", SKUPINE_SHEET_IDX)
    except Exception as e:
        print(f"Opozorilo: ni bilo mogoče naložiti config.json: {e}")

def shrani_config():
    """Shrani trenutne nastavitve v config.json."""
    import json
    cfg = {
        "NAPOVEDI_MAPA":     NAPOVEDI_MAPA,
        "REZULTATI_FILE":    REZULTATI_FILE,
        "TEKME_FILE":        TEKME_FILE,
        "IME_CELL":          IME_CELL,
        "DOMACI_COL":        DOMACI_COL,
        "GOST_COL":          GOST_COL,
        "TOCKE_COL":         TOCKE_COL,
        "IME_COLUMN":        IME_COLUMN,
        "TEKME_SHEET_IDX":   TEKME_SHEET_IDX,
        "SKUPINE_SHEET_IDX": SKUPINE_SHEET_IDX,
    }
    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        json.dump(cfg, f, ensure_ascii=False, indent=2)

# Naloži config takoj ob zagonu
nalozi_config()

# ──────────────────────────────────────────────
#  POSLOVNA LOGIKA  (enako kot original)
# ──────────────────────────────────────────────

def vrni_tip(d, g):
    if d > g: return 1
    if d < g: return 2
    return 0

def napolni_slovar_imen():
    global IMENA_VRSTICE
    app = xw.App(visible=False)
    try:
        wb  = app.books.open(REZULTATI_FILE)
        ws  = wb.sheets[0]
        imena = ws.range(f"{IME_COLUMN}3").expand("down").value
        if imena:
            IMENA_VRSTICE = {ime.strip().title(): i for i, ime in enumerate(imena, start=3)}
        wb.close()
    finally:
        app.quit()

def obdelaj_tekmo(vrsticaNapovedi, goliDomaciRez, goliGostjeRez, stolpecRez, log_cb, progress_cb=None):
    tipRez = vrni_tip(goliDomaciRez, goliGostjeRez)
    files  = [f for f in os.listdir(NAPOVEDI_MAPA) if f.endswith(".xlsx") and not f.startswith("~")]
    skupaj = len(files)
    stat = {3: 0, 1: 0, 0: 0, "napake": 0}
    app = xw.App(visible=False)
    try:
        wb  = app.books.open(REZULTATI_FILE)
        ws  = wb.sheets[0]
        for i, file in enumerate(files):
            path = os.path.join(NAPOVEDI_MAPA, file)
            try:
                wbn = openpyxl.load_workbook(path, read_only=True, data_only=True)
                wsn = wbn.worksheets[0]
                ime  = wsn[IME_CELL].value.strip().title()
                vr   = IMENA_VRSTICE.get(ime)
                if vr is None: raise Exception("Ime ni najdeno!")
                gd = int(wsn[f"{DOMACI_COL}{vrsticaNapovedi}"].value)
                gg = int(wsn[f"{GOST_COL}{vrsticaNapovedi}"].value)
                tipN = vrni_tip(gd, gg)
                if gd == goliDomaciRez and gg == goliGostjeRez:
                    tocke = 3
                elif tipN == tipRez:
                    tocke = 1
                else:
                    tocke = 0
                ws[f"{stolpecRez}{vr}"].value = tocke
                wbn.close()
                stat[tocke] += 1
                logging.info(f"{ime} | {gd}:{gg} | {tocke} tock")
                log_cb(f"✓  {ime}: {tocke} točk")
            except Exception as e:
                stat["napake"] += 1
                logging.error(f"NAPAKA {file}: {e}")
                log_cb(f"✗  {file}: {e}")
            if progress_cb: progress_cb(int((i + 1) / skupaj * 100))
        wb.save(REZULTATI_FILE)
        wb.close()
    finally:
        app.quit()
    log_cb(f"\n── Povzetek ──────────────────────")
    log_cb(f"Obdelanih: {skupaj - stat['napake']}/{skupaj}  |  3 točke: {stat[3]}  |  1 točka: {stat[1]}  |  0 točk: {stat[0]}  |  Napake: {stat['napake']}")

def obdelaj_skupino(skupina, vrsticaNapoved, stolpecNapoved, stolpecRez, ekipe, log_cb, progress_cb=None):
    files = [f for f in os.listdir(NAPOVEDI_MAPA) if f.endswith(".xlsx") and not f.startswith("~")]
    ekipeSeznam = [x.strip().upper() for x in ekipe]
    skupaj = len(files)
    stat = {8: 0, 6: 0, 4: 0, 2: 0, 0: 0, "napake": 0}
    app = xw.App(visible=False)
    try:
        wb  = app.books.open(REZULTATI_FILE)
        ws  = wb.sheets[0]
        for i, file in enumerate(files):
            path = os.path.join(NAPOVEDI_MAPA, file)
            try:
                wbn = openpyxl.load_workbook(path, read_only=True, data_only=True)
                wsn = wbn.worksheets[0]
                ime = wsn[IME_CELL].value.strip().title()
                vr  = IMENA_VRSTICE.get(ime)
                if vr is None: raise Exception("Ime ni najdeno!")
                points = 0
                for idx in range(4):
                    napoved = wsn[f"{stolpecNapoved}{vrsticaNapoved + idx}"].value
                    napoved = str(napoved).strip()
                    match, procent = razreši_z_aliasi(napoved, ekipeSeznam, ekipe)
                    if procent < PRAG_VALIDACIJE:
                        log_cb(f"  ⚠  {ime}, mesto {idx+1}: '{napoved}' → '{match.title()}' ({procent:.0f}%) – preskočeno")
                        logging.warning(f"{ime} | Skupina {skupina} | '{napoved}' → '{match}' ({procent:.0f}%) pod pragom")
                        continue
                    if match == ekipeSeznam[idx]:
                        points += 2
                        col = get_column_letter(column_index_from_string(stolpecRez) + idx)
                        ws[f"{col}{vr}"].value = 2
                stat[points] = stat.get(points, 0) + 1
                logging.info(f"{ime} | Skupina {skupina} | {points} točk")
                log_cb(f"✓  {ime}: {points} točk")
                wbn.close()
            except Exception as e:
                stat["napake"] += 1
                logging.error(f"NAPAKA {file}: {e}")
                log_cb(f"✗  {file}: {e}")
            if progress_cb: progress_cb(int((i + 1) / skupaj * 100))
        wb.save(REZULTATI_FILE)
        wb.close()
    finally:
        app.quit()
    tocke_dist = ", ".join(f"{t}t: {n}" for t, n in sorted((k, v) for k, v in stat.items() if isinstance(k, int) and v > 0))
    log_cb(f"\n── Povzetek ──────────────────────")
    log_cb(f"Obdelanih: {skupaj - stat['napake']}/{skupaj}  |  {tocke_dist}  |  Napake: {stat['napake']}")

def napolni_imena(log_cb):
    vrstica = 3
    files = [f for f in os.listdir(NAPOVEDI_MAPA) if f.endswith(".xlsx") and not f.startswith("~")]
    app = xw.App(visible=False)
    try:
        wb  = app.books.open(REZULTATI_FILE)
        ws  = wb.sheets[0]
        for file in files:
            path = os.path.join(NAPOVEDI_MAPA, file)
            wbn  = None
            try:
                wbn = openpyxl.load_workbook(path, read_only=True, data_only=True)
                ime = wbn.worksheets[0][IME_CELL].value.strip().title()
                wbn.close()
                ws[f"{IME_COLUMN}{vrstica}"].value = ime
                vrstica += 1
                log_cb(f"✓  {file}")
            except Exception as e:
                if wbn: wbn.close()
                logging.error(f"NAPAKA {file}: {e}")
                log_cb(f"✗  {file}: {e}")
        wb.save(REZULTATI_FILE)
        wb.close()
    finally:
        app.quit()

    log_cb(f"✓  Imena uspešno prebrana!")
    napolni_slovar_imen()

def sortiraj_in_ostevilci(log_cb):
    app = xw.App(visible=False)
    try:
        wb       = app.books.open(REZULTATI_FILE)
        ws       = wb.sheets[0]
        last_row = ws.used_range.last_cell.row
        last_col = ws.used_range.last_cell.column
        # Izračunaj vsoto stolpcev D naprej in zapiši v C (zamenja =SUM formulo)
        for row in range(3, last_row + 1):
            if ws.range((row, 2)).value is None: continue
            vrednosti = ws.range((row, 4), (row, last_col)).value
            if isinstance(vrednosti, (int, float)):
                vrednosti = [vrednosti]
            vsota = sum(v for v in (vrednosti or []) if isinstance(v, (int, float)))
            ws.range((row, 3)).value = vsota

        igralci  = []
        for row in range(3, last_row + 1):
            vr = ws.range((row, 1), (row, last_col)).value
            if vr[1] is None: continue
            igralci.append(vr)
        igralci.sort(key=lambda x: x[2] if isinstance(x[2], (int, float)) else 0, reverse=True)
        for idx, vr in enumerate(igralci, start=3):
            ws.range((idx, 1), (idx, last_col)).value = vr
        trenutno = 1
        prej     = None
        for row in range(3, len(igralci) + 3):
            tocke = ws.range((row, 3)).value
            if row == 3:
                ws.range((row, 1)).value = "1."
            else:
                if tocke == prej:
                    ws.range((row, 1)).value = ""
                else:
                    trenutno = row - 2
                    ws.range((row, 1)).value = f"{trenutno}."
            prej = tocke
        wb.save(REZULTATI_FILE)
        wb.close()
        log_cb("✓  Lestvica uspešno sortirana in oštevilčena.")
    except Exception as e:
        log_cb(f"✗  Napaka: {e}")
    finally:
        app.quit()
    IMENA_VRSTICE.clear()
    napolni_slovar_imen()

def preberi_tekme_iz_excela():
    """Vrne seznam dict-ov iz lista Tekme v TEKME_FILE (vključno s številko vrstice)."""
    if not TEKME_FILE or not os.path.exists(TEKME_FILE):
        return []
    wb  = openpyxl.load_workbook(TEKME_FILE, read_only=True, data_only=True)
    ws  = wb.worksheets[TEKME_SHEET_IDX]
    tekme = []
    MESECI = {"jan":1,"feb":2,"mar":3,"apr":4,"maj":5,"jun":6,
               "jul":7,"avg":8,"sep":9,"okt":10,"nov":11,"dec":12,
               "jan":1,"feb":2,"mar":3,"apr":4,"may":5,"jun":6,
               "jul":7,"aug":8,"sep":9,"oct":10,"nov":11,"dec":12}

    def format_datum(datum):
        if not datum: return ""
        try:
            if hasattr(datum, "day"):  # datetime objekt
                return f"{datum.day}.{datum.strftime('%b')}"
            return str(datum).strip()
        except Exception:
            return str(datum).strip()

    def datum_kljuc(datum):
        if not datum: return (99, 99)
        try:
            if hasattr(datum, "month"):  # datetime objekt
                return (datum.month, datum.day)
            delen = str(datum).strip().replace(".", " ").split()
            dan = int(delen[0])
            mes = delen[1].lower()[:3]
            return (MESECI.get(mes, 99), dan)
        except Exception:
            return (99, 99)

    for row_idx, row in enumerate(ws.iter_rows(min_row=1, values_only=True), start=1):
        ekipa1, sep, ekipa2, obdelana, stolpec, vrstica_napovedi, datum = (row + (None,)*7)[:7]
        if ekipa1 is None: break
        tekme.append({"ekipa1": ekipa1, "ekipa2": ekipa2,
                       "obdelana": obdelana, "stolpec": stolpec,
                       "vrstica_napovedi": vrstica_napovedi,
                       "datum": datum,
                       "datum_str": format_datum(datum),
                       "vrstica_excela": row_idx})
    wb.close()
    tekme.sort(key=lambda t: datum_kljuc(t["datum"]))
    return tekme

def oznaci_tekmo_obdelano(vrstica_excela):
    """Zapiše 'D' v stolpec obdelana (D) za dano vrstico v TEKME_FILE."""
    if not TEKME_FILE or not os.path.exists(TEKME_FILE):
        return
    wb = openpyxl.load_workbook(TEKME_FILE)
    ws = wb.worksheets[TEKME_SHEET_IDX]
    # Stolpec D (indeks 4) je stolpec "obdelana"
    ws.cell(row=vrstica_excela, column=4).value = "D"
    wb.save(TEKME_FILE)
    wb.close()


def oznaci_skupino_obdelano(vrstica_excela):
    """Zapiše 'D' v stolpec obdelana (E) za dano skupino v TEKME_FILE."""
    if not TEKME_FILE or not os.path.exists(TEKME_FILE):
        return
    wb = openpyxl.load_workbook(TEKME_FILE)
    ws = wb.worksheets[SKUPINE_SHEET_IDX]
    ws.cell(row=vrstica_excela, column=5).value = "D"
    wb.save(TEKME_FILE)
    wb.close()

def preveri_rezultati_prazni():
    """
    Vrne (True, "") če je excel rezultatov prazen (nima imen ali so vse točke 0).
    Vrne (False, razlog) sicer.
    """
    if not os.path.exists(REZULTATI_FILE):
        return True, ""
    try:
        app = xw.App(visible=False)
        try:
            wb  = app.books.open(REZULTATI_FILE)
            ws  = wb.sheets[0]
            # Preveri, ali stolpec B sploh vsebuje kakšno ime
            ime_celica = ws.range(f"{IME_COLUMN}3").value
            if ime_celica is None:
                return True, ""
            # Preveri vsoto stolpca C od vrstice 3 naprej
            vrednosti_c = ws.range("C3").expand("down").value
            wb.close()
        finally:
            app.quit()
        if vrednosti_c is None:
            return False, "Stolpec B že vsebuje imena, točke so prazne."
        if isinstance(vrednosti_c, (int, float)):
            vrednosti_c = [vrednosti_c]
        vsota = sum(v for v in vrednosti_c if isinstance(v, (int, float)))
        if vsota != 0:
            return False, f"Stolpec B že vsebuje imena in vsota točk (stolpec C) je {vsota:.0f} ≠ 0."
        return False, "Stolpec B že vsebuje imena (točke so še 0)."
    except Exception as e:
        return False, f"Napaka pri preverjanju: {e}"

def preberi_skupine_iz_excela():
    """Vrne seznam dict-ov iz lista Skupine v TEKME_FILE."""
    if not TEKME_FILE or not os.path.exists(TEKME_FILE):
        return []
    wb  = openpyxl.load_workbook(TEKME_FILE, read_only=True, data_only=True)
    ws  = wb.worksheets[SKUPINE_SHEET_IDX]
    skupine = []
    for row_idx, row in enumerate(ws.iter_rows(min_row=1, values_only=True), start=1):
        skupina, stolpec, nap_stolpec, nap_vrstica, obdelana = (row + (None,)*5)[:5]
        if skupina is None: break
        skupine.append({"skupina": skupina, "stolpec": stolpec,
                        "nap_stolpec": nap_stolpec, "nap_vrstica": nap_vrstica,
                        "obdelana": obdelana, "vrstica_excela": row_idx})
    wb.close()
    return skupine

# ──────────────────────────────────────────────
#  VALIDACIJA NAPOVEDI SKUPIN
# ──────────────────────────────────────────────

SKUPINE_SP2026 = {
    "A": ["Mehika", "Južna Afrika", "Južna Koreja", "Češka"],
    "B": ["Kanada", "Bosna in Hercegovina", "Katar", "Švica"],
    "C": ["Brazilija", "Maroko", "Škotska", "Haiti"],
    "D": ["Združene države Amerike", "Paragvaj", "Avstralija", "Turčija"],
    "E": ["Nemčija", "Slonokoščena obala", "Ekvador", "Curacao"],
    "F": ["Nizozemska", "Japonska", "Švedska", "Tunizija"],
    "G": ["Belgija", "Egipt", "Iran", "Nova Zelandija"],
    "H": ["Španija", "Urugvaj", "Savdska Arabija", "Zelenortski otoki"],
    "I": ["Francija", "Senegal", "Irak", "Norveška"],
    "J": ["Argentina", "Alžirija", "Avstrija", "Jordan"],
    "K": ["Portugalska", "DR Kongo", "Uzbekistan", "Kolumbija"],
    "L": ["Anglija", "Hrvaška", "Gana", "Panama"],
}

# Alternativna imena za ekipe (kratice, okrajšave, pogovorni izrazi)
# Ključ je kanonično ime (mora biti enako kot v SKUPINE_SP2026), vrednost seznam aliasov
ALIASI_EKIP = {
    "BOSNA IN HERCEGOVINA": ["BIH", "BOSNA"],
    "ZDRUŽENE DRŽAVE AMERIKE": ["ZDA", "USA", "AMERIKA"],
    "SLONOKOŠČENA OBALA": ["IVORY COAST", "SLONOKOŠČENA O.", "CÔTE D'IVOIRE", "COTE D'IVOIRE", "SLONOK. OBALA", "SLONOKOSCEN O.", "SL.OBALA", "SLONOKOSCENA O."],
    "JUŽNA KOREJA": ["KOREJA", "KOREA", "J KOREJA", "J. KOREJA", "J.KOREJA", "SOUTH KOREA", "J. KOREA"],
    "JUŽNA AFRIKA": ["JAR", "RSA", "J. AFRIKA", "J AFRIKA", "J.AFRIKA", "JUŽNOAFRIŠKA REPUBLIKA", "SOUTH AFRICA", "J. AFRKA"],
    "NOVA ZELANDIJA": ["N ZELANDIJA", "N. ZELANDIJA", "N.ZELANDIJA", "N.ZELENDIJA", "N. ZELENDIJA"],
    "SAVDSKA ARABIJA": ["SAU.ARABIJA", "SAV. ARABIA", "S. ARABIJA", "SAV. ARABIJA", "SAV.ARABIJA", "S ARABIJA", "S.ARABIJA", "SAV ARABIJA", "SAU ARABIJA", "SA.ARABIJA", "SAUDI ARABIA", "SAU. ARABIJA"],
    "ZELENORTSKI OTOKI": ["ZELENORTSKI O.", "Z. OTOKI", "CAPE VERDE", "ZELENORTSKI.O", "ZELENO R OTOKI", "ZELENOERTSKI", "KAPVERDSKI OTOKI", "ZEL.OTOKI", "ZELENORTSKI.O."],
    "DR KONGO": ["DRC", "CONGO DR", "D.R. KONGO", "D.R.KONGO", "D. R. KONGO", "D.R.", "CONGO", "KONGO"],
    "CURACAO": ["CURAÇAO"],
    "FRANCIJA": ["FRANCOSKA"],
    "PORTUGALSKA": ["PORTUGALIJA", "PORTUGAL"],
    "NEMČIJA": ["NEMCIJA"],
    "ČEŠKA": ["CESKA", "ČESKA"],
    "JORDAN": ["JORDANIJA"],
    "ALŽIRIJA": ["ALĐERIJA"],
    "ŠVICA": ["SVICA"],
    "KATAR": ["QATAR"],
    "KANADA": ["CANADA"],
    "MAROKO": ["MAROCO"],
    "EGIPT": ["EGYPT"],
    "AVSTRIJA": ["AUSTRIA"],
    "HRVAŠKA": ["CROATIA"],
    "IRAK": ["IRAQ"],
    "ŠKOTSKA": ["SCOTLAND"],
}

def razreši_z_aliasi(vnos, ekipe_polna, ekipe_originalne):
    """
    Najprej preveri aliase – če vnos direktno ustreza aliasu, vrni 100%.
    Sicer vrni najboljši fuzzy match.
    ekipe_polna: seznam UPPER imen (za fuzzy)
    ekipe_originalne: seznam originalnih imen (za alias lookup)
    """
    vnos_upper = vnos.strip().upper()

    # Preveri aliase za vsako ekipo
    for i, ime_orig in enumerate(ekipe_originalne):
        aliasi = ALIASI_EKIP.get(ime_orig.upper(), [])
        for alias in aliasi:
            if vnos_upper == alias.upper():
                return ekipe_polna[i], 100.0

    # Fallback na fuzzy
    match, procent, _ = process.extractOne(vnos_upper, ekipe_polna)
    return match, float(procent)


PRAG_VALIDACIJE = 75  # minimalni % ujemanja

def validiraj_napovedi_skupin(skupine_ekipe, log_cb):
    """
    skupine_ekipe: seznam dict-ov {"skupina": "A", "stolpec": "X", "nap_stolpec": "Y",
                                    "nap_vrstica": N, "ekipe": ["Ekipa1", ...]}
    Gre čez vse napovedi in preverja ujemanje imen ekip v skupinah.
    """
    files = [f for f in os.listdir(NAPOVEDI_MAPA) if f.endswith(".xlsx") and not f.startswith("~")]
    problemi = []
    skupaj_datotek = len(files)
    log_cb(f"Preverjam {skupaj_datotek} napovedi …\n")

    for file in files:
        path = os.path.join(NAPOVEDI_MAPA, file)
        problemi_datoteke = []
        try:
            wbn = openpyxl.load_workbook(path, read_only=True, data_only=True)
            wsn = wbn.worksheets[0]
            ime = wsn[IME_CELL].value.strip().title()

            for sk in skupine_ekipe:
                nst = str(sk["nap_stolpec"]).strip().upper()
                nvr = int(sk["nap_vrstica"])
                ekipe_originalne = [e.strip() for e in sk["ekipe"]]
                ekipe_polne = [e.upper() for e in ekipe_originalne]

                for idx in range(4):
                    celica = wsn[f"{nst}{nvr + idx}"].value
                    if celica is None:
                        problemi_datoteke.append(
                            f"  Skupina {sk['skupina']}, mesto {idx+1}: celica prazna")
                        continue
                    vnos = str(celica).strip()
                    match, procent = razreši_z_aliasi(vnos, ekipe_polne, ekipe_originalne)
                    if procent < PRAG_VALIDACIJE:
                        problemi_datoteke.append(
                            f"  Skupina {sk['skupina']}, mesto {idx+1}: "
                            f"'{vnos}' → najbližje '{match.title()}' ({procent:.0f}%)")

            wbn.close()

            if problemi_datoteke:
                problemi.append((ime, file, problemi_datoteke))

        except Exception as e:
            log_cb(f"✗  {file}: {e}")

    if not problemi:
        log_cb("✓  Vse napovedi so v redu – ni problematičnih vnosov.")
    else:
        log_cb(f"⚠  Najdenih {len(problemi)} datotek s problemi:\n")
        for ime, file, pp in problemi:
            log_cb(f"▸  {ime}  ({file})")
            for p in pp:
                log_cb(p)
            log_cb("")

# ──────────────────────────────────────────────
#  GUI
# ──────────────────────────────────────────────

DARK   = "#0f1117"
PANEL  = "#1a1d27"
CARD   = "#22263a"
ACCENT = "#3b82f6"
GREEN  = "#22c55e"
RED    = "#ef4444"
MUTED  = "#64748b"
FG     = "#e2e8f0"
FG2    = "#94a3b8"

FONT_BODY  = ("Inter", 10)
FONT_LABEL = ("Inter", 9)
FONT_MONO  = ("Consolas", 9)
FONT_HEAD  = ("Inter", 13, "bold")
FONT_TITLE = ("Inter", 18, "bold")


class SP2026App(tk.Tk):

    def __init__(self):
        super().__init__()
        self.title("SP 2026 – Točkovanje")
        self.geometry("960x680")
        self.minsize(860, 600)
        self.configure(bg=DARK)
        self._build()
        self._refresh_globals()
        self.after(100, lambda: self._run_threaded(self._osveži_ob_zagonu))
        self.protocol('WM_DELETE_WINDOW', self._ob_zapiranju)

    # ── GRADNJA ──────────────────────────────

    def _ob_zapiranju(self):
        shrani_config()
        self.destroy()

    def _osveži_ob_zagonu(self):
        """Ob zagonu avtomatsko naloži tekme in skupine v ozadju."""
        if TEKME_FILE and os.path.exists(TEKME_FILE):
            self.after(0, self._osveži_tekme)
            self.after(0, self._osveži_skupine)

    def _build(self):
        # Leva navigacija
        nav = tk.Frame(self, bg=PANEL, width=190)
        nav.pack(side="left", fill="y")
        nav.pack_propagate(False)

        tk.Label(nav, text="⚽  SP 2026", bg=PANEL, fg=FG,
                 font=FONT_TITLE, pady=22).pack(fill="x", padx=16)
        tk.Frame(nav, bg=MUTED, height=1).pack(fill="x", padx=16, pady=4)

        self._frames = {}
        self._nav_btns = {}

        pages = [
            ("tekma",     "▷  Vnesi tekmo"),
            ("skupina",   "◈  Zaključi skupino"),
            ("imena",     "⊕  Naloži imena"),
            ("sortiraj",  "⇅  Sortiraj lestvico"),
            ("preveri",   "🔍  Preveri napovedi"),
            ("nastavitve","⚙  Nastavitve"),
        ]

        content = tk.Frame(self, bg=DARK)
        content.pack(side="left", fill="both", expand=True)

        for key, label in pages:
            btn = tk.Button(
                nav, text=label, anchor="w",
                bg=PANEL, fg=FG2, bd=0,
                activebackground=CARD, activeforeground=FG,
                font=FONT_BODY, padx=20, pady=10,
                cursor="hand2",
                command=lambda k=key: self._show(k)
            )
            btn.pack(fill="x")
            self._nav_btns[key] = btn

        # Ustvari strani
        for key, _ in pages:
            f = tk.Frame(content, bg=DARK)
            f.place(relwidth=1, relheight=1)
            self._frames[key] = f

        self._build_tekma(self._frames["tekma"])
        self._build_skupina(self._frames["skupina"])
        self._build_imena(self._frames["imena"])
        self._build_sortiraj(self._frames["sortiraj"])
        self._build_nastavitve(self._frames["nastavitve"])
        self._build_preveri(self._frames["preveri"])

        self._show("tekma")

    def _show(self, key):
        self._frames[key].lift()
        for k, btn in self._nav_btns.items():
            btn.configure(bg=CARD if k == key else PANEL,
                          fg=FG  if k == key else FG2)

    # ── SKUPNE POMOŽNE METODE ─────────────────

    def _card(self, parent, title, row=0, col=0, colspan=1):
        f = tk.Frame(parent, bg=CARD, padx=18, pady=14)
        f.grid(row=row, column=col, columnspan=colspan,
               sticky="nsew", padx=8, pady=6)
        tk.Label(f, text=title, bg=CARD, fg=FG2,
                 font=FONT_LABEL).pack(anchor="w", pady=(0, 6))
        return f

    def _labeled_entry(self, parent, label, default="", width=22):
        tk.Label(parent, text=label, bg=CARD, fg=FG2,
                 font=FONT_LABEL).pack(anchor="w", pady=(6, 1))
        v = tk.StringVar(value=default)
        e = tk.Entry(parent, textvariable=v, bg=DARK, fg=FG,
                     insertbackground=FG, relief="flat",
                     font=FONT_BODY, width=width)
        e.pack(fill="x", ipady=5)
        return v

    def _log_box(self, parent, rows=10):
        txt = scrolledtext.ScrolledText(
            parent, height=rows, bg=DARK, fg=FG2,
            font=FONT_MONO, relief="flat",
            insertbackground=FG, state="disabled"
        )
        txt.pack(fill="both", expand=True, pady=(6, 0))
        return txt

    def _log(self, box, msg):
        box.configure(state="normal")
        box.insert("end", msg + "\n")
        box.see("end")
        box.configure(state="disabled")

    def _btn(self, parent, text, cmd, color=ACCENT):
        return tk.Button(
            parent, text=text, command=cmd,
            bg=color, fg="white", relief="flat",
            font=("Inter", 10, "bold"),
            padx=14, pady=8, cursor="hand2",
            activebackground=DARK, activeforeground=FG
        )

    def _run_threaded(self, fn, *args):
        threading.Thread(target=fn, args=args, daemon=True).start()

    # ── STRAN: TEKMA ──────────────────────────

    def _build_tekma(self, parent):
        parent.columnconfigure(0, weight=1)
        parent.columnconfigure(1, weight=1)
        parent.rowconfigure(1, weight=1)

        tk.Label(parent, text="Vnesi rezultat tekme",
                 bg=DARK, fg=FG, font=FONT_HEAD,
                 pady=16).grid(row=0, column=0, columnspan=2, sticky="w", padx=18)

        # Leva kartica – vnos
        left = self._card(parent, "Podatki tekme", row=1, col=0)

        # Dropdown iz excel datoteke
        tk.Label(left, text="Tekma iz excela (opcijsko)", bg=CARD,
                 fg=FG2, font=FONT_LABEL).pack(anchor="w", pady=(6, 1))
        self._tekma_var = tk.StringVar(value="— izberi —")
        self._tekma_combo = ttk.Combobox(
            left, textvariable=self._tekma_var,
            state="readonly", font=FONT_BODY, width=48
        )
        self._tekma_combo.pack(fill="x", ipady=4)
        self._tekma_combo.bind("<<ComboboxSelected>>", self._on_tekma_select)
        self._btn(left, "↻  Osveži seznam", self._osveži_tekme).pack(
            anchor="w", pady=(4, 10))

        ttk.Separator(left, orient="horizontal").pack(fill="x", pady=8)

        self._lbl_domaci = tk.Label(left, text="Goli domači", bg=CARD, fg=FG2, font=FONT_LABEL)
        self._lbl_domaci.pack(anchor="w", pady=(6, 1))
        self._t_domaci = tk.StringVar()
        tk.Entry(left, textvariable=self._t_domaci, bg=DARK, fg=FG,
                 insertbackground=FG, relief="flat", font=FONT_BODY, width=22).pack(fill="x", ipady=5)

        self._lbl_gostje = tk.Label(left, text="Goli gostje", bg=CARD, fg=FG2, font=FONT_LABEL)
        self._lbl_gostje.pack(anchor="w", pady=(6, 1))
        self._t_gostje = tk.StringVar()
        tk.Entry(left, textvariable=self._t_gostje, bg=DARK, fg=FG,
                 insertbackground=FG, relief="flat", font=FONT_BODY, width=22).pack(fill="x", ipady=5)

        self._btn(left, "▷  Obdelaj tekmo",
                  self._zacni_tekmo, GREEN).pack(anchor="w", pady=(14, 0))

        # Desna kartica – log
        right = self._card(parent, "Dnevnik", row=1, col=1)
        self._tekma_log = self._log_box(right, rows=16)
        self._tekma_progress = ttk.Progressbar(right, mode="determinate", maximum=100)
        self._tekma_progress.pack(fill="x", pady=(6, 0))

    def _osveži_tekme(self):
        vse_tekme = preberi_tekme_iz_excela()
        # Prikaži samo neobdelane (obdelana != "D")
        self._tekme_data = [t for t in vse_tekme if str(t["obdelana"] or "").upper() != "D"]
        vrednosti = [
            f"{t['datum_str']}  {t['ekipa1']} – {t['ekipa2']}" if t.get("datum_str") else f"{t['ekipa1']} – {t['ekipa2']}"
            for t in self._tekme_data
        ]
        self._tekma_combo["values"] = vrednosti
        self._tekma_combo.set("— izberi —")
        self._izbrana_tekma = None
        if hasattr(self, "_lbl_domaci"):
            self._lbl_domaci.configure(text="Goli domači")
            self._lbl_gostje.configure(text="Goli gostje")
        skupaj = len(vse_tekme)
        preostalo = len(self._tekme_data)
        self._log(self._tekma_log,
                  f"Naloženih {preostalo} neobdelanih tekem (od {skupaj} skupaj).")

    def _on_tekma_select(self, _=None):
        idx = self._tekma_combo.current()
        if idx < 0 or not hasattr(self, "_tekme_data"): return
        t = self._tekme_data[idx]
        self._izbrana_tekma = t
        self._lbl_domaci.configure(text=f"Goli domači  –  {t['ekipa1']}")
        self._lbl_gostje.configure(text=f"Goli gostje  –  {t['ekipa2']}")
        self._log(self._tekma_log,
                  f"Izbrana tekma: {t['ekipa1']} – {t['ekipa2']}  "
                  f"[stolpec: {t['stolpec']}, vrstica: {t.get('vrstica_napovedi', '–')}]")

    def _zacni_tekmo(self):
        try:
            gd  = int(self._t_domaci.get())
            gg  = int(self._t_gostje.get())
        except ValueError:
            messagebox.showerror("Napaka", "Preverite vnose (goli morajo biti številke).")
            return

        izbrana = getattr(self, "_izbrana_tekma", None)
        if not izbrana:
            messagebox.showerror("Napaka", "Najprej izberi tekmo iz seznama.")
            return
        vr = izbrana.get("vrstica_napovedi")
        if vr is None:
            messagebox.showerror("Napaka", "Vrstica napovedi ni določena v excelu (stolpec F).")
            return
        vr = int(vr)
        st = str(izbrana["stolpec"] or "").strip().upper()
        self._log(self._tekma_log, f"Stolpec rezultatov: {st}")
        self._log(self._tekma_log, f"\n▷  Začetek obdelave tekme …")

        self._tekma_progress["value"] = 0
        def _po_obdelavi(log_cb):
            def _progress(val): self.after(0, lambda: self._tekma_progress.configure(value=val))
            obdelaj_tekmo(vr, gd, gg, st, log_cb, _progress)
            if izbrana and TEKME_FILE:
                try:
                    oznaci_tekmo_obdelano(izbrana["vrstica_excela"])
                    self.after(0, self._log, self._tekma_log,
                               f"✓  Tekma označena kot obdelana v excelu.")
                    self.after(0, self._osveži_tekme)
                    self.after(0, self._sprazni_tekmo)
                except Exception as e:
                    self.after(0, self._log, self._tekma_log,
                               f"⚠  Ni uspelo označiti tekme: {e}")

        self._run_threaded(_po_obdelavi,
                           lambda m: self.after(0, self._log, self._tekma_log, m))

    def _sprazni_tekmo(self):
        self._t_domaci.set("")
        self._t_gostje.set("")
        self._izbrana_tekma = None
        self._lbl_domaci.configure(text="Goli domači")
        self._lbl_gostje.configure(text="Goli gostje")

    # ── STRAN: SKUPINA ────────────────────────

    def _build_skupina(self, parent):
        parent.columnconfigure(0, weight=1)
        parent.columnconfigure(1, weight=1)
        parent.rowconfigure(1, weight=1)

        tk.Label(parent, text="Zaključi skupino",
                 bg=DARK, fg=FG, font=FONT_HEAD,
                 pady=16).grid(row=0, column=0, columnspan=2, sticky="w", padx=18)

        left = self._card(parent, "Parametri skupine", row=1, col=0)

        tk.Label(left, text="Skupina iz excela (opcijsko)", bg=CARD,
                 fg=FG2, font=FONT_LABEL).pack(anchor="w", pady=(6, 1))
        self._skupina_var = tk.StringVar(value="— izberi —")
        self._skupina_combo = ttk.Combobox(
            left, textvariable=self._skupina_var,
            state="readonly", font=FONT_BODY, width=48
        )
        self._skupina_combo.pack(fill="x", ipady=4)
        self._skupina_combo.bind("<<ComboboxSelected>>", self._on_skupina_select)
        self._btn(left, "↻  Osveži seznam", self._osveži_skupine).pack(
            anchor="w", pady=(4, 10))

        ttk.Separator(left, orient="horizontal").pack(fill="x", pady=8)

        tk.Label(left, text="Ekipe (1–4, vrstni red končnega stanja)",
                 bg=CARD, fg=FG2, font=FONT_LABEL).pack(anchor="w", pady=(10, 2))
        self._s_ekipe = []
        for i in range(4):
            v = tk.StringVar(value="— izberi —")
            cb = ttk.Combobox(left, textvariable=v, state="readonly",
                              font=FONT_BODY, width=28)
            cb.pack(fill="x", ipady=4, pady=1)
            cb.bind("<<ComboboxSelected>>", self._on_ekipa_select)
            self._s_ekipe.append((v, cb))
        self._s_ekipe_combos = [pair[1] for pair in self._s_ekipe]

        self._btn(left, "◈  Obdelaj skupino",
                  self._zacni_skupino, GREEN).pack(anchor="w", pady=(14, 0))

        right = self._card(parent, "Dnevnik", row=1, col=1)
        self._skupina_log = self._log_box(right, rows=16)
        self._skupina_progress = ttk.Progressbar(right, mode="determinate", maximum=100)
        self._skupina_progress.pack(fill="x", pady=(6, 0))

    def _osveži_skupine(self):
        vse_skupine = preberi_skupine_iz_excela()
        self._skupine_data = [s for s in vse_skupine if str(s["obdelana"] or "").upper() != "D"]
        vrednosti = [f"Skupina {s['skupina']}" for s in self._skupine_data]
        self._skupina_combo["values"] = vrednosti
        self._skupina_combo.set("— izberi —")
        self._izbrana_skupina = None
        skupaj = len(vse_skupine)
        preostalo = len(self._skupine_data)
        self._log(self._skupina_log,
                  f"Naloženih {preostalo} neobdelanih skupin (od {skupaj} skupaj).")

    def _on_skupina_select(self, _=None):
        idx = self._skupina_combo.current()
        if idx < 0 or not hasattr(self, "_skupine_data"): return
        s = self._skupine_data[idx]
        self._izbrana_skupina = s
        oznaka = str(s.get("skupina") or "").strip().upper()
        ekipe = SKUPINE_SP2026.get(oznaka, [])
        self._posodobi_ekipe_combo(ekipe)

    def _posodobi_ekipe_combo(self, ekipe):
        """Nastavi vrednosti vseh 4 dropdownov glede na ekipe izbrane skupine."""
        # Poiščemo zadnje 4 Comboboxe v levi kartici
        combos = self._s_ekipe_combos
        for cb in combos:
            cb["values"] = ekipe
            cb.set("— izberi —")
        for v, _ in self._s_ekipe:
            v.set("— izberi —")

    def _on_ekipa_select(self, _=None):
        """Ob vsaki spremembi dropdowna posodobi razpoložljive možnosti v ostalih."""
        izbrane = {v.get() for v, _ in self._s_ekipe if v.get() != "— izberi —"}
        sk = getattr(self, "_izbrana_skupina", None)
        if not sk: return
        oznaka = str(sk.get("skupina") or "").strip().upper()
        vse_ekipe = SKUPINE_SP2026.get(oznaka, [])
        combos = self._s_ekipe_combos
        for (var, cb) in self._s_ekipe:
            trenutna = var.get()
            mozne = ["— izberi —"] + [
                e for e in vse_ekipe
                if e not in izbrane or e == trenutna
            ]
            cb["values"] = mozne

    def _zacni_skupino(self):
        izbrana_s = getattr(self, "_izbrana_skupina", None)
        if not izbrana_s:
            messagebox.showerror("Napaka", "Najprej izberi skupino iz seznama.")
            return
        vr = izbrana_s.get("nap_vrstica")
        if vr is None:
            messagebox.showerror("Napaka", "Začetna vrstica ni določena v excelu (stolpec D).")
            return
        nst = str(izbrana_s.get("nap_stolpec") or "").strip().upper()
        if not nst:
            messagebox.showerror("Napaka", "Stolpec napovedi ni določen v excelu (stolpec C).")
            return
        rst = str(izbrana_s.get("stolpec") or "").strip().upper()
        if not rst:
            messagebox.showerror("Napaka", "Stolpec rezultatov ni določen v excelu (stolpec B).")
            return
        vr = int(vr)
        ekipe = [v.get().strip().upper() for v, _ in self._s_ekipe]
        if any(e in ("", "— IZBERI —") for e in ekipe):
            messagebox.showerror("Napaka", "Izberi vse 4 ekipe.")
            return
        if len(set(ekipe)) < 4:
            duplikati = [e for e in ekipe if ekipe.count(e) > 1]
            messagebox.showerror("Napaka", f"Vsaka ekipa sme biti izbrana samo enkrat.\nDuplikat: {duplikati[0].title()}")
            return
        sk  = str(izbrana_s.get("skupina") or "").strip().upper()
        self._log(self._skupina_log, f"\n◈  Začetek obdelave skupine {sk} …")

        self._skupina_progress["value"] = 0
        izbrana_s_ref = izbrana_s
        def _po_skupini(log_cb):
            def _progress(val): self.after(0, lambda: self._skupina_progress.configure(value=val))
            obdelaj_skupino(sk, vr, nst, rst, ekipe, log_cb, _progress)
            if izbrana_s_ref.get("vrstica_excela") and TEKME_FILE:
                try:
                    oznaci_skupino_obdelano(izbrana_s_ref["vrstica_excela"])
                    self.after(0, self._log, self._skupina_log,
                               f"✓  Skupina {sk} označena kot obdelana v excelu.")
                    self.after(0, self._osveži_skupine)
                    
                    self.after(0, self._sprazni_skupino)
                except Exception as e:
                    self.after(0, self._log, self._skupina_log,
                               f"⚠  Ni uspelo označiti skupine: {e}")

        self._run_threaded(_po_skupini,
                           lambda m: self.after(0, self._log, self._skupina_log, m))

    def _sprazni_skupino(self):
        for v, _ in self._s_ekipe:
            v.set("— izberi —")
        self._izbrana_skupina = None
        self._skupina_combo.set("— izberi —")
        # Počisti vrednosti v dropdownih ekip
        combos = self._s_ekipe_combos
        for cb in combos:
            cb["values"] = []

    # ── STRAN: IMENA ──────────────────────────

    def _build_imena(self, parent):
        parent.columnconfigure(0, weight=1)
        parent.rowconfigure(1, weight=1)

        tk.Label(parent, text="Naloži imena v excel",
                 bg=DARK, fg=FG, font=FONT_HEAD,
                 pady=16).grid(row=0, column=0, sticky="w", padx=18)

        card = self._card(parent, "Akcija", row=1, col=0)

        tk.Label(card, text=(
            "Prebere vsa imena iz map napovedi in jih zapiše\n"
            "v excel rezultatov (stolpec B, začenši z vrstico 3).\n\n"
            "Obstoječa vsebina v stolpcu B bo prepisana."
        ), bg=CARD, fg=FG2, font=FONT_LABEL, justify="left").pack(
            anchor="w", pady=(0, 12))

        self._btn(card, "⊕  Naloži imena",
                  self._zacni_imena, ACCENT).pack(anchor="w")

        self._imenа_log = self._log_box(card, rows=16)

    def _zacni_imena(self):
        # Safeguard: preveri, ali so podatki že vneseni
        def _preveri_in_naloži():
            ok, razlog = preveri_rezultati_prazni()
            if not ok:
                # Prikaži opozorilo in ponudi možnost nadaljevanja
                nadaljuj = self.after(0, self._vprasaj_prepisovanje, razlog)
            else:
                self.after(0, self._log, self._imenа_log, "\n⊕  Nalaganje imen …")
                napolni_imena(lambda m: self.after(0, self._log, self._imenа_log, m))

        self._run_threaded(_preveri_in_naloži)

    def _vprasaj_prepisovanje(self, razlog):
        odg = messagebox.askyesno(
            "Pozor – podatki že obstajajo",
            f"{razlog}\n\nŽeliš vseeno prepisati stolpec B z imeni?\n"
            "(Točke v ostalih stolpcih ostanejo nespremenjene.)",
            icon="warning"
        )
        if odg:
            self._log(self._imenа_log, "\n⚠  Prepisovanje imen …")
            self._run_threaded(napolni_imena,
                               lambda m: self.after(0, self._log, self._imenа_log, m))
        else:
            self._log(self._imenа_log, "↩  Nalaganje preklicano.")

    # ── STRAN: SORTIRAJ ───────────────────────

    def _build_sortiraj(self, parent):
        parent.columnconfigure(0, weight=1)
        parent.rowconfigure(1, weight=1)

        tk.Label(parent, text="Sortiraj in oštevilči lestvico",
                 bg=DARK, fg=FG, font=FONT_HEAD,
                 pady=16).grid(row=0, column=0, sticky="w", padx=18)

        card = self._card(parent, "Akcija", row=1, col=0)

        tk.Label(card, text=(
            "Razvrsti vse vrstice po skupnem seštevku točk (stolpec C)\n"
            "in doda ustrezne ozake mest v stolpec A.\n\n"
            "Pri enakem številu točk se mesto podvoji, vmesna mesta so prazna."
        ), bg=CARD, fg=FG2, font=FONT_LABEL, justify="left").pack(
            anchor="w", pady=(0, 12))

        self._btn(card, "⇅  Sortiraj zdaj",
                  self._zacni_sortiranje, ACCENT).pack(anchor="w")

        self._sort_log = self._log_box(card, rows=16)

    def _zacni_sortiranje(self):
        self._log(self._sort_log, "\n⇅  Sortiranje …")
        self._run_threaded(sortiraj_in_ostevilci,
                           lambda m: self.after(0, self._log, self._sort_log, m))

    # ── STRAN: PREVERI NAPOVEDI ──────────────────

    def _build_preveri(self, parent):
        parent.columnconfigure(0, weight=1)
        parent.columnconfigure(1, weight=2)
        parent.rowconfigure(1, weight=1)

        tk.Label(parent, text="Preveri napovedi skupin",
                 bg=DARK, fg=FG, font=FONT_HEAD,
                 pady=16).grid(row=0, column=0, columnspan=2, sticky="w", padx=18)

        left = self._card(parent, "Nastavitve", row=1, col=0)

        tk.Label(left, text=(
            "Preveri, ali napovedi vseh uporabnikov\n"
            "vsebujejo prepoznavna imena ekip za vsako skupino.\n\n"
            "Referenčna imena so vgrajena v kodo (SP 2026).\n"
            "Problematični vnosi bodo izpisani v dnevniku."
        ), bg=CARD, fg=FG2, font=FONT_LABEL, justify="left").pack(anchor="w", pady=(0, 16))

        tk.Label(left, text="Prag ujemanja (%)", bg=CARD, fg=FG2,
                 font=FONT_LABEL).pack(anchor="w", pady=(0, 1))
        self._preveri_prag = tk.StringVar(value=str(PRAG_VALIDACIJE))
        tk.Entry(left, textvariable=self._preveri_prag, bg=DARK, fg=FG,
                 insertbackground=FG, relief="flat", font=FONT_BODY, width=6).pack(anchor="w", ipady=4)

        tk.Label(left, text=(
            "Priporočeno: 75\n"
            "(nižje = manj strogo, višje = strožje)"
        ), bg=CARD, fg=MUTED, font=FONT_LABEL).pack(anchor="w", pady=(2, 16))

        self._btn(left, "🔍  Zaženi preverjanje",
                  self._zacni_validacijo, ACCENT).pack(anchor="w")

        right = self._card(parent, "Rezultati preverjanja", row=1, col=1)
        self._preveri_log = self._log_box(right, rows=24)

    def _zacni_validacijo(self):
        global PRAG_VALIDACIJE
        try:
            PRAG_VALIDACIJE = int(self._preveri_prag.get())
        except ValueError:
            messagebox.showerror("Napaka", "Prag mora biti število med 0 in 100.")
            return

        # Poiščemo podatke skupin iz excela (stolpec napovedi, vrstica)
        skupine_excel = {str(s["skupina"]).strip().upper(): s
                         for s in preberi_skupine_iz_excela()}

        skupine_ekipe = []
        for oznaka, ekipe in SKUPINE_SP2026.items():
            sk_data = skupine_excel.get(oznaka)
            if sk_data is None:
                self._log(self._preveri_log,
                          f"⚠  Skupina {oznaka} ni najdena v excelu – preskočena.")
                continue
            skupine_ekipe.append({
                "skupina":     oznaka,
                "nap_stolpec": sk_data["nap_stolpec"],
                "nap_vrstica": sk_data["nap_vrstica"],
                "ekipe":       ekipe,
            })

        if not skupine_ekipe:
            messagebox.showerror("Napaka", "Nobena skupina ni bila najdena v excelu.")
            return

        self._log(self._preveri_log,
                  f"\n🔍  Začetek preverjanja {len(skupine_ekipe)} skupin …\n")
        self._run_threaded(validiraj_napovedi_skupin, skupine_ekipe,
                           lambda m: self.after(0, self._log, self._preveri_log, m))

    # ── STRAN: NASTAVITVE ─────────────────────

    def _build_nastavitve(self, parent):
        parent.columnconfigure(0, weight=1)
        parent.columnconfigure(1, weight=1)
        parent.rowconfigure(3, weight=1)

        tk.Label(parent, text="Nastavitve",
                 bg=DARK, fg=FG, font=FONT_HEAD,
                 pady=16).grid(row=0, column=0, columnspan=2, sticky="w", padx=18)

        # Kartica – poti do datotek
        paths = self._card(parent, "Poti do datotek", row=1, col=0)
        self._n_napovedi   = self._labeled_entry(paths, "Mapa napovedi",       NAPOVEDI_MAPA, 36)
        self._n_rezultati  = self._labeled_entry(paths, "Excel rezultatov",    REZULTATI_FILE, 36)
        self._n_tekme_file = self._labeled_entry(paths, "Excel tekem/skupin",  TEKME_FILE, 36)

        # Kartica – stolpci napovedi
        cols = self._card(parent, "Stolpci v napovedi", row=1, col=1)
        self._n_ime_cell   = self._labeled_entry(cols, "Celica z imenom",   IME_CELL,   12)
        self._n_domaci_col = self._labeled_entry(cols, "Stolpec domači",    DOMACI_COL, 12)
        self._n_gost_col   = self._labeled_entry(cols, "Stolpec gostje",    GOST_COL,   12)
        self._n_tocke_col  = self._labeled_entry(cols, "Stolpec točke",     TOCKE_COL,  12)

        # Kartica – stolpci rezultatov
        rcols = self._card(parent, "Stolpci v rezultatih", row=2, col=0)
        self._n_ime_col    = self._labeled_entry(rcols, "Stolpec imen",     IME_COLUMN, 12)
        self._n_tek_sheet  = self._labeled_entry(rcols, "Indeks lista – tekme",   "0", 6)
        self._n_sk_sheet   = self._labeled_entry(rcols, "Indeks lista – skupine", "1", 6)

        # Gumbi
        btn_row = tk.Frame(parent, bg=DARK)
        btn_row.grid(row=2, column=1, sticky="sw", padx=8, pady=6)
        self._btn(btn_row, "✓  Shrani nastavitve",
                  self._shrani_nastavitve, GREEN).pack(side="left", padx=(0, 8))
        self._btn(btn_row, "↺  Ponastavi",
                  self._refresh_globals, RED).pack(side="left")

        # Status
        self._n_status = tk.Label(parent, text="", bg=DARK, fg=GREEN,
                                  font=FONT_LABEL)
        self._n_status.grid(row=3, column=0, columnspan=2,
                            sticky="w", padx=18, pady=4)

    def _shrani_nastavitve(self):
        global NAPOVEDI_MAPA, REZULTATI_FILE, TEKME_FILE
        global IME_CELL, DOMACI_COL, GOST_COL, TOCKE_COL, IME_COLUMN
        global TEKME_SHEET_IDX, SKUPINE_SHEET_IDX

        NAPOVEDI_MAPA      = self._n_napovedi.get().strip()
        REZULTATI_FILE     = self._n_rezultati.get().strip()
        TEKME_FILE         = self._n_tekme_file.get().strip()
        IME_CELL           = self._n_ime_cell.get().strip().upper()
        DOMACI_COL         = self._n_domaci_col.get().strip().upper()
        GOST_COL           = self._n_gost_col.get().strip().upper()
        TOCKE_COL          = self._n_tocke_col.get().strip().upper()
        IME_COLUMN         = self._n_ime_col.get().strip().upper()
        try:
            TEKME_SHEET_IDX   = int(self._n_tek_sheet.get())
            SKUPINE_SHEET_IDX = int(self._n_sk_sheet.get())
        except ValueError:
            pass

        # Osveži slovar imen v ozadju
        self._n_status.configure(text="Nastavitve shranjene. Osveževanje slovarja imen …", fg=FG2)
        def _reload():
            try:
                napolni_slovar_imen()
                self.after(0, lambda: self._n_status.configure(
                    text=f"✓  Naloženih {len(IMENA_VRSTICE)} imen.", fg=GREEN))
            except Exception as e:
                self.after(0, lambda: self._n_status.configure(
                    text=f"✗  Napaka pri nalaganju: {e}", fg=RED))
        shrani_config()
        threading.Thread(target=_reload, daemon=True).start()

    def _refresh_globals(self):
        """Nastavi Entry vrednosti iz trenutnih globalnih spremenljivk."""
        pairs = [
            (self._n_napovedi,   NAPOVEDI_MAPA),
            (self._n_rezultati,  REZULTATI_FILE),
            (self._n_tekme_file, TEKME_FILE),
            (self._n_ime_cell,   IME_CELL),
            (self._n_domaci_col, DOMACI_COL),
            (self._n_gost_col,   GOST_COL),
            (self._n_tocke_col,  TOCKE_COL),
            (self._n_ime_col,    IME_COLUMN),
            (self._n_tek_sheet,  str(TEKME_SHEET_IDX)),
            (self._n_sk_sheet,   str(SKUPINE_SHEET_IDX)),
        ]
        # Entry varibale so dostopne le po inicializaciji
        try:
            for var, val in pairs:
                var.set(val)
        except AttributeError:
            pass  # še ni zgrajeno


# ──────────────────────────────────────────────
#  VSTOPNA TOČKA
# ──────────────────────────────────────────────

if __name__ == "__main__":
    # Nastavimo style za ttk Combobox
    style = None
    app   = SP2026App()

    s = ttk.Style(app)
    s.theme_use("clam")
    s.configure("TCombobox",
                 fieldbackground=DARK,
                 background=CARD,
                 foreground=FG,
                 selectbackground=DARK,
                 selectforeground="#ffffff",
                 insertcolor=FG,
                 arrowcolor=FG2)
    s.map("TCombobox",
          fieldbackground=[("readonly", DARK), ("disabled", PANEL)],
          foreground=[("readonly", FG), ("disabled", FG2)],
          selectbackground=[("readonly", DARK)],
          selectforeground=[("readonly", "#ffffff")])
	  
    napolni_slovar_imen()

    app.mainloop()
