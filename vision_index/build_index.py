# -*- coding: utf-8 -*-
"""Krok 1: wyciąga tekst z dokumentacji Vision do plików JSONL (jeden rekord = jedna strona PDF / jeden plik HTML).

Folder źródłowy jest tylko czytany. Wynik: <INDEKS>\\pages.jsonl, manifest.json, raport_indeksu.txt.
Ponowne uruchomienie przetwarza tylko pliki nowe lub zmienione (rozmiar + data modyfikacji).

Użycie (Windows):
  python.exe -I build_index.py
  python.exe -I build_index.py --zrodlo "D:\\inny\\folder" --indeks C:\\ServiceIndex_Vision
"""
import argparse
import glob
import html
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
from html.parser import HTMLParser

DOMYSLNE_ZRODLO = r"C:\Users\RafałMotylski\OneDrive - Altium\Dokume"
DOMYSLNY_INDEKS = r"C:\ServiceIndex_Vision"

ROZSZERZENIA_PDF = {".pdf"}
ROZSZERZENIA_HTML = {".htm", ".html", ".xhtml", ".mht"}
ROZSZERZENIA_TXT = {".txt"}

# Kategoria z nazwy folderu/pliku (pierwsze trafienie wygrywa). Dopasowanie bez wielkości liter.
KATEGORIE = [
    ("kody_bledow", r"error|b[łl][ęe]d|kod|code|alarm|event|condition"),
    ("schematy", r"schemat|schematic|wiring|electric|elektr|diagram|drawing|pcb"),
    ("mody", r"(^|[\\/ _-])mods?([\\/ _-]|$)|modyfik|modification|\bfmi\b|\bfco\b"),
    ("biuletyny", r"biulet|bulletin|\btsb\b|\btb\b|technical[ _-]?note|notice|\bfsn\b|\bfsca\b"),
    ("czesci", r"parts?[ _-]?(list|catalog)|cz[ęe][śs]ci|ipb|spare"),
    ("vdocs", r"v[ _-]?docs|library"),
    ("instrukcje", r"service|serwis|manual|instruk|procedur|guide|install|maint|pm\b"),
]


def kategoria(rel):
    sciezka = rel.lower()
    for nazwa, wzor in KATEGORIE:
        if re.search(wzor, sciezka):
            return nazwa
    return "inne"


def znajdz_pdftotext():
    sciezka = shutil.which("pdftotext")
    if sciezka:
        return sciezka
    # winget instaluje poppler do LOCALAPPDATA\Microsoft\WinGet\Packages\oschwartz10612.Poppler_*\poppler-*\Library\bin
    lad = os.environ.get("LOCALAPPDATA", "")
    wzorce = [
        os.path.join(lad, "Microsoft", "WinGet", "Packages", "*oppler*", "**", "pdftotext.exe"),
        os.path.join(lad, "Microsoft", "WinGet", "Links", "pdftotext.exe"),
        r"C:\Program Files\poppler*\**\pdftotext.exe",
    ]
    for wzor in wzorce:
        trafienia = sorted(glob.glob(wzor, recursive=True))
        if trafienia:
            return trafienia[-1]
    return None


def pdf_na_strony(pdftotext, plik):
    """Zwraca listę tekstów stron (pdftotext rozdziela strony znakiem \\f)."""
    fd, tmp = tempfile.mkstemp(suffix=".txt")
    os.close(fd)
    try:
        wynik = subprocess.run(
            [pdftotext, "-layout", "-enc", "UTF-8", plik, tmp],
            capture_output=True, timeout=600,
        )
        if wynik.returncode != 0:
            raise RuntimeError(wynik.stderr.decode("utf-8", "replace").strip()[:300] or f"kod {wynik.returncode}")
        with open(tmp, encoding="utf-8", errors="replace") as f:
            tekst = f.read()
    finally:
        try:
            os.remove(tmp)
        except OSError:
            pass
    strony = tekst.split("\f")
    if strony and not strony[-1].strip():
        strony.pop()
    return strony


class _Tekst(HTMLParser):
    POMIJANE = {"script", "style", "noscript", "head", "template"}
    BLOKOWE = {"p", "div", "br", "li", "tr", "h1", "h2", "h3", "h4", "h5", "h6", "table", "section", "td", "th", "dt", "dd", "pre"}

    def __init__(self):
        super().__init__(convert_charrefs=True)
        self.czesci, self.tytul, self._pomin, self._w_tytule = [], "", 0, False

    def handle_starttag(self, tag, attrs):
        if tag in self.POMIJANE:
            self._pomin += 1
        if tag == "title":
            self._w_tytule = True
        if tag in self.BLOKOWE:
            self.czesci.append("\n")
        if tag in ("td", "th"):
            self.czesci.append(" | ")

    def handle_endtag(self, tag):
        if tag in self.POMIJANE and self._pomin:
            self._pomin -= 1
        if tag == "title":
            self._w_tytule = False
        if tag in self.BLOKOWE:
            self.czesci.append("\n")

    def handle_data(self, data):
        if self._w_tytule:
            self.tytul += data
        elif not self._pomin:
            self.czesci.append(data)


def dekoduj(surowe):
    m = re.search(rb'charset=["\']?([A-Za-z0-9_-]+)', surowe[:4096], re.I)
    kodowania = ([m.group(1).decode("ascii", "ignore")] if m else []) + ["utf-8", "cp1250", "latin-1"]
    for kod in kodowania:
        try:
            return surowe.decode(kod)
        except (LookupError, UnicodeDecodeError):
            continue
    return surowe.decode("utf-8", "replace")


def html_na_tekst(plik):
    with open(plik, "rb") as f:
        surowe = f.read()
    p = _Tekst()
    p.feed(dekoduj(surowe))
    tekst = html.unescape("".join(p.czesci))
    tekst = re.sub(r"[ \t\u00a0]+", " ", tekst)
    tekst = re.sub(r"\n\s*\n\s*(\n\s*)+", "\n\n", tekst)
    return p.tytul.strip(), tekst.strip()


def main():
    ap = argparse.ArgumentParser(description="Ekstrakcja tekstu dokumentacji Vision.")
    ap.add_argument("--zrodlo", default=DOMYSLNE_ZRODLO)
    ap.add_argument("--indeks", default=DOMYSLNY_INDEKS)
    ap.add_argument("--od-nowa", action="store_true", help="przetwórz wszystkie pliki, ignoruj manifest")
    a = ap.parse_args()
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    except AttributeError:
        pass

    zrodlo = os.path.abspath(a.zrodlo)
    indeks = os.path.abspath(a.indeks)
    if not os.path.isdir(zrodlo):
        sys.exit(f"Brak folderu źródłowego: {zrodlo}")
    if os.path.normcase(indeks).startswith(os.path.normcase(zrodlo + os.sep)):
        sys.exit("Indeks nie może leżeć w folderze źródłowym (folder źródłowy jest tylko do odczytu).")
    os.makedirs(indeks, exist_ok=True)

    pdftotext = znajdz_pdftotext()
    if not pdftotext:
        print("UWAGA: nie znaleziono pdftotext (poppler). PDF-y zostaną pominięte.")
        print("       Instalacja: winget install oschwartz10612.Poppler")

    plik_manifest = os.path.join(indeks, "manifest.json")
    plik_stron = os.path.join(indeks, "pages.jsonl")
    manifest = {}
    stare = {}
    if not a.od_nowa and os.path.exists(plik_manifest) and os.path.exists(plik_stron):
        with open(plik_manifest, encoding="utf-8") as f:
            manifest = json.load(f)
        with open(plik_stron, encoding="utf-8") as f:
            for linia in f:
                r = json.loads(linia)
                stare.setdefault(r["rel"], []).append(r)

    nowy_manifest, rekordy = {}, []
    stat = {"pliki": 0, "nowe_lub_zmienione": 0, "z_pamieci": 0, "bledy": 0, "pominiete_typy": 0, "strony": 0, "puste_strony": 0}
    bledy, bez_tekstu, typy_pominiete = [], [], {}

    for katalog, podkatalogi, pliki in os.walk(zrodlo):
        podkatalogi.sort()
        for nazwa in sorted(pliki):
            pelna = os.path.join(katalog, nazwa)
            rel = os.path.relpath(pelna, zrodlo)
            ext = os.path.splitext(nazwa)[1].lower()
            if nazwa.startswith("~$"):
                continue
            if ext not in ROZSZERZENIA_PDF | ROZSZERZENIA_HTML | ROZSZERZENIA_TXT:
                typy_pominiete[ext or "(brak)"] = typy_pominiete.get(ext or "(brak)", 0) + 1
                stat["pominiete_typy"] += 1
                continue
            if ext in ROZSZERZENIA_PDF and not pdftotext:
                continue
            stat["pliki"] += 1
            try:
                st = os.stat(pelna)
            except OSError as e:
                bledy.append(f"{rel}: {e}")
                stat["bledy"] += 1
                continue
            podpis = [st.st_size, int(st.st_mtime)]
            if manifest.get(rel) == podpis and rel in stare:
                rekordy.extend(stare[rel])
                nowy_manifest[rel] = podpis
                stat["z_pamieci"] += 1
                continue

            kat = kategoria(rel)
            try:
                if ext in ROZSZERZENIA_PDF:
                    strony = pdf_na_strony(pdftotext, pelna)
                    nowe = [{"rel": rel, "kat": kat, "typ": "pdf", "strona": i, "tytul": nazwa, "tekst": t}
                            for i, t in enumerate(strony, 1)]
                    puste = sum(1 for r in nowe if len(r["tekst"].strip()) < 20)
                    if nowe and puste == len(nowe):
                        bez_tekstu.append(f"{rel} ({len(nowe)} str., prawdopodobnie skan — potrzebny OCR)")
                    elif puste:
                        bez_tekstu.append(f"{rel}: {puste} z {len(nowe)} stron bez tekstu")
                    stat["puste_strony"] += puste
                elif ext in ROZSZERZENIA_HTML:
                    tytul, tekst = html_na_tekst(pelna)
                    nowe = [{"rel": rel, "kat": kat, "typ": "html", "strona": 1, "tytul": tytul or nazwa, "tekst": tekst}]
                else:
                    with open(pelna, "rb") as f:
                        nowe = [{"rel": rel, "kat": kat, "typ": "txt", "strona": 1, "tytul": nazwa, "tekst": dekoduj(f.read())}]
            except Exception as e:  # jeden zły plik nie przerywa indeksowania
                bledy.append(f"{rel}: {e}")
                stat["bledy"] += 1
                continue
            rekordy.extend(nowe)
            nowy_manifest[rel] = podpis
            stat["nowe_lub_zmienione"] += 1
            print(f"[{stat['nowe_lub_zmienione']}] {kat:12} {len(nowe):4} str.  {rel}")

    stat["strony"] = len(rekordy)
    tmp = plik_stron + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        for r in rekordy:
            f.write(json.dumps(r, ensure_ascii=False) + "\n")
    os.replace(tmp, plik_stron)
    with open(plik_manifest, "w", encoding="utf-8") as f:
        json.dump(nowy_manifest, f, ensure_ascii=False, indent=0)

    kategorie = {}
    for r in rekordy:
        kategorie.setdefault(r["kat"], set()).add(r["rel"])
    with open(os.path.join(indeks, "raport_indeksu.txt"), "w", encoding="utf-8") as f:
        f.write(f"Źródło: {zrodlo}\npdftotext: {pdftotext or 'BRAK'}\n\n")
        for k, v in stat.items():
            f.write(f"{k}: {v}\n")
        f.write("\nPliki wg kategorii:\n")
        for k in sorted(kategorie):
            f.write(f"  {k}: {len(kategorie[k])}\n")
        f.write("\nPominięte typy plików (nieindeksowane):\n")
        for k, v in sorted(typy_pominiete.items(), key=lambda x: -x[1]):
            f.write(f"  {k}: {v}\n")
        f.write("\nPDF bez warstwy tekstowej (wyszukiwanie ich nie znajdzie):\n")
        f.writelines(f"  {x}\n" for x in bez_tekstu)
        f.write("\nBłędy:\n")
        f.writelines(f"  {x}\n" for x in bledy)

    print("\n" + ", ".join(f"{k}={v}" for k, v in stat.items()))
    print(f"Wynik: {plik_stron}\nRaport: {os.path.join(indeks, 'raport_indeksu.txt')}")
    print("Następny krok: build_db.py")


if __name__ == "__main__":
    main()
