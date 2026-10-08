# -*- coding: utf-8 -*-
"""Wyszukiwanie w indeksie dokumentacji Vision (SQLite FTS5).

Przykłady:
  szukaj.py 5003-01                     kod błędu (każde słowo/kod szukane jako fraza)
  szukaj.py "pipette tip" sensor        wszystkie słowa muszą wystąpić na stronie
  szukaj.py centrifuge --kat kody_bledow
  szukaj.py incubator --plik "Service Manual" -n 30
  szukaj.py --fts "centrifuge NEAR(imbalance, 10)"   surowa składnia FTS5 (OR, NOT, NEAR, prefix*)
  szukaj.py --pokaz 1234                pełny tekst rekordu o ID 1234 (z wyników)
  szukaj.py --pokaz 1234 --kontekst 1   ten rekord + sąsiednie strony
  szukaj.py --kategorie                 liczba plików/stron w kategoriach
"""
import argparse
import os
import re
import sqlite3
import sys

DOMYSLNY_INDEKS = r"C:\ServiceIndex_Vision"
L_NA_L = str.maketrans("łŁ", "lL")  # jak kolumna norm w build_db.py


def zapytanie_proste(tekst):
    """Każde słowo / kod / fraza w cudzysłowie staje się frazą FTS5 (bezpieczne dla myślników, kropek, dwukropków)."""
    czesci = re.findall(r'"([^"]+)"|(\S+)', tekst)
    frazy = []
    for fraza, slowo in czesci:
        t = (fraza or slowo).replace('"', " ").translate(L_NA_L).strip()
        if t:
            frazy.append('"' + t + '"')
    return " AND ".join(frazy)


def main():
    ap = argparse.ArgumentParser(description="Szukaj w dokumentacji Vision.", formatter_class=argparse.RawDescriptionHelpFormatter, epilog=__doc__)
    ap.add_argument("zapytanie", nargs="*")
    ap.add_argument("--indeks", default=DOMYSLNY_INDEKS)
    ap.add_argument("--fts", help="surowe zapytanie FTS5")
    ap.add_argument("--kat", help="filtr kategorii (kody_bledow, schematy, mody, biuletyny, czesci, vdocs, instrukcje, inne)")
    ap.add_argument("--plik", help="filtr: fragment ścieżki pliku")
    ap.add_argument("-n", "--limit", type=int, default=15)
    ap.add_argument("--pokaz", type=int, help="wypisz pełny tekst rekordu o tym ID")
    ap.add_argument("--kontekst", type=int, default=0, help="z --pokaz: ile sąsiednich stron dołączyć")
    ap.add_argument("--kategorie", action="store_true")
    a = ap.parse_args()
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    except AttributeError:
        pass

    baza = os.path.join(a.indeks, "vision.db")
    if not os.path.exists(baza):
        sys.exit(f"Brak bazy {baza}. Uruchom build_index.py i build_db.py.")
    con = sqlite3.connect(f"file:{baza}?mode=ro", uri=True)

    if a.kategorie:
        for kat, pliki, strony in con.execute(
                "SELECT kat, COUNT(DISTINCT rel), COUNT(*) FROM strony GROUP BY kat ORDER BY kat"):
            print(f"{kat:14} plików: {pliki:5}  stron: {strony}")
        for k, v in con.execute("SELECT klucz, wartosc FROM meta"):
            print(f"{k}: {v}")
        return

    if a.pokaz is not None:
        wiersz = con.execute("SELECT rel, strona FROM strony WHERE rowid=?", (a.pokaz,)).fetchone()
        if not wiersz:
            sys.exit(f"Brak rekordu {a.pokaz}")
        rel, strona = wiersz
        for rid, s, tekst in con.execute(
                "SELECT rowid, strona, tekst FROM strony WHERE rel=? AND strona BETWEEN ? AND ? ORDER BY strona",
                (rel, strona - a.kontekst, strona + a.kontekst)):
            print(f"===== [{rid}] {rel} — strona {s} =====")
            print(tekst)
            print()
        return

    if a.fts:
        q = a.fts
    elif a.zapytanie:
        q = zapytanie_proste(" ".join(a.zapytanie))
    else:
        ap.print_help()
        return

    sql = ("SELECT rowid, rel, strona, kat, snippet(strony, 0, '[', ']', ' … ', 24), bm25(strony, 1.0, 5.0, 3.0, 1.0) AS r "
           "FROM strony WHERE strony MATCH ?")
    par = [q]
    if a.kat:
        sql += " AND kat = ?"
        par.append(a.kat)
    if a.plik:
        sql += " AND rel LIKE ?"
        par.append(f"%{a.plik}%")
    sql += " ORDER BY r LIMIT ?"
    par.append(a.limit)
    try:
        wyniki = con.execute(sql, par).fetchall()
    except sqlite3.OperationalError as e:
        sys.exit(f"Błąd zapytania FTS5 ({q}): {e}")

    if not wyniki:
        print(f"Brak trafień dla: {q}")
        print("W tym indeksie nie ma tego tekstu — to nie znaczy, że dokumentacja tego nie opisuje (skany bez OCR, inne nazewnictwo).")
        return
    print(f"Zapytanie: {q}   trafień (pokazano): {len(wyniki)}\n")
    for rid, rel, strona, kat, fragment, _ in wyniki:
        fragment = re.sub(r"\s+", " ", fragment).strip()
        print(f"[{rid}] {rel} — str. {strona}  ({kat})")
        print(f"      {fragment}\n")
    print("Pełna strona: szukaj.py --pokaz ID [--kontekst 1]")


if __name__ == "__main__":
    main()
