# -*- coding: utf-8 -*-
"""Krok 2: buduje bazę SQLite FTS5 (vision.db) z pages.jsonl.

Baza jest budowana od zera do pliku tymczasowego i podmieniana na końcu,
więc szukaj.py działa na starej bazie aż do zakończenia budowy.

Użycie: python.exe -I build_db.py [--indeks C:\\ServiceIndex_Vision]
"""
import argparse
import datetime
import json
import os
import sqlite3
import sys

DOMYSLNY_INDEKS = r"C:\ServiceIndex_Vision"
L_NA_L = str.maketrans("łŁ", "lL")


def main():
    ap = argparse.ArgumentParser(description="Budowa bazy FTS5 dokumentacji Vision.")
    ap.add_argument("--indeks", default=DOMYSLNY_INDEKS)
    a = ap.parse_args()
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    except AttributeError:
        pass

    plik_stron = os.path.join(a.indeks, "pages.jsonl")
    if not os.path.exists(plik_stron):
        sys.exit(f"Brak {plik_stron}. Najpierw uruchom build_index.py.")
    baza = os.path.join(a.indeks, "vision.db")
    tmp = baza + ".tmp"
    if os.path.exists(tmp):
        os.remove(tmp)

    con = sqlite3.connect(tmp)
    con.executescript("""
        CREATE TABLE meta(klucz TEXT PRIMARY KEY, wartosc TEXT);
        -- norm: kopia tekstu z ł->l (unicode61 nie zdejmuje ogonka z ł), tylko gdy strona zawiera ł
        -- remove_diacritics 2: 'blad' znajduje 'błąd'; tokenchars '_': kody typu ABC_123 zostają w całości
        CREATE VIRTUAL TABLE strony USING fts5(
            tekst, tytul, rel, norm, kat UNINDEXED, typ UNINDEXED, strona UNINDEXED,
            tokenize = "unicode61 remove_diacritics 2 tokenchars '_'"
        );
    """)
    n = 0
    with open(plik_stron, encoding="utf-8") as f:
        partia = []
        for linia in f:
            r = json.loads(linia)
            norm = r["tekst"].translate(L_NA_L) if ("ł" in r["tekst"] or "Ł" in r["tekst"]) else ""
            partia.append((r["tekst"], r["tytul"], r["rel"], norm, r["kat"], r["typ"], r["strona"]))
            if len(partia) >= 2000:
                con.executemany("INSERT INTO strony VALUES (?,?,?,?,?,?,?)", partia)
                n += len(partia)
                partia = []
        con.executemany("INSERT INTO strony VALUES (?,?,?,?,?,?,?)", partia)
        n += len(partia)
    con.execute("INSERT INTO strony(strony) VALUES ('optimize')")
    con.executemany("INSERT INTO meta VALUES (?,?)", [
        ("zbudowano", datetime.datetime.now().strftime("%Y-%m-%d %H:%M")),
        ("liczba_stron", str(n)),
    ])
    con.commit()
    con.close()
    os.replace(tmp, baza)
    print(f"Zapisano {n} rekordów do {baza}")


if __name__ == "__main__":
    main()
