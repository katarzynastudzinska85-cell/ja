# Dziennik zmian instrukcji agenta

Zgodnie z rozdziałem 12 `CLAUDE.md`: zapisuj, co i dlaczego zmieniono w regułach i progach.

## v2 (06.10.2026)

- Dodano instrukcje agenta analizy alertów e-Connectivity jako `CLAUDE.md`.

## Skrypt v3 (06.10.2026)

- Dodano `live_diagnostic_agent.js` (wersja „v3_fixed”) bez zmian w logice.
- Dodano `package.json` i `.gitignore` (wyklucza `session.json`, pliki Excel i wyniki).

## Skrypt v3.1 (06.10.2026) — zgodność z CLAUDE.md

Błędy:
- Daty z dashboardu uwzględniają AM/PM (wcześniej 18:53 było odczytywane jako 06:53).
- Analiza wg typu alertu jest faktycznie wywoływana (Supply thermal, Slide (CM) Ring, Data Logger, inne).
- Błąd odczytu AAA trafia do PDF; taki alert dostaje „brak oceny technicznej” zamiast decyzji.
- Obrazy wykresów w PDF nie nachodzą na tekst.

Reguły:
- Priorytet WYSOKI / ŚREDNI / ZDALNIE / NISKI z dowodów, nie z koloru alertu (rozdz. 7).
- Luki w danych: odcinki idealnie liniowe > 3% szerokości usuwane ze statystyk w seriach z szumem;
  w seriach gładkich tylko oznaczane jako możliwe luki (rozdz. 5.3). Dlaczego: gładka temperatura
  otoczenia była w całości brana za lukę.
- Progi tylko z kolorów progów (fioletowy, żółty); płaska seria nie jest już progiem (rozdz. 5.1).
- Walidacja wykresu: J-number z tytułu (1 znak różnicy = prawdopodobny błąd OCR, odczyt niepewny;
  ≥ 2 znaki = wykres innego aparatu, wyłączony), skala, progi w zakresie osi, liczba pikseli (rozdz. 5.2).
- Data Logger: czas bez łączności (dni i godziny), wykrywanie 0/0 przez cały okres (rozdz. 6.3).
- Porównanie aparatów tego samego modelu: duty przy ≈ 27 °C (dopasowanie liniowe, próg 10 pp),
  wersja oprogramowania (rozdz. 6.1.7, 6.3.5).
- Godziny z wykresów z „ok.” (rozdz. 3).

Raport:
- Układ wg rozdz. 8: podsumowanie z tabelą priorytetów, tabela odczytów, wykresy odtworzone z lukami
  jako szary pas, „Co widać” (Fakt/Odczyt), „Co to może oznaczać” (Hipoteza), zalecenia z kryterium
  sukcesu, porównanie aparatów na wspólnej skali, alerty łączności, plan działań, metodyka i ograniczenia.
- Jedna wersja w nazwie: v3.

Testy: `npm test` (funkcje bez przeglądarki).
