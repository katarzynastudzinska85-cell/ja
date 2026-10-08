# Agent analizy alertów e-Connectivity (v3)

Skrypt otwiera dashboard e-Connectivity w Edge, czyta aktywne alerty i wykresy AAA
i tworzy raport PDF po polsku. Działa tylko w trybie odczytu. Zasady analizy: `CLAUDE.md`.

## Pierwsze uruchomienie (Windows)

1. Zainstaluj Node.js LTS ze strony https://nodejs.org (jednorazowo).
2. Rozpakuj folder, np. do `C:\eConnectivity`.
3. Skopiuj do tego folderu bazę Excel `stan Ortho na 23.04.2026.xlsx`.
   Inna nazwa pliku? Zmień ją w `live_diagnostic_agent.js`, linia 62.
4. Otwórz PowerShell w tym folderze i wpisz raz:

   ```
   npm.cmd install
   ```

## Każde kolejne uruchomienie

1. Kliknij dwukrotnie `start.cmd` (albo `start_z_podstronami.cmd`, żeby dodać treść podstron AAA).
   Zamiast tego w PowerShell, w folderze skryptu: `node live_diagnostic_agent.js`
2. Otworzy się Edge. Zaloguj się i przejdź do dashboardu z tabelami analizatorów.
3. Poczekaj, aż tabele się załadują, wróć do PowerShell i naciśnij ENTER.
4. Wynik: folder `aaa_output_live`
   - `Raport_Serwisowy_eConnectivity_LIVE_V3.pdf` — raport,
   - `Service_Triage_Data_LIVE_V3.json` — dane,
   - podfoldery aparatów z obrazami wykresów, kopią strony AAA (`aaa_strona.png`) i danymi (`data.json`).

Raport zawiera dla każdego alertu pełną treść strony AAA (rozdział „Treść strony AAA”):
teksty, tabele (wiersze wyróżnione kolorem oznaczone słownie), listy kroków z numeracją 1. / a. / i.,
a na końcu kopię całej strony jako obraz.

## Szukanie klienta po numerze J

Dwuklik nie wystarczy — numer trzeba podać. W PowerShell, w folderze skryptu:

```
node znajdz.js 76000978
```

Można podać kilka numerów naraz: `node znajdz.js 76000978 46001812`. Dane pochodzą z lokalnego pliku Excel.

## Gdy coś nie działa

- **Edge się nie otwiera:** sprawdź, gdzie jest `msedge.exe`. Jeśli w `C:\Program Files\Microsoft\Edge\Application\`,
  zmień ścieżkę w `live_diagnostic_agent.js`, linia 57.
- **„Odczyt wartości z wykresów jest wyłączony”:** uruchom ponownie `npm install`.
- **Raport bez polskich liter:** brak czcionki Arial w `C:\Windows\Fonts`.

## Opcje

- `service_knowledge.json` w folderze skryptu — lista dokumentów serwisowych dopasowywanych do alertów (opcjonalnie).
- `AAA_VISION=1` i `ANTHROPIC_API_KEY` — odczyt wykresów przez AI. **Wysyła obrazy wykresów na zewnątrz**,
  domyślnie wyłączone (CLAUDE.md, rozdz. 2).
- `AAA_LOCAL_CHARTS=0` — wyłącza lokalny odczyt wartości z wykresów.
- `AAA_FOLLOW_LINKS=1` — otwiera też podstrony podlinkowane na stronie AAA (np. TD-016837) i dodaje ich treść do raportu.
  Tylko linki z tej samej domeny, maks. 8 na alert, z pominięciem linków-akcji (Save, Apply, Run, Reset, Delete, logowanie).
  Domyślnie wyłączone. Dokumenty PDF są tylko odnotowane, bez treści.
- `AAA_SCREENSHOT=0` — bez kopii strony AAA jako obrazu.

Jak ustawić opcję w PowerShell (na jedno uruchomienie):

```
$env:AAA_FOLLOW_LINKS="1"; node live_diagnostic_agent.js
```

Komunikat „running scripts is disabled on this system” przy `npm`: Windows blokuje skrypt PowerShella `npm.ps1`.
Uruchamiaj przez `node live_diagnostic_agent.js` albo pliki `.cmd`. Do jednorazowej instalacji pakietów użyj `npm.cmd install`.

## Nie udostępniaj

`session.json` (zapisana sesja logowania), pliku Excel ani folderu `aaa_output_live` — zawierają dane klientów.
