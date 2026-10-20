# CLAUDE.md — Generator Ewidencji Przebiegu Pojazdu

## Projekt

Aplikacja webowa + desktopowa do generowania miesięcznej ewidencji przebiegu pojazdu na podstawie plikow GPS trackerow (XLS/XLSX). Wyjsciowym jest plik Excel z dwoma arkuszami (tytulowy + rozliczenie). Stworzona dla **Smart Work Sp. z o.o.**

## Stack

| Warstwa | Technologia |
|---------|-------------|
| Backend (parser + generator Excel) | Python 3.8+, Flask, openpyxl |
| Serwer HTTP / API | Node.js, Express, Multer |
| Frontend | Statyczny HTML/CSS/JS inline w `server.js` |
| Parser GPS | `src/gpsParser.js`, `src/xlsParser.js` |
| Generator Excel (Node) | `src/excelGenerator.js` |
| Build .exe | PyInstaller (`GeneratorEwidencji.spec`) |
| Deploy | Firebase App Hosting (`apphosting.yaml`), Google Cloud Run, Docker |

## Kluczowe pliki

- `app.py` — Flask backend z **wbudowanym czytnikiem XLS (OLE/BIFF8)**. Nie wymaga LibreOffice ani xlrd. Zawiera tez generator openpyxl.
- `server.js` — Express serwujacy statyczny frontend (folder `public/`) i endpointy API. Zawiera inline caly HTML strony glownej.
- `src/gpsParser.js` — logika parsowania tras GPS, agregacja km per dzien, laczenie sobot/niedziel z piatkiem.
- `src/xlsParser.js` — parser binarnych plikow `.xls`.
- `src/excelGenerator.js` — generator pliku Excel (ExcelJS) z arkuszami tytulowym i rozliczeniowym.
- `public/` — PWA (manifest.json, ikony, service worker).
- `GeneratorEwidencji.spec` — specyfikacja PyInstaller do budowania `.exe`.

## Architektura dwutorowa (Python ↔ Node)

Obecnie sa dwa backendy rownolegle:
- **`app.py` (Flask + openpyxl)** — uzywany przy budowie `.exe` przez PyInstaller. Frontend jest inline w tym samym pliku.
- **`server.js` (Express + ExcelJS)** — uzywany w wersji web/PWA i Cloud Run.

**Wazne:** przy zmianach w logice generowania ewidencji lub parserze GPS trzeba uwzglednic **obydwa backendi**.

## Zasady biznesowe (krytyczne — nie zmieniaj bez uzgodnienia)

1. **Weekendy** — km z soboty i niedzieli sumowane i przypisywane do poprzedniego **piatku roboczego**.
2. **Dni bez jazdy** — pozostaja puste (nie wpisujemy 0).
3. **Stan licznika koncowy** = poczatkowy + suma wszystkich km.
4. **Tankowanie** — uzytkownik podaje daty (np. `5.12, 12.12`); w kolumnie "Uwagi" wpisywane jest slowo "tankowanie".
5. **Adresy** — ciaglość: nowy dzien zaczyna sie od adresu docelowego poprzedniego dnia.
6. **Format wyjsciowy** — nazwa pliku: `MM. ewidencja <miesiac> <ostatni dzien miesiaca DD.MM.RRRR> <NR_REJ>.xlsx`, np. `07. ewidencja lipiec 31.07.2026 DSW3318E.xlsx` (`src/outputName.js` / `build_output_name` w `app.py` — trzymac zsynchronizowane).

## Endpointy API (server.js)

| Metoda | Path | Opis |
|--------|------|------|
| GET | `/` | Strona glowna (HTML inline) |
| POST | `/generate` | Upload pliku GPS (`gps_file`) + pola formularza → zwraca Excel. Rate limit (domyslnie 20/min/IP), maks. 10 MB. Ostrzezenia (np. brak stanu licznika) w naglowku `X-Warnings` (URL-encoded) |
| GET | `/health`, `/api/health` | Healthcheck |

Nazwa pliku wynikowego: patrz zasada #6.
Bledy odczytu pliku → HTTP 400; odpowiedzi 500 nie zawieraja stack trace (tylko w logu serwera).

Pola formularza `/generate`: `driver_name`, `trip_purpose` (maks. 200 znakow), `odometer_start`, `target_odometer`, `adjust_mileage`, `use_actual_days`, `refuel_dates` (maks. 500 znakow; niepoprawne daty sa pomijane z ostrzezeniem).

## Zmienne srodowiskowe

| Zmienna | Domyslnie | Opis |
|---------|-----------|------|
| `PORT` | 8080 | Port serwera |
| `RATE_LIMIT_MAX` / `RATE_LIMIT_WINDOW_MS` | 20 / 60000 | Limit zadan `/generate` na IP |
| `KNOWN_CITIES_FILE` | `./known_cities.json` (w Dockerze `/tmp/...`) | Plik cache kodow pocztowych → miasto |
| `NOMINATIM_CONTACT` / `NOMINATIM_USER_AGENT` | adres repo | Kontakt w User-Agent (wymog polityki Nominatim) — ustaw prawdziwy e-mail |

## Testy

`npm test` (`node --test`): `test/gpsParser.test.js`, `test/excelGenerator.test.js` (generator + bledy xlsParser), `test/server.test.js` (integracyjne API), `test/rateLimit.test.js`. Brak fixture binarnego `.xls` — pelny odczyt BIFF8 nie jest pokryty testami.

> Uwaga: na dni robocze bez jazdy generator wpisuje jawne `0` w kolumnie km (dni wolne zostaja puste) — to zachowanie kodu, niezgodne z regula #2 powyzej; do uzgodnienia z uzytkownikiem.

## Skrypty npm

```bash
npm start        # node server.js (port 8080)
npm run dev      # nodemon server.js
npm run build    # generuje ikony PWA
npm run build:exe   # ikony + PyInstaller (.exe Windows)
npm run deploy   # gcloud run deploy
```

## Uruchomienie lokalne

**Windows:** kliknij `URUCHOM.bat` (wlacza .venv + Flask na :5000 + otwiera przegladarke).

**Mac/Linux:**
```bash
chmod +x uruchom_mac_linux.sh
./uruchom_mac_linux.sh
```

**Docker:**
```bash
npm run docker:up   # port 8080
```

## Auto-detekcja formatu pliku GPS

Parser sam wykrywa, czy plik jest w starym formacie XLS czy nowym XLSX:

| Format | date | startAddr | endAddr | km (w wierszu "Razem") |
|--------|------|-----------|---------|------------------------|
| Stary XLS | col 8 | col 9 | col 12 | col 16 |
| Nowy XLSX | col 10 | col 12 | col 17 | col 23 |

Detekcja odbywa sie przez znalezienie pierwszego wiersza z obiektem `Date` w kolumnie 10. Jesli nowy tracker zmieni uklad kolumn — trzeba recznie zaktualizowac mapowanie w `gpsParser.js` (linie ~370-382).

## Wbudowany czytnik XLS (OLE/BIFF8)

Parser `.xls` jest napisany recznie w czystym JS (port z Pythona z `app.py`). Nie uzywamy `xlrd` ani LibreOffice.

Obslugiwane rekordy BIFF8: `SST`, `LABELSST`, `LABEL`, `NUMBER`, `RK`, `MULRK`.

Jesli nowy tracker produkuje pliki `.xls` z dodatkowymi typami rekordow — trzeba rozbudowac parser (`src/xlsParser.js` i analogicznie `app.py`).

## Obsluga Rich Text i Formul w XLSX

Funkcja `extractCellValue()` w `gpsParser.js` obsluguje:
- `richText` — skleja fragmenty tekstu
- `formula.result` — wyciaga wynik formuly
- `hyperlink.text` — wyciaga tekst linku

Nowe trackery czesto zapisuja adresy jako rich text — bez tego wyciagalyby puste komorki.

## Geokodowanie adresow (Nominatim)

Adresy GPS sa parsowane do postaci `ulica, miasto`.

- Jesli miasta nie ma w pliku, kod pyta **Nominatim (OpenStreetMap)** po kodzie pocztowym.
- **Ograniczenie:** 1 req/sec (sztuczne opoznienie 1100 ms w `geocodeCityNominatim`).
- **Cache trwaly:** plik `known_cities.json` (tworzony automatycznie w katalogu roboczym).
- Hardkodowane ZIPy: `43-245 → Studzionka`, `43-246 → Strumien`.

Funkcje kluczowe: `parseGpsAddr()`, `shortenAddr()`, `extractCity()`.

## Logika budowania tras (opis trasy w Excelu)

1. **Czyszczenie adresow:**
   - usuwa cyfry, kody pocztowe
   - usuwa prefixy ulic (`ulica`, `ul.`, `al.`, `aleja`, `plac`, `pl.`)
   - usuwa podzialy administracyjne (`Powiat`, `Gmina`, `Wojewodztwo`, `gm.`, `pow.`, `woj.`)

2. **Detekcja POI:** jesli pierwszy fragment adresu nie wyglada na ulice, a drugi tak — pomija POI.

3. **Ciaglosc adresow:** nowy dzien zaczyna sie od ostatniego adresu poprzedniego dnia (`lastKnownCity`).

4. **Wiele tras w jednym dniu:** route parts laczone sa srednikiem (` ; `).

## Swieta polskie i algorytm Wielkanocy

`gpsParser.js` zawiera pelna liste swiat stalych + algorytm **Meeusa/Jonesa/Butchera** do obliczania Wielkanocy, Poniedzialku Wielkanocnego i Bozego Ciala.

- Weekendy i swieta sa przerzucane na poprzedni dzien roboczy.
- Funkcja `getPreviousWorkingDay()` idzie do tylu po dniach roboczych; jesli wyjdzie poza miesiac, idzie do przodu zamiast tego.

## Format wyjsciowego Excela (szczegoly techniczne)

- **Arkusz "Tytuowa"** — naglowki, dane firmy, stany licznika.
- **Arkusz "Rozlicznie"** — 31 dni, zamrozony wiersz 6, orientacja landscape, A4, `fitToWidth=1`.
- Auto-dopasowanie szerokosci kolumn na podstawie dlugosci tekstu (z limitem 48 dla kolumny E).
- **Style:**
  - dni wolne (weekendy/swieta) → szary fill (`A6A6A6`)
  - tankowanie → pomaranczowy tekst (`B45309`) + zolty fill (`FFFBEB`)

## Dane firmy zahardkodowane

W generatorze (`excelGenerator.js`) dane Smart Work sa wpisane na stałe:
- Nazwa: `Smart Work Sp. Z o.o.`
- Adres: `58-100 Swidnica  Bystrzycka 7a`
- NIP: `8842754100`

Jesli kiedykolwiek trzeba bedzie zmienic firme — trzeba edytowac kod, nie config.

## Testowanie parsera XLS/XLSX i diagnostyka nowych plikow GPS

W razie zmian w formacie plikow GPS uzytkownik korzysta z `diagnose.js` do analizy nowego pliku:
```bash
node diagnose.js DSW3276E.xlsx
```

Wyniki zapisywane sa do `excel_analysis.txt` (przykladowy zrzut w repo).

**Przy nowym formacie od dostawcy GPS zawsze najpierw uruchamiamy `node diagnose.js <plik>`.**

## Uwagi dla deweloperow

- **Nie dodawaj zaleznosci xlrd / LibreOffice** — wbudowany czytnik XLS w `app.py` i `src/xlsParser.js` zostal napisany specjalnie, aby uniknac tych zaleznosci.
- **openpyxl** obsluguje tylko `.xlsx` (zapis). Czytnik `.xls` jest wlasny (OLE/BIFF8).
- **ExcelJS** (Node) generuje plik wyjsciowy w wersji serwerowej Node.
- **PWA** — aplikacja jest zainstalowalna na mobile; service worker cache'uje zasoby statyczne.
- Wersja **Cloud Run** uzywa `server.js` (port 8080). Wersja **Windows .exe** uzywa `app.py` + PyInstaller.
