# Architecture

**Analysis Date:** 2026-03-31

## Pattern Overview

**Overall:** Single-process Express web application — upload-process-download pipeline with no persistent state between requests.

**Key Characteristics:**
- Stateless HTTP request/response: each `POST /generate` is fully self-contained
- No database; all data lives in request memory (multer `memoryStorage`)
- All business logic is synchronous/async Node.js; no worker threads, queues, or background jobs
- The UI is a single inline HTML string served from `server.js`; no template engine, no static file directory
- Output is a binary XLSX buffer streamed directly as a file download

## Layers

**Entry Point / HTTP Server:**
- Purpose: Route HTTP requests, parse multipart form uploads, coordinate processing, return file downloads
- Location: `server.js`
- Contains: Express app setup, multer configuration, inline HTML page string, `GET /` and `POST /generate` route handlers, odometer adjustment logic, refuel date parsing
- Depends on: `src/gpsParser.js`, `src/excelGenerator.js`
- Used by: Browser clients

**GPS Parsing & Business Logic:**
- Purpose: Parse GPS report files (XLS or XLSX), build per-day trip groups, aggregate kilometers, handle Polish holiday calendar
- Location: `src/gpsParser.js`
- Contains: `parseGps()`, `aggregate()`, `aggregateActual()`, `buildRoute()`, `getPreviousWorkingDay()`, `getPolishHolidays()`, address geocoding via Nominatim, city cache I/O (`known_cities.json`)
- Depends on: `src/xlsParser.js` (for `.xls` files), `exceljs` (for `.xlsx` files), Node built-ins `fs`, `path`, `https`
- Used by: `server.js`

**Binary XLS Parser:**
- Purpose: Read legacy `.xls` (OLE/BIFF8) files without any native binary dependencies
- Location: `src/xlsParser.js`
- Contains: `xlsReadRows()` — full OLE container parser (FAT, mini-FAT, directory stream) plus BIFF8 record decoder (SST, LABELSST, LABEL, NUMBER, RK, MULRK record types)
- Depends on: Node `Buffer` only (no npm packages)
- Used by: `src/gpsParser.js`

**Excel Output Generator:**
- Purpose: Produce a two-sheet XLSX workbook (title page + daily ledger) formatted to Polish vehicle mileage log standards
- Location: `src/excelGenerator.js`
- Contains: `generateExcel()`, all cell styling helpers (`font`, `fill`, `align`, `bdr`, `thickOutline`), route string builder calls, daily row iteration
- Depends on: `src/gpsParser.js` (imports `weekdayOf`, `daysInMonth`, `getPolishHolidays`, `getPreviousWorkingDay`, `buildRoute`), `exceljs`
- Used by: `server.js`

## Data Flow

**Primary Request Lifecycle — `POST /generate`:**

1. Browser submits `multipart/form-data` with GPS file + form fields to `POST /generate`
2. `multer` buffers the entire file in memory (`req.file.buffer`); form fields land in `req.body`
3. `server.js` extracts and validates parameters (driver name, starting odometer, refuel dates, target odometer, mode flags)
4. `parseGps(buffer, filename)` in `src/gpsParser.js` dispatches by extension:
   - `.xls` → `xlsReadRows(buffer)` in `src/xlsParser.js` returns raw row arrays
   - `.xlsx` → `ExcelJS.Workbook.load(buffer)` returns raw row arrays
5. `parseGps` scans header rows for plate, car model, date range; then walks rows accumulating per-day trip groups (`dayGroups` array of `{ date, km, addresses, firstStart, lastEnd }`)
6. Back in `server.js`: `aggregate(dayGroups)` or `aggregateActual(dayGroups)` merges trips into a `Map<isoStr, {km, addresses, …}>`
   - `aggregate` shifts weekend/holiday km to the previous Polish working day
   - `aggregateActual` keeps km on the literal driving day
7. Optional proportional km adjustment: if `targetOdometer > odometer`, all daily km values are scaled so the total matches the expected distance; rounding residual is placed on the highest-km day
8. Refuel date strings are parsed and normalised to ISO format, then stored in a `Set`
9. `generateExcel(…, agg, …)` in `src/excelGenerator.js`:
   - Creates ExcelJS workbook with sheets `Tytułowa` (title page) and `Rozlicznie` (ledger)
   - Iterates every calendar day of the month; for days with km data, calls `buildRoute(addresses)` which may perform async Nominatim geocoding with a 1.1-second rate-limit delay
   - Returns a `Buffer` of the XLSX binary
10. `server.js` sets `Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet` and sends the buffer as a file download

**Page Load — `GET /`:**

1. Browser requests `/`
2. `server.js` sends the `HTML_PAGE` string constant (no-cache headers)
3. All UI interaction is handled by inline `<script>` in that HTML

## Key Abstractions

**`dayGroups` Array:**
- Purpose: Intermediate parsed representation of GPS data before aggregation
- Shape: `Array<{ date: string (ISO), km: number, addresses: string[], firstStart: string, lastEnd: string }>`
- Produced by: `parseGps()` in `src/gpsParser.js`
- Consumed by: `aggregate()` / `aggregateActual()` in `src/gpsParser.js`, then passed result to `generateExcel()`

**`agg` Map:**
- Purpose: Final per-calendar-day km and address data ready for Excel output
- Shape: `Map<isoStr, { km: number, addresses: string[], dailyAddresses: string[][], firstLoc: string|null, lastLoc: string|null }>`
- Produced by: `aggregate()` or `aggregateActual()` in `src/gpsParser.js`
- Consumed by: odometer-adjustment logic in `server.js`, then `generateExcel()` in `src/excelGenerator.js`

**`onlineCitiesCache` (module-level object in `gpsParser.js`):**
- Purpose: ZIP-code-to-city-name lookup, persisted across requests to avoid redundant Nominatim API calls
- Backing file: `known_cities.json` in `process.cwd()` (written on every new geocoding result)
- Lifetime: In-memory for the process lifetime; loaded once at module initialization

## Entry Points

**HTTP Server:**
- Location: `server.js` lines 374-376
- Triggers: `node server.js` (or `web: node server.js` via `Procfile`)
- Responsibilities: Bind to `process.env.PORT || 8080` on `0.0.0.0`, register two routes

**`GET /`:**
- Location: `server.js` lines 264-269
- Returns: Inline HTML page

**`POST /generate`:**
- Location: `server.js` lines 271-370
- Returns: XLSX file download or JSON error

## Error Handling

**Strategy:** Fail-fast with HTTP error responses; no retries, no partial results.

**Patterns:**
- Missing file → `res.status(400).json({ error: 'Brak pliku GPS' })`
- Empty GPS data after parsing → `res.status(400).json({ error: 'Nie znaleziono danych GPS w pliku.' })`
- Wrong file magic in `xlsReadRows` → throws `Error('Plik nie jest formatem XLS (OLE)…')` — caught by route try/catch
- Missing Workbook stream in OLE → throws `Error('Nie znaleziono strumienia Workbook…')`
- Unsupported extension → throws in `parseGps()`
- Geocoding errors (Nominatim) → silently return `null`; city field is left blank, processing continues
- All route-level errors caught by `try/catch` → `res.status(500).json({ error, detail: stack })`

## Cross-Cutting Concerns

**Logging:** `console.error(err)` on 500-class errors in `server.js`; `console.error` for city-cache write failures in `gpsParser.js`. No structured logging library.

**Validation:** Input validation is minimal and implicit — file type is inferred from extension, numeric fields are parsed with `parseInt`/`parseFloat` with `|| 0` fallback, date strings use regex extraction.

**Authentication:** None. The application is open with no auth layer.

**Rate Limiting:** Self-imposed only — Nominatim geocoding includes a `setTimeout(1100ms)` delay per request to comply with OSM usage policy.

---

*Architecture analysis: 2026-03-31*
