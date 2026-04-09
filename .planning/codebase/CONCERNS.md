# Codebase Concerns

**Analysis Date:** 2026-03-31

---

## Tech Debt

**HTML page embedded as a JS string literal in server.js:**
- Issue: The entire frontend HTML (≈245 lines) is hardcoded as a template literal inside `server.js` (lines 16–259). This makes the UI impossible to edit with syntax highlighting, hot-reload, or any templating toolchain.
- Files: `server.js`
- Impact: Any UI change requires editing the server file directly. HTML entities are manually escaped (`&#x2013;`, `&#x105;`, etc.) making the markup very hard to read. Bug surface grows with server restarts required for UI tweaks.
- Fix approach: Extract HTML to `public/index.html` and serve it with `express.static()` or `res.sendFile()`. Alternatively, use a minimal template engine (EJS, Handlebars).

**`http` variable bound to `https` module:**
- Issue: In `src/gpsParser.js` line 5, `const http = require('https')` — the variable is named `http` but holds the `https` module. All outbound geocoding requests (`httpsGet`) rely on this alias.
- Files: `src/gpsParser.js`
- Impact: Misleading code; any future developer adding plain HTTP calls may use the wrong module by referencing this variable.
- Fix approach: Rename to `const https = require('https')` and update `httpsGet` accordingly.

**Python predecessor still present alongside Node.js rewrite:**
- Issue: `app.py` (65,968 bytes), `requirements.txt`, `__pycache__/`, `analyze_excel*.py`, `test_*.py`, `.venv/`, `GeneratorEwidencji.exe`, and `GeneratorEwidencji.spec` all exist in the repository root. These are artifacts of the Python/Flask version.
- Files: `app.py`, `requirements.txt`, `.venv/` (directory), `__pycache__/` (directory), `GeneratorEwidencji.exe`, `GeneratorEwidencji.spec`, `analyze_excel.py`, `analyze_excel2.py`, `analyze_excel3.py`, `analyze_excel_utf8.py`, `test_parser.py`, `test_parser2.py`, `test_routes.py`, `test_routes2.py`
- Impact: Repository is bloated (23 MB+ from `GeneratorEwidencji.exe` alone). New contributors do not know which codebase is authoritative. `requirements.txt` could mislead deployment tools.
- Fix approach: Move Python artifacts to a `legacy/` branch or delete them. Add `app.py` and `*.exe` to `.gitignore` (or confirm removal).

**No lockfile pinning — `package.json` uses `^` ranges:**
- Issue: All three dependencies (`express ^4.18.2`, `multer ^1.4.5-lts.1`, `exceljs ^4.4.0`) use caret ranges. `package-lock.json` pins versions at install time but is not enforced in Dockerfile.
- Files: `package.json`, `Dockerfile`
- Impact: `npm install --omit=dev` in Docker builds resolves to whatever `^` allows at build time. A minor/patch release of any dependency could silently change behavior.
- Fix approach: Change `npm install` in Dockerfile to `npm ci` (requires lockfile to be copied first) — add `COPY package-lock.json .` before the install step.

---

## Security Concerns

**No rate limiting on `/generate` endpoint:**
- Risk: The `/generate` POST endpoint accepts file uploads up to 50 MB with no per-IP or per-session rate limiting. A single client can repeatedly post large files, exhausting CPU and memory.
- Files: `server.js` (line 11, line 271)
- Current mitigation: `multer` enforces a 50 MB per-file limit. No other limiting is applied.
- Recommendations: Add `express-rate-limit` middleware on `/generate`. Consider reducing `fileSize` limit based on real GPS file sizes (typical GPS XLS files are well under 1 MB).

**Outbound HTTPS request with static `User-Agent` containing `contact@example.com`:**
- Risk: Geocoding requests to Nominatim OSM identify the app as `GeneratorEwidencji/4.0 (contact@example.com)`. This is a placeholder email, violating Nominatim's Usage Policy which requires a valid contact.
- Files: `src/gpsParser.js` (line 168)
- Current mitigation: None.
- Recommendations: Replace placeholder email with a real contact address. Evaluate whether Nominatim usage at scale requires caching (the `onlineCitiesCache` already handles this partially).

**`known_cities.json` written to process working directory (`process.cwd()`):**
- Risk: In a containerised environment, writing to `process.cwd()` (which is `/app` per Dockerfile) writes inside the container filesystem. If the container is read-only or ephemeral, the write fails silently. If the container is not read-only, this is mutable state that resets on every redeploy.
- Files: `src/gpsParser.js` (lines 128, 143–154)
- Current mitigation: `saveKnownCities` wraps the write in a `try/catch` and only logs the error — so failures are silent to the user.
- Recommendations: Either mount a persistent volume and point `KNOWN_CITIES_FILE` to it, or drop filesystem persistence entirely and keep the cache in-process (acceptable given the small hardcoded seed in `ZIP_TO_CITY`).

**Error details exposed in 500 responses:**
- Risk: The `/generate` catch block returns `err.stack` (full Node.js stack trace) in the JSON response body under the `detail` key.
- Files: `server.js` (lines 367–369)
- Current mitigation: None.
- Recommendations: Log `err.stack` server-side only. Return only a generic user-facing message in the response body (the `error` field already does this; remove `detail: err.stack`).

**No input sanitisation on `driver_name` or `trip_purpose`:**
- Risk: These strings are passed directly into Excel cell values. ExcelJS handles value encoding, so XSS is not a direct concern, but excessively long strings (no length cap) could produce malformed or huge output files.
- Files: `server.js` (lines 281, 287)
- Current mitigation: `.trim()` is applied but no max-length check.
- Recommendations: Add a reasonable character limit (e.g., 200 chars) with a 400 response if exceeded.

---

## Performance Bottlenecks

**Synchronous Nominatim geocoding with 1.1-second forced delay per address:**
- Problem: `geocodeCityNominatim` in `src/gpsParser.js` (line 189) inserts `await new Promise(r => setTimeout(r, 1100))` before every external API call to respect OSM's 1 req/sec rate limit. For a monthly GPS report with 20+ working days, each with multiple unique addresses, total processing time can exceed 30–60 seconds.
- Files: `src/gpsParser.js` (lines 184–196)
- Cause: The delay is per-address, sequential. There is no pre-deduplification of zip codes before geocoding, so the same zip code could be queried multiple times if the in-memory cache was just populated in the same request.
- Improvement path: Deduplicate all zip codes needing geocoding before the request loop, then batch lookups with shared in-process cache. Consider using a local postal code database (e.g., a static JSON for PL zip codes) to eliminate network dependency entirely.

**`buildRoute` is async and called per-day inside `generateExcel`:**
- Problem: `generateExcel` in `src/excelGenerator.js` (lines 401–409) `await`s `buildRoute` inside a `for` loop iterating over every day of the month. Each `buildRoute` call can itself trigger multiple `parseGpsAddr` calls which may geocode.
- Files: `src/excelGenerator.js` (lines 386–411), `src/gpsParser.js` (lines 477–531)
- Cause: Sequential awaiting in a loop prevents any concurrency in address resolution.
- Improvement path: Collect all addresses first, resolve them concurrently with `Promise.all`, then build routes from the resolved data.

---

## Fragile Areas

**`xlsParser.js` is a hand-rolled OLE/BIFF8 parser with no test coverage:**
- Files: `src/xlsParser.js`
- Why fragile: The OLE compound document format and BIFF8 record format are complex binary formats. The parser only handles: SST, LABELSST, LABEL, NUMBER, RK, MULRK records. Any GPS tracker that emits FORMULA, BLANK, BOOLERR, or other numeric record types (e.g., `MULBLANK 0x00BE`) will silently produce empty cells. FAT chain parsing assumes a max of 109 DIFAT sectors (line 58), which is the standard OLE limit, but does not handle DIFAT extension sectors.
- Safe modification: Test against multiple real GPS XLS files before changing record handling. Any change to `cellKey` encoding (currently `r * 65536 + c`) breaks all cell lookups.
- Test coverage: Zero — no test files exist for any JS module.

**GPS file structure parsing relies on fixed row/column indices:**
- Files: `src/gpsParser.js` (lines 326–398)
- Why fragile: Plate and car model are read from hardcoded row index 12, columns 0 and 2. The date range is parsed from a regex applied only to the first 20 rows (line 326). Total km is read from column index 16 of the "Razem:" row. If the GPS tracker vendor changes their report layout (even a one-row header change), all metadata extraction silently returns empty/null values.
- Safe modification: Add defensive checks; log when expected fields are not found. Do not assume stable row indices without verifying against new GPS report versions.

**`getPreviousWorkingDay` has two unbounded `while (true)` loops:**
- Files: `src/gpsParser.js` (lines 90–118)
- Why fragile: If the holiday set for a given year somehow contained every day of a month (pathological input), both loops would run indefinitely. In practice this cannot happen with the current `getPolishHolidays` implementation, but there is no loop guard or iteration cap.
- Safe modification: Add a loop counter guard (e.g., break after 31 iterations) and throw a descriptive error if exceeded.

**`odometer_start` of 0 and `target_odometer` of 0 interact silently:**
- Files: `server.js` (lines 288–289, 300), `src/excelGenerator.js` (line 106)
- Why fragile: If neither odometer field is filled by the user, both default to `0`. The adjustment condition `adjustMil && targetOdometer > odometer` is `false` (0 > 0 is false), so no adjustment occurs. The ending odometer in the Excel output (`odoEnd`) also becomes `0 + totalKm`, which appears as a valid odometer reading of e.g. "247 km". No warning is shown to the user if odometer fields are empty.
- Safe modification: Validate that `odometer_start > 0` before generating output, or explicitly show "0" as a placeholder so the user knows to fill it.

---

## Scaling Limits

**Single-process Node.js with no clustering:**
- Current capacity: One event loop thread handles all requests.
- Limit: A single large GPS file that triggers many geocoding calls (30+ seconds) blocks the event loop for that duration only in the async geocoding portions, but the CPU-bound `xlsReadRows` parser is fully synchronous and will block all concurrent requests for its duration.
- Scaling path: Offload `xlsReadRows` to a `worker_threads` worker, or cluster with `node:cluster`. For cloud deployment, horizontal scaling (multiple container instances) is the simplest mitigation.

**In-process geocoding cache is not shared across instances:**
- Current capacity: `onlineCitiesCache` is a module-level JS object, populated at startup from `known_cities.json`.
- Limit: Multiple container replicas each maintain independent caches. Cache misses on one instance are not visible to others, causing redundant Nominatim calls across replicas.
- Scaling path: Use a shared external cache (Redis, or a small managed KV store) keyed by zip code.

---

## Missing Critical Features

**No authentication or access control:**
- Problem: The app is a fully open HTTP service with no login, API key, or IP allowlist. Any user with the URL can generate Excel reports.
- Blocks: Cannot restrict usage to company employees, cannot audit usage, cannot prevent abuse.

**No tests of any kind:**
- Problem: There are no Jest, Mocha, or any other JS test files in the project. The Python-era test scripts (`test_parser.py`, `test_routes.py`, etc.) are excluded from git and target the old `app.py`.
- Files affected: All of `src/` and `server.js`
- Risk: Any regression in XLS parsing, GPS aggregation, or Excel generation is invisible until a user reports a wrong output. The `xlsParser.js` binary parser is especially sensitive to regressions.

**No health check endpoint:**
- Problem: The Dockerfile and `Procfile` expose `server.js` but define no `/health` or `/readiness` route. Cloud platforms (Firebase App Hosting, Cloud Run) use HTTP health checks to determine instance readiness.
- Files: `server.js`, `Dockerfile`
- Risk: If the process starts but hangs before accepting connections, the platform has no signal to restart the instance.

**No request logging or observability:**
- Problem: Beyond `console.error` on 500 responses and a startup message, there is no structured logging, no request duration tracking, and no error aggregation.
- Files: `server.js`
- Risk: Cannot diagnose production failures, track processing time per file, or identify which GPS file formats cause errors.

---

## Deployment Risks

**`app.py` and `requirements.txt` remain in the repository root:**
- Risk: A hosting platform that auto-detects language (e.g., Google Cloud Build's buildpacks) may detect `requirements.txt` and attempt a Python build instead of using the `Dockerfile`, silently deploying the wrong application.
- Files: `requirements.txt`, `app.py`
- Mitigation: Remove or relocate these files. The `Dockerfile` should be the single source of truth for build instructions.

**Dockerfile does not copy `package-lock.json` before `npm install`:**
- Risk: `npm install --omit=dev` resolves semver ranges fresh on every build (see Tech Debt above). `npm ci` requires the lockfile, ensuring reproducible installs, but the current Dockerfile (`COPY package.json .` then `npm install`) does not enable this.
- Files: `Dockerfile`
- Mitigation: Add `COPY package-lock.json .` before the `RUN npm install` line, then change `npm install --omit=dev` to `npm ci --omit=dev`.

**`known_cities.json` is in `.gitignore` but is written at runtime inside the container:**
- Risk: On container restart or redeploy, the file is gone. Any zip-code-to-city mappings learned during the previous container lifetime (Nominatim lookups) are lost, causing repeated external API calls on the next run.
- Files: `src/gpsParser.js` (line 128), `.gitignore`
- Mitigation: Either commit the seed file, mount a persistent volume, or eliminate filesystem-backed cache in favour of in-memory only (given the tiny `ZIP_TO_CITY` seed set).

**No `.nvmrc` or `engines` field to lock Node.js version:**
- Risk: The Dockerfile pins `node:18-slim`, but local development has no enforced Node.js version. If a developer runs Node 20 or 22 locally, subtle differences in Buffer behaviour or module resolution could cause inconsistencies.
- Files: `package.json`, `Dockerfile`
- Mitigation: Add `"engines": { "node": "18.x" }` to `package.json` and add an `.nvmrc` file containing `18`.

---

## Dependencies at Risk

**`multer ^1.4.5-lts.1` — long-term maintenance status unclear:**
- Risk: Multer 1.x has been in LTS mode for several years. No active development occurs on the 1.x branch. Multer 2.x was not released as of the analysis date.
- Impact: If a security advisory is issued for Multer 1.x, patching may require a major version migration.
- Migration plan: Monitor the `multer` GitHub for v2 releases. Alternatively, evaluate `busboy` (used internally by Multer) directly for a lighter dependency.

**`exceljs ^4.4.0` — large dependency for file generation:**
- Risk: `exceljs` is a comprehensive library pulling in many sub-dependencies. It is the heaviest production dependency. Version 4.x has had issues with memory usage on large workbooks.
- Impact: For small monthly reports (≤31 rows), memory is not a concern. If usage grows to multi-month or multi-sheet reports, memory pressure could surface.
- Migration plan: No immediate action needed; monitor for v5 release.

---

*Concerns audit: 2026-03-31*
