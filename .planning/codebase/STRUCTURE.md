# Codebase Structure

**Analysis Date:** 2026-03-31

## Directory Layout

```
webapp/                         # Project root
├── server.js                   # Express server, routes, inline UI
├── package.json                # Node.js manifest (name: ewidencja-przebiegu)
├── package-lock.json           # Dependency lockfile
├── Procfile                    # Heroku/Firebase App Hosting: "web: node server.js"
├── Dockerfile                  # Container build (not used by current hosting)
├── .gitignore
├── .dockerignore
├── src/                        # All application modules
│   ├── xlsParser.js            # OLE/BIFF8 binary XLS reader
│   ├── gpsParser.js            # GPS file parsing, aggregation, geocoding
│   └── excelGenerator.js       # ExcelJS XLSX output builder
├── known_cities.json           # Runtime geocoding cache (created on first geocode)
├── app.py                      # Legacy Python/Flask predecessor (not used by server.js)
├── requirements.txt            # Python deps for app.py (Flask, xlrd, openpyxl, etc.)
├── .venv/                      # Python virtualenv (legacy, not used by Node app)
├── GPS.xls                     # Sample/test GPS input file
├── node_modules/               # Installed npm packages (not committed)
├── .planning/                  # GSD planning workspace
│   └── codebase/               # Codebase analysis documents
├── .claude/                    # Claude GSD agent definitions and commands
├── build/                      # PyInstaller build artefacts (legacy)
├── dist/                       # PyInstaller distributable (legacy)
├── __pycache__/                # Python bytecode (legacy)
├── analyze_excel*.py           # Ad-hoc analysis scripts (legacy)
├── test_parser*.py             # Ad-hoc Python test scripts (legacy)
├── test_routes*.py             # Ad-hoc Python test scripts (legacy)
├── *.txt                       # Analysis output files (legacy/dev artefacts)
├── URUCHOM.bat                 # Windows run script (legacy)
├── uruchom_mac_linux.sh        # Unix run script (legacy)
└── GeneratorEwidencji.spec     # PyInstaller spec (legacy)
```

## Directory Purposes

**`src/`:**
- Purpose: All server-side business logic modules
- Contains: Three `.js` files — the binary parser, the GPS/calendar logic, and the Excel builder
- Key files: `src/xlsParser.js`, `src/gpsParser.js`, `src/excelGenerator.js`
- Note: No subdirectories; the entire module surface is flat

**`.planning/codebase/`:**
- Purpose: GSD codebase analysis documents consumed by `/gsd:plan-phase` and `/gsd:execute-phase`
- Generated: By `/gsd:map-codebase`
- Committed: Yes

**`.claude/`:**
- Purpose: Claude agent definitions (`.claude/agents/`) and slash-command definitions (`.claude/commands/gsd/`)
- Generated: No (hand-crafted)
- Committed: Yes

**`node_modules/`:**
- Purpose: npm package installation directory
- Generated: Yes (`npm install`)
- Committed: No (in `.gitignore`)

**`.venv/`, `build/`, `dist/`, `__pycache__/`:**
- Purpose: Artefacts from the legacy Python/Flask/PyInstaller version of this tool
- Committed: Partially (`.venv` is excluded by `.gitignore`; `build/dist` appear committed)
- Note: Not used by the active Node.js application

## Key File Locations

**Entry Points:**
- `server.js`: Express application — bind port, serve HTML, handle `POST /generate`
- `Procfile`: Declares process type for Heroku/Firebase App Hosting (`web: node server.js`)

**Configuration:**
- `package.json`: Runtime + start script; only three runtime dependencies: `express`, `multer`, `exceljs`
- `package-lock.json`: Exact dependency tree lockfile

**Core Logic:**
- `src/xlsParser.js`: Self-contained OLE/BIFF8 parser — no npm dependencies, operates on a `Buffer`
- `src/gpsParser.js`: GPS file dispatch, header extraction, day-group building, Polish holiday calendar, aggregation modes, Nominatim geocoding, address formatting, city cache read/write
- `src/excelGenerator.js`: Two-sheet XLSX workbook construction including all cell formatting, column sizing, page setup

**Runtime-Generated Files:**
- `known_cities.json`: ZIP-to-city cache written to `process.cwd()` at runtime; not present until first geocoding occurs

**Legacy (not part of active application):**
- `app.py`: Original Python/Flask server
- `requirements.txt`: Python dependencies
- `analyze_excel*.py`, `test_*.py`: Development utility scripts
- `GeneratorEwidencji.exe`, `GeneratorEwidencji.spec`: PyInstaller desktop app build

## Naming Conventions

**Files:**
- Module files: camelCase with descriptive compound names — `xlsParser.js`, `gpsParser.js`, `excelGenerator.js`
- Entry point: lowercase — `server.js`
- Config files: lowercase with dots — `package.json`, `package-lock.json`, `.gitignore`

**Functions (in `src/`):**
- Public async functions: camelCase verb phrases — `parseGps`, `generateExcel`, `buildRoute`, `geocodeCityNominatim`
- Pure helpers: camelCase — `aggregate`, `aggregateActual`, `weekdayOf`, `daysInMonth`, `xlsReadRows`
- Small private helpers in `excelGenerator.js`: short lowercase — `font`, `fill`, `align`, `bdr`, `colLetter`

**Constants:**
- Module-level constants: UPPER_SNAKE_CASE — `HTML_PAGE`, `MAGIC`, `ENDOFCHAIN`, `FREESECT`, `XL_EPOCH`, `PL_DAYS`
- Regex constants: UPPER_SNAKE_CASE with `_RE` suffix — `ZIP_CITY_RE`, `ZIP_RE`, `STREET_PREFIX_RE`

**Exports:**
- Each `src/` module uses `module.exports = { ... }` with named exports
- `server.js` does not export (it is the process entry point)

## Where to Add New Code

**New parsing logic (file format support):**
- Add format-specific reader to `src/` as a new module (e.g., `src/csvParser.js`)
- Wire into the `ext` dispatch switch inside `parseGps()` in `src/gpsParser.js`

**New business logic (aggregation modes, km adjustment rules):**
- Add functions to `src/gpsParser.js` and export them
- Import and call from the `POST /generate` handler in `server.js`

**New Excel sheet or formatting:**
- Add logic inside `generateExcel()` in `src/excelGenerator.js`
- Shared style helpers (`font`, `fill`, `align`, `bdr`) are already at the top of that file — add new helpers there

**New HTTP routes:**
- Add to `server.js` after the existing routes, before the `app.listen` call

**New UI fields:**
- HTML is the `HTML_PAGE` string constant in `server.js` (lines 16-259)
- The `generate()` JS function in the inline `<script>` builds the FormData — add new `form.append(...)` calls there
- Add corresponding `req.body` extraction in the `POST /generate` handler

**Utilities shared across modules:**
- Place in `src/gpsParser.js` and re-export (current pattern: `excelGenerator.js` imports date/calendar helpers from `gpsParser.js`)
- If the utility has no relation to GPS parsing, consider a new `src/utils.js` module

## Special Directories

**`.planning/`:**
- Purpose: GSD planning workspace — roadmaps, phase plans, codebase docs
- Generated: By GSD commands
- Committed: Yes

**`node_modules/`:**
- Purpose: npm dependency installation
- Generated: Yes
- Committed: No

**`.venv/`:**
- Purpose: Legacy Python virtualenv
- Generated: Yes
- Committed: No (excluded by `.gitignore`)

---

*Structure analysis: 2026-03-31*
