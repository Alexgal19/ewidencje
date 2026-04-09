# Coding Conventions

**Analysis Date:** 2026-03-31

## Naming Patterns

**Files:**
- `camelCase` for module files: `xlsParser.js`, `gpsParser.js`, `excelGenerator.js`
- `camelCase` for the main entry point: `server.js`
- No index barrels — each module is imported directly by path

**Functions:**
- `camelCase` for all functions: `parseGps`, `buildRoute`, `generateExcel`, `xlsReadRows`, `getPolishHolidays`
- Verb-noun pattern for operations: `parseGpsAddr`, `shortenAddr`, `extractCity`, `loadKnownCities`, `saveKnownCities`
- Helper utilities use short, descriptive names: `pad2`, `toISO`, `addDays`, `weekdayOf`, `daysInMonth`
- Internal-only helpers like `side`, `bdr`, `dottedBdr`, `thickOutline`, `font`, `fill`, `align` in `src/excelGenerator.js` are NOT exported

**Variables:**
- `camelCase` for local variables: `dayGroups`, `dateFrom`, `runningOdo`, `lastKnownCity`
- Short single-letter names acceptable in tight loops: `r`, `c`, `d`, `k`, `v`, `p2`
- Constants defined at module scope in `SCREAMING_SNAKE_CASE`: `MAGIC`, `ENDOFCHAIN`, `FREESECT`, `XL_EPOCH`, `ZIP_TO_CITY`, `ZIP_CITY_RE`, `ZIP_RE`, `STREET_PREFIX_RE`, `PL_DAYS`
- Buffer/raw abbreviations: `buf`, `sd`, `rl`, `rt`, `rk`

**Parameters:**
- Descriptive names that match their role: `buffer`, `filename`, `odometer`, `refuelSet`, `tripPurpose`
- Abbreviated counterparts for internal use: `odoRaw`, `targetOdoRaw`, `refuelRaw`

## Code Style

**Formatting:**
- No `.prettierrc` or formatter config present — style is manually consistent
- 4-space indentation throughout all `.js` files
- Single quotes for string literals in Node.js modules
- Template literals used for ISO date construction and URL building
- Alignment padding with spaces to vertically align `=` signs in related assignments is used extensively in `server.js`:
  ```js
  const driver       = (req.body.driver_name      || '').trim();
  const odoRaw       = (req.body.odometer_start   || '0').trim();
  const refuelRaw    = (req.body.refuel_dates      || '').trim();
  ```

**Strict Mode:**
- `'use strict';` appears at the top of every `.js` file: `server.js`, `src/gpsParser.js`, `src/xlsParser.js`, `src/excelGenerator.js`

**Linting:**
- No project-level `.eslintrc` present
- Inline `/* eslint-disable */` and `/* eslint-enable */` directives used in `server.js` around the embedded HTML string to suppress `no-useless-escape`

## Import Organization

**Order (consistently applied in all files):**
1. Node.js built-ins: `fs`, `path`, `https`
2. Third-party packages: `express`, `multer`, `exceljs`
3. Local modules: `./src/gpsParser`, `./src/excelGenerator`, `./xlsParser`

**Style:**
- Named destructuring imports for local modules:
  ```js
  const { parseGps, aggregate, aggregateActual, getPreviousWorkingDay } = require('./src/gpsParser');
  const { generateExcel } = require('./src/excelGenerator');
  const { xlsReadRows } = require('./xlsParser');
  ```
- Assigned imports for packages: `const express = require('express');`

**No path aliases** — all paths are relative.

## Module Exports Style

**Named exports via `module.exports` object literal** at the bottom of each file:

`src/gpsParser.js` exports a flat object of all public functions:
```js
module.exports = {
    toISO, addDays, weekdayOf, daysInMonth, getPolishHolidays,
    getPreviousWorkingDay, parseGps, aggregate, aggregateActual,
    buildRoute, removeDigitsFromName, parseGpsAddr, extractCity, shortenAddr,
};
```

`src/xlsParser.js` exports a single function:
```js
module.exports = { xlsReadRows };
```

`src/excelGenerator.js` exports a single function:
```js
module.exports = { generateExcel };
```

**Internal helpers are NOT exported** — `font`, `fill`, `align`, `bdr`, `thickOutline`, `colLetter`, `parseISO`, `isoToDate`, and a local `toISO` re-declaration in `excelGenerator.js` are kept module-private.

## Error Handling

**In route handlers (`server.js`):**
- Validation errors return `400` with a JSON body: `res.status(400).json({ error: 'Brak pliku GPS' })`
- Unexpected errors are caught by a top-level `try/catch` block, logged with `console.error(err)`, and returned as `500` with both `error` and `detail` (stack trace) in the JSON body:
  ```js
  res.status(500).json({ error: String(err.message || err), detail: err.stack || '' });
  ```

**In library modules (`gpsParser.js`, `xlsParser.js`):**
- Throw `new Error(...)` with Polish-language messages for user-facing parsing failures:
  ```js
  throw new Error('Plik nie jest formatem XLS (OLE). Użyj .xls z GPS trackera.');
  throw new Error(`Nieobsługiwany format pliku: ${ext}. Użyj .xls lub .xlsx`);
  throw new Error('Nie znaleziono strumienia Workbook w pliku XLS.');
  ```
- Silent fallback `catch` blocks are used for non-critical operations (SST string decoding, city file reads, date parsing) — the caught variable `e` is discarded silently
- Network errors in `httpsGet` resolve to `null` rather than rejecting:
  ```js
  .on('error', () => resolve(null));
  ```

**In `excelGenerator.js`:**
- No explicit error handling — errors from `ExcelJS` or `buildRoute` propagate up to the server's `try/catch`

## Logging

**Framework:** `console` (Node.js built-in only)

**Patterns:**
- `console.log(...)` used only for server startup: `console.log(\`Server running on http://0.0.0.0:${PORT}\`)`
- `console.error(err)` used in route handler for unexpected errors
- `console.error('Error saving known cities:', e)` in `saveKnownCities` for non-fatal I/O errors
- No structured logging, no log levels, no timestamps

## Comments

**Style:**
- Section dividers use a consistent banner format with box-drawing characters:
  ```js
  // ── Date helpers ──────────────────────────────────────────────────────────────
  ```
- JSDoc-style block comments (`/** ... */`) on all exported functions and key parsing functions
- Inline comments explain non-obvious logic (Easter algorithm variables, BIFF8 record types, OLE structure offsets)
- No TODO, FIXME, or HACK comments present in the Node.js source files

## Function Design

**Size:** Functions vary widely. `generateExcel` in `src/excelGenerator.js` is the largest (~480 lines) and handles the entire Excel document structure as a single function. `parseGps` in `src/gpsParser.js` is similarly large (~100 lines). Helper functions are small (1–10 lines).

**Parameters:**
- Public API functions use positional parameters — no options objects
- `generateExcel` signature: `(plate, carModel, dateFrom, dateTo, driver, odometer, refuelSet, agg, tripPurpose)`
- Default parameter values used for optional args: `function buildRoute(addresses, lastKnownCity = '')`, `function xlsReadRows(xlsBytes, dateCols = [8, 11])`

**Return Values:**
- Async functions return `Promise<Buffer>` (`generateExcel`), `Promise<object>` (`parseGps`), or `Promise<string>` (`geocodeCityNominatim`, `shortenAddr`, `extractCity`)
- `buildRoute` returns a 2-tuple `[routeStr, lastKnownCity]` — a positional array, not an object
- `parseGpsAddr` returns a plain object `{ streetClean, cityName, cityKey }`
- `aggregate` / `aggregateActual` return a `Map<string, object>`

---

*Convention analysis: 2026-03-31*
