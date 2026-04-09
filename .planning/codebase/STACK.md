# Technology Stack

**Analysis Date:** 2026-03-31

## Languages

**Primary:**
- JavaScript (ES2020+, `'use strict'`) - All server and business logic

**Secondary:**
- HTML/CSS/Vanilla JS - Frontend UI embedded as a template string in `server.js`

## Runtime

**Environment:**
- Node.js 18 (slim) - specified via `FROM node:18-slim` in `Dockerfile`

**Package Manager:**
- npm (no version pinned)
- Lockfile: `package-lock.json` present

## Frameworks

**Core:**
- Express 4.18.2 - HTTP server and routing (`server.js`)

**Build/Dev:**
- None - no transpilation, bundling, or build step required
- Single `npm start` → `node server.js`

## Key Dependencies

**Critical:**
- `exceljs` ^4.4.0 - Generates `.xlsx` output workbooks and reads `.xlsx` GPS input files; used in `src/excelGenerator.js` and `src/gpsParser.js`
- `multer` ^1.4.5-lts.1 - Multipart file upload handling; in-memory storage (`memoryStorage()`), 50 MB limit; used in `server.js`

**Infrastructure:**
- None beyond the three production dependencies listed in `package.json`

**No dev dependencies** are declared in `package.json`.

## XLS Parsing

- `.xls` (BIFF8/OLE Compound Document) files are parsed with a **hand-written pure-JS parser** at `src/xlsParser.js` — no third-party XLS library is used.
- `.xlsx` files are parsed via `exceljs` (`wb.xlsx.load(buffer)`) in `src/gpsParser.js`.

## Configuration

**Environment:**
- `PORT` env var — server listens on `process.env.PORT || 8080`
- No other env vars are read at runtime

**Build:**
- `Dockerfile` — `node:18-slim` base, production-only install (`npm install --omit=dev`), copies `server.js` and `src/`
- `Procfile` — `web: node server.js` (Heroku/Firebase App Hosting compatible)

## Platform Requirements

**Development:**
- Node.js 18+
- `npm install` to restore dependencies

**Production:**
- Any container or PaaS that supports Docker or a `Procfile` (Cloud Run, Firebase App Hosting, Heroku, Railway, Render, etc.)
- No database, no filesystem persistence required at runtime (optional `known_cities.json` cache written to `process.cwd()`)

## Legacy / Reference Only

- `requirements.txt` lists Python deps (`flask`, `openpyxl`, `gunicorn`, `requests`) from a previous Python/Flask implementation (`app.py` still present in the repo). These are **not used** by the current Node.js server.

---

*Stack analysis: 2026-03-31*
