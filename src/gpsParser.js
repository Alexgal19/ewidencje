'use strict';

const fs   = require('fs');
const path = require('path');
const http = require('https');
const ExcelJS = require('exceljs');
const { xlsReadRows } = require('./xlsParser');

// ── ExcelJS cell value extractor ─────────────────────────────────────────────

/**
 * Extract a plain value from an ExcelJS cell value.
 * Handles rich text objects, formula results, hyperlinks, etc.
 */
function extractCellValue(v) {
    if (v === null || v === undefined) return null;
    if (v instanceof Date) return v;
    if (typeof v !== 'object') return v;
    // Rich text: { richText: [{ text: '...' }, ...] }
    if (Array.isArray(v.richText)) {
        return v.richText.map(rt => rt.text || '').join('');
    }
    // Formula result
    if (v.result !== undefined) {
        const r = v.result;
        if (r instanceof Date) return r;
        if (typeof r !== 'object') return r;
        if (r && Array.isArray(r.richText)) return r.richText.map(rt => rt.text || '').join('');
        return null;
    }
    // Hyperlink / shared string with text property
    if (v.text !== undefined) return v.text;
    return null;
}

// ── Date helpers ──────────────────────────────────────────────────────────────

/** Zero-pad a number to 2 digits */
function pad2(n) { return String(n).padStart(2, '0'); }

/** Build an ISO date string 'YYYY-MM-DD'. month is 1-indexed. */
function toISO(year, month, day) {
    return `${year}-${pad2(month)}-${pad2(day)}`;
}

/** Add n days to an ISO string, return ISO string */
function addDays(isoStr, n) {
    const d = new Date(isoStr + 'T00:00:00Z');
    d.setUTCDate(d.getUTCDate() + n);
    return d.toISOString().slice(0, 10);
}

/**
 * Return weekday: 0=Mon, 1=Tue, ... 6=Sun
 */
function weekdayOf(isoStr) {
    return (new Date(isoStr + 'T00:00:00Z').getUTCDay() + 6) % 7;
}

/** Return the number of days in a given month */
function daysInMonth(year, month) {
    return new Date(Date.UTC(year, month, 0)).getUTCDate();
}

// ── Polish holidays ───────────────────────────────────────────────────────────

/**
 * Returns a Set<string> of ISO date strings for Polish public holidays in a year.
 */
function getPolishHolidays(year) {
    const h = new Set();

    function add(m, d) { h.add(toISO(year, m, d)); }

    add(1, 1);   // Nowy Rok
    add(1, 6);   // Trzech Króli
    add(5, 1);   // Święto Pracy
    add(5, 3);   // Święto Konstytucji 3 Maja
    add(8, 15);  // Wniebowzięcie NMP
    add(11, 1);  // Wszystkich Świętych
    add(11, 11); // Święto Niepodległości
    add(12, 24); // Wigilia
    add(12, 25); // Boże Narodzenie 1
    add(12, 26); // Boże Narodzenie 2

    // Easter (Meeus/Jones/Butcher)
    const a = year % 19;
    const b = Math.floor(year / 100);
    const c = year % 100;
    const d = Math.floor(b / 4);
    const e = b % 4;
    const f = Math.floor((b + 8) / 25);
    const g = Math.floor((b - f + 1) / 3);
    const hh = (19 * a + b - d - g + 15) % 30;
    const i = Math.floor(c / 4);
    const k = c % 4;
    const l = (32 + 2 * e + 2 * i - hh - k) % 7;
    const m2 = Math.floor((a + 11 * hh + 22 * l) / 451);
    const month = Math.floor((hh + l - 7 * m2 + 114) / 31);
    const day   = ((hh + l - 7 * m2 + 114) % 31) + 1;

    const easterISO = toISO(year, month, day);
    h.add(easterISO);
    h.add(addDays(easterISO, 1));  // Poniedziałek Wielkanocny
    h.add(addDays(easterISO, 60)); // Boże Ciało

    return h;
}

/**
 * Get the nearest previous working day (Mon-Fri, not a holiday) for isoStr.
 * If going backwards crosses into previous month, go forward instead.
 */
function getPreviousWorkingDay(isoStr) {
    let current = isoStr;

    while (true) {
        const year = parseInt(current.slice(0, 4), 10);
        const holidays = getPolishHolidays(year);
        const wd = weekdayOf(current);
        if (wd >= 5 || holidays.has(current)) {
            current = addDays(current, -1);
        } else {
            break;
        }
    }

    // If we've drifted into the previous month, go forward instead
    const targetMonth = isoStr.slice(0, 7);
    if (current.slice(0, 7) === targetMonth) {
        return current;
    }

    // Re-start going forward from original date
    current = isoStr;
    while (true) {
        const year = parseInt(current.slice(0, 4), 10);
        const holidays = getPolishHolidays(year);
        const wd = weekdayOf(current);
        if (wd >= 5 || holidays.has(current)) {
            current = addDays(current, 1);
        } else {
            return current;
        }
    }
}

// ── Known cities cache ────────────────────────────────────────────────────────

const ZIP_TO_CITY = {
    '43-245': 'Studzionka',
    '43-246': 'Strumień',
};

const KNOWN_CITIES_FILE = path.join(process.cwd(), 'known_cities.json');

function loadKnownCities() {
    if (fs.existsSync(KNOWN_CITIES_FILE)) {
        try {
            const data = JSON.parse(fs.readFileSync(KNOWN_CITIES_FILE, 'utf-8'));
            return Object.assign({}, ZIP_TO_CITY, data);
        } catch (e) {
            // fall through
        }
    }
    return Object.assign({}, ZIP_TO_CITY);
}

function saveKnownCities(citiesDict) {
    try {
        const toSave = {};
        for (const [k, v] of Object.entries(citiesDict)) {
            if (!(k in ZIP_TO_CITY) || ZIP_TO_CITY[k] !== v) {
                toSave[k] = v;
            }
        }
        fs.writeFileSync(KNOWN_CITIES_FILE, JSON.stringify(toSave, null, 4), 'utf-8');
    } catch (e) {
        console.error('Error saving known cities:', e);
    }
}

// Module-level cache
const onlineCitiesCache = loadKnownCities();

// ── Nominatim geocoding ───────────────────────────────────────────────────────

/**
 * Simple HTTPS GET returning parsed JSON, or null on error.
 */
function httpsGet(url) {
    return new Promise((resolve) => {
        http.get(url, {
            headers: { 'User-Agent': 'GeneratorEwidencji/4.0 (contact@example.com)' },
            timeout: 5000,
        }, (res) => {
            let body = '';
            res.on('data', chunk => { body += chunk; });
            res.on('end', () => {
                try { resolve(JSON.parse(body)); }
                catch (e) { resolve(null); }
            });
        }).on('error', () => resolve(null));
    });
}

/**
 * Query Nominatim for a city name from zip code + street.
 * Returns city name string or null.
 */
async function geocodeCityNominatim(zipCode, streetRaw) {
    const query = encodeURIComponent(`${zipCode} ${streetRaw}, Poland`);
    const url = `https://nominatim.openstreetmap.org/search?q=${query}&format=jsonv2&addressdetails=1&limit=1`;
    // Respect 1 req/sec policy
    await new Promise(r => setTimeout(r, 1100));
    const data = await httpsGet(url);
    if (data && Array.isArray(data) && data.length > 0) {
        const address = data[0].address || {};
        const city = address.city || address.town || address.village || address.municipality || null;
        return city || null;
    }
    return null;
}

// ── Address parsing helpers ───────────────────────────────────────────────────

function removeDigitsFromName(text) {
    if (!text) return '';
    text = String(text);
    text = text.replace(/\b\d{2}-\d{3}\b/g, '');
    text = text.replace(/\b\d+\b/g, '');
    text = text.replace(/\b(ulica|ul\.|al\.|aleja|plac|pl\.)\s+/gi, '');
    // Remove administrative divisions: Powiat, Gmina, Województwo (incl. abbreviations pow., gm., woj.)
    text = text.replace(/\b(Powiat|Gmina|Województwo|gm\.|pow\.|woj\.)\s+.*$/gi, '');
    return text.split(/\s+/).join(' ').trim().replace(/[,.\s]+$/, '');
}

const STREET_PREFIX_RE = /\b(ulica|ul\.|al\.|aleja|plac|pl\.|straße|str\.|avenue|ave\.|road|rd\.|street|st\.)\b/i;

function looksLikeStreet(text) {
    if (!text) return false;
    if (/\d/.test(text)) return true;
    if (STREET_PREFIX_RE.test(text)) return true;
    return false;
}

const ZIP_CITY_RE = /^(\d{2}-\d{3})\s+(.+)$/;
const ZIP_RE = /^\d{2}-\d{3}$/;

/**
 * Parse a raw GPS address string.
 * Returns { streetClean, cityName, cityKey }
 */
async function parseGpsAddr(addr) {
    if (!addr) return { streetClean: '', cityName: '', cityKey: '' };

    let raw = String(addr);
    raw = raw.replace(/, Poland/g, '').replace(/, Polska/g, '');
    raw = raw.trim().replace(/,\s*$/, '');

    let parts = raw.split(',').map(p => p.trim()).filter(p => p);
    if (!parts.length) return { streetClean: '', cityName: '', cityKey: '' };

    // POI detection
    if (parts.length >= 2 && !looksLikeStreet(parts[0]) && looksLikeStreet(parts[1])) {
        parts = parts.slice(1);
    }

    const streetRaw   = parts[0];
    const streetClean = removeDigitsFromName(streetRaw);

    const rest = parts.slice(1);
    let zipCode  = '';
    let cityName = '';

    for (const part of rest) {
        if (ZIP_RE.test(part)) {
            if (!zipCode) zipCode = part;
        } else {
            const mZC = ZIP_CITY_RE.exec(part);
            if (mZC) {
                if (!zipCode) zipCode = mZC[1];
                // Prefer the city name associated with the zip code and stop looking further
                cityName = mZC[2].trim();
                break;
            } else {
                // Only set cityName if we haven't found one yet, to avoid overwriting with broader regions
                if (!cityName) {
                    const p = part.trim();
                    // Skip parts that look like administrative regions
                    if (!/^(Powiat|Gmina|Województwo|gm\.|pow\.|woj\.)/i.test(p)) {
                        cityName = p;
                    }
                }
            }
        }
    }

    if (!cityName && zipCode) {
        if (onlineCitiesCache[zipCode]) {
            cityName = onlineCitiesCache[zipCode];
        } else {
            const found = await geocodeCityNominatim(zipCode, streetRaw);
            if (found) {
                cityName = found;
                onlineCitiesCache[zipCode] = cityName;
                saveKnownCities(onlineCitiesCache);
            }
        }
    }

    const cityKey = cityName || zipCode;
    return { streetClean, cityName, cityKey };
}

async function shortenAddr(addr) {
    if (!addr) return '';
    const { streetClean, cityName } = await parseGpsAddr(addr);
    const street = removeDigitsFromName(streetClean);
    const city   = removeDigitsFromName(cityName);
    if (street && city && street.toLowerCase() !== city.toLowerCase()) {
        return `${street}, ${city}`;
    }
    return street || city;
}

async function extractCity(addr) {
    if (!addr) return '';
    const { cityName } = await parseGpsAddr(addr);
    return removeDigitsFromName((cityName || '').trim());
}

// ── GPS file parsing ──────────────────────────────────────────────────────────

/**
 * Parse a GPS report file (.xls or .xlsx).
 * Returns { plate, carModel, dateFrom, dateTo, dayGroups }
 * dateFrom and dateTo are ISO strings.
 * dayGroups entries: { date: isoStr, km: number, addresses: string[], firstStart: string, lastEnd: string }
 */
async function parseGps(buffer, filename) {
    const ext = path.extname(filename).toLowerCase();

    let rows;
    if (ext === '.xls') {
        rows = xlsReadRows(buffer, [5, 6, 7, 8, 9, 10, 11]);
    } else if (ext === '.xlsx') {
        const wb = new ExcelJS.Workbook();
        await wb.xlsx.load(buffer);
        const ws = wb.worksheets[0];
        rows = [];
        ws.eachRow({ includeEmpty: true }, (row) => {
            // Properly extract cell values, handling rich text, formulas, hyperlinks
            rows.push(row.values.slice(1).map(extractCellValue));
        });
    } else {
        throw new Error(`Nieobsługiwany format pliku: ${ext}. Użyj .xls lub .xlsx`);
    }

    let plate = null;
    let carModel = null;
    let dateFrom = null;
    let dateTo   = null;

    // ── Auto-detect column layout ────────────────────────────────────────────
    // Old XLS format: date@col8, startAddr@col9,  endAddr@col12, km@col16 in Razem row
    // New XLSX format: date@col10, startAddr@col12, endAddr@col17, km@col23 in Razem row
    //
    // Detect by finding first data row that has a Date object.
    let dateCol      = 8;
    let startAddrCol = 9;
    let endAddrCol   = 12;
    let kmCols       = [16, 15, 17, 14, 18]; // columns to search for km in "Razem" row

    const firstDateRow = rows.find(r => r && r[10] instanceof Date);
    if (firstDateRow) {
        // New XLSX format detected
        dateCol      = 10;
        startAddrCol = 12;
        endAddrCol   = 17;
        kmCols       = [23, 6, 7, 8]; // km in Razem row, col23 = total distance
        console.log('[Parser] Detected NEW XLSX format (date@col10, addr@col12/17, km@col23)');
    } else {
        console.log('[Parser] Detected OLD XLS format (date@col8, addr@col9/12, km@col16)');
    }

    // ── Scan header rows for metadata ───────────────────────────────────────
    for (let i = 0; i < Math.min(rows.length, 30); i++) {
        const row = rows[i];
        if (!row || !row.length) continue;

        // Search all cells for date range strings
        for (const cell of row) {
            if (cell == null) continue;
            const v = String(cell);
            if (!dateFrom && v.includes('Data:')) {
                const m = /(\d{4}-\d{2}-\d{2}).*?(\d{4}-\d{2}-\d{2})/.exec(v);
                if (m) { dateFrom = m[1]; dateTo = m[2]; }
            }
            // Also detect date ranges like "01.04.2026 - 30.04.2026"
            if (!dateFrom) {
                const m2 = /(\d{2}\.\d{2}\.(\d{4})).*?(\d{2}\.\d{2}\.(\d{4}))/.exec(v);
                if (m2) {
                    const parseDMY = s => { const [d, mo, y] = s.split('.'); return `${y}-${mo.padStart(2,'0')}-${d.padStart(2,'0')}`; };
                    dateFrom = parseDMY(m2[1]); dateTo = parseDMY(m2[3]);
                }
            }
        }

        // Old-format plate/carModel detection at row 12
        if (i === 12 && dateCol === 8) {
            carModel = row[0] != null ? String(row[0]).trim() : '';
            plate    = (row.length > 2 && row[2] != null) ? String(row[2]).trim() : '';
        }
    }

    const dayGroups = [];
    let current = [];

    console.log(`[Parser] Processing ${rows.length} rows from ${filename}`);

    for (const row of rows) {
        if (!row) continue;

        // Summary row detection
        const col0 = row[0] != null ? String(row[0]).trim() : '';
        const col1 = row[1] != null ? String(row[1]).trim() : '';
        const isSummary = /^(razem|suma|podsumowanie):?$/i.test(col0) || /^(razem|suma|podsumowanie):?$/i.test(col1);

        if (isSummary) {
            if (current.length) {
                const data = current.filter(r => r && r[dateCol] instanceof Date);

                if (data.length) {
                    // Grab plate from first data row if not yet set
                    if (!plate) {
                        plate = data[0][0] != null ? String(data[0][0]).trim() : '';
                    }

                    const dateVal = data[0][dateCol];
                    const dateISO = toISO(dateVal.getUTCFullYear(), dateVal.getUTCMonth() + 1, dateVal.getUTCDate());

                    // Find km in the Razem row
                    let km = 0;
                    for (const c of kmCols) {
                        const val = row[c];
                        const num = typeof val === 'number' ? val : parseFloat(String(val ?? ''));
                        if (!isNaN(num) && num > 0) { km = num; break; }
                    }

                    // Collect addresses from trip rows
                    const addresses = [];
                    for (const r of data) {
                        const s = r[startAddrCol];
                        const e = r[endAddrCol];
                        if (s != null && String(s).trim()) addresses.push(String(s));
                        if (e != null && String(e).trim()) addresses.push(String(e));
                    }

                    const firstStart = data[0][startAddrCol] != null ? String(data[0][startAddrCol]) : '';
                    const lastEnd    = data[data.length - 1][endAddrCol] != null ? String(data[data.length - 1][endAddrCol]) : '';

                    dayGroups.push({ date: dateISO, km, addresses, firstStart, lastEnd });
                }
            }
            current = [];
        } else {
            current.push(row);
        }
    }

    // Derive dateFrom/dateTo from actual trip dates if header scan didn't find them
    if (!dateFrom && dayGroups.length) {
        const sorted = [...dayGroups].sort((a, b) => a.date.localeCompare(b.date));
        dateFrom = sorted[0].date;
        dateTo   = sorted[sorted.length - 1].date;
    }

    if (dayGroups.length === 0) {
        console.warn(`[Parser] Warning: No GPS data found in ${filename}. Checked ${rows.length} rows.`);
    } else {
        console.log(`[Parser] Successfully extracted ${dayGroups.length} days of data.`);
    }

    return { plate, carModel, dateFrom, dateTo, dayGroups };
}

// ── Aggregation ───────────────────────────────────────────────────────────────

/**
 * Aggregate dayGroups, shifting weekends/holidays to the previous working day.
 * Returns Map<isoStr, { km, addresses, dailyAddresses, firstLoc, lastLoc }>
 */
function aggregate(dayGroups) {
    const agg = new Map();
    const sorted = [...dayGroups].sort((a, b) => a.date.localeCompare(b.date));

    for (const g of sorted) {
        const t = getPreviousWorkingDay(g.date);

        if (!agg.has(t)) {
            agg.set(t, { km: 0, addresses: [], dailyAddresses: [], firstLoc: null, lastLoc: null });
        }
        const entry = agg.get(t);
        entry.km += g.km;
        entry.addresses.push(...g.addresses);

        if (g.addresses.length) {
            entry.dailyAddresses.push([...g.addresses]);
        }

        if (entry.firstLoc === null && g.firstStart) {
            entry.firstLoc = g.firstStart;
        }
        if (g.lastEnd) {
            entry.lastLoc = g.lastEnd;
        }
    }

    return agg;
}

/**
 * Aggregate dayGroups without shifting – each day stays as-is.
 * Returns Map<isoStr, { km, addresses, dailyAddresses, firstLoc, lastLoc }>
 */
function aggregateActual(dayGroups) {
    const agg = new Map();
    const sorted = [...dayGroups].sort((a, b) => a.date.localeCompare(b.date));

    for (const g of sorted) {
        const d = g.date;

        if (!agg.has(d)) {
            agg.set(d, { km: 0, addresses: [], dailyAddresses: [], firstLoc: null, lastLoc: null });
        }
        const entry = agg.get(d);
        entry.km += g.km;
        entry.addresses.push(...g.addresses);

        if (g.addresses.length) {
            entry.dailyAddresses.push([...g.addresses]);
        }

        if (entry.firstLoc === null && g.firstStart) {
            entry.firstLoc = g.firstStart;
        }
        if (g.lastEnd) {
            entry.lastLoc = g.lastEnd;
        }
    }

    return agg;
}

// ── Route building ────────────────────────────────────────────────────────────

/**
 * Build a route string from an array of addresses.
 * Returns [routeStr, lastKnownCity].
 * This function calls parseGpsAddr which may do async geocoding,
 * so it is async.
 */
async function buildRoute(addresses, lastKnownCity = '') {
    if (!addresses || !addresses.length) return ['', lastKnownCity];

    const citiesSeq    = [];
    const streetsInCity = [];
    const uniqueCities = [];
    let currentCity = lastKnownCity;

    if (lastKnownCity) {
        citiesSeq.push(lastKnownCity);
        uniqueCities.push(lastKnownCity);
    }

    for (const addr of addresses) {
        const { streetClean: streetRaw, cityName, cityKey } = await parseGpsAddr(addr);

        let street = removeDigitsFromName(streetRaw);
        if (street.includes(',')) {
            street = street.split(',')[0].trim();
        }
        let city = removeDigitsFromName(cityName);

        if (!city) {
            city = currentCity;
        } else {
            currentCity = city;
        }

        if (city && !uniqueCities.includes(city)) {
            uniqueCities.push(city);
        }

        if (city && (!citiesSeq.length || citiesSeq[citiesSeq.length - 1] !== city)) {
            citiesSeq.push(city);
        }

        if (street && !streetsInCity.includes(street)) {
            if (street.toLowerCase() !== city.toLowerCase()) {
                streetsInCity.push(street);
            }
        }
    }

    const validCities = citiesSeq.filter(c => c);

    if (uniqueCities.length > 1) {
        return [validCities.join('-'), currentCity];
    } else {
        const city = uniqueCities[0] || currentCity;
        const streets = streetsInCity.filter(s => s && s.toLowerCase() !== (city || '').toLowerCase());

        if (city && streets.length) {
            return [`${city}: ${streets.join(', ')}`, currentCity];
        } else if (city) {
            return [city, currentCity];
        } else {
            return [streets.join(', '), currentCity];
        }
    }
}

module.exports = {
    toISO,
    addDays,
    weekdayOf,
    daysInMonth,
    getPolishHolidays,
    getPreviousWorkingDay,
    parseGps,
    aggregate,
    aggregateActual,
    buildRoute,
    removeDigitsFromName,
    parseGpsAddr,
    extractCity,
    shortenAddr,
};
