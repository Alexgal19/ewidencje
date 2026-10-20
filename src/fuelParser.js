'use strict';

const path = require('path');
const ExcelJS = require('exceljs');
const { xlsReadRows } = require('./xlsParser');
const { extractCellValue } = require('./gpsParser');

/**
 * Excel serial date (Windows 1900 date system) to YYYY-MM-DD.
 * e.g. 46280.2949 -> 2026-09-15
 */
function excelSerialToIsoDate(serial) {
    const num = typeof serial === 'number' ? serial : parseFloat(String(serial).replace(',', '.'));
    if (isNaN(num) || num < 20000 || num > 70000) return null;
    const utcDays = Math.floor(num - 25569);
    const date = new Date(utcDays * 86400 * 1000);
    if (isNaN(date.getTime())) return null;
    return date.toISOString().split('T')[0];
}

/**
 * Normalizes any date value (Excel serial, Date instance, ISO string, DD.MM.YYYY) to YYYY-MM-DD.
 */
function parseAnyDateToIso(val) {
    if (val == null || val === '') return null;
    if (val instanceof Date) {
        if (isNaN(val.getTime())) return null;
        return val.toISOString().split('T')[0];
    }
    if (typeof val === 'number') {
        return excelSerialToIsoDate(val);
    }
    const str = String(val).trim();

    // Check if it's a numeric string like "46280.2949537037"
    if (/^\d{5}(\.\d+)?$/.test(str)) {
        const iso = excelSerialToIsoDate(parseFloat(str));
        if (iso) return iso;
    }

    // YYYY-MM-DD or YYYY.MM.DD or YYYY/MM/DD
    const mIso = /^(\d{4})[.\-\/](\d{2})[.\-\/](\d{2})/.exec(str);
    if (mIso) {
        return `${mIso[1]}-${mIso[2]}-${mIso[3]}`;
    }

    // DD.MM.YYYY or DD-MM-YYYY or DD/MM/YYYY
    const mDmy = /^(\d{1,2})[.\-\/](\d{1,2})[.\-\/](\d{4})/.exec(str);
    if (mDmy) {
        const d = String(mDmy[1]).padStart(2, '0');
        const m = String(mDmy[2]).padStart(2, '0');
        const y = mDmy[3];
        return `${y}-${m}-${d}`;
    }

    return null;
}

/** Clean vehicle registration plate (e.g. " DSW0260A " -> "DSW0260A"). */
function cleanPlate(plate) {
    return String(plate || '').toUpperCase().replace(/[^A-Z0-9]/g, '');
}

/**
 * Checks whether the transaction represents actual vehicle engine fuel (ON, Pb95, Pb98, LPG),
 * strictly ignoring non-fuel purchases like:
 * cards, car wash, vacuum, coffee, tea, highway tolls, AdBlue, NOXY, fluids, oils.
 */
function isFuelTransaction(productName, unit, group) {
    const name = String(productName || '').toLowerCase().trim();
    const u = String(unit || '').toUpperCase().trim();
    const g = String(group || '').trim();

    // Unit of measure: If provided, must be Liters ('L' or 'LTR' or 'LITR')
    if (u && u !== 'L' && u !== 'LTR' && u !== 'LITR') {
        return false;
    }

    // Strict blacklist of non-fuel products
    const NON_FUEL_REGEX = /\b(karta|myjnia|odkurzacz|kawa|herbata|op[lł]ata|a4|przejazd|autostrada|adblue|noxy|p[lł]yn|olej|woda|nap[oó]j|hot\s*dog|przek[aą]ska|us[lł]uga)\b/i;
    if (NON_FUEL_REGEX.test(name)) {
        return false;
    }

    // Known Orlen fuel product groups:
    // 01 = ON / Efecta Diesel
    // 02 = Verva ON
    // 03 = Efecta 95
    // 04 = Eurosuper 95
    // 05 = Verva 98
    // 06 = Gaz LPG
    // 07 = CNG / LNG
    const FUEL_GROUPS = new Set(['01', '02', '03', '04', '05', '06', '07', '1', '2', '3', '4', '5', '6', '7']);
    if (FUEL_GROUPS.has(g)) {
        return true;
    }

    // Fuel name whitelist
    const FUEL_NAME_REGEX = /\b(efecta|verva|diesel|\bon\b|gaz|lpg|pb\s*95|pb\s*98|pb95|pb98|benzyna|paliwo|eurosuper|super)\b/i;
    if (FUEL_NAME_REGEX.test(name)) {
        return true;
    }

    // If unit is Liters and group is a fuel group or name matches
    return false;
}

/** Parse CSV text into array of string arrays, handling quotes and semicolons/commas. */
function parseCsvRows(text) {
    const lines = text.split(/\r?\n/).filter(line => line.trim().length > 0);
    if (!lines.length) return [];

    // Detect delimiter from first non-empty line
    const firstLine = lines[0];
    const commaCount = (firstLine.match(/,/g) || []).length;
    const semicolonCount = (firstLine.match(/;/g) || []).length;
    const tabCount = (firstLine.match(/\t/g) || []).length;

    let delimiter = ',';
    if (semicolonCount > commaCount && semicolonCount > tabCount) delimiter = ';';
    else if (tabCount > commaCount && tabCount > semicolonCount) delimiter = '\t';

    const rows = [];
    for (const line of lines) {
        const row = [];
        let cur = '';
        let inQuotes = false;
        for (let i = 0; i < line.length; i++) {
            const char = line[i];
            if (char === '"') {
                if (inQuotes && line[i + 1] === '"') {
                    cur += '"';
                    i++;
                } else {
                    inQuotes = !inQuotes;
                }
            } else if (char === delimiter && !inQuotes) {
                row.push(cur.trim());
                cur = '';
            } else {
                cur += char;
            }
        }
        row.push(cur.trim());
        rows.push(row);
    }
    return rows;
}

/**
 * Main parser: takes a file buffer or string and parses all fuel transactions grouped by vehicle plate.
 *
 * @param {Buffer|string} bufferOrString
 * @param {string} filename
 * @returns {Promise<Map<string, Set<string>>>} Map of plate -> Set of ISO dates ("YYYY-MM-DD")
 */
async function parseFuelFile(bufferOrString, filename = 'tankowania.csv') {
    const ext = path.extname(filename).toLowerCase();
    let rows = [];

    if (ext === '.xlsx') {
        const wb = new ExcelJS.Workbook();
        const buf = Buffer.isBuffer(bufferOrString) ? bufferOrString : Buffer.from(bufferOrString);
        await wb.xlsx.load(buf);
        const ws = wb.worksheets[0];
        ws.eachRow({ includeEmpty: true }, (row) => {
            rows.push(row.values.slice(1).map(extractCellValue));
        });
    } else if (ext === '.xls') {
        const buf = Buffer.isBuffer(bufferOrString) ? bufferOrString : Buffer.from(bufferOrString);
        rows = xlsReadRows(buf);
    } else {
        // Assume CSV or plain text
        const text = Buffer.isBuffer(bufferOrString) ? bufferOrString.toString('utf8') : String(bufferOrString);
        rows = parseCsvRows(text);
    }

    if (!rows || rows.length < 2) {
        return new Map();
    }

    // Find header row and column indexes
    let headerRowIdx = -1;
    let colPlate = -1;
    let colDate = -1;
    let colProduct = -1;
    let colUnit = -1;
    let colGroup = -1;

    for (let r = 0; r < Math.min(rows.length, 10); r++) {
        const row = rows[r];
        if (!row) continue;
        for (let c = 0; c < row.length; c++) {
            const val = String(row[c] || '').toLowerCase().trim();
            if (colPlate === -1 && /rejestrac|rej\b|plate|vrn|pojazd/i.test(val)) colPlate = c;
            if (colDate === -1 && /data\s*transakcji|data\s*tankowania|^data\b/i.test(val)) colDate = c;
            if (colProduct === -1 && /nazwa\s*produktu|^produkt|^towar/i.test(val)) colProduct = c;
            if (colUnit === -1 && /^jm$|jednostka|^unit$/i.test(val)) colUnit = c;
            if (colGroup === -1 && /grupa\s*produktu|^grupa|^group$/i.test(val)) colGroup = c;
        }
        if (colPlate !== -1 && colDate !== -1) {
            headerRowIdx = r;
            break;
        }
    }

    // Fallbacks if headers weren't found by name (e.g. Orlen standard fixed columns: 2=Plate, 3=Date, 7=Product, 9=Unit, 12=Group)
    if (colPlate === -1) colPlate = 2;
    if (colDate === -1) colDate = 3;
    if (colProduct === -1) colProduct = 7;
    if (colUnit === -1) colUnit = 9;
    if (colGroup === -1) colGroup = 12;

    const startRow = headerRowIdx >= 0 ? headerRowIdx + 1 : 1;
    const datesByPlate = new Map();

    for (let i = startRow; i < rows.length; i++) {
        const row = rows[i];
        if (!row || row.length <= colPlate) continue;

        const plate = cleanPlate(row[colPlate]);
        if (!plate || plate.length < 3) continue;

        const rawDate = row[colDate];
        const dateIso = parseAnyDateToIso(rawDate);
        if (!dateIso) continue;

        const product = colProduct < row.length ? row[colProduct] : '';
        const unit = colUnit < row.length ? row[colUnit] : '';
        const group = colGroup < row.length ? row[colGroup] : '';

        // Filter: ONLY fuel transactions
        if (!isFuelTransaction(product, unit, group)) {
            continue;
        }

        if (!datesByPlate.has(plate)) {
            datesByPlate.set(plate, new Set());
        }
        datesByPlate.get(plate).add(dateIso);
    }

    return datesByPlate;
}

module.exports = {
    excelSerialToIsoDate,
    parseAnyDateToIso,
    cleanPlate,
    isFuelTransaction,
    parseCsvRows,
    parseFuelFile,
};
