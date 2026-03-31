'use strict';

const ExcelJS = require('exceljs');
const {
    weekdayOf,
    daysInMonth,
    getPolishHolidays,
    getPreviousWorkingDay,
    buildRoute,
} = require('./gpsParser');

const PL_DAYS = { 0: 'pon', 1: 'wt', 2: 'śr', 3: 'czw', 4: 'pt', 5: 'sob', 6: 'niedz' };

// Column letter helper (1=A, 2=B, ...)
function colLetter(n) {
    if (n <= 26) return String.fromCharCode(64 + n);
    // two-letter columns (AA, AB, ...)
    return String.fromCharCode(64 + Math.floor((n - 1) / 26)) + String.fromCharCode(65 + ((n - 1) % 26));
}

// ── Border helpers ────────────────────────────────────────────────────────────

function side(style) {
    return style ? { style } : undefined;
}

function bdr(top, right, bottom, left) {
    const b = {};
    if (top)    b.top    = { style: top };
    if (right)  b.right  = { style: right };
    if (bottom) b.bottom = { style: bottom };
    if (left)   b.left   = { style: left };
    return b;
}

function dottedBdr() {
    return { bottom: { style: 'dotted' } };
}

function thickOutline(ci, isFirst, isLast, isTop = false, isBottom = false) {
    return {
        top:    { style: isTop    ? 'medium' : 'thin' },
        bottom: { style: isBottom ? 'medium' : 'thin' },
        left:   { style: isFirst  ? 'medium' : 'thin' },
        right:  { style: isLast   ? 'medium' : 'thin' },
    };
}

// ── Font / fill / align helpers ───────────────────────────────────────────────

function font(name, size, opts = {}) {
    const f = { name, size };
    if (opts.bold)   f.bold = true;
    if (opts.italic) f.italic = true;
    if (opts.color)  f.color = { argb: 'FF' + opts.color };
    return f;
}

function fill(rgb) {
    return { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF' + rgb } };
}

function align(horizontal, vertical, opts = {}) {
    const a = { horizontal, vertical };
    if (opts.wrap) a.wrapText = true;
    if (opts.indent) a.indent = opts.indent;
    return a;
}

// ── Parse ISO date string to { year, month, day } ─────────────────────────────

function parseISO(isoStr) {
    if (!isoStr) return null;
    const parts = isoStr.split('-').map(Number);
    return { year: parts[0], month: parts[1], day: parts[2] };
}

function isoToDate(isoStr) {
    if (!isoStr) return null;
    const { year, month, day } = parseISO(isoStr);
    return new Date(Date.UTC(year, month - 1, day));
}

// ── Main generator ────────────────────────────────────────────────────────────

/**
 * Generate an Excel workbook.
 *
 * @param {string} plate
 * @param {string} carModel
 * @param {string} dateFrom  - ISO date string
 * @param {string} dateTo    - ISO date string
 * @param {string} driver
 * @param {number} odometer
 * @param {Set<string>} refuelSet   - Set of ISO date strings
 * @param {Map<string, object>} agg - Map from isoStr to { km, addresses, dailyAddresses, firstLoc, lastLoc }
 * @param {string} tripPurpose
 * @returns {Promise<Buffer>}
 */
async function generateExcel(plate, carModel, dateFrom, dateTo,
                              driver, odometer, refuelSet, agg,
                              tripPurpose = 'dowóz/odbiór pracowników') {
    const wb = new ExcelJS.Workbook();

    const totalKm = Math.round([...agg.values()].reduce((s, v) => s + v.km, 0) * 100) / 100;
    const odoEnd  = odometer ? (Math.floor(odometer) + Math.floor(totalKm)) : 0;

    const dfParsed = parseISO(dateFrom);
    const dtParsed = parseISO(dateTo);

    // ── Tytułowa ──────────────────────────────────────────────────────────────
    const ws = wb.addWorksheet('Tytułowa');
    ws.views = [{ showGridLines: false }];

    const tytCols = [
        { key: 'A', width: 4 },
        { key: 'B', width: 20 },
        { key: 'C', width: 30 },
        { key: 'D', width: 30 },
        { key: 'E', width: 30 },
        { key: 'F', width: 4 },
    ];
    for (const { key, width } of tytCols) {
        ws.getColumn(key).width = width;
    }

    // Row 2: main title
    ws.mergeCells('B2:E2');
    const c2 = ws.getCell('B2');
    c2.value     = `Ewidencja Przebiegu Pojazdu nr. Rej.   ${plate || ''}`;
    c2.font      = font('Arial', 14, { bold: true, color: '000000' });
    c2.alignment = align('center', 'middle');
    ws.getRow(2).height = 30;

    // Row 4: subtitle
    ws.mergeCells('B4:E4');
    const c4 = ws.getCell('B4');
    c4.value     = 'wykorzystywanego wyłącznie do celów działalności gospodarczej';
    c4.font      = font('Arial', 12, { bold: true, color: '000000' });
    c4.alignment = align('center', 'middle');
    ws.getRow(4).height = 20;

    // Helper: info row (centered value + description below)
    function infoRowCentered(row, val, descr) {
        ws.mergeCells(`C${row}:D${row}`);
        const cv = ws.getCell(`C${row}`);
        cv.value     = val;
        cv.font      = font('Arial', 11, { color: '000000' });
        cv.alignment = align('center', 'bottom');
        cv.border    = dottedBdr();
        ws.getCell(`D${row}`).border = dottedBdr();

        ws.mergeCells(`C${row + 1}:D${row + 1}`);
        const cd = ws.getCell(`C${row + 1}`);
        cd.value     = descr;
        cd.font      = font('Arial', 9, { color: '000000' });
        cd.alignment = align('center', 'top');
    }

    infoRowCentered(6, 'Smart Work Sp. Z o.o.', 'Dane pracodawcy (nazwisko, imię/nazwa*)');
    ws.getRow(6).height = 25;
    infoRowCentered(9, '58-100 Świdnica  Bystrzycka 7a', 'Adres prowadzonej działalności');
    ws.getRow(9).height = 25;
    infoRowCentered(12, '8842754100', 'NIP');
    ws.getRow(12).height = 25;
    ws.getRow(14).height = 20;

    function thickBdr(top, right, bottom, left) {
        const b = {};
        if (top)    b.top    = { style: top };
        if (right)  b.right  = { style: right };
        if (bottom) b.bottom = { style: bottom };
        if (left)   b.left   = { style: left };
        return b;
    }

    // ── ROZPOCZĘCIE EWIDENCJI ────────────────────────────────────────────────
    ws.mergeCells('B15:E15');
    const c15 = ws.getCell('B15');
    c15.value     = 'ROZPOCZĘCIE EWIDENCJI';
    c15.font      = font('Arial', 10, { bold: true, color: '000000' });
    c15.alignment = align('center', 'middle');
    for (const col of ['B', 'C', 'D', 'E']) {
        ws.getCell(`${col}15`).border = thickBdr(
            'medium',
            col === 'E' ? 'medium' : undefined,
            'thin',
            col === 'B' ? 'medium' : undefined
        );
    }
    ws.getRow(15).height = 20;

    ws.mergeCells('C16:E16');
    const c16 = ws.getCell('C16');
    c16.value     = 'STAN LICZNIKA PRZEBIEGU POJAZDU\nna dzień rozpoczęcia prowadzenia ewidencji';
    c16.font      = font('Arial', 10, { color: '000000' });
    c16.alignment = align('center', 'middle', { wrap: true });
    for (const col of ['C', 'D', 'E']) {
        ws.getCell(`${col}16`).border = thickBdr(
            undefined,
            col === 'E' ? 'medium' : undefined,
            'thin',
            undefined
        );
    }
    ws.getRow(16).height = 30;

    const b16 = ws.getCell('B16');
    b16.value     = 'DATA';
    b16.font      = font('Arial', 10, { color: '000000' });
    b16.alignment = align('center', 'middle');
    b16.border    = thickBdr(undefined, 'thin', 'thin', 'medium');

    const b17 = ws.getCell('B17');
    b17.value     = dfParsed ? new Date(Date.UTC(dfParsed.year, dfParsed.month - 1, dfParsed.day)) : null;
    b17.numFmt    = 'DD.MM.YYYY';
    b17.font      = font('Arial', 14, { bold: true, color: '000000' });
    b17.alignment = align('center', 'middle');
    b17.border    = thickBdr(undefined, 'thin', 'medium', 'medium');
    ws.getRow(17).height = 25;

    ws.mergeCells('C17:E17');
    const c17 = ws.getCell('C17');
    const odoFmt = (odometer || 0).toLocaleString('pl-PL', { maximumFractionDigits: 0 }).replace(/\./g, ' ');
    c17.value     = `${Math.floor(odometer || 0).toLocaleString('en').replace(/,/g, ' ')} km`;
    c17.font      = font('Arial', 14, { bold: true, color: '000000' });
    c17.alignment = align('center', 'middle');
    for (const col of ['C', 'D', 'E']) {
        ws.getCell(`${col}17`).border = thickBdr(
            undefined,
            col === 'E' ? 'medium' : undefined,
            'medium',
            undefined
        );
    }

    ws.getRow(18).height = 20;

    // ── ZAKOŃCZENIE EWIDENCJI ────────────────────────────────────────────────
    ws.mergeCells('B19:E19');
    const c19 = ws.getCell('B19');
    c19.value     = 'ZAKOŃCZENIE EWIDENCJI';
    c19.font      = font('Arial', 10, { bold: true, color: '000000' });
    c19.alignment = align('center', 'middle');
    for (const col of ['B', 'C', 'D', 'E']) {
        ws.getCell(`${col}19`).border = thickBdr(
            'medium',
            col === 'E' ? 'medium' : undefined,
            'thin',
            col === 'B' ? 'medium' : undefined
        );
    }
    ws.getRow(19).height = 20;

    ws.mergeCells('C20:D20');
    const c20cd = ws.getCell('C20');
    c20cd.value     = 'STAN LICZNIKA PRZEBIEGU POJAZDU\nna dzień zakończenia prowadzenia ewidencji';
    c20cd.font      = font('Arial', 9, { color: '000000' });
    c20cd.alignment = align('center', 'middle', { wrap: true });
    ws.getCell('C20').border = thickBdr(undefined, undefined, 'thin', undefined);
    ws.getCell('D20').border = thickBdr(undefined, 'thin', 'thin', undefined);
    ws.getRow(20).height = 30;

    const e20 = ws.getCell('E20');
    e20.value     = 'LICZBA PRZEJECHANYCH KILOMETRÓW\nna dzień zakończenia prowadzenia ewidencji';
    e20.font      = font('Arial', 9, { color: '000000' });
    e20.alignment = align('center', 'middle', { wrap: true });
    e20.border    = thickBdr(undefined, 'medium', 'thin', 'thin');

    const b20 = ws.getCell('B20');
    b20.value     = 'DATA';
    b20.font      = font('Arial', 10, { color: '000000' });
    b20.alignment = align('center', 'middle');
    b20.border    = thickBdr(undefined, 'thin', 'thin', 'medium');

    const b21 = ws.getCell('B21');
    b21.value     = dtParsed ? new Date(Date.UTC(dtParsed.year, dtParsed.month - 1, dtParsed.day)) : null;
    b21.numFmt    = 'DD.MM.YYYY';
    b21.font      = font('Arial', 14, { bold: true, color: '000000' });
    b21.alignment = align('center', 'middle');
    b21.border    = thickBdr(undefined, 'thin', 'medium', 'medium');
    ws.getRow(21).height = 25;

    ws.mergeCells('C21:D21');
    const c21 = ws.getCell('C21');
    c21.value     = `${odoEnd.toLocaleString('en').replace(/,/g, ' ')} km`;
    c21.font      = font('Arial', 14, { bold: true, color: '000000' });
    c21.alignment = align('center', 'middle');
    ws.getCell('C21').border = thickBdr(undefined, undefined, 'medium', undefined);
    ws.getCell('D21').border = thickBdr(undefined, 'thin', 'medium', undefined);

    const e21 = ws.getCell('E21');
    e21.value     = `${Math.round(totalKm)} km`;
    e21.font      = font('Arial', 14, { bold: true, color: '000000' });
    e21.alignment = align('center', 'middle');
    e21.border    = thickBdr(undefined, 'medium', 'medium', 'thin');

    ws.getRow(22).height = 30;

    // Signature
    ws.mergeCells('C24:D24');
    const c24 = ws.getCell('C24');
    c24.value     = (driver || '').trim();
    c24.font      = font('Arial', 11, { color: '000000' });
    c24.alignment = align('center', 'bottom');
    c24.border    = dottedBdr();
    ws.getCell('D24').border = dottedBdr();
    ws.getRow(24).height = 20;

    ws.mergeCells('C25:D25');
    const c25 = ws.getCell('C25');
    c25.value     = 'podpis dysponenta';
    c25.font      = font('Arial', 9, { italic: true, color: '666666' });
    c25.alignment = align('center', 'top');

    // ── Rozlicznie ────────────────────────────────────────────────────────────
    const ws2 = wb.addWorksheet('Rozlicznie');
    ws2.views = [{ state: 'frozen', xSplit: 0, ySplit: 6, showGridLines: false }];

    const rozCols = [
        { key: 'A', width: 5 },
        { key: 'B', width: 5 },
        { key: 'C', width: 13 },
        { key: 'D', width: 24 },
        { key: 'E', width: 42 },
        { key: 'F', width: 12 },
        { key: 'G', width: 24 },
        { key: 'H', width: 20 },
    ];
    for (const { key, width } of rozCols) {
        ws2.getColumn(key).width = width;
    }

    // Header: plate number
    ws2.mergeCells('E4:G4');
    const e4 = ws2.getCell('E4');
    e4.value     = 'Nr rejestracyjny auta: ................................................................';
    e4.font      = font('Arial', 9, { bold: true, color: '000000' });
    e4.alignment = align('right', 'bottom');

    const h4 = ws2.getCell('H4');
    h4.value     = plate || '';
    h4.font      = font('Arial', 9, { bold: true, color: '000000' });
    h4.alignment = align('right', 'bottom');

    // Column headers row 6
    const HDRS = [
        'Nr\nkolejny\nwpisu',
        '',
        'Data wyjazdu',
        'Cel wyjazdu',
        'Opis trasy wyjazdu (skąd–dokąd)',
        'Liczba\nfaktycznie\nprzejechanych\nkilometrów',
        'Imię i nazwisko osoby kierującej pojazdem',
        'Uwagi',
    ];

    for (let ci = 1; ci <= HDRS.length; ci++) {
        const col = colLetter(ci);
        const c = ws2.getCell(`${col}6`);
        c.value     = HDRS[ci - 1];
        c.font      = font('Arial', 8, { color: '000000' });
        c.fill      = fill('D9D9D9');
        c.alignment = align('center', 'middle', { wrap: true });
        c.border    = thickOutline(ci, ci === 1, ci === 8, true, false);
    }
    ws2.getRow(6).height = 60;

    // ── Daily data ────────────────────────────────────────────────────────────
    const year  = dfParsed ? dfParsed.year  : new Date().getUTCFullYear();
    const month = dfParsed ? dfParsed.month : new Date().getUTCMonth() + 1;
    const daysCount = daysInMonth(year, month);
    const ROW0 = 7;

    let runningOdo = Math.floor(odometer || 0);

    // Map refuel dates to their working days
    const mappedRefuelSet = new Set();
    for (const rd of (refuelSet || new Set())) {
        mappedRefuelSet.add(getPreviousWorkingDay(rd));
    }

    let lastKnownCity = '';
    const holidays = getPolishHolidays(year);

    for (let day = 1; day <= daysCount; day++) {
        const isoDay = toISO(year, month, day);
        const row    = ROW0 + day - 1;
        const wd     = weekdayOf(isoDay);
        const isWeekend = wd >= 5;
        const isFree    = isWeekend || holidays.has(isoDay);

        const data  = agg.get(isoDay);
        const km    = (data && data.km > 0) ? Math.round(data.km * 100) / 100 : null;

        let route = '';
        if (data && km) {
            if (data.dailyAddresses && data.dailyAddresses.length > 1) {
                const routeParts = [];
                for (const dailyAddrs of data.dailyAddresses) {
                    const [rt, newCity] = await buildRoute(dailyAddrs, lastKnownCity);
                    lastKnownCity = newCity;
                    if (rt) routeParts.push(rt);
                }
                route = routeParts.join(' ;  ');
            } else {
                const [rt, newCity] = await buildRoute(data.addresses || [], lastKnownCity);
                lastKnownCity = newCity;
                route = rt;
            }
        }

        const cel   = (km ? (tripPurpose || 'dowóz/odbiór pracowników') : '');
        const uwagi = mappedRefuelSet.has(isoDay) ? 'tankowanie' : '';

        if (km) {
            runningOdo += km;
        }

        const dateVal = new Date(Date.UTC(year, month - 1, day));
        const rowData = [
            day,
            PL_DAYS[wd],
            dateVal,
            cel,
            route,
            km !== null ? km : 0,
            '',
            uwagi,
        ];

        for (let ci = 1; ci <= rowData.length; ci++) {
            const col = colLetter(ci);
            const c   = ws2.getCell(`${col}${row}`);
            let val = rowData[ci - 1];

            // Show 0 explicitly if not free day and no km
            if (ci === 6 && !km && !isFree) val = 0;

            c.value  = (val !== null && val !== undefined) ? val : '';
            c.fill   = isFree ? fill('A6A6A6') : fill('FFFFFF');
            c.border = thickOutline(ci, ci === 1, ci === 8, false, false);

            const fntCol = '000000';

            if (ci === 1) {
                c.font      = font('Arial', 9, { color: fntCol });
                c.alignment = align('center', 'middle');
            } else if (ci === 2) {
                c.font      = font('Arial', 9, { color: fntCol });
                c.alignment = align('center', 'middle');
            } else if (ci === 3) {
                c.font      = font('Arial', 9, { color: fntCol });
                if (val) c.numFmt = 'DD.MM.YYYY';
                c.alignment = align('center', 'middle');
            } else if (ci === 4) {
                c.font      = font('Arial', 9, { color: fntCol });
                c.alignment = align('center', 'middle');
            } else if (ci === 5) {
                c.font      = font('Arial', 9, { color: '000000' });
                c.alignment = align('center', 'middle', { wrap: true });
            } else if (ci === 6) {
                c.font = font('Arial', 9, { color: '000000' });
                if (val === 0) {
                    c.numFmt = '0';
                } else if (val) {
                    c.numFmt = '#,##0';
                }
                c.alignment = align('center', 'middle');
            } else if (ci === 7) {
                if (km) {
                    c.value     = (driver || '').trim();
                    c.font      = font('Arial', 9, { color: '000000' });
                } else {
                    c.font = font('Arial', 9, { color: fntCol });
                }
                c.alignment = align('center', 'middle');
            } else if (ci === 8) {
                if (uwagi) {
                    c.font = font('Arial', 8, { bold: true, color: 'B45309' });
                    c.fill = fill('FFFBEB');
                } else {
                    c.font = font('Arial', 8, { color: fntCol });
                }
                c.alignment = align('center', 'middle', { indent: 1 });
            }
        }
    }

    // Summary row
    const sr = ROW0 + daysCount;
    ws2.getRow(sr).height = 20;
    ws2.mergeCells(`A${sr}:E${sr}`);
    const asr = ws2.getCell(`A${sr}`);
    asr.value     = 'Koniec okresu rozliczeniowego';
    asr.font      = font('Arial', 9, { color: '000000' });
    asr.fill      = fill('D9D9D9');
    asr.alignment = align('center', 'middle');
    for (let ci = 1; ci <= 5; ci++) {
        const col = colLetter(ci);
        ws2.getCell(`${col}${sr}`).border = thickOutline(ci, ci === 1, false, false, true);
    }

    const cKm = ws2.getCell(`F${sr}`);
    cKm.value     = totalKm;
    cKm.numFmt    = '#,##0';
    cKm.font      = font('Arial', 10, { bold: true, color: '000000' });
    cKm.fill      = fill('FFFFFF');
    cKm.alignment = align('center', 'middle');
    cKm.border    = thickOutline(6, false, false, false, true);

    ws2.mergeCells(`G${sr}:H${sr}`);
    ws2.getCell(`G${sr}`).fill = fill('D9D9D9');
    for (let ci = 7; ci <= 8; ci++) {
        const col = colLetter(ci);
        ws2.getCell(`${col}${sr}`).border = thickOutline(ci, false, ci === 8, false, true);
    }

    // Signature row
    const sig = sr + 2;
    ws2.mergeCells(`G${sig}:H${sig}`);
    const gSig = ws2.getCell(`G${sig}`);
    gSig.value     = (driver || '').trim();
    gSig.font      = font('Arial', 11, { color: '000000' });
    gSig.alignment = align('center', 'bottom');
    ws2.getCell(`H${sig}`).font = font('Arial', 11, { color: '000000' });

    ws2.mergeCells(`G${sig + 1}:H${sig + 1}`);
    const gSig1 = ws2.getCell(`G${sig + 1}`);
    gSig1.value     = 'podpis dysponenta';
    gSig1.font      = font('Arial', 8, { italic: true, color: '6B7280' });
    gSig1.alignment = align('center', 'top');

    // Auto-adjust column widths
    for (let colIdx = 1; colIdx <= 8; colIdx++) {
        const colLet = colLetter(colIdx);
        let maxLength = 0;
        const wsCol = ws2.getColumn(colIdx);
        wsCol.eachCell({ includeEmpty: false }, (cell) => {
            if (cell.row < 5 || cell.row >= sr) return;
            if (!cell.value) return;
            const lines = String(cell.value).split('\n');
            for (const line of lines) {
                if (line.length > maxLength) maxLength = line.length;
            }
        });
        if (maxLength > 0) {
            let adjWidth;
            if (colLet === 'A' || colLet === 'B') {
                adjWidth = Math.max(5, Math.min(maxLength + 2, 8));
            } else if (colLet === 'E') {
                adjWidth = Math.min(maxLength + 2, 48);
            } else {
                adjWidth = Math.min(maxLength + 2, 25);
            }
            ws2.getColumn(colIdx).width = adjWidth;
        }
    }

    // Page setup
    ws2.pageSetup.orientation = 'landscape';
    ws2.pageSetup.paperSize   = 9; // A4
    ws2.pageSetup.fitToPage   = true;
    ws2.pageSetup.fitToHeight = 0;
    ws2.pageSetup.fitToWidth  = 1;
    ws2.pageSetup.horizontalCentered = true;
    ws2.pageSetup.margins = {
        left: 0.25,
        right: 0.25,
        top: 0.75,
        bottom: 0.75,
        header: 0.3,
        footer: 0.3,
    };

    const buffer = await wb.xlsx.writeBuffer();
    return buffer;
}

// Helper re-export for server.js usage
function toISO(year, month, day) {
    return `${year}-${String(month).padStart(2, '0')}-${String(day).padStart(2, '0')}`;
}

module.exports = { generateExcel };
