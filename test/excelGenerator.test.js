'use strict';

const { describe, it } = require('node:test');
const assert = require('node:assert');
const ExcelJS = require('exceljs');

const { generateExcel } = require('../src/excelGenerator');
const { xlsReadRows } = require('../src/xlsParser');

async function load(buf) {
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(buf);
    return wb;
}

/** Find the row on a sheet whose column A equals the given day number (rows start at 7). */
function dayRow(ws, day) { return ws.getRow(6 + day); }

describe('generateExcel', () => {
    // 2025-12-05 is a Friday; weekend km 06-07 are shifted onto it by aggregate().
    const agg = new Map([
        ['2025-12-05', { km: 60, addresses: ['Startowa 1, 00-001 Warszawa', 'Koncowa 2, 00-001 Warszawa'],
                         dailyAddresses: [], firstLoc: null, lastLoc: null }],
    ]);

    it('creates "Tytułowa" and 2nd sheet with correct totals and plate', async () => {
        const buf = await generateExcel('DSW1234E', 'Opel Astra', '2025-12-01', '2025-12-31',
            'Jan Kowalski', '', 10000, new Set(), agg, 'dowóz');
        const wb = await load(buf);
        assert.strictEqual(wb.worksheets.length, 2);
        assert.strictEqual(wb.worksheets[0].name, 'Tytułowa');
        assert.match(String(wb.worksheets[0].getCell('B2').value), /DSW1234E/);

        const ws = wb.worksheets[1];
        assert.strictEqual(dayRow(ws, 5).getCell(6).value, 60);
        assert.strictEqual(dayRow(ws, 5).getCell(7).value, 'Jan Kowalski');
        assert.strictEqual(dayRow(ws, 5).getCell(4).value, 'dowóz');
        // summary row sits right after day 31
        assert.strictEqual(ws.getRow(7 + 31).getCell(6).value, 60);
    });

    it('marks refuel day with "tankowanie" and leaves other Uwagi empty', async () => {
        const buf = await generateExcel('DSW1234E', 'Opel', '2025-12-01', '2025-12-31',
            'Jan', '', 10000, new Set(['2025-12-05']), agg);
        const ws = (await load(buf)).worksheets[1];
        assert.strictEqual(dayRow(ws, 5).getCell(8).value, 'tankowanie');
        assert.ok(!dayRow(ws, 4).getCell(8).value);
    });

    it('shifts a refuel on a weekend to the previous working day', async () => {
        const buf = await generateExcel('DSW1234E', 'Opel', '2025-12-01', '2025-12-31',
            'Jan', '', 10000, new Set(['2025-12-06']), agg);
        const ws = (await load(buf)).worksheets[1];
        assert.strictEqual(dayRow(ws, 5).getCell(8).value, 'tankowanie');
        assert.ok(!dayRow(ws, 6).getCell(8).value);
    });

    it('renders all days of the month (leap-year February)', async () => {
        const buf = await generateExcel('X', 'Y', '2024-02-01', '2024-02-29',
            'Jan', '', 0, new Set(), new Map());
        const ws = (await load(buf)).worksheets[1];
        assert.strictEqual(dayRow(ws, 29).getCell(1).value, 29);
        assert.strictEqual(ws.getRow(7 + 29).getCell(1).value, 'Koniec okresu rozliczeniowego');
    });

    it('sets compact column widths and adapts row height for long routes', async () => {
        const longRouteAgg = new Map([
            ['2025-12-05', {
                km: 150,
                addresses: [
                    'ul. Bystrzycka 7a, 58-100 Świdnica',
                    'ul. Wrocławska 10, 58-300 Wałbrzych',
                    'ul. Główna 1, 58-200 Dzierżoniów',
                    'ul. Legnicka 5, 59-220 Legnica',
                    'ul. Strzegomska 12, 58-100 Świdnica',
                    'ul. Bystrzycka 7a, 58-100 Świdnica',
                ],
                dailyAddresses: [],
                firstLoc: null,
                lastLoc: null,
            }],
        ]);

        const buf = await generateExcel(
            'DSW1234E', 'Opel Astra', '2025-12-01', '2025-12-31',
            'Khyzmyi Mykola', 'Kowalski Jan', 10000, new Set(),
            longRouteAgg, 'dowóz/odbiór pracowników'
        );
        const wb = await load(buf);
        const ws1 = wb.worksheets[0];
        const ws2 = wb.worksheets[1];

        // Sheet 1 pageSetup check
        // Sheet 1 & Sheet 2 pageSetup check (both must be A4 Landscape for printing)
        assert.strictEqual(ws1.pageSetup.orientation, 'landscape');
        assert.strictEqual(ws1.pageSetup.paperSize, 9);
        assert.strictEqual(ws2.pageSetup.orientation, 'landscape');
        assert.strictEqual(ws2.pageSetup.paperSize, 9);

        // Column widths should NOT be bloated by merged cells (A, B shouldn't be ~31)
        assert.ok(ws2.getColumn(1).width <= 7, 'Col A width should be <= 7, got ' + ws2.getColumn(1).width);
        assert.ok(ws2.getColumn(2).width <= 7, 'Col B width should be <= 7, got ' + ws2.getColumn(2).width);
        assert.ok(ws2.getColumn(4).width <= 28, 'Col D width should be <= 28, got ' + ws2.getColumn(4).width);
        assert.ok(ws2.getColumn(7).width <= 25, 'Col G width should be <= 25, got ' + ws2.getColumn(7).width);

        // Long route row height should expand and NOT be capped at 75 pt
        const row5 = dayRow(ws2, 5);
        assert.ok(row5.height >= 22, 'Row 5 height should be >= 22, got ' + row5.height);
    });

    it('marks GPS anomaly in Uwagi with warning and red highlight', async () => {
        const anomalyAgg = new Map([
            ['2026-09-04', {
                km: 0.05,
                addresses: [
                    '32, ulica Harcówka, Żory, Poland',
                    '5, ulica Rzemieślnicza, Strumień, Poland',
                ],
                dailyAddresses: [],
                firstLoc: null,
                lastLoc: null,
                isAnomaly: true,
                anomaly: {
                    date: '2026-09-04',
                    reportedKm: 0.05,
                    reason: 'Przejazd między miastami: Żory ➔ Strumień'
                }
            }],
        ]);

        const buf = await generateExcel(
            'DSW3410A', 'Renault Trafic', '2026-09-01', '2026-09-30',
            'Jan Kowalski', 'Smart Work', 10000, new Set(),
            anomalyAgg, 'dowóz/odbiór pracowników'
        );
        const wb = await load(buf);
        const ws2 = wb.worksheets[1];
        const day4Cell = dayRow(ws2, 4).getCell(8);
        assert.match(String(day4Cell.value), /⚠️ Błąd GPS \(0\.05 km\)/);
        assert.strictEqual(day4Cell.font.color.argb, 'FFDC2626');
        assert.strictEqual(day4Cell.fill.fgColor.argb, 'FFFEE2E2');
    });
});

describe('xlsReadRows', () => {
    it('rejects data that is not an OLE/XLS file', () => {
        assert.throws(() => xlsReadRows(Buffer.from('definitely not xls')), /XLS/);
    });
    it('rejects an empty buffer', () => {
        assert.throws(() => xlsReadRows(Buffer.alloc(0)));
    });
});
