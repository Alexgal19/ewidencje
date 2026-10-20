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
});

describe('xlsReadRows', () => {
    it('rejects data that is not an OLE/XLS file', () => {
        assert.throws(() => xlsReadRows(Buffer.from('definitely not xls')), /XLS/);
    });
    it('rejects an empty buffer', () => {
        assert.throws(() => xlsReadRows(Buffer.alloc(0)));
    });
});
