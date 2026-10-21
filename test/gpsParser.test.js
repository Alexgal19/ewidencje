'use strict';

const { describe, it } = require('node:test');
const assert = require('node:assert');
const ExcelJS = require('exceljs');

const {
    toISO,
    addDays,
    weekdayOf,
    daysInMonth,
    getPolishHolidays,
    getPreviousWorkingDay,
    extractCellValue,
    removeDigitsFromName,
    parseGpsAddr,
    shortenAddr,
    parseGps,
    aggregate,
    aggregateActual,
    buildRoute,
    MIN_DAILY_KM,
    rollOverSmallTrips,
} = require('../src/gpsParser');

describe('gpsParser', () => {
    describe('toISO', () => {
        it('formats 2025-12-05 correctly', () => {
            assert.strictEqual(toISO(2025, 12, 5), '2025-12-05');
        });
        it('zero-pads single digit month and day', () => {
            assert.strictEqual(toISO(2024, 1, 9), '2024-01-09');
        });
    });

    describe('addDays', () => {
        it('adds 1 day', () => {
            assert.strictEqual(addDays('2025-12-31', 1), '2026-01-01');
        });
        it('subtracts days', () => {
            assert.strictEqual(addDays('2025-01-01', -1), '2024-12-31');
        });
    });

    describe('weekdayOf', () => {
        it('2025-12-01 is Monday (0)', () => {
            assert.strictEqual(weekdayOf('2025-12-01'), 0);
        });
        it('2025-12-05 is Friday (4)', () => {
            assert.strictEqual(weekdayOf('2025-12-05'), 4);
        });
        it('2025-12-07 is Sunday (6)', () => {
            assert.strictEqual(weekdayOf('2025-12-07'), 6);
        });
    });

    describe('daysInMonth', () => {
        it('December has 31 days', () => {
            assert.strictEqual(daysInMonth(2025, 12), 31);
        });
        it('February 2024 (leap) has 29 days', () => {
            assert.strictEqual(daysInMonth(2024, 2), 29);
        });
        it('February 2025 has 28 days', () => {
            assert.strictEqual(daysInMonth(2025, 2), 28);
        });
    });

    describe('getPolishHolidays', () => {
        it('includes fixed holidays', () => {
            const h = getPolishHolidays(2025);
            assert.ok(h.has('2025-01-01'));
            assert.ok(h.has('2025-01-06'));
            assert.ok(h.has('2025-05-01'));
            assert.ok(h.has('2025-05-03'));
            assert.ok(h.has('2025-08-15'));
            assert.ok(h.has('2025-11-01'));
            assert.ok(h.has('2025-11-11'));
            assert.ok(h.has('2025-12-24'));
            assert.ok(h.has('2025-12-25'));
            assert.ok(h.has('2025-12-26'));
        });
        it('includes Easter Monday 2025', () => {
            const h = getPolishHolidays(2025);
            // Wielkanoc 2025: 20 kwietnia, Poniedzialek Wielkanocny: 21 kwietnia
            assert.ok(h.has('2025-04-21'));
        });
        it('includes Corpus Christi 2025', () => {
            const h = getPolishHolidays(2025);
            // Boze Cialo 2025: 19 czerwca (60 dni po Wielkanocy)
            assert.ok(h.has('2025-06-19'));
        });
    });

    describe('getPreviousWorkingDay', () => {
        it('Saturday goes to previous Friday', () => {
            assert.strictEqual(getPreviousWorkingDay('2025-12-06'), '2025-12-05');
        });
        it('Sunday goes to previous Friday', () => {
            assert.strictEqual(getPreviousWorkingDay('2025-12-07'), '2025-12-05');
        });
        it('Monday stays Monday', () => {
            assert.strictEqual(getPreviousWorkingDay('2025-12-01'), '2025-12-01');
        });
        it('holiday at month start goes to first working day forward (same month)', () => {
            // 2025-01-01 = holiday; going back crosses to Dec → goes forward to Jan 2
            assert.strictEqual(getPreviousWorkingDay('2025-01-01'), '2025-01-02');
        });
        it('weekend at month start goes forward within same month', () => {
            // 2025-03-01 = Saturday; going back crosses to Feb → goes forward to Mar 3 (Monday)
            assert.strictEqual(getPreviousWorkingDay('2025-03-01'), '2025-03-03');
        });
    });

    describe('extractCellValue', () => {
        it('returns null for null/undefined', () => {
            assert.strictEqual(extractCellValue(null), null);
            assert.strictEqual(extractCellValue(undefined), null);
        });
        it('returns Date as-is', () => {
            const d = new Date('2025-12-01');
            assert.strictEqual(extractCellValue(d), d);
        });
        it('extracts rich text', () => {
            const v = { richText: [{ text: 'Hello ' }, { text: 'world' }] };
            assert.strictEqual(extractCellValue(v), 'Hello world');
        });
        it('extracts formula result', () => {
            assert.strictEqual(extractCellValue({ result: 42 }), 42);
            assert.strictEqual(extractCellValue({ result: { richText: [{ text: 'X' }] } }), 'X');
        });
        it('extracts hyperlink text', () => {
            assert.strictEqual(extractCellValue({ text: 'Link' }), 'Link');
        });
        it('returns plain strings/numbers', () => {
            assert.strictEqual(extractCellValue('foo'), 'foo');
            assert.strictEqual(extractCellValue(7), 7);
        });
    });

    describe('removeDigitsFromName', () => {
        it('removes zip codes and digits', () => {
            assert.strictEqual(removeDigitsFromName('58-100 Swidnica 123'), 'Swidnica');
        });
        it('removes street prefixes', () => {
            assert.strictEqual(removeDigitsFromName('ul. Warszawska'), 'Warszawska');
        });
        it('removes administrative divisions', () => {
            assert.strictEqual(removeDigitsFromName('Bystrzycka 7a, Powiat swidnicki'), 'Bystrzycka');
        });
        it('removes house number with letter suffix (7a, 12b)', () => {
            assert.strictEqual(removeDigitsFromName('Różana 7a'), 'Różana');
            assert.strictEqual(removeDigitsFromName('Główna 12b'), 'Główna');
        });
        it('handles empty input', () => {
            assert.strictEqual(removeDigitsFromName(''), '');
            assert.strictEqual(removeDigitsFromName(null), '');
        });
    });

    describe('parseGpsAddr', () => {
        it('bare house number + ulica → correct street and city', async () => {
            const { streetClean, cityName } = await parseGpsAddr(
                '1, ulica Różana, Drogomyśl, Strumień, 43-424, Powiat Cieszyński, Poland'
            );
            assert.strictEqual(streetClean, 'Różana');
            assert.strictEqual(cityName, 'Drogomyśl');
        });
        it('house number with suffix (1a) + ulica → correct street and city', async () => {
            const { streetClean, cityName } = await parseGpsAddr(
                '1a, ulica Różana, Drogomyśl, 43-424'
            );
            assert.strictEqual(streetClean, 'Różana');
            assert.strictEqual(cityName, 'Drogomyśl');
        });
        it('standard "ul. Street N, City" format', async () => {
            const { streetClean, cityName } = await parseGpsAddr('ul. Bystrzycka 7a, Świdnica');
            assert.strictEqual(streetClean, 'Bystrzycka');
            assert.strictEqual(cityName, 'Świdnica');
        });
        it('street-like part in rest does not become city', async () => {
            // "ulica X" in rest must not be assigned as cityName
            const { cityName } = await parseGpsAddr(
                '1, ulica Różana, Drogomyśl, Powiat Cieszyński'
            );
            assert.strictEqual(cityName, 'Drogomyśl');
        });
    });

    describe('aggregate', () => {
        it('shifts weekend km to Friday', () => {
            const dayGroups = [
                { date: '2025-12-05', km: 10, addresses: ['A'], firstStart: 'A', lastEnd: 'B' }, // piatek
                { date: '2025-12-06', km: 20, addresses: ['C'], firstStart: 'C', lastEnd: 'D' }, // sobota
                { date: '2025-12-07', km: 30, addresses: ['E'], firstStart: 'E', lastEnd: 'F' }, // niedziela
            ];
            const agg = aggregate(dayGroups);
            assert.strictEqual(agg.get('2025-12-05').km, 60);
            assert.strictEqual(agg.has('2025-12-06'), false);
            assert.strictEqual(agg.has('2025-12-07'), false);
        });
        it('keeps weekdays unchanged', () => {
            const dayGroups = [
                { date: '2025-12-01', km: 5, addresses: ['X'], firstStart: 'X', lastEnd: 'Y' },
                { date: '2025-12-02', km: 8, addresses: ['Z'], firstStart: 'Z', lastEnd: 'W' },
            ];
            const agg = aggregate(dayGroups);
            assert.strictEqual(agg.get('2025-12-01').km, 5);
            assert.strictEqual(agg.get('2025-12-02').km, 8);
        });

        it('rolls over small micro-trips (< 1 km, e.g. 0.05 km) to the next driving day', () => {
            const dayGroups = [
                { date: '2026-09-01', km: 139, addresses: ['A'], firstStart: 'A', lastEnd: 'B' },
                { date: '2026-09-04', km: 0.05, addresses: ['C'], firstStart: 'C', lastEnd: 'D' },
                { date: '2026-09-07', km: 70, addresses: ['E'], firstStart: 'E', lastEnd: 'F' },
            ];
            const agg = aggregate(dayGroups);
            // Day 04 had 0.05 km -> should become 0 km with cleared addresses
            assert.strictEqual(agg.get('2026-09-04').km, 0);
            assert.strictEqual(agg.get('2026-09-04').addresses.length, 0);
            // Day 07 should receive the 0.05 km rollover (70 + 0.05 = 70.05)
            assert.strictEqual(agg.get('2026-09-07').km, 70.05);
            // Total km sum is preserved exactly: 139 + 0 + 70.05 = 209.05
            const total = [...agg.values()].reduce((s, v) => s + v.km, 0);
            assert.strictEqual(Math.round(total * 100) / 100, 209.05);
        });

        it('rolls back to previous driving day if micro-trip is on the last day of month', () => {
            const dayGroups = [
                { date: '2026-09-28', km: 100, addresses: ['A'], firstStart: 'A', lastEnd: 'B' },
                { date: '2026-09-30', km: 0.05, addresses: ['C'], firstStart: 'C', lastEnd: 'D' },
            ];
            const agg = aggregate(dayGroups);
            assert.strictEqual(agg.get('2026-09-30').km, 0);
            assert.strictEqual(agg.get('2026-09-28').km, 100.05);
        });
    });

    describe('aggregateActual', () => {
        it('does not shift days', () => {
            const dayGroups = [
                { date: '2025-12-06', km: 20, addresses: ['C'], firstStart: 'C', lastEnd: 'D' },
            ];
            const agg = aggregateActual(dayGroups);
            assert.strictEqual(agg.get('2025-12-06').km, 20);
        });
    });

    describe('buildRoute', () => {
        it('joins multiple cities with dash', async () => {
            // Addresses must have explicit street prefix so looksLikeStreet fires correctly
            const [route] = await buildRoute([
                'ulica Warszawska 1, Warszawa',
                'ulica Krakowska 5, Krakow'
            ]);
            assert.ok(route.includes('Warszawa'), `expected Warszawa in "${route}"`);
            assert.ok(route.includes('Krakow'),    `expected Krakow in "${route}"`);
            assert.ok(route.includes('-'),          `expected dash in "${route}"`);
        });
        it('returns empty for no addresses', async () => {
            const [route, city] = await buildRoute([]);
            assert.strictEqual(route, '');
            assert.strictEqual(city, '');
        });
    });

    describe('parseGps', () => {
        it('parses new XLSX format with mock workbook', async () => {
            const wb = new ExcelJS.Workbook();
            const ws = wb.addWorksheet('Raport');
            // Header-like rows (metadata detection scans first 30 rows)
            ws.addRow(['', '', '', '', '', '', '', '', '', '', 'Data: 2025-12-01 - 2025-12-05', '', '', '', '', '', '', '', '', '', '', '', '', '']);
            // Data row: col11 (index 10 in 0-based rows) = Date, col13 = startAddr, col18 = endAddr
            // Using ExcelJS 1-based column indexing:
            // col 11 (K) = date, col 13 (M) = start, col 18 (R) = end
            const rowVals = [];
            rowVals[11] = new Date(Date.UTC(2025, 11, 2)); // col K = 11
            rowVals[13] = 'Startowa 1, 00-001 Warszawa';    // col M = 13
            rowVals[18] = 'Koncowa 2, 30-001 Krakow';       // col R = 18
            ws.addRow(rowVals);
            // Razem row: col A = "Razem", col X (24) = km
            const razemVals = [];
            razemVals[1] = 'Razem';
            razemVals[24] = 45.6; // col X = 24
            ws.addRow(razemVals);

            const buf = await wb.xlsx.writeBuffer();
            const result = await parseGps(buf, 'test.xlsx');

            assert.ok(result.dayGroups.length > 0, 'expected at least one day group');
            const day = result.dayGroups[0];
            assert.strictEqual(day.date, '2025-12-02');
            assert.strictEqual(day.km, 45.6);
            assert.ok(day.addresses.length >= 2);
            assert.ok(day.firstStart.includes('Startowa'));
            assert.ok(day.lastEnd.includes('Koncowa'));
        });

        // ── km in "Razem" row: strict numeric only (mirror of _km_value in app.py) ──
        // ExcelJS sparse array index i lands in column i → parser sees it at index i-1.
        async function oldFormatBuffer(kmValue, razemExtra = {}, headerRow = null) {
            const wb = new ExcelJS.Workbook();
            const ws = wb.addWorksheet('Raport');
            if (headerRow) ws.addRow(headerRow);
            const rowVals = [];
            rowVals[9]  = new Date(Date.UTC(2025, 11, 2));   // date @ parser col 8
            rowVals[10] = 'Startowa 1, 00-001 Warszawa';      // start addr @ col 9
            rowVals[13] = 'Koncowa 2, 00-001 Warszawa';       // end addr @ col 12
            ws.addRow(rowVals);
            const razemVals = [];
            razemVals[1]  = 'Razem';
            razemVals[17] = kmValue;                          // km @ col 16
            for (const [k, v] of Object.entries(razemExtra)) razemVals[k] = v;
            ws.addRow(razemVals);
            return wb.xlsx.writeBuffer();
        }

        it('does not read parking time ("54:30:49") from the km cell', async () => {
            const result = await parseGps(await oldFormatBuffer('54:30:49'), 't.xlsx');
            assert.strictEqual(result.dayGroups.length, 1);
            assert.strictEqual(result.dayGroups[0].km, 0);
        });

        it('does not read comma-decimal text ("12,5") from the km cell', async () => {
            const result = await parseGps(await oldFormatBuffer('12,5'), 't.xlsx');
            assert.strictEqual(result.dayGroups[0].km, 0);
        });

        it('reads a numeric km cell', async () => {
            const result = await parseGps(await oldFormatBuffer(12.5), 't.xlsx');
            assert.strictEqual(result.dayGroups[0].km, 12.5);
        });

        it('prefers the "Dystans" header column over the default km column', async () => {
            const hdr = [];
            hdr[16] = 'Dystans'; // header @ parser col 15
            const result = await parseGps(
                await oldFormatBuffer('54:30:49', { 16: 33.3 }, hdr),
                't.xlsx'
            );
            assert.strictEqual(result.dayGroups[0].km, 33.3);
        });
    });
});
