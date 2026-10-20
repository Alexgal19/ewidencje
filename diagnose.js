'use strict';

const fs   = require('fs');
const path = require('path');
const { xlsReadRows } = require('./src/xlsParser');
const ExcelJS = require('exceljs');

// Helper: extract plain value from an ExcelJS cell value
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
        if (Array.isArray(r.richText)) return r.richText.map(rt => rt.text || '').join('');
        return String(r);
    }
    // Hyperlink / shared string with text property
    if (v.text !== undefined) return v.text;
    return String(v);
}

const files = fs.readdirSync('.').filter(f => /\.(xls|xlsx)$/i.test(f));
if (!files.length) {
    console.error('No .xls or .xlsx file found in current directory.');
    process.exit(1);
}

const file = files[0];
console.log(`\n=== Diagnosing: ${file} ===\n`);
const buf = fs.readFileSync(file);
const ext = path.extname(file).toLowerCase();

async function run() {
    let rows;
    if (ext === '.xls') {
        rows = xlsReadRows(buf, [5,6,7,8,9,10,11,12,13,14]);
    } else {
        const wb = new ExcelJS.Workbook();
        await wb.xlsx.load(buf);
        const ws = wb.worksheets[0];
        rows = [];
        ws.eachRow({ includeEmpty: true }, (row) => {
            rows.push(row.values.slice(1).map(extractCellValue));
        });
    }

    console.log(`Total rows: ${rows.length}`);
    console.log(`Total cols in row 1: ${rows[0] ? rows[0].length : 'N/A'}`);

    // Print first 35 rows with ALL values
    console.log('\n=== FIRST 35 ROWS (ALL COLUMNS) ===');
    for (let i = 0; i < Math.min(35, rows.length); i++) {
        const r = rows[i];
        if (!r) { console.log(`[${i}] null`); continue; }
        const formatted = r.map((v, idx) => {
            if (v instanceof Date) return `[${idx}]DATE:${v.toISOString().slice(0,10)}`;
            if (v == null) return null;
            const s = String(v).slice(0, 40);
            if (s.trim()) return `[${idx}]"${s}"`;
            return null;
        }).filter(Boolean).join('  ');
        console.log(`Row[${i}]: ${formatted || '(empty)'}`);
    }

    // Find "Razem" or summary rows
    console.log('\n=== ROWS MATCHING "Razem/Suma/Total/Razem" ===');
    let found = 0;
    for (let i = 0; i < rows.length; i++) {
        const r = rows[i];
        if (!r) continue;
        for (let c = 0; c < Math.min(r.length, 5); c++) {
            const v = r[c] != null ? String(r[c]).trim() : '';
            if (/razem|suma|total|together|podsumowanie/i.test(v)) {
                const rowStr = r.map((v,idx) => v != null ? `[${idx}]"${String(v).slice(0,25)}"` : null).filter(Boolean).join('  ');
                console.log(`  Row[${i}]: ${rowStr}`);
                found++;
                break;
            }
        }
    }
    if (!found) console.log('  (none found – this format uses PER-ROW trips, no summary rows)');

    // Find rows with Date objects and print full content
    console.log('\n=== ROWS WITH Date OBJECTS (first 10) ===');
    let dateRows = 0;
    for (let i = 0; i < rows.length && dateRows < 10; i++) {
        const r = rows[i];
        if (!r) continue;
        const dateIdx = r.findIndex(v => v instanceof Date);
        if (dateIdx !== -1) {
            const rowStr = r.map((v, idx) => {
                if (v == null) return null;
                const s = v instanceof Date ? v.toISOString().slice(0, 16) : String(v).slice(0, 40);
                if (!s.trim()) return null;
                return `[${idx}]"${s}"`;
            }).filter(Boolean).join('  ');
            console.log(`  Row[${i}]: ${rowStr}`);
            dateRows++;
        }
    }
    if (!dateRows) console.log('  (no Date objects found!)');
}

run().catch(e => {
    console.error('ERROR:', e.message);
    process.exit(1);
});
