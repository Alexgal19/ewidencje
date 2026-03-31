'use strict';

/**
 * Pure-JS reader for .xls (BIFF8/OLE) files.
 * Port of _xls_read_rows() from app.py.
 */

const MAGIC = Buffer.from([0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1]);
const ENDOFCHAIN = 0xFFFFFFFE;
const FREESECT   = 0xFFFFFFFF;

// XL epoch: 1899-12-30
const XL_EPOCH = Date.UTC(1899, 11, 30); // ms since Unix epoch

function xlDt(v, isDateCol) {
    if (!isDateCol) return v;
    if (typeof v !== 'number' || v <= 1 || v >= 3e5) return v;
    try {
        const days = Math.floor(v);
        const secs = Math.round((v - days) * 86400);
        return new Date(XL_EPOCH + days * 86400000 + secs * 1000);
    } catch (e) {
        return v;
    }
}

/**
 * Read rows from an XLS (BIFF8/OLE) buffer.
 * @param {Buffer} xlsBytes
 * @param {number[]} dateCols - 0-based column indices to treat as dates
 * @returns {Array<Array<any>>}
 */
function xlsReadRows(xlsBytes, dateCols = [8, 11]) {
    const buf = Buffer.isBuffer(xlsBytes) ? xlsBytes : Buffer.from(xlsBytes);
    const dateColSet = new Set(dateCols);

    // Check magic
    if (!buf.slice(0, 8).equals(MAGIC)) {
        throw new Error('Plik nie jest formatem XLS (OLE). Użyj .xls z GPS trackera.');
    }

    const data = buf;

    // ── OLE header ────────────────────────────────────────────────────────────
    const secSize     = Math.pow(2, data.readUInt16LE(30));
    const miniSecSize = Math.pow(2, data.readUInt16LE(32));
    const firstDir    = data.readUInt32LE(48);
    const miniCutoff  = data.readUInt32LE(56);
    const firstMiniFat = data.readUInt32LE(60);

    function getSector(sec) {
        const off = 512 + sec * secSize;
        return data.slice(off, off + secSize);
    }

    // Build FAT
    const fat = [];
    for (let i = 0; i < 109; i++) {
        const sec = data.readUInt32LE(76 + i * 4);
        if (sec === FREESECT || sec === ENDOFCHAIN) break;
        const chunk = getSector(sec);
        for (let j = 0; j + 4 <= chunk.length; j += 4) {
            fat.push(chunk.readUInt32LE(j));
        }
    }

    function follow(start) {
        const chain = [];
        let sec = start;
        while (sec !== ENDOFCHAIN && sec !== FREESECT && sec < fat.length) {
            chain.push(sec);
            sec = fat[sec];
        }
        return chain;
    }

    // Directory
    const dirChunks = follow(firstDir).map(getSector);
    const dirData = Buffer.concat(dirChunks);

    const entries = {};
    let rootEntry = null;
    const numEntries = Math.floor(dirData.length / 128);
    for (let i = 0; i < numEntries; i++) {
        const e    = dirData.slice(i * 128, (i + 1) * 128);
        const nlen = e.readUInt16LE(64);
        let name = '';
        if (nlen >= 2) {
            name = e.slice(0, Math.max(nlen - 2, 0)).toString('utf16le').replace(/\0/g, '');
        }
        const etype = e[66];
        const start = e.readUInt32LE(116);
        const size  = e.readUInt32LE(120);
        entries[i] = { name, type: etype, start, size };
        if (i === 0) rootEntry = entries[i];
    }

    // Find Workbook/Book stream
    const wbEntry = Object.values(entries).find(
        e => ['workbook', 'book'].includes(e.name.toLowerCase()) && e.type === 2
    );
    if (!wbEntry) {
        throw new Error('Nie znaleziono strumienia Workbook w pliku XLS.');
    }

    // Read stream (normal or mini)
    let stream;
    if (wbEntry.size < miniCutoff && rootEntry) {
        const miniDataChunks = follow(rootEntry.start).map(getSector);
        const miniData = Buffer.concat(miniDataChunks);

        const minifat = [];
        for (const s of follow(firstMiniFat)) {
            const chunk = getSector(s);
            for (let j = 0; j + 4 <= chunk.length; j += 4) {
                minifat.push(chunk.readUInt32LE(j));
            }
        }

        const parts = [];
        let sec = wbEntry.start;
        while (sec !== ENDOFCHAIN && sec !== FREESECT && sec < minifat.length) {
            const off = sec * miniSecSize;
            parts.push(miniData.slice(off, off + miniSecSize));
            sec = minifat[sec];
        }
        stream = Buffer.concat(parts).slice(0, wbEntry.size);
    } else {
        const chunks = follow(wbEntry.start).map(getSector);
        stream = Buffer.concat(chunks).slice(0, wbEntry.size);
    }

    // ── BIFF8 record parser ───────────────────────────────────────────────────
    const cells = new Map(); // key: `${r},${c}` → value
    const sst   = [];
    let pos     = 0;
    const sd    = stream;

    function cellKey(r, c) { return r * 65536 + c; }

    while (pos + 4 <= sd.length) {
        const rt  = sd.readUInt16LE(pos);
        const rl  = sd.readUInt16LE(pos + 2);
        const rec = sd.slice(pos + 4, pos + 4 + rl);
        pos += 4 + rl;

        // ── SST ──────────────────────────────────────────────────────────────
        if (rt === 0x00FC) {
            const sstFull = rec;
            if (sstFull.length < 8) continue;
            const count = sstFull.readUInt32LE(4);
            let p2 = 8;
            for (let idx = 0; idx < count; idx++) {
                if (p2 + 3 > sstFull.length) break;
                const cch   = sstFull.readUInt16LE(p2);
                const flags = sstFull[p2 + 2];
                p2 += 3;
                const comp = !(flags & 1);
                let rich = 0;
                if ((flags & 8) && p2 + 2 <= sstFull.length) {
                    rich = sstFull.readUInt16LE(p2);
                    p2 += 2;
                }
                if ((flags & 4) && p2 + 4 <= sstFull.length) {
                    p2 += 4;
                }
                const blen = cch * (comp ? 1 : 2);
                let s = '';
                try {
                    s = sstFull.slice(p2, p2 + blen).toString(comp ? 'latin1' : 'utf16le');
                } catch (e) {
                    s = '';
                }
                p2 += blen + rich * 4;
                sst.push(s);
            }
        }

        // ── LABELSST ─────────────────────────────────────────────────────────
        else if (rt === 0x00FD && rec.length >= 8) {
            const r   = rec.readUInt16LE(0);
            const c   = rec.readUInt16LE(2);
            const idx = rec.readUInt32LE(6);
            cells.set(cellKey(r, c), idx < sst.length ? sst[idx] : '');
        }

        // ── LABEL ─────────────────────────────────────────────────────────────
        else if (rt === 0x0204 && rec.length >= 9) {
            const r     = rec.readUInt16LE(0);
            const c     = rec.readUInt16LE(2);
            const cch   = rec.readUInt16LE(6);
            const flags = rec[8];
            const p2    = 9;
            const isUnicode = !!(flags & 1);
            const blen  = cch * (isUnicode ? 2 : 1);
            let s = '';
            try {
                s = rec.slice(p2, p2 + blen).toString(isUnicode ? 'utf16le' : 'latin1');
            } catch (e) {
                s = '';
            }
            cells.set(cellKey(r, c), s.trim());
        }

        // ── NUMBER ────────────────────────────────────────────────────────────
        else if (rt === 0x0203 && rec.length >= 14) {
            const r   = rec.readUInt16LE(0);
            const c   = rec.readUInt16LE(2);
            const val = rec.readDoubleLE(6);
            cells.set(cellKey(r, c), xlDt(val, dateColSet.has(c)));
        }

        // ── RK ───────────────────────────────────────────────────────────────
        else if (rt === 0x027E && rec.length >= 10) {
            const r  = rec.readUInt16LE(0);
            const c  = rec.readUInt16LE(2);
            const rk = rec.readUInt32LE(6);
            let val;
            if (rk & 2) {
                val = (rk >> 2) / ((rk & 1) ? 100.0 : 1.0);
            } else {
                const tmp = Buffer.alloc(8);
                tmp.writeUInt32LE(0, 0);
                tmp.writeUInt32LE((rk & 0xFFFFFFFC) >>> 0, 4);
                val = tmp.readDoubleLE(0);
                if (rk & 1) val /= 100.0;
            }
            cells.set(cellKey(r, c), xlDt(val, dateColSet.has(c)));
        }

        // ── MULRK ─────────────────────────────────────────────────────────────
        else if (rt === 0x00BD && rec.length >= 6) {
            const r    = rec.readUInt16LE(0);
            const fcol = rec.readUInt16LE(2);
            const n    = Math.floor((rec.length - 2) / 6);
            for (let i = 0; i < n; i++) {
                const c  = fcol + i;
                const rk = rec.readUInt32LE(4 + i * 6 + 2);
                let val;
                if (rk & 2) {
                    val = (rk >> 2) / ((rk & 1) ? 100.0 : 1.0);
                } else {
                    const tmp = Buffer.alloc(8);
                    tmp.writeUInt32LE(0, 0);
                    tmp.writeUInt32LE((rk & 0xFFFFFFFC) >>> 0, 4);
                    val = tmp.readDoubleLE(0);
                    if (rk & 1) val /= 100.0;
                }
                cells.set(cellKey(r, c), xlDt(val, dateColSet.has(c)));
            }
        }
    }

    if (cells.size === 0) return [];

    let maxR = 0;
    let maxC = 0;
    for (const key of cells.keys()) {
        const r = Math.floor(key / 65536);
        const c = key % 65536;
        if (r > maxR) maxR = r;
        if (c > maxC) maxC = c;
    }

    const result = [];
    for (let r = 0; r <= maxR; r++) {
        const row = [];
        for (let c = 0; c <= maxC; c++) {
            const k = cellKey(r, c);
            row.push(cells.has(k) ? cells.get(k) : undefined);
        }
        result.push(row);
    }
    return result;
}

module.exports = { xlsReadRows };
