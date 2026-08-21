'use strict';

require('dotenv').config();

const express = require('express');
const multer  = require('multer');
const path    = require('path');

const { parseGps, aggregate, aggregateActual } = require('./src/gpsParser');
const { generateExcel } = require('./src/excelGenerator');
const apiRoutes = require('./src/apiRoutes');
const { requireAuth, isSupabaseConfigured, saveEwidencja, getVehicle } = require('./src/supabase');

const app    = express();
const upload = multer({ storage: multer.memoryStorage(), limits: { fileSize: 50 * 1024 * 1024 } });

app.use(express.static(path.join(__dirname, 'public')));
app.use(express.json());

// ── API (Supabase: auth, pojazdy, kierowcy, historia) ────────────────────────
app.use('/api', apiRoutes);

// ── Routes ────────────────────────────────────────────────────────────────────

app.post('/generate',
    upload.fields([
        { name: 'gps_file', maxCount: 1 },
        { name: 'photos', maxCount: 10 }
    ]),
    (req, res, next) => {
        if (!isSupabaseConfigured()) {
            return res.status(503).json({ error: 'Supabase nie jest skonfigurowany na serwerze (brak zmiennych środowiskowych).' });
        }
        next();
    },
    requireAuth,
    async (req, res) => {
        try {
            if (!req.files || !req.files.gps_file || !req.files.gps_file[0]) {
                return res.status(400).json({ error: 'Brak pliku GPS' });
            }

            const buffer   = req.files.gps_file[0].buffer;
            const filename = req.files.gps_file[0].originalname || 'upload.xls';

            const driver       = (req.body.driver_name      || '').trim();
            const dysponent    = (req.body.dysponent_name   || '').trim();
            const odoRaw       = (req.body.odometer_start   || '0').trim();
            const refuelRaw    = (req.body.refuel_dates      || '').trim();
            const targetOdoRaw = (req.body.target_odometer  || '0').trim();
            const adjustMil    = req.body.adjust_mileage === 'true';
            const useActual    = req.body.use_actual_days   === 'true';
            const tripPurpose  = (req.body.trip_purpose     || '').trim();
            const vehicleId    = req.body.vehicle_id || null;
            const driverId     = req.body.driver_id  || null;

            const odometer       = parseInt(odoRaw.replace(/[\s,]/g, ''), 10) || 0;
            const targetOdometer = parseInt(targetOdoRaw.replace(/[\s,]/g, ''), 10) || 0;

            let { plate, carModel, dateFrom, dateTo, dayGroups } = await parseGps(buffer, filename);

            if (!dayGroups || !dayGroups.length) {
                return res.status(400).json({ error: 'Nie znaleziono danych GPS w pliku.' });
            }

            // Fall back to the DB vehicle when the GPS file has no plate/model
            let veh = null;
            if (vehicleId) {
                veh = await getVehicle(req.user.id, vehicleId);
                if (veh && !plate) {
                    plate = veh.plate;
                }
            }

            let agg = useActual ? aggregateActual(dayGroups) : aggregate(dayGroups);

            // Proportional km adjustment
            if (adjustMil && targetOdometer > odometer) {
                const expectedGpsKm = targetOdometer - odometer;
                const currentGpsKm  = [...agg.values()].reduce((s, v) => s + v.km, 0);

                if (currentGpsKm > 0 && Math.abs(expectedGpsKm - currentGpsKm) > 0.01) {
                    const ratio = expectedGpsKm / currentGpsKm;
                    for (const [k, v] of agg) {
                        if (v.km > 0) {
                            v.km = Math.round(v.km * ratio * 100) / 100;
                        }
                    }

                    // Fix rounding diff on the day with the most km
                    const newTotal = [...agg.values()].reduce((s, v) => s + v.km, 0);
                    const diff = Math.round((expectedGpsKm - newTotal) * 100) / 100;
                    if (Math.abs(diff) > 0) {
                        let maxDay = null;
                        let maxKm  = -Infinity;
                        for (const [k, v] of agg) {
                            if (v.km > 0 && v.km > maxKm) { maxKm = v.km; maxDay = k; }
                        }
                        if (maxDay !== null) {
                            agg.get(maxDay).km = Math.round((agg.get(maxDay).km + diff) * 100) / 100;
                        }
                    }
                }
            }

            // Parse refuel dates
            const year  = dateFrom ? parseInt(dateFrom.slice(0, 4), 10) : new Date().getFullYear();
            const month = dateFrom ? parseInt(dateFrom.slice(5, 7), 10) : new Date().getMonth() + 1;
            const refuelSet = new Set();

            for (const part of refuelRaw.split(/[,;\s]+/)) {
                const p = part.trim();
                if (!p) continue;
                const m = /^(\d{1,2})[.\-\/](\d{1,2})(?:[.\-\/](\d{2,4}))?/.exec(p);
                if (m) {
                    try {
                        const d  = parseInt(m[1], 10);
                        const mo = parseInt(m[2], 10);
                        let   y  = m[3] ? parseInt(m[3], 10) : year;
                        if (y < 100) y += 2000;
                        const padM = String(mo).padStart(2, '0');
                        const padD = String(d).padStart(2, '0');
                        refuelSet.add(`${y}-${padM}-${padD}`);
                    } catch (e) {
                        // ignore invalid dates
                    }
                }
            }

            const xlsBuf = await generateExcel(
                plate, carModel, dateFrom, dateTo,
                driver, dysponent, odometer, refuelSet, agg,
                tripPurpose
            );

            // Save history record + upload file to Supabase Storage (best effort)
            const totalKm = Math.round([...agg.values()].reduce((s, v) => s + v.km, 0) * 100) / 100;
            const outName = `ewidencja_${(plate || 'auto').replace(/[^a-zA-Z0-9]/g, '')}_${year}_${String(month).padStart(2, '0')}.xlsx`;

            const fallbackModel = veh && veh.make ? [veh.make, veh.model].filter(Boolean).join(' ') : (veh && veh.model) || '';

            await saveEwidencja({
                userId: req.user.id,
                vehicleId,
                driverId,
                plate,
                carModel: carModel || fallbackModel,
                driverName: driver,
                periodStart: dateFrom || null,
                periodEnd: dateTo || null,
                odometerStart: odometer,
                odometerEnd: Math.floor(odometer) + Math.floor(totalKm),
                totalKm,
                fileName: outName,
                buffer: xlsBuf,
                photos: req.files.photos || []
            });

            res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
            res.setHeader('Content-Disposition', `attachment; filename="${outName}"`);
            res.setHeader('Access-Control-Expose-Headers', 'Content-Disposition');
            res.send(xlsBuf);

        } catch (err) {
            console.error(err);
            res.status(500).json({ error: String(err.message || err), detail: err.stack || '' });
        }
    });
    app.post('/api/edit-excel', upload.single('file'), async (req, res) => {
        try {
            if (!req.file || !req.file.buffer) {
                return res.status(400).json({ error: 'Brak pliku' });
            }
            let edits = [];
            if (req.body.edits) {
                try { edits = JSON.parse(req.body.edits); } catch(e) {}
            }
            const ExcelJS = require('exceljs');
            const wb = new ExcelJS.Workbook();
            await wb.xlsx.load(req.file.buffer);

            for (const edit of edits) {
                const ws = wb.getWorksheet(edit.sheet);
                if (ws) {
                    // ExcelJS uses 1-based indexing
                    const cell = ws.getCell(edit.r + 1, edit.c + 1);
                    cell.value = edit.val;
                }
            }

            const outBuf = await wb.xlsx.writeBuffer();
            res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
            res.setHeader('Content-Disposition', `attachment; filename="edited.xlsx"`);
            res.send(outBuf);
        } catch (err) {
            console.error(err);
            res.status(500).json({ error: String(err.message || err) });
        }
    });

// ── Start ─────────────────────────────────────────────────────────────────────

const PORT = parseInt(process.env.PORT, 10) || 8080;
app.listen(PORT, '0.0.0.0', () => {
    console.log(`\n\x1b[32m✔ Server running!\x1b[0m`);
    console.log(`\x1b[36m➜ Local:\x1b[0m    http://localhost:${PORT}`);
    console.log(`\x1b[36m➜ Network:\x1b[0m  http://0.0.0.0:${PORT} (Docker)`);
    console.log(isSupabaseConfigured()
        ? `\x1b[32m✔ Supabase:\x1b[0m   connected`
        : `\x1b[33m⚠ Supabase:\x1b[0m   NOT configured – set SUPABASE_URL / ANON / SERVICE_ROLE keys in .env`);
});
