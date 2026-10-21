'use strict';

const express = require('express');
const archiver = require('archiver');
const path = require('path');
const { parseGps, aggregate, aggregateActual } = require('./gpsParser');
const { generateExcel } = require('./excelGenerator');
const { parseFuelFile, cleanPlate, parseAnyDateToIso } = require('./fuelParser');
const {
    isSupabaseConfigured,
    requireAuth,
    findVehicleByPlate,
    getLastEwidencjaForPlate,
    listDrivers,
    saveEwidencja,
} = require('./supabase');

const router = express.Router();

const BULK_ADMIN_EMAIL = 'o.holiadynets@swl.com.pl';

function createZipArchive(options = {}) {
    if (typeof archiver === 'function') {
        return archiver('zip', options);
    }
    if (archiver && archiver.ZipArchive) {
        return new archiver.ZipArchive(options);
    }
    if (archiver && archiver.default) {
        if (typeof archiver.default === 'function') {
            return archiver.default('zip', options);
        }
        if (archiver.default.ZipArchive) {
            return new archiver.default.ZipArchive(options);
        }
    }
    throw new Error('Nie znaleziono obsługiwanego archiwizatora ZIP w bibliotece archiver');
}

const MONTH_NAMES_PL = [
    'styczeń', 'luty', 'marzec', 'kwiecień', 'maj', 'czerwiec',
    'lipiec', 'sierpień', 'wrzesień', 'październik', 'listopad', 'grudzień'
];

/** "09. ewidencja wrzesień 2026 DSW86212.xlsx" */
function buildOutName(plate, year, month) {
    const mm = String(month).padStart(2, '0');
    const plateClean = cleanPlate(plate || 'auto');
    const mName = MONTH_NAMES_PL[month - 1] || 'miesiac';
    return `${mm}. ewidencja ${mName} ${year} ${plateClean}.xlsx`;
}

/**
 * Access control middleware: only o.holiadynets@swl.com.pl is authorized for bulk operations.
 */
function requireBulkAdmin(req, res, next) {
    if (!isSupabaseConfigured()) {
        return res.status(503).json({
            error: 'Baza danych Supabase nie jest skonfigurowana na serwerze (brak kluczy w .env).'
        });
    }
    const email = (req.user && req.user.email ? req.user.email : '').toLowerCase().trim();
    if (email !== BULK_ADMIN_EMAIL) {
        return res.status(403).json({
            error: `Brak uprawnień. Moduł masowego generowania jest dostępny wyłącznie dla: ${BULK_ADMIN_EMAIL}`
        });
    }
    next();
}

/**
 * POST /api/bulk/analyze
 * Analyzes uploaded GPS files and the fuel transactions file.
 * Automatically looks up previous odometer readings and drivers from Supabase history.
 */
async function handleBulkAnalyze(req, res) {
    try {
        if (!req.files || !req.files.gps_files || req.files.gps_files.length === 0) {
            return res.status(400).json({ error: 'Nie przesłano żadnych plików GPS.' });
        }

        // 1. Parse fuel transactions if provided
        let datesByPlate = new Map();
        if (req.files.fuel_file && req.files.fuel_file[0]) {
            const f = req.files.fuel_file[0];
            datesByPlate = await parseFuelFile(f.buffer, f.originalname);
        } else if (req.body.fuel_text && req.body.fuel_text.trim()) {
            datesByPlate = await parseFuelFile(req.body.fuel_text.trim(), 'fuel.csv');
        }

        const items = [];
        const userId = req.user.id;

        for (let i = 0; i < req.files.gps_files.length; i++) {
            const file = req.files.gps_files[i];
            try {
                const parsed = await parseGps(file.buffer, file.originalname);
                const { plate: rawPlate, carModel: rawModel, dateFrom, dateTo, dayGroups, anomalies = [] } = parsed;

                const fallbackPlate = path.basename(file.originalname, path.extname(file.originalname));
                const plate = cleanPlate(rawPlate || fallbackPlate);

                const agg = aggregate(dayGroups || []);
                const totalKm = Math.round([...agg.values()].reduce((s, v) => s + v.km, 0) * 100) / 100;

                const year = dateFrom ? parseInt(dateFrom.slice(0, 4), 10) : new Date().getFullYear();
                const month = dateFrom ? parseInt(dateFrom.slice(5, 7), 10) : new Date().getMonth() + 1;

                // Look up vehicle & previous ewidencja in Supabase
                let veh = null;
                let lastEwid = null;
                if (isSupabaseConfigured() && userId) {
                    try {
                        veh = await findVehicleByPlate(userId, plate);
                        lastEwid = await getLastEwidencjaForPlate(userId, plate);
                    } catch (e) {
                        console.warn('[Bulk] DB lookup warning for plate', plate, e.message);
                    }
                }

                const carModel = rawModel || (veh && (veh.make || veh.model) ? [veh.make, veh.model].filter(Boolean).join(' ') : (lastEwid ? lastEwid.car_model : ''));
                const driverName = (lastEwid && lastEwid.driver_name) ? lastEwid.driver_name : '';
                const odometerStart = (lastEwid && typeof lastEwid.odometer_end === 'number') ? lastEwid.odometer_end : 0;
                const hasPreviousOdo = Boolean(lastEwid && typeof lastEwid.odometer_end === 'number');

                // Match refuel dates within date range
                let refuelDates = [];
                if (datesByPlate.has(plate)) {
                    const allDates = datesByPlate.get(plate);
                    const matched = [];
                    for (const d of allDates) {
                        if (dateFrom && dateTo) {
                            if (d >= dateFrom && d <= dateTo) matched.push(d);
                        } else {
                            const [dy, dm] = d.split('-');
                            if (parseInt(dy, 10) === year && parseInt(dm, 10) === month) matched.push(d);
                        }
                    }
                    refuelDates = matched.sort();
                }

                items.push({
                    fileIndex: i,
                    fileName: file.originalname,
                    plate,
                    carModel: carModel || '',
                    dateFrom: dateFrom || '',
                    dateTo: dateTo || '',
                    year,
                    month,
                    totalKm,
                    driverName,
                    odometerStart,
                    hasPreviousOdo,
                    targetOdometer: '',
                    adjustMileage: false,
                    refuelDates,
                    anomalies,
                    outName: buildOutName(plate, year, month),
                    status: 'ok',
                    error: null,
                });
            } catch (err) {
                const fallbackPlate = cleanPlate(path.basename(file.originalname, path.extname(file.originalname)));
                items.push({
                    fileIndex: i,
                    fileName: file.originalname,
                    plate: fallbackPlate,
                    carModel: '',
                    dateFrom: '',
                    dateTo: '',
                    year: new Date().getFullYear(),
                    month: new Date().getMonth() + 1,
                    totalKm: 0,
                    driverName: '',
                    odometerStart: 0,
                    hasPreviousOdo: false,
                    targetOdometer: '',
                    adjustMileage: false,
                    refuelDates: [],
                    anomalies: [],
                    outName: `${fallbackPlate}.xlsx`,
                    status: 'error',
                    error: err.message || String(err),
                });
            }
        }

        res.json({
            ok: true,
            items,
            fuelPlatesCount: datesByPlate.size,
        });
    } catch (err) {
        console.error('[Bulk Analyze Error]', err);
        res.status(500).json({ error: String(err.message || err) });
    }
}

/**
 * POST /api/bulk/generate
 * Generates all ewidencja Excel files, saves each to Supabase history,
 * and streams a single ZIP archive to the client.
 */
async function handleBulkGenerate(req, res) {
    try {
        if (!req.files || !req.files.gps_files || req.files.gps_files.length === 0) {
            return res.status(400).json({ error: 'Brak plików GPS do wygenerowania.' });
        }

        let itemsConfig = [];
        try {
            itemsConfig = JSON.parse(req.body.items_config || '[]');
        } catch (e) {
            return res.status(400).json({ error: 'Nieprawidłowy format items_config (wymagany poprawny JSON).' });
        }

        const configMap = new Map();
        for (const cfg of itemsConfig) {
            configMap.set(cfg.fileIndex, cfg);
        }

        const defaultPurpose = (req.body.default_purpose || '').trim();
        const defaultDysponent = (req.body.default_dysponent || '').trim();

        // Prepare zip filename based on first valid item's month/year
        let zipMonth = new Date().getMonth() + 1;
        let zipYear = new Date().getFullYear();

        const archive = createZipArchive({ zlib: { level: 9 } });

        const errors = [];
        const allAnomalies = [];
        let generatedCount = 0;

        // Fetch user drivers for driver_id matching (best effort)
        let userDrivers = [];
        try {
            userDrivers = await listDrivers(req.user.id);
        } catch (e) {}

        for (let i = 0; i < req.files.gps_files.length; i++) {
            const file = req.files.gps_files[i];
            const cfg = configMap.get(i) || {};

            try {
                const parsed = await parseGps(file.buffer, file.originalname);
                const { plate: rawPlate, carModel: rawModel, dateFrom, dateTo, dayGroups, anomalies = [] } = parsed;
                if (anomalies.length > 0) {
                    allAnomalies.push(...anomalies);
                }

                if (!dayGroups || !dayGroups.length) {
                    throw new Error('Brak danych o trasach w pliku GPS');
                }

                const fallbackPlate = path.basename(file.originalname, path.extname(file.originalname));
                const plate = cleanPlate(cfg.plate || rawPlate || fallbackPlate);
                const carModel = (cfg.carModel || rawModel || '').trim();
                const driverName = (cfg.driverName || '').trim();
                const dysponentName = (cfg.dysponentName || defaultDysponent).trim();
                const tripPurpose = (cfg.tripPurpose || defaultPurpose).trim();

                const year = dateFrom ? parseInt(dateFrom.slice(0, 4), 10) : zipYear;
                const month = dateFrom ? parseInt(dateFrom.slice(5, 7), 10) : zipMonth;
                zipMonth = month;
                zipYear = year;

                const odometer = parseInt(String(cfg.odometerStart || '0').replace(/[\s,]/g, ''), 10) || 0;
                const targetOdo = parseInt(String(cfg.targetOdometer || '0').replace(/[\s,]/g, ''), 10) || 0;
                const adjustMil = Boolean(cfg.adjustMileage && targetOdo > odometer);

                // Zgodnie z zasadami ewidencji: weekendy i święta są ZAWSZE przesuwane na poprzedni dzień roboczy (piątek)
                let agg = aggregate(dayGroups);

                // Proportional km adjustment
                if (adjustMil && targetOdo > odometer) {
                    const expectedGpsKm = targetOdo - odometer;
                    const currentGpsKm = [...agg.values()].reduce((s, v) => s + v.km, 0);

                    if (currentGpsKm > 0 && Math.abs(expectedGpsKm - currentGpsKm) > 0.01) {
                        const ratio = expectedGpsKm / currentGpsKm;
                        for (const [k, v] of agg) {
                            if (v.km > 0) {
                                v.km = Math.round(v.km * ratio * 100) / 100;
                            }
                        }

                        const newTotal = [...agg.values()].reduce((s, v) => s + v.km, 0);
                        const diff = Math.round((expectedGpsKm - newTotal) * 100) / 100;
                        if (Math.abs(diff) > 0) {
                            let maxDay = null;
                            let maxKm = -Infinity;
                            for (const [k, v] of agg) {
                                if (v.km > 0 && v.km > maxKm) { maxKm = v.km; maxDay = k; }
                            }
                            if (maxDay !== null) {
                                agg.get(maxDay).km = Math.round((agg.get(maxDay).km + diff) * 100) / 100;
                            }
                        }
                    }
                }

                // Parse & normalize refuel dates (handles YYYY-MM-DD as well as DD.MM.YYYY)
                const refuelSet = new Set();
                const rawDates = Array.isArray(cfg.refuelDates) ? cfg.refuelDates : [];
                for (const rd of rawDates) {
                    const iso = parseAnyDateToIso(rd);
                    if (iso) refuelSet.add(iso);
                }

                const xlsBuf = await generateExcel(
                    plate, carModel, dateFrom, dateTo,
                    driverName, dysponentName, odometer, refuelSet, agg,
                    tripPurpose
                );

                const totalKm = Math.round([...agg.values()].reduce((s, v) => s + v.km, 0) * 100) / 100;
                const outName = buildOutName(plate, year, month);

                // Associate vehicle_id & driver_id if they exist in Supabase
                let vehRecord = null;
                try {
                    vehRecord = await findVehicleByPlate(req.user.id, plate);
                } catch (e) {}

                let matchedDriver = null;
                if (driverName && userDrivers.length > 0) {
                    matchedDriver = userDrivers.find(d => d.name.toLowerCase().trim() === driverName.toLowerCase());
                }

                // Save to Supabase history (best effort)
                try {
                    await saveEwidencja({
                        userId: req.user.id,
                        vehicleId: vehRecord ? vehRecord.id : null,
                        driverId: matchedDriver ? matchedDriver.id : null,
                        plate,
                        carModel: carModel || (vehRecord && (vehRecord.make || vehRecord.model) ? [vehRecord.make, vehRecord.model].filter(Boolean).join(' ') : ''),
                        driverName,
                        periodStart: dateFrom || null,
                        periodEnd: dateTo || null,
                        odometerStart: odometer,
                        odometerEnd: Math.floor(odometer) + Math.floor(totalKm),
                        totalKm,
                        fileName: outName,
                        buffer: xlsBuf,
                        photos: []
                    });
                } catch (saveErr) {
                    console.warn('[Bulk] History save warning for', plate, saveErr.message);
                }

                archive.append(xlsBuf, { name: outName });
                generatedCount++;
            } catch (itemErr) {
                console.error(`[Bulk] Error generating item ${i} (${file.originalname}):`, itemErr);
                errors.push({ file: file.originalname, error: itemErr.message || String(itemErr) });
            }
        }

        if (allAnomalies.length > 0) {
            let anomReport = '======================================================================\r\n';
            anomReport += 'RAPORT ROZBIEŻNOŚCI I ANOMALII GPS (WYKRYTE BŁĘDY CARTRACK)\r\n';
            anomReport += '======================================================================\r\n\r\n';
            anomReport += 'Wykryto następujące przypadki, w których pojazd odbył realny przejazd\r\n';
            anomReport += '(przejazd między miastami, długi czas jazdy lub wysoka prędkość),\r\n';
            anomReport += 'ale urządzenie GPS zapisało dystans bliski 0 km (np. 0.05 km):\r\n\r\n';
            for (const a of allAnomalies) {
                anomReport += `• Pojazd: ${a.plate || 'Nieznany'} | Data: ${a.date}\r\n`;
                anomReport += `  Przebieg w GPS: ${a.reportedKm} km\r\n`;
                anomReport += `  Opis anomalii: ${a.reason}\r\n\r\n`;
            }
            anomReport += 'Zalecenie: Skoryguj stany liczników fizycznych na desce rozdzielczej lub zgłoś usterkę do Cartrack.\r\n';
            anomReport += '======================================================================\r\n';
            archive.append(Buffer.from(anomReport, 'utf8'), { name: '_RAPORT_ANOMALII_GPS.txt' });
        }

        if (errors.length > 0) {
            const report = errors.map(e => `Plik: ${e.file}\nBłąd: ${e.error}\n`).join('\n---\n\n');
            archive.append(Buffer.from(report, 'utf8'), { name: '_bledy_niektorych_plikow.txt' });
        }

        const mmStr = String(zipMonth).padStart(2, '0');
        const mName = MONTH_NAMES_PL[zipMonth - 1] || 'miesiac';
        const zipName = `Ewidencje_${mmStr}_${mName}_${zipYear}.zip`;
        const asciiZipName = zipName.normalize('NFD').replace(/[\u0300-\u036f]/g, '').replace(/ł/g, 'l').replace(/[^ -~]/g, '_');

        res.setHeader('Content-Type', 'application/zip');
        res.setHeader('Content-Disposition', `attachment; filename="${asciiZipName}"; filename*=UTF-8''${encodeURIComponent(zipName)}`);
        res.setHeader('Access-Control-Expose-Headers', 'Content-Disposition, X-Generated-Count, X-Error-Count');
        res.setHeader('X-Generated-Count', String(generatedCount));
        res.setHeader('X-Error-Count', String(errors.length));

        archive.pipe(res);
        await archive.finalize();
    } catch (err) {
        console.error('[Bulk Generate Error]', err);
        if (!res.headersSent) {
            res.status(500).json({ error: String(err.message || err) });
        }
    }
}

module.exports = {
    BULK_ADMIN_EMAIL,
    requireBulkAdmin,
    handleBulkAnalyze,
    handleBulkGenerate,
    buildOutName,
    createZipArchive,
};
