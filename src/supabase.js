'use strict';

/**
 * Supabase integration – server side.
 * Uses the service role key (admin) – never expose this module to the client.
 */

const { createClient } = require('@supabase/supabase-js');

const BUCKET = 'ewidencje';

let _admin = null;

function _setAdminClientForTesting(client) {
    _admin = client;
}

function getSupabaseAdmin() {
    if (_admin) return _admin;
    const url = process.env.SUPABASE_URL;
    const key = process.env.SUPABASE_SERVICE_ROLE_KEY;
    if (!url || !key) {
        throw new Error('Brak konfiguracji Supabase. Ustaw SUPABASE_URL i SUPABASE_SERVICE_ROLE_KEY w pliku .env');
    }
    _admin = createClient(url, key, { auth: { persistSession: false, autoRefreshToken: false } });
    return _admin;
}

/** Return true when Supabase env vars are configured. */
function isSupabaseConfigured() {
    return Boolean(process.env.SUPABASE_URL && process.env.SUPABASE_ANON_KEY && process.env.SUPABASE_SERVICE_ROLE_KEY);
}

/**
 * Express middleware – requires a valid Supabase access token
 * sent as `Authorization: Bearer <token>`.
 * Sets `req.user` on success.
 */
async function requireAuth(req, res, next) {
    try {
        const token = (req.headers.authorization || '').replace(/^Bearer\s+/i, '');
        if (!token) {
            return res.status(401).json({ error: 'Brak tokenu autoryzacji' });
        }
        const { data, error } = await getSupabaseAdmin().auth.getUser(token);
        if (error || !data.user) {
            return res.status(401).json({ error: 'Sesja wygasła – zaloguj się ponownie' });
        }
        req.user = data.user;
        next();
    } catch (err) {
        next(err);
    }
}

// ── Vehicles ─────────────────────────────────────────────────────────────────

async function listVehicles(userId) {
    const { data, error } = await getSupabaseAdmin()
        .from('vehicles')
        .select('id, plate, make, model, created_at')
        .eq('user_id', userId)
        .order('plate');
    if (error) throw new Error(error.message);
    return data || [];
}

async function createVehicle(userId, { plate, make, model }) {
    const { data, error } = await getSupabaseAdmin()
        .from('vehicles')
        .insert({ user_id: userId, plate, make, model })
        .select('id, plate, make, model, created_at')
        .single();
    if (error) throw new Error(error.message);
    return data;
}

async function deleteVehicle(userId, id) {
    const { error } = await getSupabaseAdmin()
        .from('vehicles')
        .delete()
        .eq('id', id)
        .eq('user_id', userId);
    if (error) throw new Error(error.message);
}

async function getVehicle(userId, id) {
    if (!id) return null;
    const { data, error } = await getSupabaseAdmin()
        .from('vehicles')
        .select('id, plate, make, model')
        .eq('id', id)
        .eq('user_id', userId)
        .maybeSingle();
    if (error) throw new Error(error.message);
    return data || null;
}

/** Latest ewidencja row for a vehicle (last odometer + last driver). */
async function getLastEwidencjaForVehicle(userId, vehicleId) {
    if (!vehicleId) return null;
    const sb = getSupabaseAdmin();
    const { data: veh } = await sb
        .from('vehicles')
        .select('id, plate')
        .eq('id', vehicleId)
        .maybeSingle();

    const plate = veh ? veh.plate : null;
    return getLastEwidencjaForPlate(userId, plate, null, vehicleId);
}

// ── Drivers ──────────────────────────────────────────────────────────────────


/** Look up vehicle record by plate (spaces/dashes tolerant). */
async function findVehicleByPlate(userId, plate) {
    if (!plate) return null;
    const clean = String(plate).toUpperCase().replace(/[^A-Z0-9]/g, '');
    if (!clean) return null;

    // Fetch user's vehicles and match normalized plates
    const vehicles = await listVehicles(userId);
    const match = vehicles.find(v => String(v.plate || '').toUpperCase().replace(/[^A-Z0-9]/g, '') === clean);
    if (match) return match;

    // Fallback: search all vehicles in Supabase if fleet manager
    try {
        const { data } = await getSupabaseAdmin()
            .from('vehicles')
            .select('id, plate, make, model, created_at');
        if (data && data.length) {
            const gMatch = data.find(v => String(v.plate || '').toUpperCase().replace(/[^A-Z0-9]/g, '') === clean);
            if (gMatch) return gMatch;
        }
    } catch (e) {
        // ignore fallback errors
    }

    return null;
}

/**
 * Latest ewidencja row matching a registration plate or vehicle_id.
 * If beforeDate is given (e.g. '2026-09-01'), prioritizes the previous month's ewidencja
 * (period_end <= beforeDate or period_start < beforeDate).
 */
async function getLastEwidencjaForPlate(userId, plate, beforeDate = null, vehicleId = null) {
    if (!plate && !vehicleId) return null;
    const clean = String(plate || '').toUpperCase().replace(/[^A-Z0-9]/g, '');

    const sb = getSupabaseAdmin();
    let query = sb
        .from('ewidencje')
        .select('id, plate, odometer_start, odometer_end, driver_name, period_start, period_end, car_model, vehicle_id, driver_id, created_at')
        .order('period_end', { ascending: false, nullsFirst: false })
        .order('created_at', { ascending: false });

    if (userId) {
        query = query.eq('user_id', userId);
    }

    const { data, error } = await query;
    let rows = data || [];

    // Fallback: If not found under current userId, check globally for company fleet
    if (!error && rows.length === 0 && userId) {
        try {
            const { data: globalData } = await sb
                .from('ewidencje')
                .select('id, plate, odometer_start, odometer_end, driver_name, period_start, period_end, car_model, vehicle_id, driver_id, created_at')
                .order('period_end', { ascending: false, nullsFirst: false })
                .order('created_at', { ascending: false });
            if (globalData && globalData.length) {
                rows = globalData;
            }
        } catch (e) {}
    }

    if (!rows.length) return null;

    // Filter records matching plate or vehicle_id
    const matches = rows.filter(row => {
        if (vehicleId && row.vehicle_id && String(row.vehicle_id) === String(vehicleId)) return true;
        if (clean && row.plate) {
            const p = String(row.plate).toUpperCase().replace(/[^A-Z0-9]/g, '');
            if (p === clean) return true;
        }
        return false;
    });

    if (!matches.length) return null;

    // If beforeDate is given (e.g. '2026-09-01'), find prior month's record
    if (beforeDate) {
        // 1. Prior record with valid odometer_end (> 0) ending on or before beforeDate
        const priorEndWithOdo = matches.find(r => r.period_end && r.period_end <= beforeDate && Number(r.odometer_end) > 0);
        if (priorEndWithOdo) return priorEndWithOdo;

        // 2. Any prior record ending on or before beforeDate
        const priorEnd = matches.find(r => r.period_end && r.period_end <= beforeDate);
        if (priorEnd) return priorEnd;

        // 3. Record starting strictly before beforeDate with valid odometer
        const priorStart = matches.find(r => r.period_start && r.period_start < beforeDate && (Number(r.odometer_end) > 0 || Number(r.odometer_start) > 0));
        if (priorStart) return priorStart;

        // 4. Same period already generated (re-analyzing current month) -> return its starting odometer
        const samePeriod = matches.find(r => r.period_start && r.period_start === beforeDate);
        if (samePeriod) {
            return {
                ...samePeriod,
                odometer_end: (typeof samePeriod.odometer_start === 'number' && samePeriod.odometer_start > 0)
                    ? samePeriod.odometer_start
                    : samePeriod.odometer_end
            };
        }
    }

    // Default to the latest match
    return matches[0] || null;
}



async function listDrivers(userId) {
    const { data, error } = await getSupabaseAdmin()
        .from('drivers')
        .select('id, name, created_at')
        .eq('user_id', userId)
        .order('name');
    if (error) throw new Error(error.message);
    return data || [];
}

async function createDriver(userId, { name }) {
    const { data, error } = await getSupabaseAdmin()
        .from('drivers')
        .insert({ user_id: userId, name })
        .select('id, name, created_at')
        .single();
    if (error) throw new Error(error.message);
    return data;
}

async function deleteDriver(userId, id) {
    const { error } = await getSupabaseAdmin()
        .from('drivers')
        .delete()
        .eq('id', id)
        .eq('user_id', userId);
    if (error) throw new Error(error.message);
}

// ── Ewidencje (history) ──────────────────────────────────────────────────────

async function listEwidencje(userId) {
    const sb = getSupabaseAdmin();
    const { data, error } = await sb
        .from('ewidencje')
        .select('*')
        .eq('user_id', userId)
        .order('created_at', { ascending: false });
    if (error) throw new Error(error.message);

    const items = [];
    for (const row of (data || [])) {
        let downloadUrl = null;
        if (row.file_path) {
            const { data: signed } = await sb.storage
                .from(BUCKET)
                .createSignedUrl(row.file_path, 3600, { download: row.file_name || true });
            downloadUrl = signed ? signed.signedUrl : null;
        }
        
        let photoUrls = [];
        if (row.photo_paths && Array.isArray(row.photo_paths) && row.photo_paths.length > 0) {
            const { data: signedPhotos } = await sb.storage
                .from(BUCKET)
                .createSignedUrls(row.photo_paths, 3600);
            if (signedPhotos) {
                photoUrls = signedPhotos.map(p => p.signedUrl).filter(Boolean);
            }
        }
        
        items.push({ ...row, download_url: downloadUrl, photo_urls: photoUrls });
    }
    return items;
}

async function deleteEwidencja(userId, id) {
    const sb = getSupabaseAdmin();
    const { data: row } = await sb
        .from('ewidencje')
        .select('file_path, photo_paths')
        .eq('id', id)
        .eq('user_id', userId)
        .maybeSingle();

    if (row) {
        const filesToRemove = [];
        if (row.file_path) filesToRemove.push(row.file_path);
        if (row.photo_paths && Array.isArray(row.photo_paths)) {
            filesToRemove.push(...row.photo_paths);
        }
        if (filesToRemove.length > 0) {
            await sb.storage.from(BUCKET).remove(filesToRemove);
        }
    }

    const { error } = await sb
        .from('ewidencje')
        .delete()
        .eq('id', id)
        .eq('user_id', userId);
    if (error) throw new Error(error.message);
}

/**
 * Upload the generated XLSX to Storage and insert a history record.
 * Returns the created row or null when not configured.
 */
async function saveEwidencja({
    userId,
    vehicleId,
    driverId,
    plate,
    carModel,
    driverName,
    periodStart,
    periodEnd,
    odometerStart,
    odometerEnd,
    totalKm,
    fileName,
    buffer,
    photos = []
}) {
    if (!isSupabaseConfigured()) return null;

    const sb = getSupabaseAdmin();
    const safeName = fileName
        .replace(/[^a-zA-Z0-9._-]/g, '_')
        .replace(/_+/g, '_')
        .slice(0, 120);
    // Unique prefix: same plate+month generated twice must not share (and overwrite/delete) one file
    const filePath = `${userId}/${Date.now()}_${safeName}`;

    const { error: upErr } = await sb.storage
        .from(BUCKET)
        .upload(filePath, buffer, { contentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', upsert: true });
    if (upErr) {
        console.error('[Supabase] Upload failed:', upErr.message);
        return null;
    }

    const photoPaths = [];
    if (photos && photos.length > 0) {
        for (let i = 0; i < photos.length; i++) {
            const p = photos[i];
            const pName = p.originalname ? p.originalname.replace(/[^a-zA-Z0-9._-]/g, '_') : `photo_${i}.jpg`;
            const pPath = `${userId}/photos/${Date.now()}_${i}_${pName}`;
            const { error: pErr } = await sb.storage.from(BUCKET).upload(pPath, p.buffer, { contentType: p.mimetype || 'image/jpeg', upsert: true });
            if (!pErr) {
                photoPaths.push(pPath);
            }
        }
    }

    const { data, error } = await sb
        .from('ewidencje')
        .insert({
            user_id: userId,
            vehicle_id: vehicleId || null,
            driver_id: driverId || null,
            plate,
            car_model: carModel,
            driver_name: driverName,
            period_start: periodStart || null,
            period_end: periodEnd || null,
            odometer_start: odometerStart,
            odometer_end: odometerEnd,
            total_km: totalKm,
            file_name: fileName,
            file_path: filePath,
            photo_paths: photoPaths.length > 0 ? photoPaths : null
        })
        .select()
        .single();

    if (error) {
        console.error('[Supabase] History insert failed:', error.message);
        await sb.storage.from(BUCKET).remove([filePath, ...photoPaths]);
        return null;
    }
    return data;
}

async function addPhotosToHistoryEwidencja(userId, id, photos) {
    if (!isSupabaseConfigured() || !photos || !photos.length) return null;
    
    const sb = getSupabaseAdmin();
    
    // 1. Get current record
    const { data: row, error: fetchErr } = await sb
        .from('ewidencje')
        .select('photo_paths')
        .eq('id', id)
        .eq('user_id', userId)
        .single();
        
    if (fetchErr || !row) throw new Error('Nie znaleziono ewidencji');
    
    // 2. Upload new photos
    const newPhotoPaths = [];
    for (let i = 0; i < photos.length; i++) {
        const p = photos[i];
        const pName = p.originalname ? p.originalname.replace(/[^a-zA-Z0-9._-]/g, '_') : `photo_${i}.jpg`;
        const pPath = `${userId}/photos/${Date.now()}_${i}_${pName}`;
        const { error: pErr } = await sb.storage.from(BUCKET).upload(pPath, p.buffer, { contentType: p.mimetype || 'image/jpeg', upsert: true });
        if (!pErr) {
            newPhotoPaths.push(pPath);
        }
    }
    
    if (newPhotoPaths.length === 0) throw new Error('Nie udało się wgrać żadnego ze zdjęć');
    
    // 3. Update record
    const updatedPaths = Array.isArray(row.photo_paths) ? [...row.photo_paths, ...newPhotoPaths] : [...newPhotoPaths];
    const { error: updateErr } = await sb
        .from('ewidencje')
        .update({ photo_paths: updatedPaths })
        .eq('id', id)
        .eq('user_id', userId);
        
    if (updateErr) {
        await sb.storage.from(BUCKET).remove(newPhotoPaths); // rollback
        throw new Error('Błąd podczas aktualizacji rekordu w bazie: ' + updateErr.message);
    }
    
    return true;
}

module.exports = {
    BUCKET,
    getSupabaseAdmin,
    _setAdminClientForTesting,
    isSupabaseConfigured,
    requireAuth,
    listVehicles,
    createVehicle,
    deleteVehicle,
    getVehicle,
    getLastEwidencjaForVehicle,
    findVehicleByPlate,
    getLastEwidencjaForPlate,
    listDrivers,
    createDriver,
    deleteDriver,
    listEwidencje,
    deleteEwidencja,
    saveEwidencja,
    addPhotosToHistoryEwidencja,
};
