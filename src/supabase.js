'use strict';

/**
 * Supabase integration – server side.
 * Uses the service role key (admin) – never expose this module to the client.
 */

const { createClient } = require('@supabase/supabase-js');

const BUCKET = 'ewidencje';

let _admin = null;

function getSupabaseAdmin() {
    const url = process.env.SUPABASE_URL;
    const key = process.env.SUPABASE_SERVICE_ROLE_KEY;
    if (!url || !key) {
        throw new Error('Brak konfiguracji Supabase. Ustaw SUPABASE_URL i SUPABASE_SERVICE_ROLE_KEY w pliku .env');
    }
    if (!_admin) {
        _admin = createClient(url, key, { auth: { persistSession: false, autoRefreshToken: false } });
    }
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
    const { data, error } = await getSupabaseAdmin()
        .from('ewidencje')
        .select('odometer_end, driver_name, period_end')
        .eq('user_id', userId)
        .eq('vehicle_id', vehicleId)
        .order('created_at', { ascending: false })
        .limit(1)
        .maybeSingle();
    if (error) throw new Error(error.message);
    return data || null;
}

// ── Drivers ──────────────────────────────────────────────────────────────────

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
                .createSignedUrl(row.file_path, 3600);
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
    const filePath = `${userId}/${safeName}`;

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
            file_name: safeName,
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

module.exports = {
    BUCKET,
    getSupabaseAdmin,
    isSupabaseConfigured,
    requireAuth,
    listVehicles,
    createVehicle,
    deleteVehicle,
    getVehicle,
    getLastEwidencjaForVehicle,
    listDrivers,
    createDriver,
    deleteDriver,
    listEwidencje,
    deleteEwidencja,
    saveEwidencja,
};
