'use strict';

const { createClient } = require('@supabase/supabase-js');

function readEnv(name, fallbacks = []) {
    const seen = [name, ...fallbacks];
    for (const key of seen) {
        const value = process.env[key];
        if (value && String(value).trim()) {
            return String(value).trim();
        }
    }
    return '';
}

function getSupabaseClient(options = {}) {
    const useServiceRole = options.useServiceRole !== false;
    const url = readEnv('SUPABASE_URL', ['NEXT_PUBLIC_SUPABASE_URL']);
    const anonKey = readEnv('SUPABASE_ANON_KEY', ['NEXT_PUBLIC_SUPABASE_ANON_KEY']);
    const serviceKey = readEnv('SUPABASE_SERVICE_ROLE_KEY');

    const key = useServiceRole ? (serviceKey || anonKey) : anonKey;

    if (!url || !key) {
        return null;
    }

    return createClient(url, key, {
        auth: {
            persistSession: false,
            autoRefreshToken: false,
        },
    });
}

async function saveGeneratedReport(record) {
    const client = getSupabaseClient();
    if (!client) {
        return { ok: false, reason: 'missing-config' };
    }

    const tableName = readEnv('SUPABASE_REPORTS_TABLE', ['REPORTS_TABLE']) || 'reports';

    const payload = {
        driver_name: record.driver_name || '',
        plate: record.plate || '',
        car_model: record.car_model || '',
        date_from: record.date_from || null,
        date_to: record.date_to || null,
        odometer_start: Number(record.odometer_start || 0),
        odometer_end: Number(record.odometer_end || 0),
        total_km: Number(record.total_km || 0),
        trip_purpose: record.trip_purpose || '',
        file_name: record.file_name || '',
        status: record.status || 'generated',
        created_at: new Date().toISOString(),
    };

    const { data, error } = await client
        .from(tableName)
        .insert([payload])
        .select()
        .single();

    if (error) {
        throw error;
    }

    return { ok: true, data };
}

module.exports = {
    getSupabaseClient,
    saveGeneratedReport,
};
