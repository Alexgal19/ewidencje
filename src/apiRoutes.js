'use strict';

/**
 * Supabase-backed REST API – auth, vehicles, drivers, history.
 * All routes require `Authorization: Bearer <access_token>`.
 */

const express = require('express');
const {
    isSupabaseConfigured,
    requireAuth,
    listVehicles,
    createVehicle,
    deleteVehicle,
    listDrivers,
    createDriver,
    deleteDriver,
    listEwidencje,
    deleteEwidencja,
} = require('./supabase');

const router = express.Router();

/** Ensure Supabase is configured before any API call. */
router.use((req, res, next) => {
    if (!isSupabaseConfigured()) {
        return res.status(503).json({ error: 'Supabase nie jest skonfigurowany na serwerze (brak zmiennych środowiskowych).' });
    }
    next();
});

// ── Config (public – anon key for client-side auth) ─────────────────────────

router.get('/config', (req, res) => {
    res.json({
        supabaseUrl: process.env.SUPABASE_URL,
        supabaseAnonKey: process.env.SUPABASE_ANON_KEY,
    });
});

// ── Auth ─────────────────────────────────────────────────────────────────────

router.get('/me', requireAuth, (req, res) => {
    res.json({
        id: req.user.id,
        email: req.user.email,
        createdAt: req.user.created_at,
    });
});

// ── Vehicles ─────────────────────────────────────────────────────────────────

router.get('/vehicles', requireAuth, async (req, res) => {
    try {
        res.json(await listVehicles(req.user.id));
    } catch (err) {
        res.status(500).json({ error: err.message });
    }
});

router.post('/vehicles', requireAuth, async (req, res) => {
    try {
        const plate = String(req.body.plate || '').trim();
        const make  = String(req.body.make  || '').trim();
        const model = String(req.body.model || '').trim();
        if (!plate) {
            return res.status(400).json({ error: 'Numer rejestracyjny jest wymagany' });
        }
        res.json(await createVehicle(req.user.id, { plate, make, model }));
    } catch (err) {
        res.status(500).json({ error: err.message });
    }
});

router.delete('/vehicles/:id', requireAuth, async (req, res) => {
    try {
        await deleteVehicle(req.user.id, req.params.id);
        res.json({ ok: true });
    } catch (err) {
        res.status(500).json({ error: err.message });
    }
});

// ── Drivers ──────────────────────────────────────────────────────────────────

router.get('/drivers', requireAuth, async (req, res) => {
    try {
        res.json(await listDrivers(req.user.id));
    } catch (err) {
        res.status(500).json({ error: err.message });
    }
});

router.post('/drivers', requireAuth, async (req, res) => {
    try {
        const name = String(req.body.name || '').trim();
        if (!name) {
            return res.status(400).json({ error: 'Imię i nazwisko jest wymagane' });
        }
        res.json(await createDriver(req.user.id, { name }));
    } catch (err) {
        res.status(500).json({ error: err.message });
    }
});

router.delete('/drivers/:id', requireAuth, async (req, res) => {
    try {
        await deleteDriver(req.user.id, req.params.id);
        res.json({ ok: true });
    } catch (err) {
        res.status(500).json({ error: err.message });
    }
});

// ── History (ewidencje) ──────────────────────────────────────────────────────

router.get('/ewidencje', requireAuth, async (req, res) => {
    try {
        res.json(await listEwidencje(req.user.id));
    } catch (err) {
        res.status(500).json({ error: err.message });
    }
});

router.delete('/ewidencje/:id', requireAuth, async (req, res) => {
    try {
        await deleteEwidencja(req.user.id, req.params.id);
        res.json({ ok: true });
    } catch (err) {
        res.status(500).json({ error: err.message });
    }
});

module.exports = router;
