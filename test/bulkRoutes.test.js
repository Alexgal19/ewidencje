'use strict';

const { describe, it, before, after } = require('node:test');
const assert = require('node:assert');
const { requireBulkAdmin, buildOutName, BULK_ADMIN_EMAIL, createZipArchive } = require('../src/bulkRoutes');

describe('bulkRoutes', () => {
    const origUrl = process.env.SUPABASE_URL;
    const origAnon = process.env.SUPABASE_ANON_KEY;
    const origServ = process.env.SUPABASE_SERVICE_ROLE_KEY;

    before(() => {
        process.env.SUPABASE_URL = 'https://mock.supabase.co';
        process.env.SUPABASE_ANON_KEY = 'mock_anon';
        process.env.SUPABASE_SERVICE_ROLE_KEY = 'mock_service_role';
    });

    after(() => {
        process.env.SUPABASE_URL = origUrl;
        process.env.SUPABASE_ANON_KEY = origAnon;
        process.env.SUPABASE_SERVICE_ROLE_KEY = origServ;
    });

    it('builds proper output name for vehicles', () => {
        const name = buildOutName('DSW0260A', 2026, 9);
        assert.strictEqual(name, '09. ewidencja wrzesień 2026 DSW0260A.xlsx');
    });

    it('rejects unauthorized users with 403', () => {
        let status = null;
        let responseJson = null;
        const res = {
            status: (s) => {
                status = s;
                return {
                    json: (obj) => { responseJson = obj; }
                };
            }
        };

        const reqOther = { user: { email: 'inna.osoba@firma.pl' } };
        let nextCalled = false;
        requireBulkAdmin(reqOther, res, () => { nextCalled = true; });

        assert.strictEqual(status, 403);
        assert.strictEqual(nextCalled, false);
        assert.ok(responseJson.error.includes(BULK_ADMIN_EMAIL));
    });

    it('allows o.holiadynets@swl.com.pl (case-insensitive)', () => {
        let status = null;
        const res = {
            status: (s) => ({ json: () => {} })
        };

        const reqAdmin = { user: { email: 'O.Holiadynets@swl.com.pl' } };
        let nextCalled = false;
        requireBulkAdmin(reqAdmin, res, () => { nextCalled = true; });

        assert.strictEqual(status, null);
        assert.strictEqual(nextCalled, true);
    });

    it('returns 503 when Supabase is not configured', () => {
        delete process.env.SUPABASE_URL;
        let status = null;
        let responseJson = null;
        const res = {
            status: (s) => {
                status = s;
                return { json: (o) => { responseJson = o; } };
            }
        };
        const req = { user: { email: BULK_ADMIN_EMAIL } };
        requireBulkAdmin(req, res, () => {});
        assert.strictEqual(status, 503);
        assert.ok(responseJson.error.includes('Supabase'));
        process.env.SUPABASE_URL = 'https://mock.supabase.co';
    });

    it('creates zip archive instance with append and finalize support', async () => {
        const za = createZipArchive({ zlib: { level: 9 } });
        assert.ok(za, 'Archive instance should be defined');
        assert.strictEqual(typeof za.append, 'function', 'Should have append method');
        assert.strictEqual(typeof za.finalize, 'function', 'Should have finalize method');
        assert.strictEqual(typeof za.pipe, 'function', 'Should have pipe method');

        const chunks = [];
        const { Writable } = require('stream');
        const out = new Writable({
            write(chunk, enc, cb) {
                chunks.push(chunk);
                cb();
            }
        });
        za.pipe(out);
        za.append(Buffer.from('test content'), { name: 'test.txt' });
        await za.finalize();
        const total = Buffer.concat(chunks);
        assert.ok(total.length > 0, 'Should output valid zip bytes');
    });
});
