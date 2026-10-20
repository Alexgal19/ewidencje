const { describe, it } = require('node:test');
const assert = require('node:assert');
const { parseGpsAddr, buildRoute } = require('../src/gpsParser');

describe('district names (osiedle) are not written directly', () => {
    it('osiedle as first segment resolves to the city', async () => {
        const r = await parseGpsAddr('osiedle Powstańców Śląskich, Żory, Poland');
        assert.strictEqual(r.streetClean, '');
        assert.strictEqual(r.cityName, 'Żory');
    });
    it('osiedle after the street is ignored as a city', async () => {
        const r = await parseGpsAddr('Rybnicka 5, osiedle Powstańców Śląskich, Żory');
        assert.strictEqual(r.cityName, 'Żory');
    });
    it('route contains only the city, never the osiedle', async () => {
        const [route] = await buildRoute([
            'Rynek 1, 44-240 Żory',
            'osiedle Powstańców Śląskich, 44-240 Żory',
        ]);
        assert.ok(!/osiedle|Powsta/i.test(route), route);
        assert.match(route, /Żory/);
    });
});