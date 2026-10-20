'use strict';

const { describe, it } = require('node:test');
const assert = require('node:assert');
const {
    excelSerialToIsoDate,
    parseAnyDateToIso,
    cleanPlate,
    isFuelTransaction,
    parseFuelFile,
} = require('../src/fuelParser');

describe('fuelParser', () => {
    it('converts Excel serial dates accurately', () => {
        assert.strictEqual(excelSerialToIsoDate(46266.21), '2026-09-01');
        assert.strictEqual(excelSerialToIsoDate(46280.29), '2026-09-15');
        assert.strictEqual(excelSerialToIsoDate(46295.18), '2026-09-30');
    });

    it('parses various date formats to ISO', () => {
        assert.strictEqual(parseAnyDateToIso('2026-09-15'), '2026-09-15');
        assert.strictEqual(parseAnyDateToIso('15.09.2026'), '2026-09-15');
        assert.strictEqual(parseAnyDateToIso('15-09-2026'), '2026-09-15');
        assert.strictEqual(parseAnyDateToIso('46280.29495'), '2026-09-15');
    });

    it('identifies fuel correctly and rejects non-fuel items', () => {
        // True fuels
        assert.strictEqual(isFuelTransaction('EFECTA', 'L', '01'), true);
        assert.strictEqual(isFuelTransaction('VERVA ON', 'L', '02'), true);
        assert.strictEqual(isFuelTransaction('EFECTA 95', 'L', '03'), true);
        assert.strictEqual(isFuelTransaction('VERVA 98', 'L', '05'), true);
        assert.strictEqual(isFuelTransaction('GAZ LPG', 'L', '06'), true);

        // Non-fuels
        assert.strictEqual(isFuelTransaction('karta', 'szt', '80'), false);
        assert.strictEqual(isFuelTransaction('MYJNIA', 'JD', '20'), false);
        assert.strictEqual(isFuelTransaction('ODKURZACZ', 'JD', '20'), false);
        assert.strictEqual(isFuelTransaction('KAWA Z', 'SZT', '55'), false);
        assert.strictEqual(isFuelTransaction('HERBATA', 'SZT', '55'), false);
        assert.strictEqual(isFuelTransaction('A4', 'szt', '36'), false);
        assert.strictEqual(isFuelTransaction('PRZEJAZD', 'szt', '36'), false);
        assert.strictEqual(isFuelTransaction('ADBLUE', 'L', '81'), false);
        assert.strictEqual(isFuelTransaction('NOXY', 'SZT', '81'), false);
    });

    it('parses CSV input with multiple vehicles and transactions', async () => {
        const csv = `Flota,Numer karty,Rejestracja,Data transakcji,Stacja,Miejscowość,EAN,Nazwa produktu,Pobrana ilość,JM,Cena,Wartość transakcji,Grupa produktu,Kategoria pojazdu,Licznik KM,MPK,Dane kosztowe,Dane dodatkowe,Imię,Nazwisko,VAT,Numer faktury,Dane dodatkowe 1,Dane dodatkowe 2,ID Transakcji
19488,78971511948801185,DSW0260A,46280.2949537037,4469.0,Łosice,5906041020166,EFECTA    ,55.22,L,8.84,488.1448,01,O,0.0,XXX,,,,,23   ,0491897505,0,XXX,634910857.0
19488  ,78971511948801185,DSW0260A,46295.1880940162,9001.0,Płock,2000000003221,karta     ,1.0,szt,1.23,1.23,80,O,0.0,,,,,,23   ,0491946400,,,632483385.0
19488,78971511948801185,DSW0260A,46293.22201388889,4349.0,Twardogóra,5906041020166,EFECTA    ,70.25,L,9.14,642.085,01,O,0.0,XXX,,,,,23   ,0491946400,0,XXX,637126820.0
19488,78971511948801904,DWR4756V,46276.848240740735,4533.0,Świdnica,2000000004129,ADBLUE    ,9.77,L,4.99,48.7523,81,O,0.0,XXX,,,,,23   ,0491897505,0,XXX,634543261.0
19488,78971511948801706,WE7AN73,46282.42612268518,100.0,Świdnica,5906041020067,GAZ LPG   ,36.82,L,2.92,107.5144,06,O,0.0,XXX,,,,,23   ,0491946400,0,XXX,635343030.0`;

        const datesByPlate = await parseFuelFile(csv, 'tankowanie.csv');
        assert.ok(datesByPlate.has('DSW0260A'));
        assert.ok(datesByPlate.has('WE7AN73'));
        assert.ok(!datesByPlate.has('DWR4756V')); // only ADBLUE, so ignored

        const dsw0260aDates = Array.from(datesByPlate.get('DSW0260A'));
        assert.strictEqual(dsw0260aDates.length, 2); // 46280 and 46293, karta on 46295 ignored
        assert.ok(dsw0260aDates.includes('2026-09-15'));
        assert.ok(dsw0260aDates.includes('2026-09-28'));
    });
});
