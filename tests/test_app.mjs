/**
 * Tests for app.js — mirrors the Python test_main.py test catalog.
 * Run with: node --test tests/test_app.mjs
 * Prerequisite: npm install xlsx
 */

import { describe, it, mock, beforeEach } from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';

const __dirname = dirname(fileURLToPath(import.meta.url));

// ─── Load SheetJS (XLSX) as a global, as app.js expects it ───────────────
const XLSX = await import('xlsx');
globalThis.XLSX = XLSX;

// ─── Load app.js functions into scope ────────────────────────────────────
// app.js is written for browsers (no exports). We extract functions via eval.
const appSource = readFileSync(join(__dirname, '..', 'web', 'app.js'), 'utf-8');

// Remove 'use strict' and DOM event listener at the end, then eval to get functions
const cleanedSource = appSource
  .replace(/^'use strict';\s*/, '')
  .replace(/document\.addEventListener\('DOMContentLoaded'[\s\S]*$/, '');

// Provide stubs for browser globals
globalThis.document = {
  getElementById: () => ({
    classList: { add() {}, remove() {} },
    innerHTML: '',
    textContent: '',
    value: '',
    appendChild() {},
  }),
  addEventListener() {},
  createElement: () => ({ className: '', textContent: '' }),
};

const moduleScope = {};
const wrappedSource = `
  ${cleanedSource}
  Object.assign(moduleScope, {
    WIRE_FEE_USD, esc, fmt, parseDateToIso, formatDate, subtractDays,
    readTransactions, getNbpRate, fetchNbpRates, groupBySaleDate,
    renderResults,
    excelSerialToIso, xtbDateToIso, isXtbWorkbook, readXtbReport,
    xtbGroupBySaleDate, xtbRate, xtbYearSummaries, findPrevBusinessRate,
  });
`;
eval(wrappedSource);

const {
  WIRE_FEE_USD, esc, fmt, parseDateToIso, formatDate, subtractDays,
  readTransactions, getNbpRate, fetchNbpRates, groupBySaleDate,
  excelSerialToIso, xtbDateToIso, isXtbWorkbook, readXtbReport,
  xtbGroupBySaleDate, xtbRate, xtbYearSummaries, findPrevBusinessRate,
} = moduleScope;

// ─── Load example XLSX ───────────────────────────────────────────────────
const EXAMPLE_FILE = join(__dirname, 'input', 'example', 'example_2025.xlsx');
let exampleBuffer;
try {
  exampleBuffer = readFileSync(EXAMPLE_FILE);
} catch {
  console.error(
    'Example file not found. Run: cd tests/input && mkdir -p example && python3 generate_example.py'
  );
  process.exit(1);
}

function loadTransactions() {
  return readTransactions(exampleBuffer.buffer.slice(
    exampleBuffer.byteOffset,
    exampleBuffer.byteOffset + exampleBuffer.byteLength
  ));
}


// ═══════════════════════════════════════════════════════════════════════════
// readTransactions
// ═══════════════════════════════════════════════════════════════════════════

describe('readTransactions', () => {
  it('reads only Sell records', () => {
    const txns = loadTransactions();
    assert.ok(txns.every(t => t['Record Type'] === 'Sell'));
  });

  it('skips Summary row', () => {
    const txns = loadTransactions();
    assert.ok(!txns.some(t => t['Record Type'] === 'Summary'));
  });

  it('returns correct transaction count', () => {
    const txns = loadTransactions();
    assert.equal(txns.length, 4); // 2 RSU + 1 ESPP + 1 ESPP
  });

  it('RSU transactions have zero acquisition cost', () => {
    const txns = loadTransactions();
    const rsu = txns.filter(t => t['Plan Type'] === 'RS');
    assert.equal(rsu.length, 2);
    for (const t of rsu) {
      assert.equal(t['Acquisition Cost'], 0);
    }
  });

  it('ESPP transactions have nonzero acquisition cost', () => {
    const txns = loadTransactions();
    const espp = txns.filter(t => t['Plan Type'] === 'ESPP');
    assert.equal(espp.length, 2);
    for (const t of espp) {
      assert.ok(t['Acquisition Cost'] > 0);
    }
  });

  it('all transactions have required fields', () => {
    const required = ['Record Type', 'Symbol', 'Plan Type', 'Quantity',
                      'Date Sold', 'Total Proceeds', 'Proceeds Per Share'];
    const txns = loadTransactions();
    for (const t of txns) {
      for (const field of required) {
        assert.ok(field in t, `Missing field: ${field}`);
      }
    }
  });
});


// ═══════════════════════════════════════════════════════════════════════════
// parseDateToIso
// ═══════════════════════════════════════════════════════════════════════════

describe('parseDateToIso', () => {
  it('converts Date object to ISO string', () => {
    const d = new Date(2025, 5, 9); // June 9, 2025 (month is 0-based)
    assert.equal(parseDateToIso(d), '2025-06-09');
  });

  it('converts MM/DD/YYYY string to ISO', () => {
    assert.equal(parseDateToIso('06/09/2025'), '2025-06-09');
  });

  it('passes through YYYY-MM-DD string', () => {
    assert.equal(parseDateToIso('2025-03-15'), '2025-03-15');
  });

  it('returns null for dashes', () => {
    assert.equal(parseDateToIso('--'), null);
  });

  it('returns null for null/undefined', () => {
    assert.equal(parseDateToIso(null), null);
    assert.equal(parseDateToIso(undefined), null);
  });

  it('throws on unknown format', () => {
    assert.throws(() => parseDateToIso('some-date'), /Nieznany format daty/);
  });
});


// ═══════════════════════════════════════════════════════════════════════════
// formatDate
// ═══════════════════════════════════════════════════════════════════════════

describe('formatDate', () => {
  it('formats Date object as DD.MM.YYYY', () => {
    const d = new Date(2025, 5, 9); // June 9
    assert.equal(formatDate(d), '09.06.2025');
  });

  it('formats US string as DD.MM.YYYY', () => {
    assert.equal(formatDate('06/09/2025'), '09.06.2025');
  });

  it('formats ISO string as DD.MM.YYYY', () => {
    assert.equal(formatDate('2025-06-09'), '09.06.2025');
  });

  it('returns dash for null/blank', () => {
    assert.equal(formatDate(null), '—');
    assert.equal(formatDate('--'), '—');
  });
});


// ═══════════════════════════════════════════════════════════════════════════
// subtractDays
// ═══════════════════════════════════════════════════════════════════════════

describe('subtractDays', () => {
  it('subtracts days correctly', () => {
    assert.equal(subtractDays('2025-04-10', 1), '2025-04-09');
  });

  it('crosses month boundary', () => {
    assert.equal(subtractDays('2025-03-01', 1), '2025-02-28');
  });

  it('crosses year boundary', () => {
    assert.equal(subtractDays('2025-01-01', 1), '2024-12-31');
  });
});


// ═══════════════════════════════════════════════════════════════════════════
// groupBySaleDate
// ═══════════════════════════════════════════════════════════════════════════

describe('groupBySaleDate', () => {
  it('groups by date sold', () => {
    const txns = loadTransactions();
    const groups = groupBySaleDate(txns);
    assert.equal(groups.size, 3);
  });

  it('RSU package has two transactions', () => {
    const txns = loadTransactions();
    const groups = groupBySaleDate(txns);
    const firstKey = [...groups.keys()][0];
    assert.equal(groups.get(firstKey).length, 2);
  });

  it('ESPP packages have one transaction each', () => {
    const txns = loadTransactions();
    const groups = groupBySaleDate(txns);
    const keys = [...groups.keys()];
    assert.equal(groups.get(keys[1]).length, 1); // June ESPP
    assert.equal(groups.get(keys[2]).length, 1); // December ESPP
  });

  it('preserves insertion order (chronological)', () => {
    const txns = loadTransactions();
    const groups = groupBySaleDate(txns);
    const keys = [...groups.keys()];
    const sorted = [...keys].sort();
    assert.deepEqual(keys, sorted);
  });
});


// ═══════════════════════════════════════════════════════════════════════════
// Tax calculation logic (RSU)
// ═══════════════════════════════════════════════════════════════════════════

describe('RSU tax logic', () => {
  it('RSU proceeds sum is correct', () => {
    const txns = loadTransactions();
    const rsu = txns.filter(t => t['Plan Type'] === 'RS');
    const total = rsu.reduce((s, t) => s + t['Total Proceeds'], 0);
    // 50 * 52 + 30 * 52 = 4160
    assert.ok(Math.abs(total - 4160.0) < 0.01);
  });

  it('RSU cost is zero', () => {
    const txns = loadTransactions();
    const rsu = txns.filter(t => t['Plan Type'] === 'RS');
    const totalCost = rsu.reduce((s, t) => s + t['Acquisition Cost'], 0);
    assert.equal(totalCost, 0.0);
  });
});


// ═══════════════════════════════════════════════════════════════════════════
// Tax calculation logic (ESPP)
// ═══════════════════════════════════════════════════════════════════════════

describe('ESPP tax logic', () => {
  it('ESPP proceeds sum is correct', () => {
    const txns = loadTransactions();
    const espp = txns.filter(t => t['Plan Type'] === 'ESPP');
    const total = espp.reduce((s, t) => s + t['Total Proceeds'], 0);
    // 120 * 53 + 80 * 48 = 10200
    assert.ok(Math.abs(total - 10200.0) < 0.01);
  });

  it('ESPP cost sum is correct', () => {
    const txns = loadTransactions();
    const espp = txns.filter(t => t['Plan Type'] === 'ESPP');
    const totalCost = espp.reduce((s, t) => s + t['Acquisition Cost'], 0);
    // 120 * 38.29 + 80 * 35.70 = 7450.80
    assert.ok(Math.abs(totalCost - 7450.80) < 0.01);
  });

  it('ESPP taxable amount is correct', () => {
    const txns = loadTransactions();
    const espp = txns.filter(t => t['Plan Type'] === 'ESPP');
    const proceeds = espp.reduce((s, t) => s + t['Total Proceeds'], 0);
    const cost = espp.reduce((s, t) => s + t['Acquisition Cost'], 0);
    const taxable = proceeds - cost;
    // 10200 - 7450.80 = 2749.20
    assert.ok(Math.abs(taxable - 2749.20) < 0.01);
  });

  it('ESPP transactions have valid purchase date', () => {
    const txns = loadTransactions();
    const espp = txns.filter(t => t['Plan Type'] === 'ESPP');
    for (const t of espp) {
      assert.notEqual(t['Purchase Date'], '--');
      const iso = parseDateToIso(t['Purchase Date']);
      assert.ok(iso !== null);
      assert.match(iso, /^\d{4}-\d{2}-\d{2}$/);
    }
  });
});


// ═══════════════════════════════════════════════════════════════════════════
// NBP rate fetching (mocked)
// ═══════════════════════════════════════════════════════════════════════════

describe('NBP rates', () => {
  beforeEach(() => {
    mock.restoreAll();
  });

  it('getNbpRate returns rate and effective date', async () => {
    const mockResponse = {
      ok: true,
      status: 200,
      json: async () => ({
        rates: [{ mid: 4.05, effectiveDate: '2025-04-09' }],
      }),
    };
    mock.method(globalThis, 'fetch', async () => mockResponse);

    const result = await getNbpRate('2025-04-10');
    assert.equal(result.rate, 4.05);
    assert.equal(result.effectiveDate, '2025-04-09');
  });

  it('getNbpRate retries on 404 (weekend/holiday)', async () => {
    let callCount = 0;
    mock.method(globalThis, 'fetch', async () => {
      callCount++;
      if (callCount <= 2) {
        return { ok: false, status: 404 };
      }
      return {
        ok: true,
        status: 200,
        json: async () => ({
          rates: [{ mid: 3.80, effectiveDate: '2025-04-03' }],
        }),
      };
    });

    const result = await getNbpRate('2025-04-07');
    assert.equal(result.rate, 3.80);
    assert.equal(callCount, 3);
  });

  it('getNbpRate throws after 7 retries', async () => {
    mock.method(globalThis, 'fetch', async () => ({ ok: false, status: 404 }));

    await assert.rejects(
      () => getNbpRate('2025-04-10'),
      /Nie udało się pobrać kursu NBP/
    );
  });

  it('fetchNbpRates deduplicates dates', async () => {
    let callCount = 0;
    mock.method(globalThis, 'fetch', async () => {
      callCount++;
      return {
        ok: true,
        status: 200,
        json: async () => ({
          rates: [{ mid: 4.0, effectiveDate: '2025-04-09' }],
        }),
      };
    });

    const dates = ['2025-04-10', '2025-04-10', '2025-04-10'];
    const rates = await fetchNbpRates(dates);
    assert.equal(callCount, 1);
    assert.ok(rates.has('2025-04-10'));
  });
});


// ═══════════════════════════════════════════════════════════════════════════
// fetchNbpRates — batched range queries (NBP throttles rapid single requests)
// ═══════════════════════════════════════════════════════════════════════════

describe('fetchNbpRates (batched range queries)', () => {
  beforeEach(() => {
    mock.restoreAll();
  });

  const mockRange = (rates) => ({
    ok: true,
    status: 200,
    json: async () => ({ rates }),
  });

  it('fetches many dates with a single request', async () => {
    const mockRates = [
      { effectiveDate: '2025-04-04', mid: 3.8 },
      { effectiveDate: '2025-04-07', mid: 3.9 },
      { effectiveDate: '2025-04-09', mid: 4.0 },
      { effectiveDate: '2025-04-10', mid: 4.05 },
    ];
    let callCount = 0;
    mock.method(globalThis, 'fetch', async () => {
      callCount++;
      return mockRange(mockRates);
    });

    const dates = ['2025-04-11', '2025-04-10', '2025-04-10', '2025-04-07'];
    const rates = await fetchNbpRates(dates);

    assert.equal(callCount, 1);
    assert.deepEqual([...rates.keys()], ['2025-04-07', '2025-04-10', '2025-04-11']);
  });

  it('uses the previous business day, skipping weekends', async () => {
    const mockRates = [
      { effectiveDate: '2025-04-04', mid: 3.8 },
      { effectiveDate: '2025-04-07', mid: 3.9 },
      { effectiveDate: '2025-04-09', mid: 4.0 },
      { effectiveDate: '2025-04-10', mid: 4.05 },
    ];
    mock.method(globalThis, 'fetch', async () => mockRange(mockRates));

    const rates = await fetchNbpRates(['2025-04-07', '2025-04-10', '2025-04-11']);

    // 05–06.04 to weekend → dla 07.04 kursem jest notowanie z 04.04
    assert.equal(rates.get('2025-04-07').rate, 3.8);
    assert.equal(rates.get('2025-04-07').effectiveDate, '2025-04-04');
    // dla 10.04 — notowanie z 09.04 (nie własne)
    assert.equal(rates.get('2025-04-10').rate, 4.0);
    // dla 11.04 — notowanie z 10.04
    assert.equal(rates.get('2025-04-11').rate, 4.05);
    assert.equal(rates.get('2025-04-11').effectiveDate, '2025-04-10');
  });

  it('calls onProgress once per unique date', async () => {
    mock.method(globalThis, 'fetch', async () =>
      mockRange([{ effectiveDate: '2025-04-09', mid: 4.0 }]));

    const seen = [];
    await fetchNbpRates(
      ['2025-04-10', '2025-04-11', '2025-04-10'],
      iso => seen.push(iso),
    );

    assert.deepEqual(seen, ['2025-04-10', '2025-04-11']);
  });

  it('splits into chunks when dates span more than one request', async () => {
    const mockRates = [
      { effectiveDate: '2025-01-02', mid: 3.9 },
      { effectiveDate: '2026-03-03', mid: 4.1 },
    ];
    let callCount = 0;
    mock.method(globalThis, 'fetch', async () => {
      callCount++;
      return mockRange(mockRates);
    });

    const rates = await fetchNbpRates(['2025-01-05', '2026-03-04']);

    assert.equal(callCount, 2);
    assert.equal(rates.get('2025-01-05').rate, 3.9);
    assert.equal(rates.get('2026-03-04').rate, 4.1);
  });

  it('builds a range URL with the lowercase currency code', async () => {
    const urls = [];
    mock.method(globalThis, 'fetch', async url => {
      urls.push(String(url));
      return mockRange([{ effectiveDate: '2025-04-09', mid: 4.0 }]);
    });

    await fetchNbpRates(['2025-04-10'], null, 'USD');

    assert.equal(
      urls[0],
      'https://api.nbp.pl/api/exchangerates/rates/a/usd' +
      '/2025-04-03/2025-04-10/?format=json',
    );
  });

  it('throws the standard error when no quote is available', async () => {
    mock.method(globalThis, 'fetch', async () => mockRange([]));

    await assert.rejects(
      () => fetchNbpRates(['2025-04-10']),
      /Nie udało się pobrać kursu NBP/,
    );
  });
});

describe('findPrevBusinessRate', () => {
  const asc = [
    { effectiveDate: '2025-04-01', mid: 3.7 },
    { effectiveDate: '2025-04-04', mid: 3.8 },
  ];

  it('returns the closest earlier quote', () => {
    assert.equal(findPrevBusinessRate(asc, '2025-04-08').effectiveDate, '2025-04-04');
    assert.equal(findPrevBusinessRate(asc, '2025-04-04').effectiveDate, '2025-04-01');
  });

  it('accepts a quote exactly 7 days back', () => {
    const one = [{ effectiveDate: '2025-04-01', mid: 3.7 }];
    assert.equal(findPrevBusinessRate(one, '2025-04-08').mid, 3.7);
  });

  it('rejects a quote 8+ days back', () => {
    const one = [{ effectiveDate: '2025-04-01', mid: 3.7 }];
    assert.equal(findPrevBusinessRate(one, '2025-04-09'), null);
  });

  it('never returns a quote from the same or later day', () => {
    // 01.04 jest jedynym wcześniejszym notowaniem — dla samej daty 01.04 brak
    assert.equal(findPrevBusinessRate(
      [{ effectiveDate: '2025-04-01', mid: 3.7 }], '2025-04-01'), null);
  });
});


// ═══════════════════════════════════════════════════════════════════════════
// PLN conversion: different rates for revenue vs ESPP cost
// ═══════════════════════════════════════════════════════════════════════════

describe('PLN conversion', () => {
  it('ESPP cost uses purchase-date rate (dates differ)', () => {
    const txns = loadTransactions();
    const groups = groupBySaleDate(txns);
    const keys = [...groups.keys()];

    const esppTxns = groups.get(keys[1]);
    assert.equal(esppTxns[0]['Plan Type'], 'ESPP');

    const saleIso = parseDateToIso(esppTxns[0]['Date Sold']);
    const purchaseIso = parseDateToIso(esppTxns[0]['Purchase Date']);

    assert.notEqual(saleIso, purchaseIso);
    assert.ok(purchaseIso < saleIso);
  });

  it('different rates produce different PLN values', () => {
    const saleRate = 3.75;
    const costRate = 3.70;
    const proceedsUsd = 6360.0;
    const costUsd = 4594.80;

    const proceedsPln = proceedsUsd * saleRate;
    const costPln = costUsd * costRate;

    const taxablePln = proceedsPln - costPln;
    const singleRateTaxable = (proceedsUsd - costUsd) * saleRate;

    assert.ok(Math.abs(taxablePln - singleRateTaxable) > 0.01);
  });
});


// ═══════════════════════════════════════════════════════════════════════════
// Wire fee is informational only
// ═══════════════════════════════════════════════════════════════════════════

describe('Wire fee', () => {
  it('wire fee is not deducted from taxable amount', () => {
    const txns = loadTransactions();
    const rsu = txns.filter(t => t['Plan Type'] === 'RS');
    const proceeds = rsu.reduce((s, t) => s + t['Total Proceeds'], 0);
    const taxable = proceeds; // RSU cost = 0
    assert.equal(taxable, proceeds);
    assert.notEqual(taxable, proceeds - WIRE_FEE_USD);
  });
});


// ═══════════════════════════════════════════════════════════════════════════
// Helpers
// ═══════════════════════════════════════════════════════════════════════════

describe('esc (XSS prevention)', () => {
  it('escapes HTML special characters', () => {
    assert.equal(esc('<script>alert("xss")</script>'),
      '&lt;script&gt;alert(&quot;xss&quot;)&lt;/script&gt;');
  });

  it('handles null/undefined', () => {
    assert.equal(esc(null), '');
    assert.equal(esc(undefined), '');
  });
});

describe('fmt', () => {
  it('formats to 2 decimal places by default', () => {
    assert.equal(fmt(3.14159), '3.14');
  });

  it('formats to custom decimal places', () => {
    assert.equal(fmt(3.14159, 4), '3.1416');
  });
});


// ═══════════════════════════════════════════════════════════════════════════
// XTB report: fixture
// ═══════════════════════════════════════════════════════════════════════════

/** Excel serial (1900 date system) for an ISO date + UTC time. */
const excelSerial = (iso, h = 14, m = 30) =>
  Date.UTC(+iso.slice(0, 4), +iso.slice(5, 7) - 1, +iso.slice(8, 10)) / 86400000 +
  25569 + (h * 60 + m) / 1440;

const XTB_CLOSED_HEADERS = [
  'Instrument', 'Ticker', 'Category', 'Type', 'Volume', 'Open Price',
  'Open Time (UTC)', 'Close Price', 'Close Time (UTC)', 'Product',
  'Profit/Loss', 'Gross Profit', 'Purchase Value', 'Sale Value',
  'Stop Loss', 'Take Profit', 'Commission', 'Margin', 'Swap', 'Rollover',
  'Open Conversion Rate', 'Close Conversion Rate', 'Close Origin',
  'Position ID', 'Comment',
];

function closedRow(o) {
  const row = new Array(XTB_CLOSED_HEADERS.length).fill(null);
  row[0] = o.instrument;
  row[1] = o.ticker;
  row[2] = o.category ?? 'STOCK';
  row[3] = o.type ?? 'BUY';
  row[4] = o.volume;
  row[5] = o.openPrice;
  row[6] = o.openSerial;
  row[7] = o.closePrice;
  row[8] = o.closeSerial;
  row[9] = 'My Trades';
  row[10] = o.profit;
  row[11] = o.profit;
  row[12] = o.purchaseValue;
  row[13] = o.saleValue;
  row[20] = 1;
  row[21] = 1;
  row[22] = 'Web';
  row[23] = o.positionId;
  return row;
}

function buildXtbFixtureBuffer() {
  const closed = [
    ['Account number', '12345678'],
    ['Closed Positions', ''],
    ['Date from (UTC)', 46022.958333333336],
    ['Date to (UTC)', 46346.876124675924],
    XTB_CLOSED_HEADERS,
    // Long position bought in 2025, sold 2026-01-05
    closedRow({
      instrument: 'Nvidia', ticker: 'NVDA.US', volume: 8,
      openPrice: 126.5, openSerial: excelSerial('2025-07-31'),
      closePrice: 191.7, closeSerial: excelSerial('2026-01-05'),
      profit: 521.6, purchaseValue: 1012.0, saleValue: 1533.6,
      positionId: '1632540131',
    }),
    // Long position bought 2025-12-02, sold 2026-01-05 (loss)
    closedRow({
      instrument: 'Amazon', ticker: 'AMZN.US', volume: 5,
      openPrice: 232.53, openSerial: excelSerial('2025-12-02'),
      closePrice: 228.63, closeSerial: excelSerial('2026-01-05', 15, 30),
      profit: -19.5, purchaseValue: 1162.65, saleValue: 1143.15,
      positionId: '1949822424',
    }),
    // Short position: sale leg is the opening trade, purchase leg closes it
    closedRow({
      instrument: 'Short Co', ticker: 'SHORT.US', type: 'SELL', volume: 2,
      openPrice: 50, openSerial: excelSerial('2026-02-10'),
      closePrice: 45, closeSerial: excelSerial('2026-02-12'),
      profit: 10.0, purchaseValue: 90.0, saleValue: 100.0,
      positionId: '2000000001',
    }),
    ['Profit/loss', null, null, null, null, null, null, null, null, null,
      512.1, 512.1],
  ];

  const cash = [
    ['Account number', '12345678'],
    ['Cash Operations', ''],
    ['Date from (UTC)', 46022.958333333336],
    ['Date to (UTC)', 46346.876124675924],
    ['Type', 'Instrument', 'Ticker', 'Category', 'Time', 'Amount', 'ID',
      'Comment', 'Product', 'Position ID'],
    ['Free funds interest', null, null, null, excelSerial('2026-01-14', 16, 53),
      0.04, '1505505859', 'Free-funds Interest 2026-01', 'My Trades', ''],
    ['Free funds interest tax', null, null, null,
      excelSerial('2026-01-14', 16, 46), -0.01, '1505384989',
      'Free-funds Interest Tax 2026-01', 'My Trades', ''],
    ['Stock sell', 'Nvidia', 'NVDA.US', 'STOCK',
      excelSerial('2026-01-05', 14, 30), 1533.6, '1074378267',
      'CLOSE BUY 8/9 @ 191.70', 'My Trades', '1632540131'],
    ['Stock purchase', 'Nvidia', 'NVDA.US', 'STOCK',
      excelSerial('2025-07-31', 14, 30), -1012.0, '1074378268',
      'OPEN BUY 8 @ 126.50', 'My Trades', '1632540131'],
    ['Total', null, null, null, null, -878.37],
  ];

  const openPositions = [
    ['Account number', '12345678'],
    ['Open Positions', ''],
    ['Data as of report generated', 46346.87616181713],
    ['Product', 'Metric', 'Amount', 'Currency'],
    ['My Trades', 'Open position value', 4955.5, 'USD'],
    ['My Trades', 'Open position profit', -303.6, 'USD'],
  ];

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(closed), 'Closed Positions');
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(cash), 'Cash Operations');
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(openPositions), 'Open Positions');

  const buf = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });
  return buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength);
}

const xtbReport = readXtbReport(
  buildXtbFixtureBuffer(), 'USD_12345678_2025-12-31_2026-11-20.xlsx');


// ═══════════════════════════════════════════════════════════════════════════
// XTB: excelSerialToIso / xtbDateToIso
// ═══════════════════════════════════════════════════════════════════════════

describe('excelSerialToIso', () => {
  it('converts a real serial from the XTB report', () => {
    assert.equal(excelSerialToIso(46027.60971064815), '2026-01-05');
    assert.equal(excelSerialToIso(46022.958333333336), '2025-12-31');
    assert.equal(excelSerialToIso(46055.38998842592), '2026-02-02');
  });

  it('drops the time fraction (UTC)', () => {
    assert.equal(excelSerialToIso(46027.0), '2026-01-05');
    assert.equal(excelSerialToIso(46027.99999), '2026-01-05');
  });

  it('throws on non-numeric values', () => {
    assert.throws(() => excelSerialToIso('abc'), /Nieprawidłowa data/);
  });
});

describe('xtbDateToIso', () => {
  it('converts Excel serial numbers', () => {
    assert.equal(xtbDateToIso(46027.60971064815), '2026-01-05');
  });

  it('converts ISO strings with a time part', () => {
    assert.equal(xtbDateToIso('2026-01-05 14:32:10'), '2026-01-05');
    assert.equal(xtbDateToIso('2026-01-05T14:32'), '2026-01-05');
  });

  it('falls back to parseDateToIso for plain strings', () => {
    assert.equal(xtbDateToIso('06/09/2025'), '2025-06-09');
    assert.equal(xtbDateToIso('2025-03-15'), '2025-03-15');
  });

  it('returns null for blank values', () => {
    assert.equal(xtbDateToIso(null), null);
    assert.equal(xtbDateToIso(''), null);
    assert.equal(xtbDateToIso('--'), null);
    assert.equal(xtbDateToIso(undefined), null);
  });
});


// ═══════════════════════════════════════════════════════════════════════════
// XTB: isXtbWorkbook (auto-detection)
// ═══════════════════════════════════════════════════════════════════════════

describe('isXtbWorkbook', () => {
  it('detects an XTB report', () => {
    const wb = XLSX.read(new Uint8Array(buildXtbFixtureBuffer()), { type: 'array' });
    assert.equal(isXtbWorkbook(wb), true);
  });

  it('does not detect the E*TRADE example as XTB', () => {
    const wb = XLSX.read(new Uint8Array(exampleBuffer.buffer.slice(
      exampleBuffer.byteOffset, exampleBuffer.byteOffset + exampleBuffer.byteLength
    )), { type: 'array' });
    assert.equal(isXtbWorkbook(wb), false);
  });
});


// ═══════════════════════════════════════════════════════════════════════════
// XTB: readXtbReport
// ═══════════════════════════════════════════════════════════════════════════

describe('readXtbReport', () => {
  it('reads all closed positions and skips the Profit/loss row', () => {
    assert.equal(xtbReport.positions.length, 3);
    assert.ok(!xtbReport.positions.some(p => /profit/i.test(p.instrument)));
  });

  it('detects the account currency', () => {
    assert.equal(xtbReport.currency, 'USD');
  });

  it('maps long (BUY) legs: cost = open, revenue = close', () => {
    const p = xtbReport.positions[0];
    assert.equal(p.ticker, 'NVDA.US');
    assert.equal(p.type, 'BUY');
    assert.equal(p.costIso, '2025-07-31');
    assert.equal(p.saleIso, '2026-01-05');
    assert.equal(p.costValue, 1012.0);
    assert.equal(p.revenueValue, 1533.6);
    assert.equal(p.volume, 8);
  });

  it('maps short (SELL) legs: revenue = open, cost = close', () => {
    const p = xtbReport.positions[2];
    assert.equal(p.type, 'SELL');
    assert.equal(p.saleIso, '2026-02-10'); // sale is the opening trade
    assert.equal(p.costIso, '2026-02-12'); // purchase closes the short
    assert.equal(p.revenueValue, 100.0);
    assert.equal(p.costValue, 90.0);
    assert.equal(p.profit, 10.0);
  });

  it('parses free funds interest and its withheld tax', () => {
    assert.equal(xtbReport.interestOps.length, 1);
    assert.equal(xtbReport.interestOps[0].iso, '2026-01-14');
    assert.equal(xtbReport.interestOps[0].amount, 0.04);
    assert.equal(xtbReport.interestTaxOps.length, 1);
    assert.equal(xtbReport.interestTaxOps[0].amount, -0.01);
  });

  it('ignores stock purchase/sell cash operations', () => {
    // Only the two "Free funds interest*" rows are collected
    assert.equal(
      xtbReport.interestOps.length + xtbReport.interestTaxOps.length, 2);
  });
});


// ═══════════════════════════════════════════════════════════════════════════
// XTB: grouping and tax calculation
// ═══════════════════════════════════════════════════════════════════════════

describe('xtbGroupBySaleDate', () => {
  it('groups by sale date and sorts chronologically', () => {
    const groups = xtbGroupBySaleDate(xtbReport.positions);
    assert.deepEqual([...groups.keys()], ['2026-01-05', '2026-02-10']);
    assert.equal(groups.get('2026-01-05').length, 2);
    assert.equal(groups.get('2026-02-10').length, 1);
  });
});

describe('xtbRate', () => {
  it('uses the NBP rates map for foreign currency accounts', () => {
    const rates = new Map([['2026-01-05', { rate: 4.0, effectiveDate: '2026-01-02' }]]);
    assert.equal(xtbRate(rates, 'USD', '2026-01-05').rate, 4.0);
    assert.throws(() => xtbRate(rates, 'USD', '2020-01-01'), /Brak kursu NBP/);
  });

  it('uses rate 1 for PLN accounts', () => {
    assert.equal(xtbRate(null, 'PLN', '2026-01-05').rate, 1);
  });
});

describe('xtbYearSummaries', () => {
  const rates = new Map([
    ['2025-07-31', { rate: 3.80, effectiveDate: '2025-07-30' }],
    ['2025-12-02', { rate: 3.90, effectiveDate: '2025-12-01' }],
    ['2026-01-05', { rate: 4.00, effectiveDate: '2026-01-02' }],
    ['2026-02-10', { rate: 4.10, effectiveDate: '2026-02-09' }],
    ['2026-02-12', { rate: 4.20, effectiveDate: '2026-02-11' }],
  ]);

  it('converts revenue at the sale-date rate and cost at the purchase-date rate', () => {
    const [y] = xtbYearSummaries(xtbReport.positions, rates, 'USD');
    assert.equal(y.year, '2026');
    // 1533.6 + 1143.15 + 100
    assert.ok(Math.abs(y.revenue - 2776.75) < 0.001);
    // 1533.6*4.0 + 1143.15*4.0 + 100*4.1
    assert.ok(Math.abs(y.revenuePln - 11117.0) < 0.001);
    // 1012.0 + 1162.65 + 90
    assert.ok(Math.abs(y.cost - 2264.65) < 0.001);
    // 1012*3.8 + 1162.65*3.9 + 90*4.2
    assert.ok(Math.abs(y.costPln - 8757.935) < 0.001);
  });

  it('taxable amount matches the report Profit/Loss', () => {
    const [y] = xtbYearSummaries(xtbReport.positions, rates, 'USD');
    // 521.6 - 19.5 + 10.0 = 512.1
    assert.ok(Math.abs(y.taxable - 512.1) < 0.001);
    assert.ok(Math.abs(y.taxablePln - 2359.065) < 0.001);
  });

  it('computes 19% tax rounded to 2 decimals', () => {
    const [y] = xtbYearSummaries(xtbReport.positions, rates, 'USD');
    // 2359.065 * 0.19 = 448.22235 → 448.22
    assert.equal(y.taxPln, 448.22);
  });

  it('loss year carries no tax', () => {
    const lossPositions = [{
      instrument: 'X', ticker: 'X.US', type: 'BUY', volume: 1,
      openPrice: 200, closePrice: 100, costValue: 200, revenueValue: 100,
      profit: -100, saleIso: '2025-01-10', costIso: '2024-06-01',
    }];
    const lossRates = new Map([
      ['2025-01-10', { rate: 4.0, effectiveDate: '2025-01-09' }],
      ['2024-06-01', { rate: 4.0, effectiveDate: '2024-05-31' }],
    ]);
    const [y] = xtbYearSummaries(lossPositions, lossRates, 'USD');
    assert.ok(y.taxablePln < 0);
    assert.equal(y.taxPln, 0);
  });

  it('PLN accounts use rate 1 (no conversion)', () => {
    const [y] = xtbYearSummaries(xtbReport.positions, null, 'PLN');
    assert.equal(y.revenuePln, y.revenue);
    assert.equal(y.costPln, y.cost);
  });

  it('splits summaries by tax year', () => {
    const twoYears = [
      ...xtbReport.positions,
      {
        instrument: 'Old', ticker: 'OLD.US', type: 'BUY', volume: 1,
        openPrice: 10, closePrice: 20, costValue: 10, revenueValue: 20,
        profit: 10, saleIso: '2025-03-01', costIso: '2024-03-01',
      },
    ];
    const moreRates = new Map(rates);
    moreRates.set('2025-03-01', { rate: 4.0, effectiveDate: '2025-02-28' });
    moreRates.set('2024-03-01', { rate: 4.0, effectiveDate: '2024-02-29' });

    const summaries = xtbYearSummaries(twoYears, moreRates, 'USD');
    assert.deepEqual(summaries.map(s => s.year), ['2025', '2026']);
  });
});
