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
    parseAmount, parseDateAny, parseTimeAny, normCurrency,
    matchLegsFifo, isIbkrRows, readIbkrRows, isT212Rows, readT212Rows,
    isRevolutRows, readRevolutRows, isEtoroWorkbook, readEtoroReport,
    rateFor, brokerYearSummaries, brokerCurrencyTotals,
  });
`;
eval(wrappedSource);

const {
  WIRE_FEE_USD, esc, fmt, parseDateToIso, formatDate, subtractDays,
  readTransactions, getNbpRate, fetchNbpRates, groupBySaleDate,
  excelSerialToIso, xtbDateToIso, isXtbWorkbook, readXtbReport,
  xtbGroupBySaleDate, xtbRate, xtbYearSummaries, findPrevBusinessRate,
  parseAmount, parseDateAny, parseTimeAny, normCurrency,
  matchLegsFifo, isIbkrRows, readIbkrRows, isT212Rows, readT212Rows,
  isRevolutRows, readRevolutRows, isEtoroWorkbook, readEtoroReport,
  rateFor, brokerYearSummaries, brokerCurrencyTotals,
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

// ═══════════════════════════════════════════════════════════════
// Multi-broker support (IBKR, Trading 212, Revolut, eToro)
// ═══════════════════════════════════════════════════════════════

describe('parseAmount', () => {
  it('parses plain numbers and numeric cells', () => {
    assert.equal(parseAmount(12.5), 12.5);
    assert.equal(parseAmount('123.45'), 123.45);
  });

  it('strips currency prefixes and spaces', () => {
    assert.equal(parseAmount('EUR 104.64'), 104.64);
    assert.equal(parseAmount('USD 1 234.56'), 1234.56);
  });

  it('handles both thousand/decimal separator styles', () => {
    assert.equal(parseAmount('1,234.56'), 1234.56);
    assert.equal(parseAmount('1.234,56'), 1234.56);
    assert.equal(parseAmount('1,234'), 1234);
    assert.equal(parseAmount('12,34'), 12.34);
  });

  it('handles negatives (minus and parentheses)', () => {
    assert.equal(parseAmount('-5'), -5);
    assert.equal(parseAmount('(123.45)'), -123.45);
    assert.equal(parseAmount(''), 0);
    assert.equal(parseAmount(null), 0);
  });
});

describe('parseDateAny / parseTimeAny / normCurrency', () => {
  it('parses ISO timestamps to date strings', () => {
    assert.equal(parseDateAny('2025-01-06T09:02:14.772Z'), '2025-01-06');
    assert.equal(parseDateAny('2026-07-24 01:16:11+00:00'), '2026-07-24');
  });

  it('parses IBKR "YYYY-MM-DD, HH:MM:SS"', () => {
    assert.equal(parseDateAny('2017-07-08, 08:53:11'), '2017-07-08');
  });

  it('parses E*TRADE dates, serials and Date objects', () => {
    assert.equal(parseDateAny('03/15/2025'), '2025-03-15');
    assert.equal(parseDateAny(45000), excelSerialToIso(45000));
    assert.equal(parseDateAny(new Date(Date.UTC(2025, 0, 6))), '2025-01-06');
    assert.equal(parseDateAny('--'), null);
    assert.equal(parseDateAny(''), null);
  });

  it('parseTimeAny keeps the time for same-day ordering', () => {
    assert.equal(parseTimeAny('2025-01-06T09:02:14.772Z'), '2025-01-06 09:02:14');
    assert.equal(parseTimeAny('2025-01-06'), '2025-01-06');
  });

  it('normCurrency maps GBX to GBP', () => {
    assert.equal(normCurrency('GBX'), 'GBP');
    assert.equal(normCurrency('usd'), 'USD');
    assert.equal(normCurrency(null), 'USD');
  });
});

describe('matchLegsFifo', () => {
  const leg = (iso, qty, value, extra = {}) => ({
    instrument: 'X', ticker: 'X', category: 'Akcje', currency: 'USD',
    iso, ts: iso, qty, value, fee: 0, ...extra,
  });

  it('matches long trades FIFO with partial fills', () => {
    const positions = matchLegsFifo([
      leg('2025-01-02', 10, 100),  // buy 10 @ 10
      leg('2025-02-02', 10, 200),  // buy 10 @ 20
      leg('2025-03-02', -15, 450), // sell 15 @ 30
    ]);
    assert.equal(positions.length, 2);
    assert.equal(positions[0].volume, 10);
    assert.equal(positions[0].costValue, 100);
    assert.equal(positions[0].revenueValue, 300);
    assert.equal(positions[0].costIso, '2025-01-02');
    assert.equal(positions[0].saleIso, '2025-03-02');
    assert.equal(positions[1].volume, 5);
    assert.equal(positions[1].costValue, 100); // 5 @ 20
    assert.equal(positions[1].revenueValue, 150);
    assert.equal(positions[1].costIso, '2025-02-02');
  });

  it('adds both fees to the cost of the matched position', () => {
    const positions = matchLegsFifo([
      leg('2025-01-02', 10, 100, { fee: 2 }),
      leg('2025-03-02', -10, 200, { fee: 3 }),
    ]);
    assert.equal(positions.length, 1);
    assert.equal(positions[0].costValue, 105); // 100 + 2 + 3
    assert.equal(positions[0].revenueValue, 200);
  });

  it('handles short positions (sell first, buy back later)', () => {
    const positions = matchLegsFifo([
      leg('2025-03-02', -5, 500),
      leg('2025-04-02', 5, 400),
    ]);
    assert.equal(positions.length, 1);
    assert.equal(positions[0].type, 'SELL'); // short position
    assert.equal(positions[0].saleIso, '2025-03-02'); // sale happened first
    assert.equal(positions[0].costIso, '2025-04-02');
    assert.equal(positions[0].revenueValue, 500);
    assert.equal(positions[0].costValue, 400);
  });

  it('ignores legs that never close', () => {
    const positions = matchLegsFifo([leg('2025-01-02', 10, 100)]);
    assert.equal(positions.length, 0);
  });

  it('orders same-day legs by time', () => {
    const positions = matchLegsFifo([
      { ...leg('2025-01-02', -1, 20), ts: '2025-01-02 15:00:00' },
      { ...leg('2025-01-02', 1, 10), ts: '2025-01-02 10:00:00' },
    ]);
    assert.equal(positions.length, 1);
    assert.equal(positions[0].type, 'BUY'); // long: bought 10:00, sold 15:00
  });
});

describe('IBKR Activity Statement', () => {
  const ibkrRows = [
    ['Statement', 'Header', 'Field Name', 'Field Value'],
    ['Statement', 'Data', 'BaseCurrency', 'USD'],
    ['Trades', 'Header', 'DataDiscriminator', 'Asset Category', 'Currency', 'Symbol', 'Date/Time', 'Quantity', 'T. Price', 'C. Price', 'Proceeds', 'Comm/Fee', 'Basis', 'Realized P/L', 'MTM P/L', 'Code'],
    ['Trades', 'Data', 'Order', 'Stocks', 'USD', 'AAPL', '2025-03-03, 15:30:00', 10, 100, 101, -1000, -1, 1001, 0, 0, 'O'],
    ['Trades', 'Data', 'Order', 'Stocks', 'USD', 'AAPL', '2025-06-02, 15:30:00', -10, 120, 119, 1200, -1, -1001, 198, 0, 'C'],
    ['Trades', 'Data', 'Closed Lot', 'Stocks', 'USD', 'AAPL', '2025-06-02, 15:30:00', -10, 100, 119, 1200, -1, -1001, 198, 0, 'C'],
    ['Trades', 'SubTotal', '', 'Stocks', 'USD', 'AAPL', '', 0, 120, 119, 200, -2, 0, 198, 0, ''],
    ['Trades', 'Total', '', 'Stocks', 'USD', '', '', 0, 120, 119, 200, -2, 0, 198, 0, ''],
    ['Dividends', 'Header', 'Currency', 'Date', 'Description', 'Amount'],
    ['Dividends', 'Data', 'USD', '2025-05-15', 'AAPL Cash Dividend', 2.5],
    ['Withholding Tax', 'Header', 'Currency', 'Date', 'Description', 'Amount'],
    ['Withholding Tax', 'Data', 'USD', '2025-05-15', 'AAPL Cash Dividend', -0.38],
  ];

  it('detects IBKR activity statement rows', () => {
    assert.equal(isIbkrRows(ibkrRows), true);
    assert.equal(isIbkrRows([['Action', 'Time']]), false);
  });

  it('reads trades, skips Closed Lot/SubTotal/Total rows', () => {
    const report = readIbkrRows(ibkrRows);
    assert.equal(report.positions.length, 1);
    const [p] = report.positions;
    assert.equal(p.volume, 10);
    assert.equal(p.costValue, 1002);   // 1000 + buy 1 + sell 1 commission
    assert.equal(p.revenueValue, 1200);
    assert.equal(p.costIso, '2025-03-03');
    assert.equal(p.saleIso, '2025-06-02');
    assert.equal(p.currency, 'USD');
  });

  it('reads dividends and withholding tax (informational)', () => {
    const report = readIbkrRows(ibkrRows);
    assert.deepEqual(report.dividendOps, [
      { iso: '2025-05-15', amount: 2.5, currency: 'USD' },
    ]);
    assert.deepEqual(report.dividendTaxOps, [
      { iso: '2025-05-15', amount: 0.38, currency: 'USD' },
    ]);
  });

  it('skips non-stock asset categories', () => {
    const rows = [
      ['Trades', 'Header', 'DataDiscriminator', 'Asset Category', 'Currency', 'Symbol', 'Date/Time', 'Quantity', 'T. Price', 'C. Price', 'Proceeds', 'Comm/Fee', 'Basis', 'Realized P/L', 'MTM P/L', 'Code'],
      ['Trades', 'Data', 'Order', 'Equity and Index Options', 'USD', 'TSLA 27MAR26 300 C', '2025-03-03, 15:30:00', 1, 500, 510, -500, -2, 502, 0, 0, 'O'],
      ['Trades', 'Data', 'Order', 'Stocks', 'USD', 'TSLA', '2025-03-03, 15:30:00', 1, 100, 101, -100, -1, 101, 0, 0, 'O'],
    ];
    const report = readIbkrRows(rows);
    assert.equal(report.positions.length, 0); // buy legs only, options skipped
  });
});

describe('Trading 212 export', () => {
  const t212Rows = [
    ['Action', 'Time', 'ISIN', 'Ticker', 'Name', 'No. of shares', 'Price / share', 'Currency (Price / share)', 'Exchange rate', 'Result', 'Total', 'Withholding tax', 'Charge amount', 'Currency conversion fee', 'Notes', 'ID'],
    ['Market buy', '2025-02-03 10:00:00+00:00', 'US0378331005', 'AAPL', 'Apple', 5, 150, 'USD', 4.1, null, -750, null, 0, 1.2, '', '1'],
    ['Market sell', '2025-07-14 10:00:00+00:00', 'US0378331005', 'AAPL', 'Apple', 5, 200, 'USD', 4.0, 250, 1000, null, 0, 1.5, '', '2'],
    ['Dividend (Ordinary)', '2025-05-15 10:00:00+00:00', 'US0378331005', 'AAPL', 'Apple', 5, null, 'USD', 4.1, null, 2.13, 0.38, null, null, '', '3'],
    ['Market buy', '2025-03-03 10:00:00+00:00', 'GB0002374006', 'HSBA', 'HSBC', 10, 500, 'GBX', 5.2, null, -50, null, 0, null, '', '4'],
    ['Deposit', '2025-01-01 09:00:00+00:00', null, null, null, null, null, null, null, null, 10000, null, null, null, '', '5'],
  ];

  it('detects Trading 212 rows', () => {
    assert.equal(isT212Rows(t212Rows), true);
    assert.equal(isT212Rows(ibkrFixtureRows()), false);
  });

  it('matches buy/sell into a position', () => {
    const report = readT212Rows(t212Rows);
    assert.equal(report.positions.length, 1);
    const [p] = report.positions;
    assert.equal(p.volume, 5);
    assert.equal(p.costValue, 750);
    assert.equal(p.revenueValue, 1000);
    assert.equal(p.costIso, '2025-02-03');
    assert.equal(p.saleIso, '2025-07-14');
    assert.equal(p.currency, 'USD');
  });

  it('converts GBX (pence) to GBP and skips deposits', () => {
    const report = readT212Rows(t212Rows);
    // The HSBC buy stays open (no sell) — not in positions…
    assert.ok(!report.positions.some(p => p.ticker === 'HSBA'));
    // …but a closed GBX trade would be denominated in GBP.
    const rows = [t212Rows[0], t212Rows[4], [
      'Market sell', '2025-04-03 10:00:00+00:00', 'GB0002374006', 'HSBA', 'HSBC', 10, 600, 'GBX', 5.2, null, 60, null, 0, null, '', '6',
    ]];
    const [p] = readT212Rows(rows).positions;
    assert.equal(p.currency, 'GBP');
    assert.equal(p.costValue, 50);  // 10 × 500 GBX = 5 000 pence = £50
    assert.equal(p.revenueValue, 60);
  });

  it('reads dividends with withholding tax (informational)', () => {
    const report = readT212Rows(t212Rows);
    assert.deepEqual(report.dividendOps, [
      { iso: '2025-05-15', amount: 2.51, currency: 'USD' },
    ]);
    assert.deepEqual(report.dividendTaxOps, [
      { iso: '2025-05-15', amount: 0.38, currency: 'USD' },
    ]);
  });
});

function ibkrFixtureRows() {
  return [
    ['Statement', 'Header', 'Field Name', 'Field Value'],
    ['Trades', 'Header', 'DataDiscriminator', 'Asset Category', 'Currency', 'Symbol', 'Date/Time', 'Quantity', 'T. Price'],
  ];
}

describe('Revolut statement', () => {
  const revolutRows = [
    ['Date', 'Ticker', 'Type', 'Quantity', 'Price per share', 'Total Amount', 'Currency', 'FX Rate', 'Fees', 'Taxes'],
    ['2025-01-06T09:02:14.772Z', 'AAPL', 'BUY - MARKET', 2, 'EUR 100.00', 200, 'EUR', 1, 0.5, 0],
    ['2025-04-02T14:05:00.000Z', 'AAPL', 'SELL - MARKET', 2, 'EUR 120.00', 240, 'EUR', 1, 0.5, 0],
    ['2025-03-10T00:00:00.000Z', 'AAPL', 'DIVIDEND', null, null, 1.85, 'EUR', 1, 0, 0.28],
    ['2025-01-01T00:00:00.000Z', null, 'CASH TOP-UP', null, null, 1000, 'EUR', 1, 0, 0],
  ];

  it('detects Revolut rows', () => {
    assert.equal(isRevolutRows(revolutRows), true);
    assert.equal(isRevolutRows(t212FixtureRows()), false);
  });

  it('matches BUY/SELL rows into a position with fees in cost', () => {
    const report = readRevolutRows(revolutRows);
    assert.equal(report.positions.length, 1);
    const [p] = report.positions;
    assert.equal(p.volume, 2);
    assert.equal(p.costValue, 201);    // 2 × 100 + 0.5 + 0.5 fees
    assert.equal(p.revenueValue, 240); // 2 × 120
    assert.equal(p.costIso, '2025-01-06');
    assert.equal(p.saleIso, '2025-04-02');
    assert.equal(p.currency, 'EUR');
  });

  it('reads dividends and skips cash top-ups', () => {
    const report = readRevolutRows(revolutRows);
    assert.deepEqual(report.dividendOps, [
      { iso: '2025-03-10', amount: 2.13, currency: 'EUR' },
    ]);
    assert.deepEqual(report.dividendTaxOps, [
      { iso: '2025-03-10', amount: 0.28, currency: 'EUR' },
    ]);
    assert.equal(report.interestOps.length, 0);
  });
});

function t212FixtureRows() {
  return [
    ['Action', 'Time', 'ISIN', 'Ticker', 'Name', 'No. of shares', 'Price / share', 'Currency (Price / share)', 'Total'],
  ];
}

describe('eToro account statement', () => {
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet([
    ['Account Summary'],
    ['Account Base Currency', 'USD'],
  ]), 'Account Summary');
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet([
    ['Position ID', 'Instrument', 'Direction', 'Size (units)', 'Open Date', 'Close Date', 'Open Rate', 'Close Rate', 'Leverage', 'Fees and Dividends', 'Profit'],
    [1, 'Apple', 'Buy', 10, '2025-01-06', '2025-04-02', 100, 120, 1, 0, 200],
    [2, 'Tesla', 'Sell', 5, '2025-02-03', '2025-03-03', 200, 180, 1, 0, 100],
    [3, 'Bitcoin', 'Buy', 1, '2025-05-05', null, 50000, null, 1, 0, 0],
  ]), 'Closed Positions');
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet([
    ['Position ID', 'Instrument', 'Date', 'Net Amount', 'Withholding Tax'],
    [1, 'Apple', '2025-03-10', 1.72, 0.28],
  ]), 'Dividends');

  it('detects an eToro workbook', () => {
    assert.equal(isEtoroWorkbook(workbook), true);
    assert.equal(isEtoroWorkbook({ SheetNames: ['Closed Positions', 'Cash Operations'] }), false);
  });

  it('reads long and short closed positions, skips open ones', () => {
    const report = readEtoroReport(workbook);
    assert.equal(report.positions.length, 2);

    const [long, short] = report.positions;
    assert.equal(long.type, 'BUY');
    assert.equal(long.costValue, 1000);    // 10 × 100 at open
    assert.equal(long.revenueValue, 1200); // 10 × 120 at close
    assert.equal(long.costIso, '2025-01-06');
    assert.equal(long.saleIso, '2025-04-02');

    assert.equal(short.type, 'SELL');
    assert.equal(short.saleIso, '2025-02-03'); // short: sale at open
    assert.equal(short.revenueValue, 1000);    // 5 × 200
    assert.equal(short.costIso, '2025-03-03');
    assert.equal(short.costValue, 900);        // 5 × 180 at close
  });

  it('uses the account base currency from Account Summary', () => {
    const report = readEtoroReport(workbook);
    assert.equal(report.positions[0].currency, 'USD');
  });

  it('reads dividends (gross = net + withholding)', () => {
    const report = readEtoroReport(workbook);
    assert.deepEqual(report.dividendOps, [
      { iso: '2025-03-10', amount: 2, currency: 'USD' },
    ]);
    assert.deepEqual(report.dividendTaxOps, [
      { iso: '2025-03-10', amount: 0.28, currency: 'USD' },
    ]);
  });
});

describe('rateFor & brokerYearSummaries (multi-currency)', () => {
  const rates = new Map([
    ['USD|2025-01-06', { rate: 4.0, effectiveDate: '2025-01-03' }],
    ['USD|2025-04-02', { rate: 4.2, effectiveDate: '2025-04-01' }],
    ['EUR|2025-01-06', { rate: 4.1, effectiveDate: '2025-01-03' }],
    ['EUR|2025-04-02', { rate: 4.3, effectiveDate: '2025-04-01' }],
  ]);

  it('rateFor reads multi-currency keys, legacy keys and PLN', () => {
    assert.equal(rateFor(rates, 'USD', '2025-01-06').rate, 4.0);
    assert.equal(rateFor(rates, 'PLN', '2025-01-06').rate, 1);
    const legacy = new Map([['2025-01-06', { rate: 3.9, effectiveDate: '2025-01-03' }]]);
    assert.equal(rateFor(legacy, 'USD', '2025-01-06').rate, 3.9);
  });

  it('converts each side at its own date and currency', () => {
    const positions = [
      {
        instrument: 'A', ticker: 'A', category: 'Akcje', currency: 'USD',
        type: 'BUY', volume: 1, costValue: 100, revenueValue: 120,
        costIso: '2025-01-06', saleIso: '2025-04-02',
      },
      {
        instrument: 'B', ticker: 'B', category: 'Akcje', currency: 'EUR',
        type: 'BUY', volume: 1, costValue: 50, revenueValue: 60,
        costIso: '2025-01-06', saleIso: '2025-04-02',
      },
    ];
    const [y] = brokerYearSummaries(positions, rates);
    assert.equal(y.year, '2025');
    // 120×4.2 + 60×4.3 = 504 + 258 = 762
    assert.equal(y.revenuePln, 762);
    // 100×4.0 + 50×4.1 = 400 + 205 = 605
    assert.equal(y.costPln, 605);
    assert.equal(y.taxablePln, 157);
    assert.equal(y.taxPln, Math.round(157 * 0.19 * 100) / 100);
    // per-currency subtotals are kept in original currencies
    assert.deepEqual(y.byCurrency.get('USD'), { revenue: 120, cost: 100 });
    assert.deepEqual(y.byCurrency.get('EUR'), { revenue: 60, cost: 50 });
  });

  it('loss years have zero tax', () => {
    const positions = [{
      instrument: 'A', ticker: 'A', category: 'Akcje', currency: 'USD',
      type: 'BUY', volume: 1, costValue: 100, revenueValue: 50,
      costIso: '2025-01-06', saleIso: '2025-04-02',
    }];
    const [y] = brokerYearSummaries(positions, rates);
    assert.ok(y.taxablePln < 0);
    assert.equal(y.taxPln, 0);
  });
});
