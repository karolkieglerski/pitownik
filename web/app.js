'use strict';

/* ═══════════════════════════════════════════════════════════════
   pitownik — app.js
   Kalkulator PIT-38 na podstawie pliku Gains & Losses z E*TRADE
   ═══════════════════════════════════════════════════════════════ */

const WIRE_FEE_USD = 20.0; // $10 E*TRADE + $10 bank fee

// ─────────────────────────────────────────────────────────────
// Helpers
// ─────────────────────────────────────────────────────────────

/** Escape HTML special characters to prevent XSS. */
function esc(val) {
  return String(val ?? '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

/** Format a number to fixed decimal places. */
function fmt(num, dec = 2) {
  return Number(num).toFixed(dec);
}

// ─────────────────────────────────────────────────────────────
// Date utilities
// ─────────────────────────────────────────────────────────────

/**
 * Convert a date value (JS Date object or "MM/DD/YYYY" string) to
 * an ISO "YYYY-MM-DD" string. Returns null for blank/dashes.
 */
function parseDateToIso(dateVal) {
  if (!dateVal || dateVal === '--') return null;

  if (dateVal instanceof Date) {
    // SheetJS returns local Date objects; read local components.
    const y = dateVal.getFullYear();
    const m = String(dateVal.getMonth() + 1).padStart(2, '0');
    const d = String(dateVal.getDate()).padStart(2, '0');
    return `${y}-${m}-${d}`;
  }

  const s = String(dateVal).trim();

  // MM/DD/YYYY  (E*TRADE export format when stored as text)
  const mmddyyyy = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (mmddyyyy) {
    return `${mmddyyyy[3]}-${mmddyyyy[1].padStart(2, '0')}-${mmddyyyy[2].padStart(2, '0')}`;
  }

  // Already YYYY-MM-DD
  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;

  throw new Error(`Nieznany format daty: ${dateVal}`);
}

/** Format an ISO date string or date value as "DD.MM.YYYY". */
function formatDate(isoOrDateVal) {
  const iso = typeof isoOrDateVal === 'string' && /^\d{4}-\d{2}-\d{2}$/.test(isoOrDateVal)
    ? isoOrDateVal
    : parseDateToIso(isoOrDateVal);
  if (!iso) return '—';
  const [y, m, d] = iso.split('-');
  return `${d}.${m}.${y}`;
}

/** Subtract n calendar days from an ISO date string; return new ISO string. */
function subtractDays(isoDate, n) {
  const [y, m, d] = isoDate.split('-').map(Number);
  const dt = new Date(Date.UTC(y, m - 1, d));
  dt.setUTCDate(dt.getUTCDate() - n);
  return [
    String(dt.getUTCFullYear()),
    String(dt.getUTCMonth() + 1).padStart(2, '0'),
    String(dt.getUTCDate()).padStart(2, '0'),
  ].join('-');
}

// ─────────────────────────────────────────────────────────────
// XLSX parsing
// ─────────────────────────────────────────────────────────────

/**
 * Parse an ArrayBuffer containing a .xlsx file and return only
 * the "Sell" transaction rows as an array of plain objects.
 */
function readTransactions(arrayBuffer) {
  const workbook = XLSX.read(new Uint8Array(arrayBuffer), {
    type: 'array',
    cellDates: true,  // parse date cells as JS Date objects
  });
  const worksheet = workbook.Sheets[workbook.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(worksheet, {
    raw: true,        // keep numbers as numbers
    cellDates: true,
    defval: null,     // use null for empty cells
  });
  return rows.filter(row => row['Record Type'] === 'Sell');
}

// ─────────────────────────────────────────────────────────────
// XTB report parsing
// ─────────────────────────────────────────────────────────────

/**
 * Convert an Excel serial date number (1900 date system) to an ISO
 * "YYYY-MM-DD" string. XTB report times are UTC, so the date part is
 * taken in UTC as well.
 */
function excelSerialToIso(serial) {
  const n = Number(serial);
  if (!Number.isFinite(n)) {
    throw new Error(`Nieprawidłowa data (Excel serial): ${serial}`);
  }
  const days = Math.floor(n);
  return new Date((days - 25569) * 86400000).toISOString().slice(0, 10);
}

/**
 * Convert a date cell value from an XTB report to an ISO "YYYY-MM-DD"
 * string. XTB stores dates as Excel serial numbers with a UTC time
 * fraction; text values fall back to parseDateToIso().
 */
function xtbDateToIso(val) {
  if (val === null || val === undefined || val === '' || val === '--') return null;
  if (typeof val === 'number') return excelSerialToIso(val);

  const s = String(val).trim();

  // "YYYY-MM-DD HH:MM[:SS]" — drop the time part
  const isoWithTime = s.match(/^(\d{4}-\d{2}-\d{2})[ T]\d{1,2}:\d{2}/);
  if (isoWithTime) return isoWithTime[1];

  return parseDateToIso(val);
}

/** Find an XTB worksheet by case-insensitive name; null if absent. */
function xtbSheet(workbook, name) {
  const wanted = name.trim().toLowerCase();
  const sheetName = workbook.SheetNames.find(n => n.trim().toLowerCase() === wanted);
  return sheetName ? workbook.Sheets[sheetName] : null;
}

/** True if the workbook looks like an XTB account statement. */
function isXtbWorkbook(workbook) {
  const names = workbook.SheetNames.map(n => n.trim().toLowerCase());
  if (names.includes('closed positions') || names.includes('cash operations')) return true;

  // Fallback: look for XTB column headers in the first rows of sheet 1.
  const ws = workbook.Sheets[workbook.SheetNames[0]];
  if (!ws) return false;
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null });
  const flat = rows.slice(0, 10).flat().map(v => String(v ?? '').trim());
  return flat.includes('Close Time (UTC)') && flat.includes('Position ID');
}

/**
 * Parse one XTB worksheet into an array of objects keyed by its header
 * row. XTB sheets start with a few rows of report metadata, so the header
 * row is located by an exact marker cell value.
 */
function xtbSheetRecords(worksheet, headerMarker) {
  const rows = XLSX.utils.sheet_to_json(worksheet, {
    header: 1, raw: true, defval: null, blankrows: true,
  });
  const headerIdx = rows.findIndex(
    r => r.some(c => String(c ?? '').trim() === headerMarker),
  );
  if (headerIdx === -1) {
    throw new Error(`Nie znaleziono wiersza nagłówka "${headerMarker}"`);
  }

  const headers = rows[headerIdx].map(h => String(h ?? '').trim());
  const records = [];
  for (const row of rows.slice(headerIdx + 1)) {
    const rec = {};
    let hasValue = false;
    headers.forEach((h, i) => {
      const v = row[i] ?? null;
      rec[h] = v;
      if (v !== null && v !== '') hasValue = true;
    });
    if (hasValue) records.push(rec);
  }
  return records;
}

/** Detect the account currency of an XTB report ("USD", "PLN", ...). */
function xtbCurrency(workbook, fileName) {
  const ws = xtbSheet(workbook, 'Open Positions');
  if (ws) {
    const rows = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null });
    const headerIdx = rows.findIndex(
      r => r.some(c => String(c ?? '').trim() === 'Currency'),
    );
    if (headerIdx !== -1) {
      const col = rows[headerIdx]
        .findIndex(c => String(c ?? '').trim() === 'Currency');
      for (const row of rows.slice(headerIdx + 1)) {
        const m = String(row[col] ?? '').trim().match(/^[A-Za-z]{3}$/);
        if (m) return m[0].toUpperCase();
      }
    }
  }

  // Report file names look like "USD_51672638_2025-12-31_2026-11-20.xlsx"
  const m = String(fileName ?? '').match(/^([A-Za-z]{3})_/);
  return m ? m[1].toUpperCase() : 'USD';
}

/**
 * Parse an ArrayBuffer containing an XTB account statement (.xlsx).
 * Returns { currency, positions, interestOps, interestTaxOps } where each
 * position carries both legs of the closed trade:
 *   revenueValue / saleIso  — the sale leg (przychód),
 *   costValue   / costIso   — the purchase leg (koszt uzyskania).
 */
function readXtbReport(arrayBuffer, fileName) {
  const workbook = XLSX.read(new Uint8Array(arrayBuffer), {
    type: 'array',
    cellDates: false, // keep dates as raw Excel serial numbers
  });

  const closedWs = xtbSheet(workbook, 'Closed Positions');
  if (!closedWs) throw new Error('Nie znaleziono arkusza "Closed Positions"');

  const positions = xtbSheetRecords(closedWs, 'Close Time (UTC)')
    .filter(r => {
      const name = String(r['Instrument'] ?? '').trim();
      return name !== '' && !/^(profit\/loss|total)$/i.test(name);
    })
    .map(r => {
      const type = String(r['Type'] ?? 'BUY').trim().toUpperCase();
      const isShort = type === 'SELL';
      // "Sale Value" / "Purchase Value" are named after the transaction
      // leg. For a long (BUY) position the sale is the closing trade and
      // the purchase is the opening trade; for a short (SELL) position the
      // sale is the opening trade and the purchase closes the position.
      const saleIso = xtbDateToIso(
        r[isShort ? 'Open Time (UTC)' : 'Close Time (UTC)']);
      const costIso = xtbDateToIso(
        r[isShort ? 'Close Time (UTC)' : 'Open Time (UTC)']);
      if (!saleIso || !costIso) {
        throw new Error(
          `Pozycja ${String(r['Ticker'] ?? '?').trim()} bez daty otwarcia/zamknięcia`);
      }

      return {
        instrument: String(r['Instrument'] ?? '').trim(),
        ticker: String(r['Ticker'] ?? '').trim(),
        category: String(r['Category'] ?? '').trim(),
        type,
        volume: Number(r['Volume'] ?? 0) || 0,
        openPrice: Number(r['Open Price'] ?? 0) || 0,
        closePrice: Number(r['Close Price'] ?? 0) || 0,
        costValue: Number(r['Purchase Value'] ?? 0) || 0,
        revenueValue: Number(r['Sale Value'] ?? 0) || 0,
        profit: Number(r['Profit/Loss'] ?? 0) || 0,
        saleIso,
        costIso,
      };
    });

  const interestOps = [];
  const interestTaxOps = [];
  const cashWs = xtbSheet(workbook, 'Cash Operations');
  if (cashWs) {
    for (const r of xtbSheetRecords(cashWs, 'Amount')) {
      const type = String(r['Type'] ?? '').trim().toLowerCase();
      if (type !== 'free funds interest' && type !== 'free funds interest tax') continue;
      const op = {
        iso: xtbDateToIso(r['Time']),
        amount: Number(r['Amount'] ?? 0) || 0,
      };
      if (type === 'free funds interest') interestOps.push(op);
      else interestTaxOps.push(op);
    }
  }

  return {
    currency: xtbCurrency(workbook, fileName),
    positions,
    interestOps,
    interestTaxOps,
  };
}

/** Group XTB positions by sale date (ISO string key), sorted chronologically. */
function xtbGroupBySaleDate(positions) {
  const groups = new Map();
  for (const p of positions) {
    if (!groups.has(p.saleIso)) groups.set(p.saleIso, []);
    groups.get(p.saleIso).push(p);
  }
  return new Map([...groups.entries()].sort((a, b) => a[0].localeCompare(b[0])));
}

/**
 * NBP rate entry for one XTB row. Accounts held in PLN need no conversion
 * (rate 1). Otherwise the rates map holds the NBP rate from the day before
 * the transaction date — the same rule as in the E*TRADE flow.
 */
function xtbRate(rates, currency, iso) {
  if (currency === 'PLN') return { rate: 1, effectiveDate: null };
  const entry = rates.get(iso);
  if (!entry) throw new Error(`Brak kursu NBP dla daty ${iso}`);
  return entry;
}

/**
 * Per-tax-year XTB summary: revenue, cost and taxable amount in the
 * account currency and PLN plus the 19% tax. Each side is converted at
 * the NBP rate of its own transaction date (revenue — sale date, cost —
 * purchase date), exactly like ESPP costs in the E*TRADE flow.
 */
function xtbYearSummaries(positions, rates, currency) {
  const byYear = new Map();
  for (const p of positions) {
    const year = p.saleIso.slice(0, 4);
    if (!byYear.has(year)) {
      byYear.set(year, { year, revenue: 0, revenuePln: 0, cost: 0, costPln: 0 });
    }
    const y = byYear.get(year);
    const saleRate = xtbRate(rates, currency, p.saleIso).rate;
    const costRate = xtbRate(rates, currency, p.costIso).rate;
    y.revenue += p.revenueValue;
    y.revenuePln += p.revenueValue * saleRate;
    y.cost += p.costValue;
    y.costPln += p.costValue * costRate;
  }

  return [...byYear.values()]
    .sort((a, b) => a.year.localeCompare(b.year))
    .map(y => {
      const taxablePln = y.revenuePln - y.costPln;
      return {
        ...y,
        taxable: y.revenue - y.cost,
        taxablePln,
        taxPln: Math.round(Math.max(0, taxablePln) * 0.19 * 100) / 100,
      };
    });
}

// ─────────────────────────────────────────────────────────────
// NBP API
// ─────────────────────────────────────────────────────────────

/**
 * Fetch the average (table A) rate from NBP against PLN for the last
 * business day before the given ISO date. Currency defaults to USD
 * (E*TRADE flow); XTB reports pass the account currency.
 * Retries up to 7 days back on 404 (weekends / public holidays).
 * Returns { rate: number, effectiveDate: string }.
 */
async function getNbpRate(isoDate, currency = 'usd') {
  let day = subtractDays(isoDate, 1);
  const code = String(currency).toLowerCase();

  for (let attempt = 0; attempt < 7; attempt++) {
    const url =
      `https://api.nbp.pl/api/exchangerates/rates/a/${code}/${day}/?format=json`;
    let resp;
    try {
      resp = await fetch(url);
    } catch (err) {
      throw new Error(`Błąd połączenia z NBP API: ${err.message}`);
    }

    if (resp.ok) {
      const data = await resp.json();
      return {
        rate: data.rates[0].mid,
        effectiveDate: data.rates[0].effectiveDate,
      };
    }

    if (resp.status === 404) {
      day = subtractDays(day, 1); // try the previous day
    } else {
      throw new Error(`NBP API zwróciło błąd ${resp.status} dla daty ${day}`);
    }
  }

  throw new Error(
    `Nie udało się pobrać kursu NBP dla daty ${isoDate} (sprawdzono 7 dni wstecz)`,
  );
}

/**
 * Fetch NBP rates for an array of ISO date strings.
 * Deduplicates and sorts dates; calls onProgress(iso, {rate, effectiveDate})
 * after each successful lookup. Currency defaults to USD.
 *
 * Uses NBP *range* queries — one request per ~year of data instead of one
 * request per date. Rapid single-date requests are throttled by NBP and
 * fail with connection errors ("Load failed") on files with many dates.
 * Returns a Map of ISO date string → { rate, effectiveDate }.
 */
async function fetchNbpRates(isoDateArray, onProgress, currency = 'usd') {
  const unique = [...new Set(isoDateArray.filter(Boolean))].sort();
  const ratesMap = new Map();
  if (unique.length === 0) return ratesMap;

  const code = String(currency).toLowerCase();
  const available = await fetchNbpRateRange(
    code,
    subtractDays(unique[0], 7),
    unique[unique.length - 1],
  );

  for (const iso of unique) {
    const entry = findPrevBusinessRate(available, iso);
    if (!entry) {
      throw new Error(
        `Nie udało się pobrać kursu NBP dla daty ${iso} (sprawdzono 7 dni wstecz)`,
      );
    }
    const result = { rate: entry.mid, effectiveDate: entry.effectiveDate };
    ratesMap.set(iso, result);
    if (onProgress) onProgress(iso, result);
  }

  return ratesMap;
}

/**
 * Last NBP rate strictly before isoDate, looking back at most 7 days
 * (the same window as the single-date getNbpRate() fallback).
 * `ratesAsc` is an ascending array of { effectiveDate, mid }.
 */
function findPrevBusinessRate(ratesAsc, isoDate) {
  const limit = subtractDays(isoDate, 7);
  let found = null;
  for (const r of ratesAsc) {
    if (r.effectiveDate < isoDate && r.effectiveDate >= limit) found = r;
  }
  return found;
}

/**
 * Fetch NBP table A rates for [fromIso, toIso] using range queries
 * (max 367 days per request; consecutive chunks overlap so that the
 * 7-day "previous business day" window stays covered). The end date is
 * clamped to today — NBP rejects future ranges with an error.
 * Returns an ascending array of { effectiveDate, mid }.
 */
async function fetchNbpRateRange(currencyCode, fromIso, toIso) {
  const today = new Date().toISOString().slice(0, 10);
  const end = toIso > today ? today : toIso;

  const CHUNK = 365;  // NBP allows up to 367 days per range query
  const OVERLAP = 7;  // keeps the 7-day lookup window covered across chunks

  const out = [];
  let chunkStart = fromIso;
  while (chunkStart <= end) {
    let chunkEnd = subtractDays(chunkStart, -CHUNK);
    if (chunkEnd > end) chunkEnd = end;

    for (const r of await fetchNbpChunk(currencyCode, chunkStart, chunkEnd)) {
      if (!out.some(x => x.effectiveDate === r.effectiveDate)) out.push(r);
    }

    if (chunkEnd === end) break;
    chunkStart = subtractDays(chunkEnd, OVERLAP - 1);
  }

  out.sort((a, b) => a.effectiveDate.localeCompare(b.effectiveDate));
  return out;
}

/**
 * One NBP range request with retries on transient network errors and
 * rate limiting. A 404 means "no quotes in this range" (e.g. weekends
 * only) and yields an empty list.
 */
async function fetchNbpChunk(currencyCode, fromIso, toIso) {
  const url =
    `https://api.nbp.pl/api/exchangerates/rates/a/${currencyCode}` +
    `/${fromIso}/${toIso}/?format=json`;

  for (let attempt = 1; attempt <= 3; attempt++) {
    let resp;
    try {
      resp = await fetch(url);
    } catch (err) {
      if (attempt < 3) {
        await new Promise(r => setTimeout(r, 400 * attempt));
        continue;
      }
      throw new Error(`Błąd połączenia z NBP API: ${err.message}`);
    }

    if (resp.ok) {
      const data = await resp.json();
      return data.rates ?? [];
    }
    if (resp.status === 404) return [];
    if ((resp.status === 429 || resp.status >= 500) && attempt < 3) {
      await new Promise(r => setTimeout(r, 400 * attempt));
      continue;
    }
    throw new Error(
      `NBP API zwróciło błąd ${resp.status} dla zakresu ${fromIso}–${toIso}`);
  }
  return [];
}

// ─────────────────────────────────────────────────────────────
// Grouping
// ─────────────────────────────────────────────────────────────

/** Group transactions by sale date (ISO string key). Preserves order. */
function groupBySaleDate(transactions) {
  const groups = new Map();
  for (const t of transactions) {
    const key = parseDateToIso(t['Date Sold']);
    if (!groups.has(key)) groups.set(key, []);
    groups.get(key).push(t);
  }
  return groups;
}

// ─────────────────────────────────────────────────────────────
// HTML rendering
// ─────────────────────────────────────────────────────────────

function renderTransactionsTable(txns, isRsu, saleRate) {
  if (isRsu) {
    const rows = txns.map((t, i) => {
      const proc = t['Total Proceeds'] ?? 0;
      return `<tr>
        <td class="num">${i + 1}</td>
        <td>${esc(t['Symbol'])}</td>
        <td class="num">${fmt(t['Quantity'] ?? 0, 3)}</td>
        <td class="num">${fmt(proc)}</td>
        <td class="num">${fmt(proc * saleRate)}</td>
      </tr>`;
    }).join('');

    const totQty  = txns.reduce((s, t) => s + (t['Quantity']      ?? 0), 0);
    const totProc = txns.reduce((s, t) => s + (t['Total Proceeds'] ?? 0), 0);

    return `<table class="transactions-table">
      <thead><tr>
        <th>Lp</th><th>Symbol</th><th>Ilość</th>
        <th>Przychód USD</th><th>Przychód PLN</th>
      </tr></thead>
      <tbody>${rows}</tbody>
      <tfoot><tr>
        <td></td><td></td>
        <td class="num">${fmt(totQty, 3)}</td>
        <td class="num">${fmt(totProc)}</td>
        <td class="num">${fmt(totProc * saleRate)}</td>
      </tr></tfoot>
    </table>`;
  }

  // ESPP table
  const rows = txns.map((t, i) => {
    const proc = t['Total Proceeds']  ?? 0;
    const acq  = t['Acquisition Cost'] ?? 0;
    return `<tr>
      <td class="num">${i + 1}</td>
      <td>${esc(t['Symbol'])}</td>
      <td class="num">${fmt(t['Quantity'] ?? 0, 3)}</td>
      <td class="num">${fmt(proc)}</td>
      <td class="num">${fmt(acq)}</td>
      <td class="num">${fmt(proc - acq)}</td>
    </tr>`;
  }).join('');

  const totQty  = txns.reduce((s, t) => s + (t['Quantity']        ?? 0), 0);
  const totProc = txns.reduce((s, t) => s + (t['Total Proceeds']  ?? 0), 0);
  const totCost = txns.reduce((s, t) => s + (t['Acquisition Cost'] ?? 0), 0);

  return `<table class="transactions-table">
    <thead><tr>
      <th>Lp</th><th>Symbol</th><th>Ilość</th>
      <th>Przychód USD</th><th>Koszt USD</th><th>Zysk USD</th>
    </tr></thead>
    <tbody>${rows}</tbody>
    <tfoot><tr>
      <td></td><td></td>
      <td class="num">${fmt(totQty, 3)}</td>
      <td class="num">${fmt(totProc)}</td>
      <td class="num">${fmt(totCost)}</td>
      <td class="num">${fmt(totProc - totCost)}</td>
    </tr></tfoot>
  </table>`;
}

function renderSaleGroup(dateSoldIso, txns, nbpRates) {
  const { rate: saleRate, effectiveDate: saleRateDate } =
    nbpRates.get(dateSoldIso);

  const isEspp  = txns.every(t => t['Plan Type'] === 'ESPP');
  const isRsu   = txns.every(t => t['Plan Type'] === 'RS');
  const planTypes = [
    ...new Set(
      txns.map(t => (t['Plan Type'] === 'RS' ? 'RSU' : (t['Plan Type'] ?? '?')))
    ),
  ].sort().join(', ');
  const typeClass = isRsu ? 'rsu' : isEspp ? 'espp' : 'mixed';

  const totalProceeds    = txns.reduce((s, t) => s + (t['Total Proceeds'] ?? 0), 0);
  const totalProceedsPln = totalProceeds * saleRate;
  const netTransfer      = totalProceeds - WIRE_FEE_USD;
  const netTransferPln   = netTransfer * saleRate;
  const pricePerShare    = txns[0]['Proceeds Per Share'] ?? 0;

  let packageCostUsd = 0;
  let packageCostPln = 0;
  let costRate       = null;
  let costRateDate   = null;

  if (isEspp) {
    packageCostUsd = txns.reduce((s, t) => s + (t['Acquisition Cost'] ?? 0), 0);
    const purchaseIso = parseDateToIso(txns[0]['Purchase Date']);
    ({ rate: costRate, effectiveDate: costRateDate } = nbpRates.get(purchaseIso));
    packageCostPln = packageCostUsd * costRate;
  }

  const taxableUsd = totalProceeds  - packageCostUsd;
  const taxablePln = totalProceedsPln - packageCostPln;

  const nbpCostRow = isEspp
    ? `<div class="meta-row">
        <span class="meta-label">Kurs NBP koszt (${esc(costRateDate)}):</span>
        <span class="meta-value">${fmt(costRate, 4)} PLN/USD</span>
       </div>`
    : '';

  const esppCostSummaryRow = isEspp
    ? `<tr>
        <td>Koszt ESPP</td>
        <td class="num">${fmt(packageCostUsd)}</td>
        <td class="num">${fmt(packageCostPln)}</td>
        <td class="note">(kurs z ${esc(costRateDate)})</td>
       </tr>`
    : '';

  return `
  <div class="sale-group ${typeClass}">
    <div class="sale-header">
      <div class="meta-row">
        <span class="meta-label">Data sprzedaży:</span>
        <span class="meta-value">${esc(formatDate(dateSoldIso))}</span>
      </div>
      <div class="meta-row">
        <span class="meta-label">Typ:</span>
        <span class="meta-value">
          <span class="type-badge ${typeClass}">${esc(planTypes)}</span>
        </span>
      </div>
      <div class="meta-row">
        <span class="meta-label">Cena akcji:</span>
        <span class="meta-value">${fmt(pricePerShare)} USD</span>
      </div>
      <div class="meta-row">
        <span class="meta-label">Kurs NBP przychód (${esc(saleRateDate)}):</span>
        <span class="meta-value">${fmt(saleRate, 4)} PLN/USD</span>
      </div>
      ${nbpCostRow}
      <div class="meta-row">
        <span class="meta-label">Liczba transakcji:</span>
        <span class="meta-value">${txns.length}</span>
      </div>
    </div>

    <div class="table-wrap">
      ${renderTransactionsTable(txns, isRsu, saleRate)}
    </div>

    <div class="group-summary">
      <table class="summary-subtable">
        <thead>
          <tr><th></th><th>USD</th><th>PLN</th><th></th></tr>
        </thead>
        <tbody>
          <tr>
            <td>Przychód (podat.)</td>
            <td class="num">${fmt(totalProceeds)}</td>
            <td class="num">${fmt(totalProceedsPln)}</td>
            <td></td>
          </tr>
          <tr class="muted-row">
            <td>Opłata za przelew</td>
            <td class="num">${fmt(WIRE_FEE_USD)}</td>
            <td class="num">${fmt(WIRE_FEE_USD * saleRate)}</td>
            <td class="note">(informacyjnie)</td>
          </tr>
          <tr class="muted-row">
            <td>Przelew na konto</td>
            <td class="num">${fmt(netTransfer)}</td>
            <td class="num">${fmt(netTransferPln)}</td>
            <td class="note">(informacyjnie)</td>
          </tr>
          ${esppCostSummaryRow}
          <tr class="taxable-row">
            <td><strong>Do opodatkowania</strong></td>
            <td class="num"><strong>${fmt(taxableUsd)}</strong></td>
            <td class="num"><strong>${fmt(taxablePln)}</strong></td>
            <td></td>
          </tr>
        </tbody>
      </table>
    </div>
  </div>`;
}

function renderResults(transactions, nbpRates) {
  const groups = groupBySaleDate(transactions);

  let grandRsuUsd      = 0, grandRsuPln      = 0;
  let grandEsppProcUsd = 0, grandEsppProcPln = 0;
  let grandEsppCostUsd = 0, grandEsppCostPln = 0;

  const groupHtmlParts = [];

  for (const [dateSoldIso, txns] of groups) {
    const { rate: saleRate } = nbpRates.get(dateSoldIso);

    // Accumulate per-transaction to handle mixed RSU+ESPP groups correctly
    for (const t of txns) {
      const proceeds    = t['Total Proceeds'] ?? 0;
      const proceedsPln = proceeds * saleRate;

      if (t['Plan Type'] === 'ESPP') {
        grandEsppProcUsd += proceeds;
        grandEsppProcPln += proceedsPln;
        const costUsd     = t['Acquisition Cost'] ?? 0;
        const purchaseIso = parseDateToIso(t['Purchase Date']);
        const { rate: costRate } = nbpRates.get(purchaseIso);
        grandEsppCostUsd += costUsd;
        grandEsppCostPln += costUsd * costRate;
      } else {
        grandRsuUsd += proceeds;
        grandRsuPln += proceedsPln;
      }
    }

    groupHtmlParts.push(renderSaleGroup(dateSoldIso, txns, nbpRates));
  }

  const grandProcUsd  = grandRsuUsd + grandEsppProcUsd;
  const grandProcPln  = grandRsuPln + grandEsppProcPln;
  const grandTaxUsd   = grandProcUsd - grandEsppCostUsd;
  const grandTaxPln   = grandProcPln - grandEsppCostPln;
  const taxPln        = Math.round(grandTaxPln * 0.19 * 100) / 100;

  const annualHtml = `
  <div class="annual-summary">
    <div class="annual-title">Podsumowanie roczne (do PIT-38)</div>
    <table class="annual-table">
      <thead>
        <tr><th></th><th>USD</th><th>PLN</th><th></th></tr>
      </thead>
      <tbody>
        <tr>
          <td>Przychód RSU</td>
          <td class="num">${fmt(grandRsuUsd)}</td>
          <td class="num">${fmt(grandRsuPln)}</td>
          <td></td>
        </tr>
        <tr>
          <td>Przychód ESPP</td>
          <td class="num">${fmt(grandEsppProcUsd)}</td>
          <td class="num">${fmt(grandEsppProcPln)}</td>
          <td></td>
        </tr>
        <tr>
          <td>Koszt uzyskania (ESPP)</td>
          <td class="num">${fmt(grandEsppCostUsd)}</td>
          <td class="num">${fmt(grandEsppCostPln)}</td>
          <td></td>
        </tr>
        <tr class="separator-row"><td colspan="4"></td></tr>
        <tr class="highlight-row">
          <td>Przychód łączny</td>
          <td class="num">${fmt(grandProcUsd)}</td>
          <td class="num">${fmt(grandProcPln)}</td>
          <td class="note">← poz. 22 PIT-38</td>
        </tr>
        <tr class="highlight-row">
          <td>Koszty łączne</td>
          <td class="num">${fmt(grandEsppCostUsd)}</td>
          <td class="num">${fmt(grandEsppCostPln)}</td>
          <td class="note">← poz. 23 PIT-38</td>
        </tr>
        <tr class="highlight-row">
          <td>Do opodatkowania</td>
          <td class="num">${fmt(grandTaxUsd)}</td>
          <td class="num">${fmt(grandTaxPln)}</td>
          <td class="note">← poz. 29 PIT-ZG</td>
        </tr>
        <tr class="tax-final-row">
          <td>Podatek (19%)</td>
          <td></td>
          <td class="num tax-amount">${fmt(taxPln)} PLN</td>
          <td></td>
        </tr>
      </tbody>
    </table>
  </div>`;

  const esppProfitUsd = grandEsppProcUsd - grandEsppCostUsd;
  const esppProfitPln = grandEsppProcPln - grandEsppCostPln;
  const esppInfoHtml  = esppProfitUsd > 0 ? `
  <div class="espp-info">
    <p>💡 Dzięki uczestnictwie w programie ESPP zarobiłeś w tym roku:</p>
    <p class="profit-figure">
      ${fmt(esppProfitUsd)} USD = ${fmt(esppProfitPln)} PLN
      <span class="muted-text">(przed podatkiem)</span>
    </p>
    <p class="profit-figure">
      ${fmt(esppProfitPln - esppProfitPln * 0.19)} PLN
      <span class="muted-text">(po podatku 19%)</span>
    </p>
  </div>` : '';

  document.getElementById('results').innerHTML =
    groupHtmlParts.join('') + annualHtml + esppInfoHtml;
}

// ─────────────────────────────────────────────────────────────
// XTB rendering
// ─────────────────────────────────────────────────────────────

function renderXtbPositionsTable(positions, rates, currency) {
  const showPln = currency !== 'PLN';

  const rows = positions.map((p, i) => {
    const costRate = xtbRate(rates, currency, p.costIso).rate;
    const saleRate = xtbRate(rates, currency, p.saleIso).rate;
    return `<tr>
      <td class="num">${i + 1}</td>
      <td>${esc(p.instrument)}</td>
      <td>${esc(p.ticker)}</td>
      <td class="num">${fmt(p.volume, 3)}</td>
      <td>${esc(formatDate(p.costIso))}</td>
      ${showPln ? `<td class="num">${fmt(costRate, 4)}</td>` : ''}
      <td class="num">${fmt(p.costValue)}</td>
      ${showPln ? `<td class="num">${fmt(p.costValue * costRate)}</td>` : ''}
      <td class="num">${fmt(p.revenueValue)}</td>
      ${showPln ? `<td class="num">${fmt(p.revenueValue * saleRate)}</td>` : ''}
    </tr>`;
  }).join('');

  const totVolume  = positions.reduce((s, p) => s + p.volume, 0);
  const totCost    = positions.reduce((s, p) => s + p.costValue, 0);
  const totRevenue = positions.reduce((s, p) => s + p.revenueValue, 0);
  const totCostPln = positions.reduce(
    (s, p) => s + p.costValue * xtbRate(rates, currency, p.costIso).rate, 0);
  const totRevenuePln = positions.reduce(
    (s, p) => s + p.revenueValue * xtbRate(rates, currency, p.saleIso).rate, 0);

  return `<table class="transactions-table">
    <thead><tr>
      <th>Lp</th><th>Instrument</th><th>Ticker</th><th>Ilość</th>
      <th>Data kosztu</th>
      ${showPln ? '<th>Kurs kosztu</th>' : ''}
      <th>Koszt ${esc(currency)}</th>
      ${showPln ? '<th>Koszt PLN</th>' : ''}
      <th>Przychód ${esc(currency)}</th>
      ${showPln ? '<th>Przychód PLN</th>' : ''}
    </tr></thead>
    <tbody>${rows}</tbody>
    <tfoot><tr>
      <td></td><td></td><td></td>
      <td class="num">${fmt(totVolume, 3)}</td>
      <td></td>
      ${showPln ? '<td></td>' : ''}
      <td class="num">${fmt(totCost)}</td>
      ${showPln ? `<td class="num">${fmt(totCostPln)}</td>` : ''}
      <td class="num">${fmt(totRevenue)}</td>
      ${showPln ? `<td class="num">${fmt(totRevenuePln)}</td>` : ''}
    </tr></tfoot>
  </table>`;
}

function renderXtbSaleGroup(saleIso, positions, rates, currency) {
  const showPln = currency !== 'PLN';
  const { rate: saleRate, effectiveDate: saleRateDate } =
    xtbRate(rates, currency, saleIso);

  let revenue = 0, revenuePln = 0, cost = 0, costPln = 0;
  for (const p of positions) {
    const sr = xtbRate(rates, currency, p.saleIso).rate;
    const cr = xtbRate(rates, currency, p.costIso).rate;
    revenue += p.revenueValue;
    revenuePln += p.revenueValue * sr;
    cost += p.costValue;
    costPln += p.costValue * cr;
  }

  const categories = [...new Set(positions.map(p => p.category).filter(Boolean))]
    .sort().join(', ') || '?';

  const nbpSaleRow = showPln
    ? `<div class="meta-row">
        <span class="meta-label">Kurs NBP przychód (${esc(saleRateDate)}):</span>
        <span class="meta-value">${fmt(saleRate, 4)} PLN/${esc(currency)}</span>
       </div>`
    : '';

  return `
  <div class="sale-group xtb">
    <div class="sale-header">
      <div class="meta-row">
        <span class="meta-label">Data sprzedaży:</span>
        <span class="meta-value">${esc(formatDate(saleIso))}</span>
      </div>
      <div class="meta-row">
        <span class="meta-label">Typ:</span>
        <span class="meta-value">
          <span class="type-badge xtb">${esc(categories)}</span>
        </span>
      </div>
      ${nbpSaleRow}
      <div class="meta-row">
        <span class="meta-label">Liczba pozycji:</span>
        <span class="meta-value">${positions.length}</span>
      </div>
    </div>

    <div class="table-wrap">
      ${renderXtbPositionsTable(positions, rates, currency)}
    </div>

    <div class="group-summary">
      <table class="summary-subtable">
        <thead>
          <tr><th></th><th>${esc(currency)}</th>${showPln ? '<th>PLN</th>' : ''}<th></th></tr>
        </thead>
        <tbody>
          <tr>
            <td>Przychód (podatk.)</td>
            <td class="num">${fmt(revenue)}</td>
            ${showPln ? `<td class="num">${fmt(revenuePln)}</td>` : ''}
            <td></td>
          </tr>
          <tr>
            <td>Koszty uzyskania</td>
            <td class="num">${fmt(cost)}</td>
            ${showPln ? `<td class="num">${fmt(costPln)}</td>` : ''}
            <td class="note">(kursy NBP z dni poprzedzających zakup)</td>
          </tr>
          <tr class="taxable-row">
            <td><strong>Do opodatkowania</strong></td>
            <td class="num"><strong>${fmt(revenue - cost)}</strong></td>
            ${showPln ? `<td class="num"><strong>${fmt(revenuePln - costPln)}</strong></td>` : ''}
            <td></td>
          </tr>
        </tbody>
      </table>
    </div>
  </div>`;
}

function renderXtbAnnualSummaries(years, currency) {
  const showPln = currency !== 'PLN';

  return years.map(y => {
    const lossRow = y.taxablePln <= 0
      ? `<tr><td colspan="${showPln ? 4 : 3}" class="note">
           Strata w tym roku — nie powstaje obowiązek podatkowy.
         </td></tr>`
      : '';

    return `
    <div class="annual-summary">
      <div class="annual-title">Podsumowanie roczne ${esc(y.year)} (do PIT-38)</div>
      <table class="annual-table">
        <thead>
          <tr><th></th><th>${esc(currency)}</th>${showPln ? '<th>PLN</th>' : ''}<th></th></tr>
        </thead>
        <tbody>
          <tr>
            <td>Przychód (sprzedaż)</td>
            <td class="num">${fmt(y.revenue)}</td>
            ${showPln ? `<td class="num">${fmt(y.revenuePln)}</td>` : ''}
            <td></td>
          </tr>
          <tr>
            <td>Koszty uzyskania</td>
            <td class="num">${fmt(y.cost)}</td>
            ${showPln ? `<td class="num">${fmt(y.costPln)}</td>` : ''}
            <td></td>
          </tr>
          <tr class="separator-row"><td colspan="${showPln ? 4 : 3}"></td></tr>
          <tr class="highlight-row">
            <td>Przychód łączny</td>
            <td class="num">${fmt(y.revenue)}</td>
            ${showPln ? `<td class="num">${fmt(y.revenuePln)}</td>` : ''}
            <td class="note">← poz. 22 PIT-38</td>
          </tr>
          <tr class="highlight-row">
            <td>Koszty łączne</td>
            <td class="num">${fmt(y.cost)}</td>
            ${showPln ? `<td class="num">${fmt(y.costPln)}</td>` : ''}
            <td class="note">← poz. 23 PIT-38</td>
          </tr>
          <tr class="highlight-row">
            <td>Do opodatkowania</td>
            <td class="num">${fmt(y.taxable)}</td>
            ${showPln ? `<td class="num">${fmt(y.taxablePln)}</td>` : ''}
            <td class="note">← poz. 29 PIT-ZG</td>
          </tr>
          ${lossRow}
          <tr class="tax-final-row">
            <td>Podatek (19%)</td>
            <td></td>
            ${showPln ? '<td></td>' : ''}
            <td class="num tax-amount">${fmt(y.taxPln)} PLN</td>
          </tr>
        </tbody>
      </table>
    </div>`;
  }).join('');
}

function renderXtbInterestInfo(interestOps, interestTaxOps, rates, currency) {
  if (interestOps.length === 0 && interestTaxOps.length === 0) return '';

  const showPln = currency !== 'PLN';
  const byYear = new Map();
  const add = (op, key) => {
    if (!op.iso) return;
    const year = op.iso.slice(0, 4);
    if (!byYear.has(year)) {
      byYear.set(year, { interest: 0, interestPln: 0, tax: 0, taxPln: 0 });
    }
    const y = byYear.get(year);
    const rate = xtbRate(rates, currency, op.iso).rate;
    y[key] += op.amount;
    y[key === 'tax' ? 'taxPln' : 'interestPln'] += op.amount * rate;
  };
  interestOps.forEach(op => add(op, 'interest'));
  interestTaxOps.forEach(op => add(op, 'tax'));

  const rows = [...byYear.entries()]
    .sort((a, b) => a[0].localeCompare(b[0]))
    .map(([year, y]) => `<tr>
      <td>${esc(year)}</td>
      <td class="num">${fmt(y.interest)}</td>
      ${showPln ? `<td class="num">${fmt(y.interestPln)}</td>` : ''}
      <td class="num">${fmt(Math.abs(y.tax))}</td>
      ${showPln ? `<td class="num">${fmt(Math.abs(y.taxPln))}</td>` : ''}
    </tr>`).join('');

  return `
  <div class="xtb-info">
    <p>ℹ️ Odsetki od wolnych środków
       <span class="muted-text">(informacyjnie — XTB pobrało już od nich podatek 19%)</span></p>
    <table class="summary-subtable">
      <thead>
        <tr>
          <th>Rok</th><th>Odsetki ${esc(currency)}</th>
          ${showPln ? '<th>Odsetki PLN</th>' : ''}
          <th>Podatek pobrany ${esc(currency)}</th>
          ${showPln ? '<th>Podatek pobrany PLN</th>' : ''}
        </tr>
      </thead>
      <tbody>${rows}</tbody>
    </table>
    <p class="note">Odsetki nie zostały uwzględnione w rozliczeniu podatku powyżej.</p>
  </div>`;
}

function renderXtbResults(report, rates) {
  const { currency, positions, interestOps, interestTaxOps } = report;
  const groups = xtbGroupBySaleDate(positions);

  const groupHtml = [...groups.entries()]
    .map(([saleIso, ps]) => renderXtbSaleGroup(saleIso, ps, rates, currency))
    .join('');

  const annualHtml = renderXtbAnnualSummaries(
    xtbYearSummaries(positions, rates, currency), currency);

  const interestHtml = renderXtbInterestInfo(
    interestOps, interestTaxOps, rates, currency);

  document.getElementById('results').innerHTML =
    groupHtml + annualHtml + interestHtml;

  return groups;
}

// ─────────────────────────────────────────────────────────────
// App state & event handling
// ─────────────────────────────────────────────────────────────

function resetApp() {
  document.getElementById('uploadSection').classList.remove('hidden');
  document.getElementById('loadingSection').classList.add('hidden');
  document.getElementById('errorSection').classList.add('hidden');
  document.getElementById('resultsSection').classList.add('hidden');
  document.getElementById('fileInput').value = '';
  document.getElementById('results').innerHTML = '';
  document.getElementById('nbpLog').innerHTML = '';
}

function showError(msg) {
  document.getElementById('loadingSection').classList.add('hidden');
  document.getElementById('errorSection').classList.remove('hidden');
  document.getElementById('errorBox').textContent = msg;
}

/** Full processing flow for an XTB account statement. */
async function processXtbFile(file, arrayBuffer) {
  let report;
  try {
    report = readXtbReport(arrayBuffer, file.name);
  } catch (err) {
    showError(`Błąd podczas parsowania pliku XTB: ${err.message}`);
    return;
  }

  const { currency, positions, interestOps, interestTaxOps } = report;

  if (positions.length === 0) {
    showError(
      'Nie znaleziono zamkniętych pozycji (arkusz "Closed Positions") w pliku. ' +
      'Upewnij się, że wczytałeś raport z XTB.',
    );
    return;
  }

  // Collect all ISO dates that need NBP rates
  const allIsoDates = [];
  for (const p of positions) {
    allIsoDates.push(p.saleIso, p.costIso);
  }
  for (const op of [...interestOps, ...interestTaxOps]) {
    if (op.iso) allIsoDates.push(op.iso);
  }

  let nbpRates = null;
  if (currency !== 'PLN') {
    const uniqueCount = new Set(allIsoDates).size;
    document.getElementById('loadingText').textContent =
      `Pobieranie kursów NBP (${uniqueCount} ${uniqueCount === 1 ? 'data' : 'daty/dat'})…`;

    const nbpLog = document.getElementById('nbpLog');
    try {
      nbpRates = await fetchNbpRates(
        allIsoDates,
        (iso, { rate, effectiveDate }) => {
          const div = document.createElement('div');
          div.className = 'nbp-log-entry';
          div.textContent =
            `Kurs NBP ${currency}/PLN na ${effectiveDate}: ${fmt(rate, 4)}` +
            ` (dla daty ${formatDate(iso)})`;
          nbpLog.appendChild(div);
        },
        currency,
      );
    } catch (err) {
      showError(`Błąd NBP API: ${err.message}`);
      return;
    }
  }

  let groups;
  try {
    groups = renderXtbResults(report, nbpRates);
  } catch (err) {
    showError(`Błąd podczas obliczania: ${err.message}`);
    return;
  }

  document.getElementById('resultsFilename').textContent =
    `${file.name} · XTB · ${positions.length} pozycji · ${groups.size} dni sprzedaży`;

  document.getElementById('loadingSection').classList.add('hidden');
  document.getElementById('resultsSection').classList.remove('hidden');
}

async function processFile(file) {
  if (!file.name.toLowerCase().endsWith('.xlsx')) {
    showError('Wybierz plik w formacie .xlsx (Gains & Losses z E*TRADE).');
    return;
  }

  document.getElementById('uploadSection').classList.add('hidden');
  document.getElementById('loadingSection').classList.remove('hidden');
  document.getElementById('loadingText').textContent = 'Wczytywanie pliku…';
  document.getElementById('nbpLog').innerHTML = '';

  let arrayBuffer;
  try {
    arrayBuffer = await file.arrayBuffer();
  } catch (err) {
    showError(`Nie udało się wczytać pliku: ${err.message}`);
    return;
  }

  // Auto-detect the report type: XTB account statement vs E*TRADE Gains & Losses.
  let isXtb;
  try {
    const probe = XLSX.read(new Uint8Array(arrayBuffer), {
      type: 'array',
      cellDates: false,
    });
    isXtb = isXtbWorkbook(probe);
  } catch (err) {
    showError(`Błąd podczas parsowania pliku XLSX: ${err.message}`);
    return;
  }

  if (isXtb) {
    await processXtbFile(file, arrayBuffer);
    return;
  }

  let transactions;
  try {
    transactions = readTransactions(arrayBuffer);
  } catch (err) {
    showError(`Błąd podczas parsowania pliku XLSX: ${err.message}`);
    return;
  }

  if (transactions.length === 0) {
    showError(
      'Nie znaleziono transakcji sprzedaży (Record Type = "Sell") w pliku. ' +
      'Upewnij się, że wczytałeś plik Gains & Losses z E*TRADE.',
    );
    return;
  }

  // Collect all ISO dates that need NBP rates
  const allIsoDates = [];
  for (const t of transactions) {
    const saleIso = parseDateToIso(t['Date Sold']);
    if (saleIso) allIsoDates.push(saleIso);

    if (
      t['Plan Type'] === 'ESPP' &&
      t['Purchase Date'] &&
      t['Purchase Date'] !== '--'
    ) {
      const purchaseIso = parseDateToIso(t['Purchase Date']);
      if (purchaseIso) allIsoDates.push(purchaseIso);
    }
  }

  const uniqueCount = new Set(allIsoDates).size;
  document.getElementById('loadingText').textContent =
    `Pobieranie kursów NBP (${uniqueCount} ${uniqueCount === 1 ? 'data' : 'daty/dat'})…`;

  const nbpLog = document.getElementById('nbpLog');
  let nbpRates;
  try {
    nbpRates = await fetchNbpRates(allIsoDates, (iso, { rate, effectiveDate }) => {
      const div = document.createElement('div');
      div.className = 'nbp-log-entry';
      div.textContent =
        `Kurs NBP USD/PLN na ${effectiveDate}: ${fmt(rate, 4)}` +
        ` (dla daty ${formatDate(iso)})`;
      nbpLog.appendChild(div);
    });
  } catch (err) {
    showError(`Błąd NBP API: ${err.message}`);
    return;
  }

  try {
    const groups = groupBySaleDate(transactions);
    document.getElementById('resultsFilename').textContent =
      `${file.name} · ${transactions.length} transakcji · ${groups.size} pakietów`;
    renderResults(transactions, nbpRates);
  } catch (err) {
    showError(`Błąd podczas obliczania: ${err.message}`);
    return;
  }

  document.getElementById('loadingSection').classList.add('hidden');
  document.getElementById('resultsSection').classList.remove('hidden');
}

document.addEventListener('DOMContentLoaded', () => {
  const fileInput  = document.getElementById('fileInput');
  const uploadZone = document.getElementById('uploadZone');

  fileInput.addEventListener('change', e => {
    const file = e.target.files[0];
    if (file) processFile(file);
  });

  uploadZone.addEventListener('dragover', e => {
    e.preventDefault();
    uploadZone.classList.add('drag-over');
  });

  uploadZone.addEventListener('dragleave', () => {
    uploadZone.classList.remove('drag-over');
  });

  uploadZone.addEventListener('drop', e => {
    e.preventDefault();
    uploadZone.classList.remove('drag-over');
    const file = e.dataTransfer.files[0];
    if (file) processFile(file);
  });

  // Allow clicking anywhere on the upload zone to trigger file picker.
  // The <label for="fileInput"> already opens the dialog natively,
  // so we must skip it (and the input itself) to avoid double-open.
  uploadZone.addEventListener('click', e => {
    if (e.target === fileInput || e.target.closest('label[for="fileInput"]')) return;
    fileInput.click();
  });
});
