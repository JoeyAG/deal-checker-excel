/*
 * core.js — the Deal Checker add-in's logic, with no Excel or Office code in it,
 * so it can be tested on its own. taskpane.js connects it to the workbook.
 */
var DealCore = (function () {
  'use strict';

  const PREFIX = 'DEALCHECKER:v1:';
  const NEW_COLS = ['Delivery Date', 'eBay Order Line'];
  const TEST_DAYS = 30;
  const PART_CATEGORIES = new Set(['CPU', 'GPU', 'RAM', 'Storage', 'PSU', 'Motherboard', 'Cooling', 'Case']);

  // ---------------------------------------------------------------------------
  // Small helpers
  // ---------------------------------------------------------------------------

  const norm = s => String(s == null ? '' : s).trim().toLowerCase();

  // Excel stores dates as days since 30 Dec 1899.
  function isoToSerial(iso) {
    if (!iso) return '';
    const t = Date.parse(iso + 'T00:00:00Z');
    return Number.isFinite(t) ? Math.round(t / 864e5) + 25569 : '';
  }
  function serialToIso(n) {
    if (typeof n !== 'number' || !Number.isFinite(n) || n <= 0) return null;
    return new Date(Math.round((n - 25569) * 864e5)).toISOString().slice(0, 10);
  }
  function todaySerial(now) {
    const d = now ? new Date(now) : new Date();
    return Math.floor(Date.UTC(d.getFullYear(), d.getMonth(), d.getDate()) / 864e5) + 25569;
  }

  function colIndex(headers) {
    const map = {};
    headers.forEach((h, i) => { map[norm(h)] = i; });
    return name => (norm(name) in map ? map[norm(name)] : -1);
  }

  // Rows as objects keyed by header name, keeping their position in the table.
  function toObjects(headers, values) {
    return values.map((row, i) => {
      const o = { _row: i };
      headers.forEach((h, j) => { o[String(h).trim()] = row[j]; });
      return o;
    });
  }

  // ---------------------------------------------------------------------------
  // Payload from the Chrome extension
  // ---------------------------------------------------------------------------

  function parsePayload(text) {
    const t = String(text || '').trim();
    const at = t.indexOf(PREFIX);
    if (at < 0) throw new Error('That isn’t something copied by the Deal Checker extension. In Chrome, click “Copy rows for Excel” on your eBay Purchases page first.');
    let data;
    try { data = JSON.parse(t.slice(at + PREFIX.length)); } catch (e) {
      throw new Error('The copied text is incomplete. Copy it again from Chrome, then paste.');
    }
    return { rows: Array.isArray(data.rows) ? data.rows : [], watchlist: Array.isArray(data.watchlist) ? data.watchlist : [], exportedAt: data.exportedAt };
  }

  // ---------------------------------------------------------------------------
  // The Purchases table
  // ---------------------------------------------------------------------------

  // A column is "calculated" when every row holds the same formula (in R1C1 form).
  // Returns, per column, { a1, r1c1 } or null.
  function detectCalculated(formulasA1, formulasR1C1) {
    if (!formulasR1C1.length) return [];
    const cols = formulasR1C1[0].length;
    const out = [];
    for (let c = 0; c < cols; c++) {
      const first = formulasR1C1[0][c];
      let calc = typeof first === 'string' && first.startsWith('=');
      for (let r = 1; calc && r < formulasR1C1.length; r++) if (formulasR1C1[r][c] !== first) calc = false;
      out.push(calc && formulasR1C1.length > 1 ? { a1: formulasA1[0][c], r1c1: first } : null);
    }
    return out;
  }

  // A formula that's identical on every row in A1 form (structured references like
  // Purchases[[#This Row],[Cost]]) can be written as-is into new rows.
  const rowIndependent = f => !/(^|[^A-Za-z0-9_\]])\$?[A-Z]{1,3}\$?\d+(?![\w(])/.test(String(f).replace(/"[^"]*"/g, '""'));

  function nextPartNumber(partIds) {
    let max = 0;
    partIds.forEach(id => {
      const m = String(id || '').trim().match(/^P(\d+)$/i);
      if (m) max = Math.max(max, +m[1]);
    });
    return max + 1;
  }

  function missingColumns(headers) {
    const idx = colIndex(headers);
    return NEW_COLS.filter(c => idx(c) < 0);
  }

  // "Days left to Test", counted from delivery as eBay does (falls back to purchase date).
  function daysLeftFormula(tableName) {
    const r = c => `${tableName}[[#This Row],[${c}]]`;
    return `=IF(AND(${r('Tested?')}="N",${r('Purchased From')}="eBay"),${TEST_DAYS}-(TODAY()-IF(ISNUMBER(${r('Delivery Date')}),${r('Delivery Date')},${r('Purchase Date')})),"")`;
  }

  /*
   * Works out what to write for an import.
   *   table: { name, headers, values, formulasA1, formulasR1C1 }
   *   rows:  rows from the extension
   * Returns { add: 2D values for new rows, patchR1C1: [{col, r1c1}], updates: [{row, col, value}],
   *           skipped, preview: [...] }
   */
  function planImport(table, rows) {
    const idx = colIndex(table.headers);
    const calc = detectCalculated(table.formulasA1, table.formulasR1C1);
    const keyCol = idx('eBay Order Line');
    const delCol = idx('Delivery Date');
    const existing = new Map();
    if (keyCol >= 0) table.values.forEach((r, i) => { const k = String(r[keyCol] || '').trim(); if (k) existing.set(k, i); });

    let next = nextPartNumber(table.values.map(r => r[idx('Part ID')]));
    const add = [];
    const preview = [];
    const updates = [];
    let skipped = 0;
    const seen = new Set();

    for (const r of rows) {
      if (!r || !r.key || seen.has(r.key)) continue;
      seen.add(r.key);
      if (existing.has(r.key)) {
        // Already imported: just record a delivery date that's now known (or confirmed).
        const i = existing.get(r.key);
        const newDel = isoToSerial(r.deliveryDate);
        if (delCol >= 0 && newDel && table.values[i][delCol] !== newDel && !r.deliveryEstimated) {
          updates.push({ row: i, col: delCol, value: newDel, key: r.key });
        }
        skipped++;
        continue;
      }
      const partId = 'P' + next++;
      const fields = {
        'part id': partId,
        'brand': r.brand || 'No Brand',
        'description': r.description || r.title || '',
        'category': r.category || 'Other',
        'other notes': r.notes || '',
        'cost': typeof r.cost === 'number' ? r.cost : '',
        'collection postcode': '',
        'miles': 'N/A',
        'cost per mile': '',
        'purchase date': isoToSerial(r.purchaseDate),
        'purchased from': 'eBay',
        'estimated part value at purchase date': typeof r.marketValue === 'number' ? r.marketValue : '',
        'days in stock': '',
        'available?': 'Y',
        'allocation': '',
        'tested?': r.tested == null ? '' : r.tested,
        'sold?': '',
        'delivery date': isoToSerial(r.deliveryDate),
        'ebay order line': r.key
      };
      add.push(table.headers.map((h, c) => {
        if (calc[c]) return rowIndependent(calc[c].a1) ? calc[c].a1 : '';
        const v = fields[norm(h)];
        return v === undefined ? '' : v;
      }));
      preview.push({ partId, key: r.key, brand: fields.brand, description: fields.description, category: fields.category,
        cost: r.cost, marketValue: r.marketValue, purchaseDate: r.purchaseDate, deliveryDate: r.deliveryDate,
        deliveryEstimated: !!r.deliveryEstimated, notes: r.notes || '', valueSource: r.valueSource || null });
    }
    const patchR1C1 = calc.map((c, i) => (c && !rowIndependent(c.a1) ? { col: i, r1c1: c.r1c1 } : null)).filter(Boolean);
    return { add, patchR1C1, updates, skipped, preview };
  }

  // ---------------------------------------------------------------------------
  // Dashboard
  // ---------------------------------------------------------------------------

  /*
   * Parts bought on eBay that still need testing, with days left in eBay's 30-day window.
   * estimatedKeys: order lines whose Delivery Date is only an estimate.
   */
  function testDeadlines(purchases, today, estimatedKeys) {
    const est = estimatedKeys || {};
    const out = [];
    for (const p of purchases) {
      if (norm(p['Tested?']) !== 'n' || norm(p['Purchased From']) !== 'ebay') continue;
      const del = typeof p['Delivery Date'] === 'number' ? p['Delivery Date'] : null;
      const bought = typeof p['Purchase Date'] === 'number' ? p['Purchase Date'] : null;
      const base = del || bought;
      if (!base) continue;
      const isEst = !!(p['eBay Order Line'] && est[p['eBay Order Line']]);
      const notArrived = del && del > today;
      const daysLeft = TEST_DAYS - (today - base);
      let status = 'ok';
      if (notArrived) status = 'coming';
      else if (daysLeft < 0) status = 'over';
      else if (daysLeft <= 7) status = 'urgent';
      else if (daysLeft <= 14) status = 'soon';
      out.push({
        row: p._row, partId: p['Part ID'], name: [p['Brand'], p['Description']].filter(x => x && x !== 'No Brand').join(' '),
        category: p['Category'], daysLeft, deadline: serialToIso(base + TEST_DAYS), status, fromDelivery: !!del, estimated: isEst,
        allocation: p['Allocation'] || ''
      });
    }
    const order = { over: 0, urgent: 1, soon: 2, ok: 3, coming: 4 };
    return out.sort((a, b) => order[a.status] - order[b.status] || a.daysLeft - b.daysLeft);
  }

  function num(v) { return typeof v === 'number' && Number.isFinite(v) ? v : null; }

  function buildsSummary(builds, purchases) {
    return builds.filter(b => b['Build ID']).map(b => {
      const id = String(b['Build ID']).trim();
      const parts = purchases.filter(p => String(p['Allocation'] || '').trim() === id);
      const untested = parts.filter(p => norm(p['Tested?']) === 'n');
      const flag = k => norm(b[k]) === 'y';
      const stage = flag('Benchmarked?') ? 'Benchmarked' : flag('Built?') ? 'Built' : flag('Ready to Build?') ? 'Ready to build'
        : flag('All Parts Required to Complete?') ? 'All parts in' : 'Collecting parts';
      return {
        row: b._row, id, description: b['Build Description'] || '', cost: num(b['Build Cost']),
        missing: String(b['Missing Parts'] || '').split(/[,;/]+/).map(s => s.trim()).filter(Boolean),
        waitingOn: b['Waiting on'] || '', stage, parts: parts.length, untested: untested.length,
        partNames: parts.map(p => `${p['Part ID']} ${p['Description'] || ''}`.trim())
      };
    });
  }

  function stockSummary(purchases) {
    const groups = new Map();
    for (const p of purchases) {
      if (norm(p['Available?']) !== 'y' || String(p['Allocation'] || '').trim() || norm(p['Sold?']) === 'y') continue;
      const cat = p['Category'] || 'Other';
      if (!groups.has(cat)) groups.set(cat, { category: cat, count: 0, cost: 0, value: 0, valued: 0 });
      const g = groups.get(cat);
      g.count++;
      g.cost += num(p['Total Cost']) != null ? p['Total Cost'] : (num(p['Cost']) || 0);
      if (num(p['Estimated Part Value at Purchase Date']) != null) { g.value += p['Estimated Part Value at Purchase Date']; g.valued++; }
    }
    return [...groups.values()].sort((a, b) => b.cost - a.cost);
  }

  function missingValues(purchases) {
    return purchases.filter(p => PART_CATEGORIES.has(p['Category']) &&
      (p['Estimated Part Value at Purchase Date'] === '' || p['Estimated Part Value at Purchase Date'] == null));
  }

  // Watchlist sheet rows from the extension's list.
  function watchlistValues(list, now) {
    const header = ['Part', 'Median Sold', 'Sales', 'Alert At', 'Updated', 'Find Deals'];
    const rows = list.map(w => [
      w.label || '',
      typeof w.median === 'number' ? Math.round(w.median * 100) / 100 : '',
      w.n || '',
      typeof w.alertAt === 'number' ? w.alertAt : '',
      w.updated ? isoToSerial(new Date(w.updated).toISOString().slice(0, 10)) : isoToSerial(new Date(now || Date.now()).toISOString().slice(0, 10)),
      w.dealsUrl ? `=HYPERLINK("${String(w.dealsUrl).replace(/"/g, '""')}","Find deals")` : ''
    ]);
    return { header, rows };
  }

  return {
    PREFIX, NEW_COLS, isoToSerial, serialToIso, todaySerial, toObjects, colIndex, parsePayload,
    detectCalculated, rowIndependent, nextPartNumber, missingColumns, daysLeftFormula, planImport,
    testDeadlines, buildsSummary, stockSummary, missingValues, watchlistValues
  };
})();

if (typeof module !== 'undefined' && module.exports) module.exports = DealCore;
