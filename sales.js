/*
 * sales.js — reads eBay Seller Hub report files and works out what to write in Sales.
 *
 * Two eBay reports work, and you can use either or both:
 *   - Orders report   (Seller Hub > Reports > Downloads > Orders, CSV)
 *   - Transaction report (Seller Hub > Payments > Reports > Transaction report, CSV)
 * The transaction report also carries final value fees, promoted listing fees and
 * postage labels bought through eBay, so profit comes out right.
 *
 * Headers are matched loosely (eBay changes them now and then), and any lines above
 * the header row are skipped. No Excel code here.
 */
var DealSales = (function () {
  'use strict';

  const norm = s => String(s == null ? '' : s).trim().toLowerCase();
  const key = s => norm(s).replace(/[–—]/g, '-').replace(/[‘’“”"']/g, '').replace(/[^a-z0-9]+/g, ' ').trim();

  // ---------------------------------------------------------------------------
  // CSV
  // ---------------------------------------------------------------------------

  function parseCsv(text) {
    const t = String(text || '').replace(/^﻿/, '');
    const rows = [];
    let row = [];
    let cell = '';
    let q = false;
    for (let i = 0; i < t.length; i++) {
      const ch = t[i];
      if (q) {
        if (ch === '"') { if (t[i + 1] === '"') { cell += '"'; i++; } else q = false; } else cell += ch;
      } else if (ch === '"') q = true;
      else if (ch === ',') { row.push(cell); cell = ''; }
      else if (ch === '\n' || ch === '\r') {
        if (ch === '\r' && t[i + 1] === '\n') i++;
        row.push(cell); rows.push(row); row = []; cell = '';
      } else cell += ch;
    }
    if (cell !== '' || row.length) { row.push(cell); rows.push(row); }
    return rows.filter(r => r.some(c => String(c).trim() !== ''));
  }

  // ---------------------------------------------------------------------------
  // Columns we use, with the names eBay has used for them
  // ---------------------------------------------------------------------------

  const COLS = {
    order: ['order number', 'order id', 'order no'],
    legacyOrder: ['legacy order id'],
    item: ['item number', 'item id'],
    title: ['item title', 'title'],
    sku: ['custom label', 'custom label sku', 'sku'],
    qty: ['quantity'],
    soldFor: ['sold for', 'item price'],
    subtotal: ['item subtotal'],
    postage: ['postage and packaging', 'postage packaging', 'shipping and handling', 'shipping handling', 'postage'],
    total: ['total price', 'total'],
    saleDate: ['sale date', 'order creation date', 'paid on date', 'transaction creation date', 'date'],
    type: ['type'],
    description: ['description'],
    net: ['net amount'],
    gross: ['gross transaction amount', 'gross amount'],
    fvfFixed: ['final value fee fixed'],
    fvfVar: ['final value fee variable'],
    regFee: ['regulatory operating fee'],
    inadFee: ['very high item not as described fee'],
    belowFee: ['below standard performance fee'],
    intlFee: ['international fee'],
    promoted: ['sold via promoted listings'],
    vat: ['ebay collected tax', 'ebay collected vat']
  };

  function headerMap(header) {
    const ks = header.map(key);
    const map = {};
    for (const [name, alts] of Object.entries(COLS)) {
      let at = -1;
      for (const a of alts) { at = ks.indexOf(a); if (at >= 0) break; }
      if (at < 0) for (const a of alts) { at = ks.findIndex(k => k.startsWith(a + ' ')); if (at >= 0) break; }
      if (at >= 0) map[name] = at;
    }
    return map;
  }

  function findHeader(rows) {
    for (let i = 0; i < Math.min(rows.length, 30); i++) {
      const ks = rows[i].map(key);
      if (ks.includes('order number') && (ks.includes('item title') || ks.includes('item number') || ks.includes('item id'))) return i;
    }
    return -1;
  }

  // ---------------------------------------------------------------------------
  // Values
  // ---------------------------------------------------------------------------

  // "£1,234.50", "-12.34", "GBP 5.00", "--" -> number (or null)
  function money(v) {
    const s = String(v == null ? '' : v).replace(/[£$€,\s]|GBP/gi, '');
    if (!s || /^-+$/.test(s)) return null;
    const m = s.match(/^\(?(-?\d+(?:\.\d+)?)\)?$/);
    if (!m) return null;
    const n = parseFloat(m[1]);
    return /^\(/.test(s) ? -n : n;
  }

  const MONTHS = { jan: 1, feb: 2, mar: 3, apr: 4, may: 5, jun: 6, jul: 7, aug: 8, sep: 9, sept: 9, oct: 10, nov: 11, dec: 12 };
  const pad = n => String(n).padStart(2, '0');
  const iso = (y, m, d) => (y > 0 && m >= 1 && m <= 12 && d >= 1 && d <= 31 ? `${y}-${pad(m)}-${pad(d)}` : null);
  const year = y => (String(y).length === 2 ? 2000 + +y : +y);

  // eBay UK dates: "08 Oct 2026", "8-Oct-26", "Oct 8, 2026", "2026-10-08", "08/10/2026" (day first).
  function parseDate(v) {
    const s = norm(v);
    if (!s) return null;
    let m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
    if (m) return iso(+m[1], +m[2], +m[3]);
    m = s.match(/^(\d{1,2})[\s-]([a-z]{3,4})[a-z]*[\s-,]+(\d{2,4})/);
    if (m && MONTHS[m[2]]) return iso(year(m[3]), MONTHS[m[2]], +m[1]);
    m = s.match(/^([a-z]{3,4})[a-z]*[\s-](\d{1,2}),?[\s-](\d{2,4})/);
    if (m && MONTHS[m[1]]) return iso(year(m[3]), MONTHS[m[1]], +m[2]);
    m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})/);
    if (m) return iso(year(m[3]), +m[2], +m[1]);
    return null;
  }

  const idFrom = s => {
    const m = String(s || '').trim().match(/^\s*([bp])\s?-?(\d{1,6})\b/i);
    return m ? m[1].toUpperCase() + +m[2] : null;
  };

  // ---------------------------------------------------------------------------
  // Reading a report into sales
  // ---------------------------------------------------------------------------

  /*
   * Returns { kind: 'orders'|'transactions', sales: [...], ignored, problems: [] }
   * Each sale: { order, item, title, sku, id, qty, date, itemPrice, postage, soldPrice,
   *              ebayFees, promoFee, shippingCost, refunded }
   */
  function readReport(text) {
    const rows = parseCsv(text);
    const h = findHeader(rows);
    if (h < 0) throw new Error('This doesn’t look like an eBay Orders or Transaction report. In Seller Hub, download the Orders report (Reports > Downloads) or the Transaction report (Payments > Reports) as CSV.');
    const map = headerMap(rows[h]);
    const kind = map.type != null && (map.net != null || map.gross != null) ? 'transactions' : 'orders';
    const get = (r, k) => (map[k] == null ? '' : r[map[k]]);
    const sales = new Map();
    const fees = [];
    let ignored = 0;

    const saleKey = (order, item) => `${order}|${item || ''}`;

    for (const r of rows.slice(h + 1)) {
      const order = String(get(r, 'order') || '').trim();
      const type = key(get(r, 'type'));
      if (!order || !/\d/.test(order)) { ignored++; continue; }

      if (kind === 'transactions' && type && type !== 'order') {
        const amount = Math.abs(money(get(r, 'net')) ?? money(get(r, 'gross')) ?? 0);
        const desc = key(get(r, 'description'));
        if (/refund/.test(type)) fees.push({ order, item: String(get(r, 'item') || '').trim(), kind: 'refund', amount });
        else if (/shipping label|postage label|label/.test(type) || /label/.test(desc)) fees.push({ order, kind: 'shipping', amount });
        else if (/promot|ad fee|advert/.test(desc) || /promot|ad fee/.test(type)) fees.push({ order, item: String(get(r, 'item') || '').trim(), kind: 'promo', amount });
        else if (/fee/.test(type)) fees.push({ order, item: String(get(r, 'item') || '').trim(), kind: 'fee', amount });
        else ignored++;
        continue;
      }

      const title = String(get(r, 'title') || '').trim();
      if (!title) { ignored++; continue; } // order summary lines in the Orders report
      const item = String(get(r, 'item') || '').trim();
      const qty = parseInt(get(r, 'qty'), 10) || 1;
      const subtotal = money(get(r, 'subtotal'));
      const soldFor = money(get(r, 'soldFor'));
      const itemPrice = subtotal != null ? subtotal : soldFor != null ? soldFor * qty : null;
      const postage = money(get(r, 'postage')) || 0;
      const feeCols = ['fvfFixed', 'fvfVar', 'regFee', 'inadFee', 'belowFee', 'intlFee'].map(c => Math.abs(money(get(r, c)) || 0));
      const sku = String(get(r, 'sku') || '').trim();
      const k = saleKey(order, item);
      const prev = sales.get(k);
      if (prev) { prev.qty += qty; prev.itemPrice = (prev.itemPrice || 0) + (itemPrice || 0); prev.ebayFees += feeCols.reduce((a, b) => a + b, 0); continue; }
      sales.set(k, {
        order, item, title, sku, id: idFrom(sku), qty,
        date: parseDate(get(r, 'saleDate')),
        itemPrice, postage,
        ebayFees: feeCols.reduce((a, b) => a + b, 0),
        promoFee: 0, shippingCost: 0, refunded: false,
        promotedFlag: /^y/i.test(String(get(r, 'promoted') || '')),
        source: kind
      });
    }

    // Fees that sit on their own lines belong to the sale with the same order (and item, if given).
    const list = [...sales.values()];
    fees.forEach(f => {
      const matches = list.filter(s => s.order === f.order && (!f.item || !s.item || s.item === f.item));
      const target = matches[0];
      if (!target) return;
      if (f.kind === 'refund') { matches.forEach(s => { s.refunded = true; }); return; }
      const share = f.amount / (f.item ? 1 : matches.length);
      (f.item ? [target] : matches).forEach(s => {
        if (f.kind === 'shipping') s.shippingCost += share;
        else if (f.kind === 'promo') s.promoFee += share;
        else s.ebayFees += share;
      });
    });
    const round = v => Math.round(v * 100) / 100;
    list.forEach(s => {
      s.soldPrice = s.itemPrice == null ? null : round(s.itemPrice + (s.postage || 0));
      s.ebayFees = round(s.ebayFees);
      s.promoFee = round(s.promoFee);
      s.shippingCost = round(s.shippingCost);
    });
    return { kind, sales: list, ignored, hasFees: kind === 'transactions' };
  }

  // Combine an Orders report and a Transaction report for the same sales.
  function merge(a, b) {
    const out = new Map();
    const k = s => `${s.order}|${s.item}`;
    a.concat(b).forEach(s => {
      const prev = out.get(k(s));
      if (!prev) { out.set(k(s), Object.assign({}, s)); return; }
      ['sku', 'id', 'date', 'title'].forEach(f => { if (!prev[f] && s[f]) prev[f] = s[f]; });
      if (prev.itemPrice == null) prev.itemPrice = s.itemPrice;
      if (prev.soldPrice == null) prev.soldPrice = s.soldPrice;
      ['ebayFees', 'promoFee', 'shippingCost'].forEach(f => { prev[f] = Math.max(prev[f] || 0, s[f] || 0); });
      prev.refunded = prev.refunded || s.refunded;
      if (s.source === 'transactions') prev.source = 'both';
    });
    return [...out.values()];
  }

  // ---------------------------------------------------------------------------
  // Matching sales to your Build / Part IDs
  // ---------------------------------------------------------------------------

  const words = s => new Set(norm(s).replace(/[^a-z0-9]+/g, ' ').split(' ').filter(w => w.length > 1 && !/^(gb|pc|the|and|with|for|gaming|desktop|computer|windows|win|ram|ssd|tb|used)$/.test(w)));

  // Best guess at which Sales row / build a title belongs to.
  function guess(title, candidates) {
    const tw = words(title);
    let best = null;
    candidates.forEach(c => {
      const cw = words(c.text);
      if (!cw.size) return;
      let hit = 0;
      cw.forEach(w => { if (tw.has(w) || [...tw].some(x => x.includes(w) || w.includes(x))) hit++; });
      const score = hit / cw.size;
      if (!best || score > best.score) best = { id: c.id, score };
    });
    return best && best.score >= 0.6 ? best : null;
  }

  /*
   * salesTable: { headers, values } of the Sales table
   * builds / purchases: rows as objects
   * sales: from readReport / merge
   * Returns { lines: [{ sale, id, how, row (Sales row or -1), action: 'add'|'update'|'skip', reason }] }
   */
  function planSales(salesTable, builds, purchases, sales) {
    const h = salesTable.headers.map(norm);
    const col = n => h.indexOf(norm(n));
    const idCol = col('Build / Part ID');
    const orderCol = col('eBay Order');
    const soldCol = col('Sold?');
    const descCol = col('Description');
    const rows = salesTable.values;
    const byId = new Map();
    rows.forEach((r, i) => { const id = idFrom(r[idCol]); if (id && !byId.has(id)) byId.set(id, i); });
    const byOrder = new Map();
    if (orderCol >= 0) rows.forEach((r, i) => String(r[orderCol] || '').split(/[,;\s]+/).filter(Boolean).forEach(o => byOrder.set(o, i)));

    const knownIds = new Set(builds.map(b => idFrom(b['Build ID'])).filter(Boolean).concat(purchases.map(p => idFrom(p['Part ID'])).filter(Boolean)));
    const candidates = rows.map((r, i) => ({ id: idFrom(r[idCol]), text: descCol >= 0 ? String(r[descCol] || '') : '', i }))
      .filter(c => c.id && norm(rows[c.i][soldCol]) !== 'y' && c.text && c.text !== '-');
    builds.forEach(b => { const id = idFrom(b['Build ID']); if (id && !candidates.some(c => c.id === id)) candidates.push({ id, text: String(b['Build Description'] || '') }); });

    const lines = sales.map(sale => {
      if (byOrder.has(sale.order)) {
        return { sale, id: idFrom(rows[byOrder.get(sale.order)][idCol]), how: 'order', row: byOrder.get(sale.order), action: 'update', reason: 'Already recorded; fees and postage will be filled in if missing.' };
      }
      let id = sale.id && knownIds.has(sale.id) ? sale.id : null;
      let how = id ? 'label' : null;
      let score = null;
      if (!id) {
        const g = guess(sale.title, candidates);
        if (g) { id = g.id; how = 'title'; score = g.score; }
      }
      if (!id) return { sale, id: null, how: null, row: -1, action: 'skip', reason: sale.sku ? `Custom label “${sale.sku}” isn’t a Build or Part ID in the sheet.` : 'No custom label, and the title didn’t match a build. Type its Build or Part ID in the box to record it.' };
      const row = byId.has(id) ? byId.get(id) : -1;
      if (row >= 0 && norm(rows[row][soldCol]) === 'y') return { sale, id, how, score, row, action: 'skip', reason: `${id} is already marked sold.` };
      if (sale.refunded) return { sale, id, how, score, row, action: 'skip', reason: 'This order was refunded.' };
      return { sale, id, how, score, row, action: row >= 0 ? 'update' : 'add' };
    });
    return { lines };
  }

  function profitFormula(table) {
    const r = c => `${table}[[#This Row],[${c}]]`;
    return `=IFERROR(IF(${r('Sold?')}="Y",${r('Sold Price')}-${r('Cost')}-N(${r('Promotion Fee')})-N(${r('Shipping Cost')})-N(${r('eBay Fees')}),""),"")`;
  }

  /*
   * Cell values for one line. Existing values are kept unless they're blank (or for a
   * new sale, where Sold?/Date/Price always come from eBay).
   */
  function fieldsFor(line, existing, todaySerialFn) {
    const s = line.sale;
    const serial = d => (d ? Math.round(Date.parse(d + 'T00:00:00Z') / 864e5) + 25569 : '');
    const f = {
      'Build / Part ID': line.id,
      'Sold?': 'Y',
      'Date of Sale': serial(s.date),
      'Sold Via': 'eBay',
      'Sold Price': s.soldPrice == null ? '' : s.soldPrice,
      'eBay Order': s.order
    };
    if (s.promoFee) f['Promotion Fee'] = s.promoFee;
    if (s.shippingCost) f['Shipping Cost'] = s.shippingCost;
    if (s.ebayFees) f['eBay Fees'] = s.ebayFees;
    if (!existing || norm(existing['Listed?']) !== 'y') f['Listed?'] = 'Y';
    if (line.action === 'update' && line.how === 'order') {
      // Second report for the same order: only fill gaps.
      const keep = {};
      ['Promotion Fee', 'Shipping Cost', 'eBay Fees'].forEach(k => {
        if (f[k] != null && (existing[k] === '' || existing[k] == null || existing[k] === 0)) keep[k] = f[k];
      });
      return keep;
    }
    return f;
  }

  return { parseCsv, headerMap, findHeader, money, parseDate, idFrom, readReport, merge, guess, planSales, fieldsFor, profitFormula };
})();

if (typeof module !== 'undefined' && module.exports) module.exports = DealSales;
