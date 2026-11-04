/*
 * finance.js — the money picture from Purchases & Parts, Builds and Sales.
 * Pure functions (no Office.js), so it runs in Node tests too.
 *
 * Every purchase ends up in exactly one place:
 *   spare      available, not allocated, not sold
 *   inBuilds   allocated to a build that hasn't sold
 *   sold       marked sold, or allocated to a build that sold
 *   elsewhere  not available and not allocated: faulty, returned, kit, upgrades…
 */
const DealFinance = (() => {
  const norm = v => String(v == null ? '' : v).trim().toLowerCase();
  const num = v => (typeof v === 'number' && Number.isFinite(v) ? v : null);
  const yes = v => v === true || norm(v) === 'y' || norm(v) === 'yes';
  const idOf = v => String(v == null ? '' : v).trim().toUpperCase();
  const costOf = p => (num(p['Total Cost']) != null ? p['Total Cost'] : (num(p['Cost']) || 0));
  const valueOf = p => num(p['Estimated Part Value at Purchase Date']);
  const round2 = n => Math.round(n * 100) / 100;
  const MONTHS = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];

  // Excel serial -> { y, m } (1900 date system)
  function ym(serial) {
    const d = new Date(Date.UTC(1899, 11, 30) + Math.round(serial) * 86400000);
    return { y: d.getUTCFullYear(), m: d.getUTCMonth() };
  }
  const monthKey = serial => { const x = ym(serial); return `${x.y}-${String(x.m + 1).padStart(2, '0')}`; };

  // Why a part isn't in stock, from its notes: groups faulty and returned parts together.
  function elsewhereReason(p) {
    const note = String(p['Other Notes'] || '').trim();
    if (/faulty|dead|doa|broken/i.test(note)) return 'Faulty';
    if (/refund|return|not delivered/i.test(note)) return 'Returned or refunded';
    if (note) return note.length > 40 ? note.slice(0, 38) + '…' : note;
    return p['Category'] ? `${p['Category']} (not for resale)` : 'Not for resale';
  }

  function stage(b, sale) {
    if (sale && yes(sale['Listed?'])) return 'Listed';
    if (yes(b['Imaged?'])) return 'Imaged';
    if (yes(b['Benchmarked?'])) return 'Benchmarked';
    if (yes(b['Built?'])) return 'Built';
    if (yes(b['Ready to Build?'])) return 'Ready to build';
    if (yes(b['All Parts Required to Complete?'])) return 'All parts in';
    return 'Collecting parts';
  }

  /*
   * opts: { today (serial), feePct (avg eBay+promo fee %, or null), agedDays (60),
   *         months (6), deadlines (DealCore.testDeadlines output) }
   */
  function overview(purchases, builds, sales, opts) {
    opts = Object.assign({ agedDays: 60, months: 6, feePct: null, deadlines: [] }, opts || {});
    const today = opts.today;
    purchases = (purchases || []).filter(p => String(p['Part ID'] || '').trim());
    builds = (builds || []).filter(b => String(b['Build ID'] || '').trim());
    sales = (sales || []).filter(s => String(s['Build / Part ID'] || '').trim());

    const saleById = new Map();
    sales.forEach(s => { const id = idOf(s['Build / Part ID']); if (!saleById.has(id) || yes(s['Sold?'])) saleById.set(id, s); });
    const soldRows = sales.filter(s => yes(s['Sold?']));
    const soldIds = new Set(soldRows.map(s => idOf(s['Build / Part ID'])));

    // ---- where every part is -------------------------------------------------
    const spare = { count: 0, cost: 0, value: 0, valued: 0, valuedCost: 0, byCategory: [] };
    const inBuilds = { count: 0, cost: 0, value: 0, builds: 0 };
    const sold = { count: 0, cost: 0 };
    const elsewhere = { count: 0, cost: 0, groups: [] };
    const cats = new Map();
    const away = new Map();
    const partsByBuild = new Map();
    const aged = [];
    purchases.forEach(p => {
      const cost = costOf(p);
      const val = valueOf(p);
      const alloc = idOf(p['Allocation']);
      if (yes(p['Sold?']) || (alloc && soldIds.has(alloc))) { sold.count++; sold.cost += cost; return; }
      if (alloc) {
        inBuilds.count++; inBuilds.cost += cost; if (val != null) inBuilds.value += val;
        partsByBuild.set(alloc, (partsByBuild.get(alloc) || 0) + cost);
        return;
      }
      if (yes(p['Available?'])) {
        spare.count++; spare.cost += cost;
        if (val != null) { spare.value += val; spare.valued++; spare.valuedCost += cost; }
        const cat = p['Category'] || 'Other';
        if (!cats.has(cat)) cats.set(cat, { category: cat, count: 0, cost: 0, value: 0, valued: 0 });
        const g = cats.get(cat);
        g.count++; g.cost += cost; if (val != null) { g.value += val; g.valued++; }
        const bought = num(p['Purchase Date']);
        if (today != null && bought != null && today - bought >= opts.agedDays) {
          aged.push({ partId: String(p['Part ID']).trim(), name: [p['Brand'] === 'No Brand' ? '' : p['Brand'], p['Description']].filter(Boolean).join(' '), category: p['Category'] || '', days: Math.round(today - bought), cost });
        }
        return;
      }
      elsewhere.count++; elsewhere.cost += cost;
      const why = elsewhereReason(p);
      if (!away.has(why)) away.set(why, { reason: why, count: 0, cost: 0 });
      const a = away.get(why); a.count++; a.cost += cost;
    });
    inBuilds.builds = partsByBuild.size;
    spare.byCategory = [...cats.values()].sort((a, b) => b.cost - a.cost);
    elsewhere.groups = [...away.values()].sort((a, b) => b.cost - a.cost);
    aged.sort((a, b) => b.days - a.days);

    // ---- builds on the way to sale --------------------------------------------
    const fee = num(opts.feePct) || 0;
    const pipeline = builds.filter(b => !soldIds.has(idOf(b['Build ID']))).map(b => {
      const id = String(b['Build ID']).trim();
      const s = saleById.get(idOf(id));
      const cost = num(b['Build Cost']) != null ? b['Build Cost'] : (partsByBuild.get(idOf(id)) || 0);
      const price = s && num(s['Listing Price']) > 0 ? s['Listing Price'] : null;
      const listedOn = s && yes(s['Listed?']) && num(s['Listed Date']) ? s['Listed Date'] : null;
      return {
        id, description: String(b['Build Description'] || ''), stage: stage(b, s), cost, price,
        complete: stage(b, s) !== 'Collecting parts',
        profit: price != null ? price - cost : null,
        profitAfterFees: price != null ? price * (1 - fee / 100) - cost : null,
        listed: !!(s && yes(s['Listed?'])), listedDays: listedOn != null && today != null ? Math.round(today - listedOn) : null
      };
    });
    // Single parts listed for sale on their own (Sales rows with a Part ID).
    const partListings = sales.filter(s => !yes(s['Sold?']) && /^P\d+/i.test(String(s['Build / Part ID']).trim())).map(s => {
      const cost = num(s['Cost']) || 0;
      const price = num(s['Listing Price']) > 0 ? s['Listing Price'] : null;
      return { id: String(s['Build / Part ID']).trim(), description: String(s['Description'] || ''), cost, price, profit: price != null ? price - cost : null, listed: yes(s['Listed?']) };
    });
    // Profit only counts builds with all their parts in: a build still collecting parts
    // looks more profitable than it will be.
    const priced = pipeline.filter(b => b.price != null && b.complete);
    const provisional = pipeline.filter(b => b.price != null && !b.complete);
    const sumOf = (list, f) => list.reduce((a, x) => a + (f(x) || 0), 0);
    const pipe = {
      builds: pipeline,
      parts: partListings,
      count: pipeline.length,
      cost: sumOf(pipeline, b => b.cost),
      priced: priced.length,
      pricedCost: sumOf(priced, b => b.cost),
      listValue: sumOf(priced, b => b.price),
      profit: sumOf(priced, b => b.profit),
      profitAfterFees: sumOf(priced, b => b.profitAfterFees),
      unpriced: pipeline.filter(b => b.price == null).map(b => b.id),
      provisional: provisional.map(b => b.id),
      provisionalProfit: sumOf(provisional, b => b.profit),
      listed: pipeline.filter(b => b.listed).length,
      feePct: fee || null
    };
    pipe.marginPct = pipe.pricedCost > 0 ? pipe.profit / pipe.pricedCost * 100 : null;

    // ---- sales so far ---------------------------------------------------------
    const realised = { count: soldRows.length, revenue: 0, cost: 0, fees: { ebay: 0, promo: 0, postage: 0, materials: 0 }, profit: 0, days: [] };
    const byVia = new Map();
    soldRows.forEach(s => {
      const rev = num(s['Sold Price']) || 0;
      const cost = num(s['Cost']) || 0;
      const f = { ebay: num(s['eBay Fees']) || 0, promo: num(s['Promotion Fee']) || 0, postage: num(s['Shipping Cost']) || 0, materials: num(s['Shipping Material Cost']) || 0 };
      const feeSum = f.ebay + f.promo + f.postage + f.materials;
      const profit = num(s['Profit']) != null ? s['Profit'] : rev - cost - feeSum;
      realised.revenue += rev; realised.cost += cost; realised.profit += profit;
      Object.keys(f).forEach(k => { realised.fees[k] += f[k]; });
      if (num(s['Date of Sale']) && num(s['Listed Date'])) realised.days.push(s['Date of Sale'] - s['Listed Date']);
      const via = String(s['Sold Via'] || '').trim() || 'Not recorded';
      if (!byVia.has(via)) byVia.set(via, { via, count: 0, revenue: 0, profit: 0 });
      const v = byVia.get(via); v.count++; v.revenue += rev; v.profit += profit;
    });
    realised.feesTotal = realised.fees.ebay + realised.fees.promo + realised.fees.postage + realised.fees.materials;
    realised.returnPct = realised.cost > 0 ? realised.profit / realised.cost * 100 : null;
    realised.avgProfit = realised.count ? realised.profit / realised.count : null;
    realised.avgDaysToSell = realised.days.length ? Math.round(realised.days.reduce((a, b) => a + b, 0) / realised.days.length) : null;
    realised.byVia = [...byVia.values()].sort((a, b) => b.revenue - a.revenue);
    delete realised.days;

    // ---- month by month -------------------------------------------------------
    const months = [];
    if (today != null) {
      const now = ym(today);
      for (let i = opts.months - 1; i >= 0; i--) {
        const d = new Date(Date.UTC(now.y, now.m - i, 1));
        months.push({ key: `${d.getUTCFullYear()}-${String(d.getUTCMonth() + 1).padStart(2, '0')}`, label: MONTHS[d.getUTCMonth()], year: d.getUTCFullYear(), spend: 0, revenue: 0, profit: 0, bought: 0, soldCount: 0 });
      }
      const at = new Map(months.map(m => [m.key, m]));
      purchases.forEach(p => { const d = num(p['Purchase Date']); const m = d != null && at.get(monthKey(d)); if (m) { m.spend += costOf(p); m.bought++; } });
      soldRows.forEach(s => {
        const d = num(s['Date of Sale']); const m = d != null && at.get(monthKey(d));
        if (!m) return;
        const rev = num(s['Sold Price']) || 0;
        m.revenue += rev; m.soldCount++;
        m.profit += num(s['Profit']) != null ? s['Profit'] : rev - (num(s['Cost']) || 0) - (num(s['eBay Fees']) || 0) - (num(s['Promotion Fee']) || 0) - (num(s['Shipping Cost']) || 0) - (num(s['Shipping Material Cost']) || 0);
      });
    }

    // ---- how well you buy ------------------------------------------------------
    const valuedBuys = purchases.filter(p => valueOf(p) != null && valueOf(p) > 0 && costOf(p) > 0);
    const buyCost = sumOf(valuedBuys, costOf);
    const buyValue = sumOf(valuedBuys, valueOf);
    const buying = {
      parts: purchases.length,
      spent: sumOf(purchases, costOf),
      valued: valuedBuys.length,
      under: buyValue - buyCost,
      underPct: buyValue > 0 ? (buyValue - buyCost) / buyValue * 100 : null
    };

    // ---- money at risk: untested eBay parts near or past the 30 days ------------
    const costById = new Map(purchases.map(p => [String(p['Part ID']).trim(), costOf(p)]));
    const risk = { over: { count: 0, cost: 0 }, urgent: { count: 0, cost: 0 } };
    (opts.deadlines || []).forEach(d => {
      const bucket = d.status === 'over' ? risk.over : d.status === 'urgent' ? risk.urgent : null;
      if (!bucket) return;
      bucket.count++; bucket.cost += costById.get(String(d.partId).trim()) || 0;
    });

    const agedTotal = { count: aged.length, cost: sumOf(aged, a => a.cost) };
    const r = {
      stock: { spare, inBuilds, tiedUp: spare.cost + inBuilds.cost, aged, agedTotal, agedDays: opts.agedDays },
      sold, elsewhere, pipeline: pipe, realised, months, buying, risk,
      thisMonth: months.length ? months[months.length - 1] : null,
      lastMonth: months.length > 1 ? months[months.length - 2] : null
    };
    return roundAll(r);
  }

  // Pennies, not floating-point dust.
  function roundAll(x) {
    if (typeof x === 'number') return round2(x);
    if (Array.isArray(x)) return x.map(roundAll);
    if (x && typeof x === 'object') { Object.keys(x).forEach(k => { x[k] = roundAll(x[k]); }); return x; }
    return x;
  }

  return { overview, elsewhereReason, monthKey };
})();

if (typeof module !== 'undefined' && module.exports) module.exports = DealFinance;
