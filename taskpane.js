/*
 * taskpane.js — Rigboard: connects the panel to your Business Tracker workbook (Office.js).
 * The decisions live in core.js, planner.js, listing.js, sales.js and finance.js; this file reads
 * and writes the tables and draws the panel.
 *
 * Tables are found by their column headings, not their names, so renaming a
 * table (e.g. "Table2") doesn't break anything.
 */
(function () {
  'use strict';
  const C = DealCore;
  const PL = DealPlanner;
  const LI = DealListing;
  const SA = DealSales;
  const SETTINGS_EST = 'dealChecker.estimatedDelivery';
  const SETTINGS_PREFS = 'dealChecker.prefs';
  const HUB_RE = /^https:\/\/script\.google\.com\/macros\/s\/[\w-]+\/exec$/;
  const SALES_NEW_COLS = ['eBay Fees', 'eBay Order'];
  const APP_NAME = 'Rigboard';
  const APP_VERSION = '2.0';
  const gbp = new Intl.NumberFormat('en-GB', { style: 'currency', currency: 'GBP' });
  const gbp0 = new Intl.NumberFormat('en-GB', { style: 'currency', currency: 'GBP', maximumFractionDigits: 0 });
  const money = v => (typeof v === 'number' && Number.isFinite(v) ? gbp.format(v) : '—');
  const pounds = v => (typeof v === 'number' && Number.isFinite(v) ? gbp0.format(v) : '—');
  const nice = iso => (iso ? new Date(iso + 'T12:00:00').toLocaleDateString('en-GB', { day: 'numeric', month: 'short' }) : '—');
  const esc = s => String(s == null ? '' : s).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  const plural = (n, one, many) => `${n} ${n === 1 ? one : (many || one + 's')}`;
  const lc = s => String(s == null ? '' : s).trim().toLowerCase();

  const FI = DealFinance;
  const view = document.getElementById('view');
  const setup = document.getElementById('setup');
  let tab = 'home';
  let lastTab = 'home';         // where the ⚙ Settings "Done" button goes back to
  let buildsSeg = 'list';       // Builds: list | shop | alerts
  let sellSeg = 'listing';      // Sell: listing | photo
  let model = null;
  let pending = null;        // parsed purchases payload waiting for confirmation
  let pendingSales = null;   // eBay sales waiting for confirmation
  let hubCache = null;       // last reply from the Deal Alerts sheet
  const sellEdits = {};      // your edits to listing drafts, per build
  let sellId = null;

  // ---------------------------------------------------------------------------
  // Preferences (saved in the workbook, so your business partner gets them too)
  // ---------------------------------------------------------------------------

  function prefs() {
    let p = {};
    try { p = Office.context.document.settings.get(SETTINGS_PREFS) || {}; } catch (e) { p = {}; }
    return Object.assign({ hubUrl: '', hubKey: '', marginPct: 25, os: 'Windows 11 Home', greatPct: 25 }, p);
  }
  async function savePrefs(changes) {
    const next = Object.assign(prefs(), changes);
    Office.context.document.settings.set(SETTINGS_PREFS, next);
    await new Promise(res => Office.context.document.settings.saveAsync(() => res()));
    return next;
  }

  // ---------------------------------------------------------------------------
  // Little helpers for the panel
  // ---------------------------------------------------------------------------

  function openUrl(url) {
    try {
      if (Office.context.ui && Office.context.ui.openBrowserWindow) { Office.context.ui.openBrowserWindow(url); return; }
    } catch (e) { /* fall through */ }
    window.open(url, '_blank', 'noopener');
  }
  // Links in the panel open in your browser.
  view.addEventListener('click', e => {
    const a = e.target.closest('a[data-ext]');
    if (!a) return;
    e.preventDefault();
    openUrl(a.href);
  });
  const ext = (url, text) => `<a href="${esc(url)}" data-ext>${text}</a>`;

  async function copyText(text, btn) {
    let ok = false;
    try { await navigator.clipboard.writeText(text); ok = true; } catch (e) {
      const ta = document.createElement('textarea');
      ta.value = text;
      ta.style.position = 'fixed'; ta.style.opacity = '0';
      document.body.appendChild(ta);
      ta.select();
      try { ok = document.execCommand('copy'); } catch (x) { ok = false; }
      ta.remove();
    }
    if (btn) {
      const was = btn.textContent;
      btn.textContent = ok ? 'Copied' : 'Couldn’t copy';
      setTimeout(() => { btn.textContent = was; }, 1500);
    }
    return ok;
  }

  function download(name, text, type) {
    const url = URL.createObjectURL(text instanceof Blob ? text : new Blob([text], { type: type || 'text/csv;charset=utf-8' }));
    const a = document.createElement('a');
    a.href = url; a.download = name;
    document.body.appendChild(a); a.click(); a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 5000);
  }

  // ---------------------------------------------------------------------------
  // Reading the workbook
  // ---------------------------------------------------------------------------

  const has = (headers, names) => names.every(n => headers.some(h => String(h).trim().toLowerCase() === n.toLowerCase()));

  async function findTables(ctx) {
    const tables = ctx.workbook.tables;
    tables.load('items/name');
    await ctx.sync();
    const heads = tables.items.map(t => { const r = t.getHeaderRowRange(); r.load('values'); return { t, r }; });
    await ctx.sync();
    const found = {};
    heads.forEach(({ t, r }) => {
      const h = r.values[0];
      if (!found.purchases && has(h, ['Part ID', 'Purchased From', 'Tested?'])) found.purchases = t.name;
      else if (!found.builds && has(h, ['Build ID', 'Missing Parts'])) found.builds = t.name;
      else if (!found.sales && has(h, ['Build / Part ID', 'Sold?', 'Sold Price'])) found.sales = t.name;
      else if (!found.watch && t.name === 'Watchlist') found.watch = t.name;
      else if (!found.games && has(h, ['Build ID', 'Game', 'Average FPS'])) found.games = t.name;
      else if (!found.benches && has(h, ['Build ID', 'Benchmark', 'Score'])) found.benches = t.name;
    });
    return found;
  }

  async function readTable(ctx, name, withFormulas) {
    const t = ctx.workbook.tables.getItem(name);
    const head = t.getHeaderRowRange();
    head.load('values');
    const body = t.getDataBodyRange();
    body.load(withFormulas ? 'values,formulas,formulasR1C1,rowCount' : 'values,rowCount');
    await ctx.sync();
    return {
      name,
      headers: head.values[0].map(h => String(h).trim()),
      values: body.values,
      formulasA1: withFormulas ? body.formulas : [],
      formulasR1C1: withFormulas ? body.formulasR1C1 : []
    };
  }

  function estimatedKeys() {
    try { return Office.context.document.settings.get(SETTINGS_EST) || {}; } catch (e) { return {}; }
  }

  async function loadModel() {
    return Excel.run(async ctx => {
      const found = await findTables(ctx);
      if (!found.purchases) throw new Error('Couldn’t find the Purchases & Parts table. It needs columns called “Part ID”, “Purchased From” and “Tested?”.');
      const m = { found, purchases: await readTable(ctx, found.purchases, true) };
      m.builds = found.builds ? await readTable(ctx, found.builds, false) : null;
      m.sales = found.sales ? await readTable(ctx, found.sales, true) : null;
      m.watch = found.watch ? await readTable(ctx, found.watch, false) : null;
      m.games = found.games ? await readTable(ctx, found.games, false) : null;
      m.benches = found.benches ? await readTable(ctx, found.benches, false) : null;
      return m;
    });
  }

  async function refresh() {
    try {
      model = await loadModel();
      renderSetup();
      render();
      autoSync();
    } catch (e) {
      view.innerHTML = `<div class="err-box"><strong>Couldn’t read the workbook</strong><br>${esc(e.message || e)}</div>`;
    }
  }

  const purchases = () => C.toObjects(model.purchases.headers, model.purchases.values);
  const builds = () => (model.builds ? C.toObjects(model.builds.headers, model.builds.values) : []);
  const gameRows = () => (model.games ? C.toObjects(model.games.headers, model.games.values) : []);
  const benchRows = () => (model.benches ? C.toObjects(model.benches.headers, model.benches.values) : []);
  const salesRows = () => (model.sales ? C.toObjects(model.sales.headers, model.sales.values) : []);

  // ---------------------------------------------------------------------------
  // Writing: add rows to a table, keeping its formula columns and formats
  // ---------------------------------------------------------------------------

  /*
   * fresh: the table as just read (with formulas). rows: arrays in header order.
   * Formula columns get the column's formula; formats are copied from the first row.
   */
  async function addRows(ctx, fresh, rows, formatCols) {
    if (!rows.length) return;
    const t = ctx.workbook.tables.getItem(fresh.name);
    const calc = C.detectCalculated(fresh.formulasA1, fresh.formulasR1C1);
    const startRow = fresh.values.length;
    t.rows.add(null, rows.map(r => r.map((v, c) => (calc[c] && C.rowIndependent(calc[c].a1) ? calc[c].a1 : (calc[c] ? '' : v)))));
    await ctx.sync();
    const body = t.getDataBodyRange();
    const n = rows.length;
    const newCells = c => body.getCell(startRow, c).getResizedRange(n - 1, 0);
    calc.forEach((cf, c) => {
      if (!cf) return;
      if (C.rowIndependent(cf.a1)) newCells(c).formulas = Array.from({ length: n }, () => [cf.a1]);
      else newCells(c).formulasR1C1 = Array.from({ length: n }, () => [cf.r1c1]);
    });
    if (formatCols && formatCols.length) {
      // Same format as the first row; dates fall back to dd/mm/yyyy so they don't show as numbers.
      const cols = formatCols.map(name => ({ name, c: fresh.headers.findIndex(h => lc(h) === lc(name)) })).filter(x => x.c >= 0);
      const cells = startRow > 0 ? cols.map(x => { const cell = body.getCell(0, x.c); cell.load('numberFormat'); return Object.assign({ cell }, x); }) : cols;
      await ctx.sync();
      cells.forEach(x => {
        let fmt = x.cell ? x.cell.numberFormat[0][0] : null;
        if ((!fmt || fmt === 'General') && /date/i.test(x.name)) fmt = 'dd/mm/yyyy';
        if (fmt && fmt !== 'General') newCells(x.c).numberFormat = Array.from({ length: n }, () => [fmt]);
      });
    }
  }

  // ---------------------------------------------------------------------------
  // One-time sheet setup: Delivery Date + eBay Order Line columns, corrected countdown
  // ---------------------------------------------------------------------------

  function renderSetup() {
    const missing = C.missingColumns(model.purchases.headers);
    if (!missing.length) { setup.hidden = true; return; }
    setup.hidden = false;
    setup.innerHTML = `<p><strong>One-time setup</strong><br>Adds two columns to the end of Purchases &amp; Parts (<b>Delivery Date</b> and <b>eBay Order Line</b>)
      and changes <b>Days left to Test</b> to count from delivery, as eBay does. Nothing else in the sheet changes.</p>
      <button class="btn" id="do-setup" type="button">Set up the sheet</button>`;
    document.getElementById('do-setup').addEventListener('click', upgradeSheet);
  }

  async function upgradeSheet(e) {
    const btn = e.target;
    btn.disabled = true;
    btn.textContent = 'Setting up…';
    try {
      await Excel.run(async ctx => {
        const t = ctx.workbook.tables.getItem(model.purchases.name);
        const missing = C.missingColumns(model.purchases.headers);
        missing.forEach(name => t.columns.add(null, null, name));
        await ctx.sync();

        const body = t.getDataBodyRange();
        body.load('rowCount');
        await ctx.sync();
        const n = body.rowCount;
        if (n > 0) {
          if (model.purchases.headers.some(h => h.toLowerCase() === 'days left to test')) {
            const f = C.daysLeftFormula(model.purchases.name);
            t.columns.getItem('Days left to Test').getDataBodyRange().formulas = Array.from({ length: n }, () => [f]);
          }
          const pd = t.columns.getItem('Purchase Date').getDataBodyRange().getCell(0, 0);
          pd.load('numberFormat');
          await ctx.sync();
          t.columns.getItem('Delivery Date').getDataBodyRange().numberFormat = Array.from({ length: n }, () => [pd.numberFormat[0][0]]);
        }
        await ctx.sync();
      });
      await refresh();
    } catch (err) {
      btn.disabled = false;
      btn.textContent = 'Set up the sheet';
      setup.insertAdjacentHTML('beforeend', `<div class="err-box">${esc(err.message || err)}</div>`);
    }
  }

  // ---------------------------------------------------------------------------
  // Selecting rows in the sheet
  // ---------------------------------------------------------------------------

  async function selectRow(tableName, row) {
    try {
      await Excel.run(async ctx => {
        const t = ctx.workbook.tables.getItem(tableName);
        t.worksheet.activate();
        t.getDataBodyRange().getRow(row).select();
        await ctx.sync();
      });
    } catch (e) { /* the sheet may have changed; a refresh fixes it */ }
  }

  // ---------------------------------------------------------------------------
  // Views
  // ---------------------------------------------------------------------------

  function render() {
    if (!model) return;
    document.querySelectorAll('.tabs button').forEach(b => b.setAttribute('aria-selected', String(b.dataset.tab === tab)));
    document.getElementById('open-settings').setAttribute('aria-pressed', String(tab === 'settings'));
    ({ home: renderHome, builds: renderBuilds, sell: renderSell, money: renderMoney, import: renderImport, settings: renderSettings }[tab] || renderHome)();
  }

  function go(nextTab, opts) {
    if (nextTab !== 'settings') lastTab = nextTab;
    tab = nextTab;
    if (opts && opts.builds) buildsSeg = opts.builds;
    if (opts && opts.sell) sellSeg = opts.sell;
    if (opts && opts.sellId) sellId = opts.sellId;
    render();
    window.scrollTo(0, 0);
  }

  // Small segmented switch at the top of a tab.
  const segHtml = (items, cur) => `<div class="seg" role="tablist">${items.map(([k, label, n]) =>
    `<button type="button" role="tab" data-seg="${k}" aria-selected="${k === cur}">${label}${n ? ` <span class="n">${n}</span>` : ''}</button>`).join('')}</div>`;
  const wireSeg = set => view.querySelectorAll('[data-seg]').forEach(b => b.addEventListener('click', () => { set(b.dataset.seg); render(); }));
  // Buttons and links anywhere that jump to another tab: data-go="builds:shop", "sell:photo", "settings"
  view.addEventListener('click', e => {
    const b = e.target.closest('[data-go]');
    if (!b) return;
    const [t, sub, id] = b.dataset.go.split(':');
    go(t, t === 'builds' ? { builds: sub } : t === 'sell' ? { sell: sub || 'listing', sellId: id } : null);
  });

  // Today's money picture, shared by Home and Money.
  function money2() {
    const today = C.todaySerial();
    return FI.overview(purchases(), builds(), salesRows(), {
      today, feePct: LI.feeShare(salesRows()), deadlines: C.testDeadlines(purchases(), today, estimatedKeys())
    });
  }

  function deadlineRow(d) {
    const days = d.status === 'coming' ? `<span>Not arrived</span><small>expected ${nice(C.serialToIso(C.isoToSerial(d.deadline) - 30))}</small>`
      : d.status === 'over' ? `${Math.abs(d.daysLeft)}d over<small>ended ${nice(d.deadline)}</small>`
        : `${d.daysLeft}d<small>by ${nice(d.deadline)}</small>`;
    const tags = [
      d.allocation ? `<span class="tag">${esc(d.allocation)}</span>` : '',
      d.estimated ? '<span class="tag warn">est. delivery</span>' : '',
      !d.fromDelivery ? '<span class="tag warn">from purchase date</span>' : ''
    ].join('');
    return `<li class="s-${d.status}"><button class="row" type="button" data-row="${d.row}">
      <span><span class="name">${esc(d.partId)} ${esc(d.name)}</span><span class="sub">${tags}${esc(d.category || '')}</span></span>
      <span class="days">${days}</span></button></li>`;
  }

  // ---- Home: what needs doing, and the headline numbers ----------------------------

  function renderHome() {
    const today = C.todaySerial();
    const list = C.testDeadlines(purchases(), today, estimatedKeys());
    const by = s => list.filter(d => d.status === s);
    const urgent = by('over').length + by('urgent').length;
    const noValue = C.missingValues(purchases()).length;
    const f = money2();
    const section = (title, items) => items.length ? `<h2><span>${title}</span><span>${items.length}</span></h2><ul class="list">${items.map(deadlineRow).join('')}</ul>` : '';
    const later = by('ok');

    // Next up: the few things worth doing today, each one click away.
    const toList = f.pipeline.builds.filter(b => !b.listed && ['Built', 'Benchmarked', 'Imaged'].includes(b.stage));
    let fits = 0;
    if (model.builds) { try { fits = PL.plan(builds(), purchases()).fromStock; } catch (e) { fits = 0; } }
    const todo = [];
    if (urgent) todo.push(`<button class="todo urgent" type="button" data-jump="tests"><b>Test ${plural(urgent, 'part')}</b><span>${by('over').length ? `${by('over').length} past eBay’s 30 days` : 'eBay’s 30 days end this week'}</span></button>`);
    if (toList.length > 2) {
      const val = toList.reduce((a, b) => a + (b.profit || 0), 0);
      todo.push(`<button class="todo" type="button" data-go="sell:listing:${esc(toList[0].id)}"><b>List ${plural(toList.length, 'built PC')}</b><span>${esc(toList.map(b => b.id).join(', '))}${val ? ` · about ${pounds(val)} profit between them` : ''}</span></button>`);
    } else toList.forEach(b => todo.push(`<button class="todo" type="button" data-go="sell:listing:${esc(b.id)}"><b>List ${esc(b.id)}</b><span>${esc(b.stage.toLowerCase())}${b.price ? ` · ${pounds(b.price)}, about ${pounds(b.profit)} profit` : ' · no price yet'}</span></button>`));
    if (fits) todo.push(`<button class="todo" type="button" data-go="builds:shop"><b>Use ${plural(fits, 'spare part')}</b><span>already in stock and fit${fits === 1 ? 's' : ''} a build</span></button>`);
    if (f.pipeline.unpriced.length) todo.push(`<button class="todo" type="button" data-go="money"><b>Price ${plural(f.pipeline.unpriced.length, 'build')}</b><span>${esc(f.pipeline.unpriced.join(', '))} ${f.pipeline.unpriced.length === 1 ? 'has' : 'have'} no Listing Price in Sales</span></button>`);
    const deals = (hubCache && hubCache.alerts || []).slice(0, 3);

    view.innerHTML = `
      <div class="kpis">
        <button class="kpi" type="button" data-go="money"><span>Tied up in stock</span><b>${pounds(f.stock.tiedUp)}</b><small>${pounds(f.stock.inBuilds.cost)} in builds · ${pounds(f.stock.spare.cost)} spare</small></button>
        <button class="kpi" type="button" data-go="money"><span>Potential profit</span><b class="pos">${pounds(f.pipeline.profit)}</b><small>on ${plural(f.pipeline.priced, 'priced build')}${f.pipeline.feePct ? ' after fees' : ''}</small></button>
      </div>
      ${todo.length ? `<h2><span>Next up</span></h2><div class="todos">${todo.join('')}</div>` : '<p class="msg"><strong>All caught up</strong><span class="muted">Nothing to test, list or allocate right now.</span></p>'}
      ${deals.length ? `<h2><span>Latest deals</span><button class="link small" type="button" data-go="builds:alerts">All alerts</button></h2><ul class="list">${deals.map(dealRow).join('')}</ul>` : ''}
      <div id="tests">
      ${list.length ? `<h2 class="sec"><span>Testing deadlines</span><span class="muted small">30 days from delivery</span></h2>` : ''}
      ${section('Past the 30 days', by('over'))}
      ${section('Test within 7 days', by('urgent'))}
      ${section('Within 14 days', by('soon'))}
      ${later.length ? `<details class="more" ${urgent || by('soon').length ? '' : 'open'}><summary>Later (${later.length})</summary><ul class="list">${later.map(deadlineRow).join('')}</ul></details>` : ''}
      ${section('On the way', by('coming'))}
      </div>
      ${noValue ? `<p class="small muted" style="margin-top:14px">${plural(noValue, 'part has', 'parts have')} no market value yet. Parts you check with the Chrome extension before buying get one automatically.</p>` : ''}`;
    view.querySelectorAll('[data-row]').forEach(b => b.addEventListener('click', () => selectRow(model.purchases.name, +b.dataset.row)));
    const jump = view.querySelector('[data-jump="tests"]');
    if (jump) jump.addEventListener('click', () => document.getElementById('tests').scrollIntoView({ behavior: 'smooth' }));
    if (isHubSet() && !hubCache) loadHub().then(() => { if (tab === 'home') renderHome(); }).catch(() => {});
  }

  const when = ts => (ts ? new Date(ts).toLocaleString('en-GB', { day: 'numeric', month: 'short', hour: '2-digit', minute: '2-digit' }) : '');
  const dealRow = a => `<li><div class="row" style="cursor:default">
      <span><span class="name">${ext(a.url, esc(a.title))}</span><span class="sub">${esc(a.label)} · ${esc(when(a.found))}</span></span>
      <span class="days">${money(a.total)}<small>alert ≤ ${pounds(a.alertAt)}</small></span></div></li>`;

  function noBuilds() {
    view.innerHTML = '<p class="msg"><strong>No Builds table found</strong><span class="muted">It needs columns called “Build ID” and “Missing Parts”.</span></p>';
  }

  // ---- Builds: the builds themselves, what to buy for them, and deal alerts ---------

  function buildsTop() {
    let toBuy = 0;
    try { model.plan = PL.plan(builds(), purchases()); toBuy = model.plan.shopping.reduce((a, g) => a + g.count, 0); } catch (e) { toBuy = 0; }
    const deals = hubCache && hubCache.alerts ? hubCache.alerts.length : 0;
    return segHtml([['list', 'Builds', builds().filter(b => String(b['Build ID'] || '').trim()).length], ['shop', 'Shopping', toBuy], ['alerts', 'Alerts', deals]], buildsSeg);
  }

  function renderBuilds() {
    if (!model.builds) return noBuilds();
    if (buildsSeg === 'shop') renderShop();
    else if (buildsSeg === 'alerts') renderWatch();
    else renderBuildList();
    wireSeg(s => { buildsSeg = s; });
  }

  function renderBuildList() {
    const list = C.buildsSummary(builds(), purchases());
    const fin = new Map(money2().pipeline.builds.map(b => [b.id, b]));
    const order = { 'Collecting parts': 0, 'All parts in': 1, 'Ready to build': 2, 'Built': 3, 'Benchmarked': 4 };
    list.sort((a, b) => order[a.stage] - order[b.stage] || String(a.id).localeCompare(String(b.id), undefined, { numeric: true }));
    const stages = Object.keys(order).map(s => [s, list.filter(b => b.stage === s).length]);
    view.innerHTML = buildsTop() + `
      <div class="pipe">${stages.map(([s, n]) => `<div class="${n ? '' : 'none'}"><b>${n}</b><span>${esc(s)}</span></div>`).join('')}</div>
      ${list.map(b => {
        const f = fin.get(b.id);
        const priceBit = f && f.price ? ` · <b>${pounds(f.price)}</b> ${f.complete ? `<span class="pos">+${pounds(f.profit)}</span>` : `<span title="Still collecting parts, so the profit will drop">(+${pounds(f.profit)} so far)</span>`}` : f ? ' · <span class="warn-text">no price</span>' : ' · sold';
        return `<button class="build" type="button" data-row="${b.row}">
        <div class="head"><span class="id">${esc(b.id)}</span><span class="stage">${esc(f && f.listed ? 'Listed' : b.stage)}</span></div>
        <div class="desc">${esc(b.description) || '<span class="muted">No description</span>'}</div>
        <div class="small muted">${plural(b.parts, 'part')} · cost ${pounds(b.cost)}${priceBit}${b.untested ? ` · <span class="tag warn">${b.untested} untested</span>` : ''}${b.waitingOn ? ` · waiting on ${esc(b.waitingOn)}` : ''}</div>
        ${b.missing.length ? `<div class="chips">${b.missing.map(m => `<span class="chip">needs ${esc(m)}</span>`).join('')}</div>` : ''}
      </button>`;
      }).join('') || '<p class="muted">No builds yet.</p>'}
      <p class="small muted">Click a build to find it in the sheet. Prices come from <b>Listing Price</b> in Sales.</p>`;
    view.querySelectorAll('[data-row]').forEach(b => b.addEventListener('click', () => selectRow(model.builds.name, +b.dataset.row)));
  }

  // ---------------------------------------------------------------------------
  // Money: where it is, what it could make, what it has made
  // ---------------------------------------------------------------------------

  const COL_A = '#2a78d6';   // chart colours (checked for colour-blind separation)
  const COL_B = '#eb6834';

  function renderMoney() {
    const f = money2();
    const st = f.stock;
    const pipe = f.pipeline;
    const rz = f.realised;
    const tm = f.thisMonth || { spend: 0 };
    const lm = f.lastMonth;
    const pct = v => (v == null ? '—' : `${Math.round(v)}%`);
    const pctOf = (a, b) => (b > 0 ? Math.max(0, a / b * 100) : 0);

    // Where the money is: one bar, builds vs spare
    const split = st.tiedUp > 0 ? `
      <div class="split" role="img" aria-label="${pounds(st.inBuilds.cost)} in builds, ${pounds(st.spare.cost)} in spare parts">
        <span style="width:${pctOf(st.inBuilds.cost, st.tiedUp)}%;background:${COL_A}" title="In builds: ${pounds(st.inBuilds.cost)}"></span>
        <span style="width:${pctOf(st.spare.cost, st.tiedUp)}%;background:${COL_B}" title="Spare parts: ${pounds(st.spare.cost)}"></span>
      </div>
      <div class="legend">
        <div><i style="background:${COL_A}"></i><span>In ${plural(st.inBuilds.builds, 'build')}</span><b>${pounds(st.inBuilds.cost)}</b><small>${st.inBuilds.count} parts · worth ${pounds(st.inBuilds.value)}</small></div>
        <div><i style="background:${COL_B}"></i><span>Spare parts</span><b>${pounds(st.spare.cost)}</b><small>${st.spare.count} parts · worth ${pounds(st.spare.value)}${st.spare.valued < st.spare.count ? '*' : ''}</small></div>
      </div>` : '<p class="muted">No stock.</p>';

    const pipeRows = pipe.builds.slice().sort((a, b) => (b.price != null) - (a.price != null) || String(a.id).localeCompare(String(b.id), undefined, { numeric: true }));
    const buildRow = b => `<tr class="${b.complete ? '' : 'prov'}">
        <td><b>${esc(b.id)}</b><small>${esc(b.listed ? `Listed${b.listedDays != null ? ` ${b.listedDays}d` : ''}` : b.stage)}</small></td>
        <td class="n">${pounds(b.cost)}</td>
        <td class="n">${b.price != null ? pounds(b.price) : `<button class="link small" type="button" data-go="sell:listing:${esc(b.id)}">set</button>`}</td>
        <td class="n ${b.profit > 0 ? 'pos' : b.profit < 0 ? 'neg' : ''}">${b.profit != null ? (b.complete ? '' : '(') + pounds(b.profit) + (b.complete ? '' : ')') : '—'}</td></tr>`;

    const cats = st.spare.byCategory;
    const catRow = x => `<tr><td>${esc(x.category)}</td><td class="n">${x.count}</td><td class="n">${pounds(x.cost)}</td><td class="n">${x.valued ? pounds(x.value) : '—'}</td></tr>`;

    view.innerHTML = `
      <div class="kpis">
        <div class="kpi"><span>Tied up in stock</span><b>${pounds(st.tiedUp)}</b><small>${plural(st.inBuilds.count + st.spare.count, 'part')} not yet sold</small></div>
        <div class="kpi"><span>Potential profit</span><b class="pos">${pounds(pipe.feePct ? pipe.profitAfterFees : pipe.profit)}</b><small>${plural(pipe.priced, 'build')} with all parts${pipe.feePct ? `, after ${pipe.feePct}% fees` : ', before fees'}</small></div>
        <div class="kpi"><span>Profit so far</span><b class="${rz.profit > 0 ? 'pos' : rz.profit < 0 ? 'neg' : ''}">${rz.count ? pounds(rz.profit) : '£0'}</b><small>${rz.count ? `${plural(rz.count, 'sale')} · ${pct(rz.returnPct)} return` : 'No sales recorded yet'}</small></div>
        <div class="kpi"><span>Spent this month</span><b>${pounds(tm.spend)}</b><small>${lm ? `${esc(lm.label)}: ${pounds(lm.spend)}` : ''}</small></div>
      </div>

      <h2 class="sec"><span>Where your money is</span></h2>
      ${split}
      ${st.spare.valued < st.spare.count ? `<p class="small muted">* ${st.spare.count - st.spare.valued} spare parts have no Estimated Part Value, so the value leaves them out.</p>` : ''}

      <h2 class="sec"><span>Builds</span><span>${pipe.count}</span></h2>
      ${pipe.count ? `<table class="grid money"><tr><th>Build</th><th class="n">Cost</th><th class="n">Price</th><th class="n">Profit</th></tr>
        ${pipeRows.map(buildRow).join('')}
        <tr class="total"><td>${plural(pipe.priced, 'ready build')}</td><td class="n">${pounds(pipe.pricedCost)}</td><td class="n">${pounds(pipe.listValue)}</td><td class="n pos">${pounds(pipe.profit)}</td></tr></table>
        <p class="small muted">Price is the <b>Listing Price</b> in Sales. Profit is before eBay fees and postage${pipe.feePct && pipe.priced ? ` (your sales so far average ${pipe.feePct}% fees, which would leave ${pounds(pipe.profitAfterFees)})` : ''}.
          ${pipe.provisional.length ? `Figures in brackets are for builds still collecting parts (${esc(pipe.provisional.join(', '))}): their profit drops as you buy the rest, so they’re not in the total.` : ''}
          ${pipe.marginPct != null ? `Ready builds are priced at ${Math.round(pipe.marginPct)}% over cost on average.` : ''}</p>` : '<p class="muted">No unsold builds.</p>'}
      ${pipe.parts.length ? `<details class="more"><summary>Parts for sale on their own (${pipe.parts.length})</summary><table class="grid money"><tr><th>Part</th><th class="n">Cost</th><th class="n">Price</th><th class="n">Profit</th></tr>
        ${pipe.parts.map(p => `<tr><td><b>${esc(p.id)}</b><small>${esc(p.description.slice(0, 40))}</small></td><td class="n">${pounds(p.cost)}</td><td class="n">${pounds(p.price)}</td><td class="n">${p.profit != null ? pounds(p.profit) : '—'}</td></tr>`).join('')}</table></details>` : ''}

      <h2 class="sec"><span>Sales</span><span>${rz.count || ''}</span></h2>
      ${rz.count ? `<table class="grid money">
          <tr><td>Sold for</td><td class="n">${money(rz.revenue)}</td></tr>
          <tr><td>Cost of what sold</td><td class="n">− ${money(rz.cost)}</td></tr>
          ${rz.fees.ebay ? `<tr><td class="ind">eBay fees</td><td class="n">− ${money(rz.fees.ebay)}</td></tr>` : ''}
          ${rz.fees.promo ? `<tr><td class="ind">Promoted listings</td><td class="n">− ${money(rz.fees.promo)}</td></tr>` : ''}
          ${rz.fees.postage ? `<tr><td class="ind">Postage</td><td class="n">− ${money(rz.fees.postage)}</td></tr>` : ''}
          ${rz.fees.materials ? `<tr><td class="ind">Packaging</td><td class="n">− ${money(rz.fees.materials)}</td></tr>` : ''}
          <tr class="total"><td>Profit</td><td class="n ${rz.profit >= 0 ? 'pos' : 'neg'}">${money(rz.profit)}</td></tr></table>
        <div class="minis">
          <div><b>${pct(rz.returnPct)}</b><span>return on cost</span></div>
          <div><b>${pounds(rz.avgProfit)}</b><span>profit per sale</span></div>
          <div><b>${rz.avgDaysToSell != null ? rz.avgDaysToSell + 'd' : '—'}</b><span>listed to sold</span></div>
        </div>
        ${rz.byVia.length > 1 ? `<p class="small muted">${rz.byVia.map(v => `${esc(v.via)}: ${plural(v.count, 'sale')}, ${pounds(v.profit)} profit`).join(' · ')}</p>` : ''}
        <p class="small muted">Profit uses your Sales sheet’s Profit column where it’s filled in.</p>`
        : `<p class="small muted">Nothing marked sold yet. When something sells, record it on the <button class="link" type="button" data-go="import">Import</button> tab from eBay’s report (or set <b>Sold?</b> to Y in Sales) and revenue, fees and profit show here.</p>`}

      <h2 class="sec"><span>Last ${f.months.length} months</span></h2>
      ${monthsChart(f.months)}

      <h2 class="sec"><span>Spare parts by category</span><span>${st.spare.count}</span></h2>
      <table class="grid money"><tr><th>Category</th><th class="n">Parts</th><th class="n">Cost</th><th class="n">Value</th></tr>
        ${cats.slice(0, 6).map(catRow).join('')}
        ${cats.length > 6 ? `<tr class="fold"><td colspan="4"><details><summary>${cats.length - 6} more categories</summary><table class="grid inner">${cats.slice(6).map(catRow).join('')}</table></details></td></tr>` : ''}
        <tr class="total"><td>Total</td><td class="n">${st.spare.count}</td><td class="n">${pounds(st.spare.cost)}</td><td class="n">${pounds(st.spare.value)}</td></tr></table>

      <h2 class="sec"><span>Worth a look</span></h2>
      <ul class="notes">
        ${f.risk.over.count || f.risk.urgent.count ? `<li class="bad"><b>${pounds(f.risk.over.cost + f.risk.urgent.cost)} untested</b> near or past eBay’s 30 days: ${f.risk.over.count ? `${plural(f.risk.over.count, 'part')} (${pounds(f.risk.over.cost)}) past` : ''}${f.risk.over.count && f.risk.urgent.count ? ', ' : ''}${f.risk.urgent.count ? `${plural(f.risk.urgent.count, 'part')} (${pounds(f.risk.urgent.cost)}) due this week` : ''}. <button class="link" type="button" data-go="home">See them</button></li>` : ''}
        ${st.agedTotal.count ? `<li class="warn"><b>${pounds(st.agedTotal.cost)} in older spare parts</b>: ${plural(st.agedTotal.count, 'part')} bought ${st.agedDays}+ days ago and not in a build.
          <details><summary>Oldest</summary><ul class="plain">${st.aged.slice(0, 8).map(a => `<li>${esc(a.partId)} ${esc(a.name)} <span class="muted">· ${a.days}d · ${pounds(a.cost)}</span></li>`).join('')}</ul></details></li>` : ''}
        ${f.elsewhere.count ? `<li><b>${pounds(f.elsewhere.cost)} spent on parts that aren’t stock</b> (faulty, returned, tools, upgrades for others):
          <ul class="plain">${f.elsewhere.groups.slice(0, 5).map(g => `<li>${esc(g.reason)}: ${pounds(g.cost)} <span class="muted">(${plural(g.count, 'part')})</span></li>`).join('')}</ul></li>` : ''}
        ${f.buying.underPct != null ? `<li class="good"><b>You buy ${Math.round(f.buying.underPct)}% under market value</b> on average: ${pounds(f.buying.under)} below value across ${plural(f.buying.valued, 'valued part')}.</li>` : ''}
        <li class="muted small">All-time: ${pounds(f.buying.spent)} spent on ${plural(f.buying.parts, 'part')}${rz.count ? `, ${pounds(rz.revenue)} back from sales` : ''}.</li>
      </ul>`;
  }

  // Spend vs sales per month: grouped columns, one axis.
  function monthsChart(months) {
    if (!months.length) return '';
    const max = Math.max(1, ...months.map(m => Math.max(m.spend, m.revenue)));
    const step = [50, 100, 250, 500, 1000, 2500, 5000, 10000].find(s => max / s <= 4) || 25000;
    const top = Math.ceil(max / step) * step;
    const W = 320, H = 150, L = 36, B = 20, T = 8, plotH = H - B - T, slot = (W - L) / months.length, bw = Math.min(14, slot / 2 - 4);
    const y = v => T + plotH - v / top * plotH;
    const bar = (x, v, col, label) => {
      if (v <= 0) return '';
      const h = Math.max(2, v / top * plotH), r = Math.min(3, h / 2), yy = T + plotH - h;
      return `<path d="M${x},${T + plotH}V${yy + r}Q${x},${yy} ${x + r},${yy}H${x + bw - r}Q${x + bw},${yy} ${x + bw},${yy + r}V${T + plotH}Z" fill="${col}"><title>${esc(label)}</title></path>`;
    };
    const ticks = [];
    for (let v = 0; v <= top; v += step) ticks.push(v);
    const hasSales = months.some(m => m.revenue > 0);
    return `<div class="chart">
      <svg viewBox="0 0 ${W} ${H}" role="img" aria-label="Spending${hasSales ? ' and sales' : ''} per month">
        ${ticks.map(v => `<line x1="${L}" x2="${W}" y1="${y(v)}" y2="${y(v)}" class="grid-l"/><text x="${L - 5}" y="${y(v) + 3.5}" text-anchor="end" class="ax">${v >= 1000 ? '£' + v / 1000 + 'k' : '£' + v}</text>`).join('')}
        ${months.map((m, i) => {
          const cx = L + slot * i + slot / 2;
          return bar(cx - bw - 1, m.spend, COL_A, `${m.label} ${m.year}: spent ${pounds(m.spend)} on ${plural(m.bought, 'part')}`) +
            bar(cx + 1, m.revenue, COL_B, `${m.label} ${m.year}: sold ${pounds(m.revenue)} (${plural(m.soldCount, 'sale')}), profit ${pounds(m.profit)}`) +
            `<text x="${cx}" y="${H - 5}" text-anchor="middle" class="ax">${m.label}</text>`;
        }).join('')}
      </svg>
      <div class="legend row"><div><i style="background:${COL_A}"></i><span>Spent on parts</span></div><div><i style="background:${COL_B}"></i><span>Sales</span></div></div>
      <details class="more"><summary>As a table</summary><table class="grid money"><tr><th>Month</th><th class="n">Spent</th><th class="n">Sales</th><th class="n">Profit</th></tr>
        ${months.map(m => `<tr><td>${m.label} ${m.year}</td><td class="n">${pounds(m.spend)}</td><td class="n">${pounds(m.revenue)}</td><td class="n">${m.soldCount ? pounds(m.profit) : '—'}</td></tr>`).join('')}</table></details>
    </div>`;
  }

  // ---------------------------------------------------------------------------
  // Shop: the build-aware shopping list
  // ---------------------------------------------------------------------------

  const KIND = PL.KIND_NAME;

  function renderShop() {
    const top = buildsTop();
    const plan = model.plan;
    const stockBuilds = plan.builds.filter(b => b.needs.some(n => n.stock));
    const toBuy = plan.shopping.reduce((a, g) => a + g.count, 0);
    const watched = new Map(((hubCache && hubCache.watches) || []).map(w => [lc(w.label), w]));
    const hubOn = isHubSet();

    const autoOn = hubOn && prefs().autoAlerts;
    view.innerHTML = top + `
      <div class="big">
        <div class="stat"><b>${plan.fromStock}</b><span>spare parts fit a build</span></div>
        <div class="stat"><b>${toBuy}</b><span>parts still to buy</span></div>
      </div>
      <p class="small muted" style="margin:0 0 4px">Worked out from each build’s <b>Missing Parts</b>, its CPU or board, and your spare stock.
        ${autoOn ? '<span class="tag">automatic alerts on</span>' : `<button class="link" type="button" data-go="builds:alerts">${hubOn ? 'Turn on automatic alerts' : 'Get deal alerts for these'}</button>`}</p>

      <h2><span>To buy</span><span>${plan.shopping.length}</span></h2>
      ${plan.shopping.length ? `<ul class="list">${plan.shopping.map((g, i) => {
        const w = watched.get(lc(watchLabel(g)));
        return `<li><div class="row" style="cursor:default">
          <span><span class="name">${esc(g.label)} <span class="muted">×${g.count}</span></span>
          <span class="sub">${esc(g.detail !== g.label ? g.detail : KIND[g.kind] || '')}<br>for ${esc(g.builds.join(', '))}${g.typical ? ` · usually ${pounds(g.typical)} (${plural(g.seen, 'past buy')})` : ''}</span>
          <span class="acts">${ext(PL.ebayUrl(g.item, { max: w && w.alertAt }), 'Find')} · ${g.item.cat !== 'generic' ? ext(PL.ebayUrl(g.item, { sold: true }), 'Sold prices') + ' · ' : ''}
          ${w ? `<span class="tag">watching ≤ ${pounds(w.alertAt)}</span>` : `<button class="link" type="button" data-watch="${i}">Watch</button>`}</span>
          <span class="watchform" id="wf-${i}" hidden>Alert me at or under £ <input type="number" min="1" step="1" value="${g.typical ? Math.floor(g.typical * (1 - prefs().greatPct / 100)) : ''}" aria-label="Alert price">
            <button class="btn sm" type="button" data-addwatch="${i}">Add</button><span class="small" id="wr-${i}"></span></span></span>
          </div></li>`;
      }).join('')}</ul>` : '<p class="muted">Nothing to buy: stock covers every build’s missing parts.</p>'}

      <h2><span>Use from stock</span><span>${stockBuilds.length}</span></h2>
      ${stockBuilds.map(b => `<div class="build" style="cursor:default">
          <div class="head"><span class="id">${esc(b.id)}</span><span class="stage">${esc(b.platform.name)}</span></div>
          <ul class="plain">${b.needs.filter(n => n.stock).map(n => `<li><b>${esc(KIND[n.kind] || n.text)}</b>: ${n.stock.map(s => `${esc(s.partId)} ${esc(s.text)}`).join(' + ')}
            ${n.stock.some(s => s.check) ? '<span class="tag warn">check it fits</span>' : ''}${n.smaller ? '<span class="tag warn">smaller than suggested</span>' : ''}</li>`).join('')}</ul>
          <div class="actions" style="margin-top:6px"><button class="btn sm" type="button" data-alloc="${esc(b.id)}">Allocate to ${esc(b.id)}</button><span class="small" id="ar-${esc(b.id)}"></span></div>
        </div>`).join('') || '<p class="muted">No unallocated stock fits a build right now.</p>'}

      <details class="more"><summary>Build by build</summary>
        ${plan.builds.map(b => `<div class="build" style="cursor:default">
          <div class="head"><span class="id">${esc(b.id)}</span><span class="stage">${esc(b.platform.name)}${b.platform.from === 'suggested' || b.platform.from === 'guess' ? ' (suggested)' : ''}</span></div>
          <div class="desc">${esc(b.description)}</div>
          ${b.needs.length ? `<ul class="plain">${b.needs.map(n => `<li>${n.stock ? '✓' : '○'} <b>${esc(KIND[n.kind] || n.text)}</b>: ${esc(n.label)}${n.note ? ` <span class="muted">(${esc(n.note)})</span>` : ''}${n.stock ? ` <span class="muted">→ ${n.stock.map(s => esc(s.partId)).join(' + ')}</span>` : ''}</li>`).join('')}</ul>` : '<div class="small muted">Nothing listed in Missing Parts.</div>'}
          ${b.notes.map(t => `<div class="small" style="color:var(--urgent)">${esc(t)}</div>`).join('')}
        </div>`).join('')}
      </details>`;

    view.querySelectorAll('[data-watch]').forEach(b => b.addEventListener('click', () => {
      if (!isHubSet()) { go('settings'); return; }
      const f = document.getElementById('wf-' + b.dataset.watch);
      f.hidden = !f.hidden;
      if (!f.hidden) f.querySelector('input').focus();
    }));
    view.querySelectorAll('[data-addwatch]').forEach(b => b.addEventListener('click', () => addShopWatch(plan.shopping[+b.dataset.addwatch], +b.dataset.addwatch)));
    view.querySelectorAll('[data-alloc]').forEach(b => b.addEventListener('click', () => allocateBuild(b.dataset.alloc, b)));
    if (hubOn && !hubCache) loadHub().then(() => { if (tab === 'builds') render(); }).catch(() => {});
  }

  // ---- automatic alerts for builds --------------------------------------------

  const PRS = DealPairings;
  let autoBusy = false;
  let autoMsg = '';
  let autoTried = '';

  function autoFor(plan) {
    const auto = PL.needsForAlerts(plan, purchases());
    auto.hash = JSON.stringify([auto.needs.map(n => [n.id, n.builds]), auto.parts.map(g => [g.label, g.count])]);
    return auto;
  }

  function optionsHtml(n) {
    if (n.type === 'gpu') return esc(PRS.acceptableGpus(n.cpu).map(g => DealSpecs.label(g)).join(', '));
    const byPf = new Map();
    PRS.acceptableCpus(n.gpu, { ddr: n.ddr, platform: n.platform }).forEach(c => {
      if (!byPf.has(c.platform)) byPf.set(c.platform, []);
      byPf.get(c.platform).push(c.name);
    });
    return [...byPf.entries()].map(([pf, names]) => `<div><b>${esc(PRS.PLATFORMS[pf].name)}</b>: ${esc(names.join(', '))}</div>`).join('');
  }

  function autoHtml(plan) {
    const auto = autoFor(plan);
    const p = prefs();
    const on = !!p.autoAlerts;
    const count = auto.needs.length + auto.parts.length;
    return `<div class="auto ${on ? 'on' : ''}">
      <div class="head"><b>Automatic alerts for builds</b><span class="tag ${on ? '' : 'warn'}">${on ? 'on' : 'off'}</span></div>
      ${count ? `<ul class="plain">${auto.needs.map((n, i) => `<li><b>${esc(n.label)}</b>
          <details class="opts"><summary>${n.type === 'gpu' ? PRS.acceptableGpus(n.cpu).length + ' cards that suit it' : PRS.acceptableCpus(n.gpu, { ddr: n.ddr, platform: n.platform }).length + ' CPUs that pair well'}${n.type === 'combo' ? ', with any board that fits' : ''}</summary>${optionsHtml(n)}</details></li>`).join('')}
          ${auto.parts.map(g => `<li>${esc(g.label)} <span class="muted">×${g.count}</span></li>`).join('')}</ul>`
        : '<div class="small muted">Nothing to look for: stock covers every build.</div>'}
      <div class="small muted">Alerts when a CPU + board bundle is ${hubSettings().bundlePct}% or more under its parts’ value, or a single CPU or card ${hubSettings().partPct}% under (<button class="link" type="button" data-go="settings">change</button>). Values come from your purchase history and the sold prices the Chrome extension collects.</div>
      <div class="actions" style="margin-top:6px">
        ${on ? `<button class="btn sm ghost" type="button" id="auto-update">Update now</button><button class="link" type="button" id="auto-off">Turn off</button>`
          : `<button class="btn sm" type="button" id="auto-on" ${count ? '' : 'disabled'}>Turn on automatic alerts</button>`}
        <span class="small" id="auto-msg">${esc(autoMsg || (on && p.autoAt ? 'Updated ' + new Date(p.autoAt).toLocaleString('en-GB', { day: 'numeric', month: 'short', hour: '2-digit', minute: '2-digit' }) : ''))}</span>
      </div></div>`;
  }

  function hubSettings() {
    return Object.assign({ bundlePct: 35, partPct: 25 }, (hubCache && hubCache.settings) || {});
  }

  function wireAuto(plan) {
    if (!isHubSet()) return;
    const go = () => pushAuto(plan, true);
    const on = document.getElementById('auto-on');
    const upd = document.getElementById('auto-update');
    const off = document.getElementById('auto-off');
    if (on) on.addEventListener('click', go);
    if (upd) upd.addEventListener('click', go);
    if (off) off.addEventListener('click', () => pushAuto(plan, false));
  }

  // Keeps the alerts sheet in step as builds change (after every refresh), whichever tab is open.
  function autoSync() {
    if (!model || !model.builds || !isHubSet() || autoBusy) return;
    const p = prefs();
    if (!p.autoAlerts) return;
    let plan;
    try { plan = PL.plan(builds(), purchases()); } catch (e) { return; }
    const h = autoFor(plan).hash;
    if (h !== p.autoHash && h !== autoTried) { autoTried = h; pushAuto(plan, true, true); }
  }

  async function pushAuto(plan, turnOn, quiet) {
    if (autoBusy) return;
    autoBusy = true;
    const msg = document.getElementById('auto-msg');
    if (msg) msg.textContent = turnOn ? 'Updating your alerts sheet…' : 'Turning off…';
    try {
      const auto = autoFor(plan);
      const keep = new Set();
      if (turnOn) {
        await hubCall({ action: 'needs.replace', needs: auto.needs });
        const seeds = auto.seeds.filter(x => x.median);
        if (seeds.length) await hubCall({ action: 'prices.put', prices: seeds });
        const entries = auto.parts.map(g => Object.assign(PL.watchEntry(g, null, prefs().greatPct), { source: 'auto' }));
        entries.forEach(e => keep.add(e.label.toLowerCase()));
        if (entries.length) hubCache = await hubCall({ action: 'put', watches: entries });
      } else {
        await hubCall({ action: 'needs.replace', needs: [] });
      }
      if (!hubCache) hubCache = await hubCall({ action: 'list' });
      const stale = (hubCache.watches || []).filter(w => w.source === 'auto' && !keep.has(String(w.label).toLowerCase())).map(w => w.label);
      if (stale.length) hubCache = await hubCall({ action: 'remove', labels: stale });
      else hubCache = await hubCall({ action: 'list' });
      await savePrefs(turnOn ? { autoAlerts: true, autoHash: auto.hash, autoAt: Date.now() } : { autoAlerts: false, autoHash: '' });
      autoMsg = turnOn ? `Updated: ${plural(auto.needs.length, 'build search', 'build searches')} and ${plural(auto.parts.length, 'part watch', 'part watches')}.` : 'Automatic alerts are off.';
    } catch (e) {
      autoMsg = e.message || String(e);
    } finally {
      autoBusy = false;
    }
    if (tab === 'builds' && !(quiet && !autoMsg)) render();
  }

  const watchLabel = g => (g.item.cat === 'generic' ? g.label : DealSpecs.label(g.item));

  async function addShopWatch(group, i) {
    const out = document.getElementById('wr-' + i);
    const v = parseFloat(document.querySelector(`#wf-${i} input`).value);
    out.textContent = ' Adding…';
    try {
      const entry = PL.watchEntry(group, Number.isFinite(v) && v > 0 ? v : null, prefs().greatPct);
      if (entry.alertAt == null) { out.textContent = ' Enter an alert price (there’s no past price for this yet).'; return; }
      hubCache = await hubCall({ action: 'put', watches: [entry] });
      out.textContent = ' Added. You’ll get a notification when one is listed at or under this.';
      setTimeout(() => { if (tab === 'builds') render(); }, 1200);
    } catch (e) {
      out.textContent = ' ' + (e.message || e);
    }
  }

  // Writes Allocation for the stock parts suggested for a build, and takes those
  // parts out of the build's Missing Parts.
  async function allocateBuild(buildId, btn) {
    const plan = model.plan || PL.plan(builds(), purchases());
    const b = plan.builds.find(x => x.id === buildId);
    if (!b) return;
    const picks = b.needs.filter(n => n.stock);
    const out = document.getElementById('ar-' + buildId);
    btn.disabled = true;
    out.textContent = ' Allocating…';
    try {
      const done = [];
      const skipped = [];
      await Excel.run(async ctx => {
        const pt = await readTable(ctx, model.purchases.name, false);
        const ix = C.colIndex(pt.headers);
        const idC = ix('Part ID'), alC = ix('Allocation');
        const t = ctx.workbook.tables.getItem(pt.name);
        const filled = new Set();
        picks.forEach(n => {
          let all = true;
          n.stock.forEach(s => {
            const row = pt.values.findIndex(r => String(r[idC]).trim() === s.partId);
            if (row < 0 || String(pt.values[row][alC] || '').trim()) { skipped.push(s.partId); all = false; return; }
            t.getDataBodyRange().getCell(row, alC).values = [[buildId]];
            done.push(s.partId);
          });
          if (all) filled.add(n.kind);
        });
        await ctx.sync();
        if (model.builds && filled.size) {
          const bt = await readTable(ctx, model.builds.name, false);
          const bx = C.colIndex(bt.headers);
          const row = bt.values.findIndex(r => String(r[bx('Build ID')]).trim() === buildId);
          if (row >= 0 && bx('Missing Parts') >= 0) {
            const before = PL.missingKinds(bt.values[row][bx('Missing Parts')]);
            const left = before.filter(m => !filled.has(m.kind)).map(m => m.text).join(', ');
            const bb = ctx.workbook.tables.getItem(bt.name).getDataBodyRange();
            bb.getCell(row, bx('Missing Parts')).values = [[left]];
            if (!left && bx('All Parts Required to Complete?') >= 0) bb.getCell(row, bx('All Parts Required to Complete?')).values = [['Y']];
            await ctx.sync();
          }
        }
      });
      await refresh();
      tab = 'builds'; buildsSeg = 'shop';
      render();
      const msg = document.createElement('div');
      msg.className = 'ok-box';
      msg.innerHTML = `<strong>${esc(buildId)}</strong>: allocated ${esc(done.join(', ')) || 'nothing'}.${skipped.length ? ` Skipped ${esc(skipped.join(', '))} (already allocated).` : ''} Missing Parts updated.`;
      view.prepend(msg);
    } catch (e) {
      btn.disabled = false;
      out.textContent = ' ' + (e.message || e);
    }
  }

  // ---------------------------------------------------------------------------
  // Sell: listing drafts from builds
  // ---------------------------------------------------------------------------

  function renderSell() {
    if (!model.builds) return noBuilds();
    const all = builds().filter(b => String(b['Build ID'] || '').trim());
    if (!all.length) { view.innerHTML = '<p class="muted">No builds yet.</p>'; return; }
    const sales = salesRows();
    const soldIds = new Set(sales.filter(s => lc(s['Sold?']) === 'y').map(s => SA.idFrom(s['Build / Part ID'])));
    const ready = all.filter(b => lc(b['Built?']) === 'y' && !soldIds.has(SA.idFrom(b['Build ID'])));
    const rest = all.filter(b => !ready.includes(b) && !soldIds.has(SA.idFrom(b['Build ID'])));
    if (!sellId || !all.some(b => b['Build ID'] === sellId)) sellId = (ready[0] || rest[0] || all[0])['Build ID'];
    const build = all.find(b => b['Build ID'] === sellId);
    const p = prefs();
    const fee = LI.feeShare(sales);
    const d = LI.draft(build, purchases(), { marginPct: p.marginPct, os: p.os, feePct: fee || 0 });
    const ed = sellEdits[sellId] || {};
    const title = ed.title != null ? ed.title : d.title;
    const saleRow = sales.find(s => SA.idFrom(s['Build / Part ID']) === SA.idFrom(sellId));
    // Price: your edit, else the Listing Price in Sales, else the suggestion.
    const listedPrice = saleRow && typeof saleRow['Listing Price'] === 'number' && saleRow['Listing Price'] > 0 ? saleRow['Listing Price'] : null;
    const price = ed.price != null ? ed.price : (listedPrice || d.price);
    // The image always shows the Sales sheet's Listing Price when there is one.
    sheetPrice = listedPrice || (typeof price === 'number' && price > 0 ? price : null);
    sheetPriceFrom = listedPrice ? 'sales' : 'panel';
    const desc = ed.description != null ? ed.description : buildDescription(build, d, sheetPrice);
    const listed = saleRow && lc(saleRow['Listed?']) === 'y';
    const specText = d.specifics.map(([k, v]) => `${k}: ${v}`).join('\n');
    const opt = b => `<option value="${esc(b['Build ID'])}" ${b['Build ID'] === sellId ? 'selected' : ''}>${esc(b['Build ID'])} · ${esc(String(b['Build Description'] || '').slice(0, 40))}</option>`;

    const head = `
      <label class="field pick"><span>Build</span><select id="sell-pick">
        ${ready.length ? `<optgroup label="Built, not sold">${ready.map(opt).join('')}</optgroup>` : ''}
        ${rest.length ? `<optgroup label="Still being built">${rest.map(opt).join('')}</optgroup>` : ''}
      </select></label>
      ${segHtml([['listing', 'Listing'], ['photo', 'Spec sheet image']], sellSeg)}
      ${d.missing.length ? `<div class="warn-box">No ${esc(d.missing.join(', '))} found for ${esc(sellId)}. Allocate the parts in Purchases &amp; Parts (or add them to the Build Description) and the draft fills in.</div>` : ''}
      ${listed ? `<div class="ok-box small" style="margin:0 0 10px">Listed on ${esc(saleRow['Listed Date'] ? nice(C.serialToIso(saleRow['Listed Date'])) : '?')} at ${money(saleRow['Listing Price'])}.</div>` : ''}`;
    if (sellSeg === 'photo') {
      view.innerHTML = head + sheetHtml(build);
      document.getElementById('sell-pick').addEventListener('change', e => { sellId = e.target.value; renderSell(); });
      wireSeg(x => { sellSeg = x; });
      wireSheet(build);
      return;
    }
    view.innerHTML = head + `

      <div class="field"><span>Title <small class="muted" id="t-count">${title.length}/80</small></span>
        <div class="with-btn"><input id="s-title" maxlength="80" value="${esc(title)}"><button class="btn ghost sm" type="button" data-copy="title">Copy</button></div></div>

      <div class="field"><span>Custom label (SKU)</span>
        <div class="with-btn"><input value="${esc(sellId)}" readonly><button class="btn ghost sm" type="button" data-copy="sku">Copy</button></div>
        <small class="muted">Put this in eBay’s <b>Custom label (SKU)</b> box. It’s how sales get matched back to ${esc(sellId)} automatically.</small></div>

      <div class="field"><span>Price</span>
        <div class="with-btn"><span class="pfx">£</span><input id="s-price" type="number" min="1" step="1" value="${price == null ? '' : price}"><button class="btn ghost sm" type="button" data-copy="price">Copy</button></div>
        <small class="muted">${listedPrice && ed.price == null ? `From Listing Price in Sales. ` : ''}Cost ${money(d.cost)} + ${p.marginPct}% margin${fee ? ` + ${fee}% fees` : ''} → ${pounds(d.priceFromCost)}${d.priceFromValue ? ` · parts’ value → ${pounds(d.priceFromValue)}` : ''}${typeof price === 'number' && price > 0 ? ` · profit ${pounds(price - d.cost)}` : ''}${d.similarUrl ? ` · ${ext(d.similarUrl, 'similar PCs sold')}` : ''}</small></div>

      <div class="field"><span>Item specifics <button class="link" type="button" data-copy="specs">Copy all</button></span>
        <table class="grid specs">${d.specifics.map(([k, v]) => `<tr><td>${esc(k)}</td><td>${esc(v)}</td></tr>`).join('')}</table>
        <small class="muted">Category: Desktops &amp; All-In-Ones (${LI.EBAY_CATEGORY}). Condition: Used. eBay may offer slightly different wording; pick the closest.</small></div>

      <div class="field"><span>Description <button class="link" type="button" data-copy="desc">Copy</button></span>
        <textarea id="s-desc" rows="12">${esc(desc)}</textarea></div>

      <div class="actions">
        <button class="btn" type="button" id="s-listed">${listed ? 'Update listing price' : 'Mark as listed'}</button>
        <button class="link" type="button" id="s-csv">Seller Hub drafts file</button>
        <button class="link" type="button" id="s-reset">Reset draft</button>
      </div>
      <div id="s-out"></div>
      <p class="small muted">Description wording, margin and Windows edition are in <button class="link" type="button" data-go="settings">Settings</button>.</p>`;

    const keep = (k, v) => { sellEdits[sellId] = Object.assign({}, sellEdits[sellId], { [k]: v }); };
    document.getElementById('sell-pick').addEventListener('change', e => { sellId = e.target.value; renderSell(); });
    wireSeg(x => { sellSeg = x; });
    const t = document.getElementById('s-title');
    t.addEventListener('input', () => { keep('title', t.value); document.getElementById('t-count').textContent = `${t.value.length}/80`; });
    const pr = document.getElementById('s-price');
    pr.addEventListener('input', () => keep('price', pr.value === '' ? null : +pr.value));
    const ds = document.getElementById('s-desc');
    ds.addEventListener('input', () => keep('description', ds.value));
    view.querySelectorAll('[data-copy]').forEach(b => b.addEventListener('click', () => {
      const what = b.dataset.copy;
      copyText({ title: t.value, sku: sellId, price: pr.value, specs: specText, desc: ds.value }[what], b);
    }));
    document.getElementById('s-reset').addEventListener('click', () => { delete sellEdits[sellId]; renderSell(); });
    document.getElementById('s-listed').addEventListener('click', e => markListed(sellId, +pr.value, e.target));
    document.getElementById('s-csv').addEventListener('click', () => {
      const dd = Object.assign({}, d, { title: t.value, description: ds.value, price: pr.value === '' ? null : +pr.value });
      const csv = LI.draftsCsv([dd]);
      try { download(`ebay-drafts-${sellId}.csv`, csv); } catch (x) { /* some Excel versions block downloads */ }
      document.getElementById('s-out').innerHTML = `<div class="ok-box small">Saved <b>ebay-drafts-${esc(sellId)}.csv</b> to your Downloads. In Seller Hub: <b>Reports → Uploads → Upload template</b>, then finish the draft (photos) under <b>Listings → Drafts</b>.
        If nothing downloaded, <button class="link" type="button" id="s-csv-copy">copy the file’s text</button> into Notepad and save it as a .csv.</div>`;
      document.getElementById('s-csv-copy').addEventListener('click', ev => copyText(csv, ev.target));
    });
  }

  // ---- spec sheet image (for the listing photos) ----------------------------------

  const SS = DealSpecSheet;
  const sheetEdits = {};   // your changes to the specification lines, per build
  const photoKey = id => 'dealChecker.photo.' + id;
  const photoCache = {};
  let sheetPrice = null;      // the price shown on the image
  let sheetPriceFrom = 'panel';

  // The listing description from your template (or the default).
  function buildDescription(build, d, price) {
    const m = SS.model(build, purchases(), benchRows(), gameRows(), { os: prefs().os, osValue: +prefs().osValue || 0, price });
    const vars = LI.descriptionVars(d.parts, {
      os: prefs().os, title: (sellEdits[d.id] || {}).title || d.title, buildId: d.id,
      games: m.games, benchmarks: m.benchmarks, price, partsValue: m.totalValue
    });
    return LI.fillTemplate(prefs().descTemplate || LI.DEFAULT_TEMPLATE, vars);
  }

  function sheetModel(build) {
    const p = prefs();
    const m = SS.model(build, purchases(), benchRows(), gameRows(), { os: p.os, osValue: +p.osValue || 0, price: sheetPrice });
    const ed = sheetEdits[m.id] || {};
    m.specs = m.specs.map(r => Object.assign({}, r, { text: ed[r.label] != null ? ed[r.label] : r.text }));
    return m;
  }

  const sheetOpts = () => {
    const p = prefs();
    const foot = p.sheetFooter != null ? String(p.sheetFooter).split('|').map(x => x.trim()).filter(Boolean) : null;
    return { showValues: p.sheetValues !== false, footer: foot };
  };

  function sheetHtml(build) {
    const m = sheetModel(build);
    const p = prefs();
    const missing = [!model.benches && 'Build Benchmarking', !model.games && 'Build Game Testing'].filter(Boolean);
    const notes = [];
    if (p.sheetValues !== false && m.unvalued.length) notes.push(`No Estimated Part Value for some ${m.unvalued.join(', ')} parts, so the total leaves them out.`);
    if (p.sheetValues !== false && m.price && m.totalValue && m.price >= m.totalValue) notes.push(`Your price (£${m.price}) is above the parts’ value (£${m.totalValue}), so no saving is shown. If Windows is activated with a retail licence, adding its value in Settings counts it.`);
    return `<div class="sheet-sec">
      <div class="photo-row">
        <label class="file"><input type="file" id="sh-folder" webkitdirectory directory multiple><span class="btn ghost sm">Find photos in a folder…</span></label>
        <label class="file"><input type="file" id="sh-photo" accept="image/*"><span class="link">${loadPhoto(m.id) ? 'or choose a different photo' : 'or choose one photo'}</span></label>
        ${loadPhoto(m.id) ? '<button class="link" type="button" id="sh-nophoto">Remove photo</button>' : ''}
      </div>
      ${loadPhoto(m.id) ? '' : '<p class="small muted">Pick your <b>Computing Business Files</b> folder once: every build folder (“Build 67”, “B69”…) with a <b>Thumbnail.jpg</b> is matched to its build and remembered on this computer.</p>'}
      <div id="sh-found"></div>
      ${notes.map(n => `<div class="warn-box">${esc(n)}</div>`).join('')}
      <div class="sheet-preview"><canvas id="sh-canvas" aria-label="Spec sheet preview"></canvas></div>
      <p class="small muted">${sheetPrice ? (sheetPriceFrom === 'sales' ? `Price on the image: £${Math.round(sheetPrice)}, the Listing Price for ${esc(m.id)} in Sales.` : `${esc(m.id)} has no Listing Price in Sales yet, so the image uses the Listing tab’s price (£${Math.round(sheetPrice)}). <b>Mark as listed</b> saves it to Sales.`) : 'No price yet: set one on the Listing tab or in Sales.'}
        ${missing.length ? ` Benchmarks and FPS come from the <b>${missing.join(' and ')}</b> tables, which aren’t in this workbook.` : ''}</p>
      <div class="actions">
        <button class="btn" type="button" id="sh-download">Download image</button>
        <button class="btn ghost" type="button" id="sh-copy">Copy image</button>
        <button class="link" type="button" id="sh-place">Put it on a sheet</button>
      </div>
      <div id="sh-out"></div>
      <details class="more"><summary>Edit the lines on the image</summary>
        ${m.specs.map(r => `<label class="field"><span>${esc(r.label)}</span><input data-spec="${esc(r.label)}" value="${esc(r.text)}" placeholder="(left off the image when empty)"></label>`).join('')}
        <button class="link" type="button" id="sh-reset">Reset lines to the sheet’s parts</button>
        <p class="small muted">Values, the Windows licence and the points along the bottom are in <button class="link" type="button" data-go="settings">Settings</button>.</p>
      </details>
      ${!m.benchmarks.length && !m.games.length ? `<p class="small muted">No benchmarks or game results for ${esc(m.id)} yet: add rows with Build ID ${esc(m.id)} to those sheets and click ⟳.</p>` : ''}
    </div>`;
  }

  function loadPhoto(id) {
    if (photoCache[id] !== undefined) return photoCache[id];
    try { photoCache[id] = localStorage.getItem(photoKey(id)) || null; } catch (e) { photoCache[id] = null; }
    return photoCache[id];
  }

  function savePhoto(id, data) {
    photoCache[id] = data;
    try { localStorage.setItem(photoKey(id), data); return true; } catch (x) { return false; }   // full: still used this session
  }

  // Shrinks the photo (long side 1400px) so it can be remembered. The browser turns
  // phone photos the right way up from their orientation tag.
  function readPhoto(file) {
    return new Promise((resolve, reject) => {
      const img = new Image();
      const url = URL.createObjectURL(file);
      img.onload = () => {
        const k = Math.min(1, 1400 / Math.max(img.naturalWidth, img.naturalHeight));
        const c = document.createElement('canvas');
        c.width = Math.round(img.naturalWidth * k);
        c.height = Math.round(img.naturalHeight * k);
        c.getContext('2d').drawImage(img, 0, 0, c.width, c.height);
        URL.revokeObjectURL(url);
        resolve(c.toDataURL('image/jpeg', 0.86));
      };
      img.onerror = () => { URL.revokeObjectURL(url); reject(new Error(`${file.name} isn’t a picture Excel can read. Save it as a JPG or PNG.`)); };
      img.src = url;
    });
  }

  function imageFrom(dataUrl) {
    return new Promise(resolve => {
      if (!dataUrl) { resolve(null); return; }
      const img = new Image();
      img.onload = () => resolve(img);
      img.onerror = () => resolve(null);
      img.src = dataUrl;
    });
  }

  async function drawSheet(build) {
    const canvas = document.getElementById('sh-canvas');
    if (!canvas) return null;
    await SS.loadFonts();
    const m = sheetModel(build);
    const photo = await imageFrom(loadPhoto(m.id));
    SS.draw(canvas, m, photo, sheetOpts());
    return canvas;
  }

  const canvasBlob = canvas => new Promise(res => canvas.toBlob(res, 'image/png'));

  // A folder picked with "Find photos in a folder…": every build's Thumbnail.jpg.
  async function takePhotoFolder(files, current) {
    const found = SS.findThumbnails(files);
    const ids = new Set(builds().map(b => String(b['Build ID']).trim()));
    const mine = Object.keys(found).filter(id => ids.has(id));
    const box = document.getElementById('sh-found');
    if (!Object.keys(found).length) {
      box.innerHTML = '<div class="warn-box">No Thumbnail.jpg found in folders named like “Build 69” or “B69” in that folder.</div>';
      return;
    }
    box.innerHTML = `<div class="small muted">Reading ${mine.length} photo${mine.length === 1 ? '' : 's'}…</div>`;
    const done = [];
    const failed = [];
    let full = false;
    for (const id of mine) {
      try {
        const data = await readPhoto(found[id]);
        if (!savePhoto(id, data)) full = true;
        done.push(id);
      } catch (e) { failed.push(`${id} (${e.message})`); }
    }
    const missing = [...ids].filter(id => !found[id]);
    renderSell();
    document.getElementById('sh-found').innerHTML = `<div class="ok-box small">Photos found for ${done.length ? esc(done.join(', ')) : 'none of your current builds'}.
      ${missing.length ? `<br>No Thumbnail.jpg yet for ${esc(missing.join(', '))}.` : ''}${failed.length ? `<br>Couldn’t read ${esc(failed.join('; '))}.` : ''}
      ${full ? '<br>Too many photos to remember them all on this computer; the rest work until Excel closes.' : ''}</div>`;
  }

  function wireSheet(build) {
    const id = String(build['Build ID']).trim();
    const out = document.getElementById('sh-out');
    if (!document.getElementById('sh-canvas')) return;
    drawSheet(build);
    const redraw = (() => { let t = null; return () => { clearTimeout(t); t = setTimeout(() => drawSheet(build), 250); }; })();
    document.getElementById('sh-folder').addEventListener('change', e => { if (e.target.files && e.target.files.length) takePhotoFolder(e.target.files, id); });
    document.getElementById('sh-photo').addEventListener('change', async e => {
      const f = e.target.files && e.target.files[0];
      if (!f) return;
      try { savePhoto(id, await readPhoto(f)); renderSell(); } catch (x) { out.innerHTML = `<div class="err-box">${esc(x.message)}</div>`; }
    });
    const rm = document.getElementById('sh-nophoto');
    if (rm) rm.addEventListener('click', () => { photoCache[id] = null; try { localStorage.removeItem(photoKey(id)); } catch (x) { /* fine */ } renderSell(); });
    view.querySelectorAll('[data-spec]').forEach(inp => inp.addEventListener('input', () => {
      sheetEdits[id] = Object.assign({}, sheetEdits[id], { [inp.dataset.spec]: inp.value });
      redraw();
    }));
    document.getElementById('sh-reset').addEventListener('click', () => { delete sheetEdits[id]; renderSell(); });
    document.getElementById('sh-download').addEventListener('click', async () => {
      const c = await drawSheet(build);
      const blob = await canvasBlob(c);
      try { download(`${id}-spec-sheet.png`, blob); } catch (x) { /* some Excel versions block downloads */ }
      out.innerHTML = `<div class="ok-box small">Saved <b>${esc(id)}-spec-sheet.png</b> to your Downloads (${c.width} × ${c.height}px). If nothing downloaded, use <b>Copy image</b> or <b>Put it on a sheet</b>.</div>`;
    });
    document.getElementById('sh-copy').addEventListener('click', async () => {
      const c = await drawSheet(build);
      try {
        const blob = await canvasBlob(c);
        await navigator.clipboard.write([new ClipboardItem({ 'image/png': blob })]);
        out.innerHTML = '<div class="ok-box small">Copied. Paste it into Paint or Photos and save it from there.</div>';
      } catch (x) {
        out.innerHTML = '<div class="err-box small">Excel didn’t allow copying a picture here. Use <b>Put it on a sheet</b> instead.</div>';
      }
    });
    document.getElementById('sh-place').addEventListener('click', async e => {
      const c = await drawSheet(build);
      e.target.disabled = true;
      try {
        await placeOnSheet(id, c);
        out.innerHTML = `<div class="ok-box small">Added to the <b>Spec Sheets</b> sheet. Right-click the picture and choose <b>Save as Picture…</b> to save it.</div>`;
      } catch (x) {
        out.innerHTML = `<div class="err-box small">${esc(x.message || x)}</div>`;
      }
      e.target.disabled = false;
    });
  }

  // Puts the image on a "Spec Sheets" worksheet (one picture per build, newest at the bottom).
  async function placeOnSheet(id, canvas) {
    if (!Office.context.requirements || !Office.context.requirements.isSetSupported('ExcelApi', '1.9')) {
      throw new Error('This version of Excel can’t add pictures from the panel. Use Download image or Copy image instead.');
    }
    const b64 = canvas.toDataURL('image/png').split(',')[1];
    await Excel.run(async ctx => {
      let ws = ctx.workbook.worksheets.getItemOrNullObject('Spec Sheets');
      await ctx.sync();
      if (ws.isNullObject) ws = ctx.workbook.worksheets.add('Spec Sheets');
      const shapes = ws.shapes;
      shapes.load('items/name,items/top,items/height');
      await ctx.sync();
      const name = 'SpecSheet_' + id;
      shapes.items.filter(sh => sh.name === name).forEach(sh => sh.delete());
      const others = shapes.items.filter(sh => sh.name !== name);
      const top = others.reduce((a, sh) => Math.max(a, sh.top + sh.height), 0) + (others.length ? 20 : 10);
      const pic = shapes.addImage(b64);
      pic.name = name;
      pic.left = 10;
      pic.top = top;
      pic.width = 600;   // points; keeps its proportions
      ws.activate();
      await ctx.sync();
    });
  }

  async function markListed(id, price, btn) {
    const out = document.getElementById('s-out');
    if (!model.sales) { out.innerHTML = '<div class="err-box">Couldn’t find the Sales table (needs “Build / Part ID”, “Sold?” and “Sold Price” columns).</div>'; return; }
    if (!(price > 0)) { out.innerHTML = '<div class="err-box">Enter the listing price first.</div>'; return; }
    btn.disabled = true;
    try {
      let added = false;
      await Excel.run(async ctx => {
        const st = await readTable(ctx, model.sales.name, true);
        const ix = C.colIndex(st.headers);
        const today = C.todaySerial();
        const row = st.values.findIndex(r => SA.idFrom(r[ix('Build / Part ID')]) === SA.idFrom(id));
        if (row >= 0) {
          const body = ctx.workbook.tables.getItem(st.name).getDataBodyRange();
          const set = (col, v) => { if (ix(col) >= 0) body.getCell(row, ix(col)).values = [[v]]; };
          set('Listed?', 'Y');
          if (!st.values[row][ix('Listed Date')]) set('Listed Date', today);
          set('Listing Price', price);
          await ctx.sync();
        } else {
          const f = { 'build / part id': id, 'listed?': 'Y', 'listed date': today, 'listing price': price, 'sold?': 'N' };
          await addRows(ctx, st, [st.headers.map(h => (f[lc(h)] !== undefined ? f[lc(h)] : ''))], ['Listed Date', 'Listing Price', 'Date of Sale', 'Sold Price']);
          added = true;
        }
      });
      await refresh();
      tab = 'sell'; sellSeg = 'listing';
      render();
      document.getElementById('s-out').innerHTML = `<div class="ok-box small">${esc(id)} ${added ? 'added to' : 'updated in'} Sales: listed today at ${money(price)}.</div>`;
    } catch (e) {
      btn.disabled = false;
      out.innerHTML = `<div class="err-box">${esc(e.message || e)}</div>`;
    }
  }

  // ---------------------------------------------------------------------------
  // Watch: deal alerts (the Google Sheet hub) and the watchlist
  // ---------------------------------------------------------------------------

  function isHubSet() { const p = prefs(); return HUB_RE.test(p.hubUrl) && p.hubKey.length >= 16; }

  async function hubCall(body, override) {
    const p = Object.assign(prefs(), override || {});
    if (!HUB_RE.test(p.hubUrl) || !p.hubKey) throw new Error('Connect your Deal Alerts sheet in Settings (⚙) first.');
    let res;
    try {
      res = await fetch(p.hubUrl, { method: 'POST', body: JSON.stringify(Object.assign({ key: p.hubKey }, body)) });
    } catch (e) {
      // The browser hides the reason; almost always the web app isn't open to "Anyone" or needs re-authorising.
      throw new Error('Couldn’t reach the alerts sheet. Open the web app URL in your browser: it should show {"ok":true,"hub":"deal-alerts"…}. If you see a Google sign-in or an authorisation message instead, redeploy the web app with Who has access: Anyone (Deploy → Manage deployments → edit → New version).');
    }
    const text = await res.text();
    let data;
    try { data = JSON.parse(text); } catch (e) {
      const title = (text.match(/<title>([^<]*)<\/title>/i) || [])[1] || '';
      const body = text.replace(/<style[\s\S]*?<\/style>|<script[\s\S]*?<\/script>/gi, ' ').replace(/<[^>]+>/g, ' ').replace(/\s+/g, ' ').trim().slice(0, 160);
      throw new Error(`The alerts sheet sent back a web page instead of data. Google said: “${(title + ' ' + body).trim().slice(0, 200)}”. In Apps Script, open Executions (left sidebar) to see the error.`);
    }
    if (!data.ok) throw new Error(data.error || 'The alerts sheet refused the request.');
    if (!Array.isArray(data.watches)) {
      // A reply without the watchlist means the sheet's doPost didn't run: an old deployment, or the request arrived as a plain page visit.
      throw new Error(`The alerts sheet answered, but without your watchlist (it said: ${text.slice(0, 90)}). In Apps Script, check the whole new Code.gs is saved, then Deploy → Manage deployments → edit → Version: New version → Deploy, and connect again.`);
    }
    data.alerts = Array.isArray(data.alerts) ? data.alerts : [];
    return data;
  }
  async function loadHub() { hubCache = await hubCall({ action: 'list' }); return hubCache; }

  // Builds › Alerts: automatic build searches, latest deals and the parts you watch.
  function renderWatch() {
    const top = buildsTop();
    if (!isHubSet()) {
      view.innerHTML = top + `
        <div class="auto"><b>Deal alerts on your phone</b>
          <p class="small" style="margin:4px 0 8px">Your Deal Alerts Google Sheet checks eBay every 15 minutes. It can look for CPU + board bundles and cards that suit each build automatically, and for any part you watch, and notify your phone when one is cheap enough.</p>
          <button class="btn sm" type="button" data-go="settings">Connect it in Settings</button></div>
        ${sheetWatchHtml()}`;
      return;
    }
    if (!hubCache) {
      view.innerHTML = top + '<p class="msg"><span class="spinner" aria-hidden="true"></span>Reading your alerts sheet…</p>';
      loadHub().then(() => { if (tab === 'builds' && buildsSeg === 'alerts') render(); }).catch(e => {
        if (!(tab === 'builds' && buildsSeg === 'alerts')) return;
        view.innerHTML = top + `<div class="err-box"><strong>Couldn’t reach your alerts sheet</strong><br>${esc(e.message || e)}</div>
          <div class="actions"><button class="link" id="h-retry" type="button">Try again</button><button class="link" type="button" data-go="settings">Check the connection in Settings</button></div>`;
        wireSeg(s => { buildsSeg = s; });
        document.getElementById('h-retry').addEventListener('click', () => render());
      });
      return;
    }
    const h = hubCache;
    const plan = model.plan;
    view.innerHTML = top + `
      ${h.paused ? '<div class="warn-box">Alerts are paused on the sheet (Deal alerts → Resume).</div>' : ''}
      ${autoHtml(plan)}
      <h2><span>Latest deals found</span><span>${h.alerts.length || ''}</span></h2>
      ${h.alerts.length ? `<ul class="list">${h.alerts.slice(0, 8).map(dealRow).join('')}</ul>` : '<p class="small muted">None yet. They appear here and on your phone as they’re found.</p>'}
      <details class="more"><summary>Parts you watch (${h.watches.length})</summary>
      ${h.watches.length ? `<ul class="list">${h.watches.map((w, i) => `<li><div class="row" style="cursor:default">
        <span><span class="name">${esc(w.label)}${w.active === false ? ' <span class="tag warn">off</span>' : ''}</span>
        <span class="sub">${w.median ? `median ${money(w.median)} from ${w.n || 0} sales` : 'not priced yet (the Chrome extension prices it next time you open an eBay listing)'}${w.source ? ` · from ${esc(w.source === 'excel' ? 'Excel' : w.source === 'chrome' ? 'Chrome' : w.source)}` : ''}
        · <button class="link" type="button" data-unwatch="${i}">Remove</button></span></span>
        <span class="days">${pounds(w.alertAt)}<small>${w.alertFixed ? 'your price' : 'alert at or under'}</small></span></div></li>`).join('')}</ul>`
        : '<p class="muted small">Nothing watched yet. Use <b>Watch</b> on the Shopping list, or the Watch button in Chrome.</p>'}
        <p class="small muted">Change alert prices or switch parts off on the Watchlist tab of your Deal Alerts sheet.</p>
      </details>
      <p class="small muted"><button class="link" type="button" id="h-reload">Refresh alerts</button></p>`;
    wireAuto(plan);
    document.getElementById('h-reload').addEventListener('click', () => { hubCache = null; render(); });
    view.querySelectorAll('[data-unwatch]').forEach(b => b.addEventListener('click', async () => {
      b.textContent = 'Removing…';
      try { hubCache = await hubCall({ action: 'remove', labels: [h.watches[+b.dataset.unwatch].label] }); render(); } catch (x) { b.textContent = x.message; }
    }));
  }

  // The Watchlist sheet that comes across with purchase imports (works without the alerts sheet).
  function sheetWatchHtml() {
    if (!model.watch || !model.watch.values.length || !String(model.watch.values[0][0] || '').trim()) return '';
    const rows = C.toObjects(model.watch.headers, model.watch.values).filter(w => w['Part']);
    return `<h2><span>From your last Chrome import</span><span>${rows.length}</span></h2><ul class="list">${rows.map(w => `<li><div class="row" style="cursor:default">
        <span><span class="name">${esc(w['Part'])}</span><span class="sub">median ${money(w['Median Sold'])} from ${esc(w['Sales'] || 0)} sales · updated ${nice(C.serialToIso(w['Updated']))}</span></span>
        <span class="days">${pounds(w['Alert At'])}<small>alert at or under</small></span></div></li>`).join('')}</ul>`;
  }

  // ---------------------------------------------------------------------------
  // Settings (⚙): everything you set once. Saved in the workbook, so your partner gets it too.
  // ---------------------------------------------------------------------------

  function renderSettings() {
    const p = prefs();
    const hubOn = isHubSet();
    const h = hubCache;
    const hs = hubSettings();
    const fee = LI.feeShare(salesRows());
    view.innerHTML = `
      <div class="set-head"><h1>Settings</h1><span class="small ok-text" id="set-msg" aria-live="polite"></span><button class="btn sm" type="button" id="set-done">Done</button></div>

      <section class="set">
        <h2><span>Deal alerts</span>${hubOn ? '<span class="tag">connected</span>' : ''}</h2>
        ${hubOn ? `
          <p class="small" style="margin:0 0 8px">${h ? `${plural(h.watches.length, 'part')} watched${h.paused ? ' · <span class="tag warn">paused on the sheet</span>' : ''} · ` : ''}<button class="link" type="button" id="h-test">Send a test notification</button> · <button class="link" type="button" id="h-forget">Disconnect</button></p>
          <div class="field inline"><span>CPU + board bundle</span><input type="number" id="pct-bundle" min="5" max="80" value="${hs.bundlePct}" aria-label="Bundle percentage"> % under value</div>
          <div class="field inline"><span>Single CPU / card</span><input type="number" id="pct-part" min="5" max="80" value="${hs.partPct}" aria-label="Single part percentage"> % under value</div>
          <p class="small muted" style="margin:0 0 8px">How far under the parts’ value a CPU + board bundle, or a single CPU or graphics card, must be before your phone is alerted.</p>
          <div class="actions" style="margin-top:0"><button class="btn sm ghost" type="button" id="pct-save">Save alert levels</button><span class="small" id="pct-msg"></span></div>`
        : `<p class="small muted" style="margin:0 0 8px">Your Deal Alerts Google Sheet checks eBay every 15 minutes and notifies your phone. Both values are in the sheet: <b>Deal alerts → Show connection details</b>.</p>
          <label class="field"><span>Web app URL</span><input id="h-url" placeholder="https://script.google.com/macros/s/…/exec" value="${esc(p.hubUrl)}"></label>
          <label class="field"><span>Key</span><input id="h-key" type="password" value="${esc(p.hubKey)}"></label>
          <div class="actions" style="margin-top:0"><button class="btn sm" id="h-save" type="button">Connect</button><span id="h-out" class="small"></span></div>`}
      </section>

      <section class="set">
        <h2><span>Pricing</span></h2>
        <div class="field inline"><span>Margin on cost</span><input id="p-margin" type="number" min="0" max="300" value="${p.marginPct}"> %</div>
        <div class="field inline"><span>Windows</span><select id="p-os">${['Windows 11 Home', 'Windows 11 Pro'].map(o => `<option ${o === p.os ? 'selected' : ''}>${o}</option>`).join('')}</select></div>
        <div class="field inline"><span>Windows licence value</span>£ <input id="sh-osval" type="number" min="0" step="1" value="${+p.osValue || ''}" placeholder="0"></div>
        <p class="small muted" style="margin:0">Margin is for the suggested price on the Sell tab. Only give Windows a value if it’s a genuine retail licence: it’s added to the parts’ value on the image. ${fee ? `eBay fees: ${fee}% of the sale price on average, from your recorded sales.` : 'eBay fees are allowed for automatically once you record some eBay sales.'}</p>
      </section>

      <section class="set">
        <h2><span>Spec sheet image</span></h2>
        <label class="small chk-line"><input type="checkbox" id="sh-values" ${p.sheetValues !== false ? 'checked' : ''}> Show part values, total value and the saving</label>
        <label class="field"><span>Points along the bottom (separate with |)</span><input id="sh-footer" value="${esc(p.sheetFooter != null ? p.sheetFooter : 'Professionally built | Stress-tested & benchmarked | Windows 11 activated')}"></label>
      </section>

      <section class="set">
        <h2><span>Listing description</span>${p.descTemplate ? '<span class="tag">your wording</span>' : ''}</h2>
        <p class="small muted" style="margin:0 0 6px">Your wording for every listing. {Placeholders} are filled in for each build; a paragraph whose placeholder has nothing to show (e.g. no benchmarks yet) is left out.</p>
        <textarea id="tpl" rows="14" spellcheck="true">${esc(p.descTemplate || LI.DEFAULT_TEMPLATE)}</textarea>
        <div class="actions"><button class="btn sm ghost" type="button" id="tpl-save">Save wording</button><button class="link" type="button" id="tpl-default">Back to the default</button><span class="small" id="tpl-msg"></span></div>
        <details class="more"><summary>Placeholders</summary><div class="small muted">${LI.PLACEHOLDERS.map(([k, d]) => `<code>{${esc(k)}}</code> ${esc(d)}`).join(' · ')}</div></details>
      </section>
      <p class="small muted">Settings are saved in this workbook, so your business partner gets them too. ${esc(APP_NAME)} ${esc(APP_VERSION)}.</p>`;

    const saved = (el, text) => { if (el) { el.textContent = ' ' + text; setTimeout(() => { el.textContent = ''; }, 2500); } };
    document.getElementById('set-done').addEventListener('click', () => go(lastTab || 'home'));
    // Alerts
    if (!hubOn) {
      document.getElementById('h-save').addEventListener('click', async () => {
        const url = document.getElementById('h-url').value.trim();
        const key = document.getElementById('h-key').value.trim();
        const out = document.getElementById('h-out');
        if (!HUB_RE.test(url)) { out.textContent = 'The URL should look like https://script.google.com/macros/s/…/exec'; return; }
        out.textContent = 'Connecting…';
        try {
          hubCache = await hubCall({ action: 'list' }, { hubUrl: url, hubKey: key });
          await savePrefs({ hubUrl: url, hubKey: key, greatPct: hubCache.greatPct || 25 });
        } catch (e) { out.textContent = e.message || String(e); return; }
        renderSettings();
      });
    } else {
      if (!h) loadHub().then(() => { if (tab === 'settings') renderSettings(); }).catch(() => {});
      document.getElementById('h-test').addEventListener('click', async e => {
        e.target.textContent = 'Sending…';
        try { await hubCall({ action: 'test', from: 'Excel' }); e.target.textContent = 'Sent'; } catch (x) { e.target.textContent = x.message; }
      });
      document.getElementById('h-forget').addEventListener('click', async () => { await savePrefs({ hubUrl: '', hubKey: '' }); hubCache = null; renderSettings(); });
      document.getElementById('pct-save').addEventListener('click', async e => {
        const b = +document.getElementById('pct-bundle').value;
        const pp = +document.getElementById('pct-part').value;
        const msg = document.getElementById('pct-msg');
        if (!(b >= 5 && b <= 80 && pp >= 5 && pp <= 80)) { msg.textContent = ' Use a number between 5 and 80.'; return; }
        e.target.disabled = true;
        try { hubCache = await hubCall({ action: 'settings.put', bundlePct: b, partPct: pp }); saved(msg, 'Saved.'); } catch (x) { msg.textContent = ' ' + (x.message || x); }
        e.target.disabled = false;
      });
    }
    // Pricing and image (saved as soon as they change)
    const onChange = (id, fn) => document.getElementById(id).addEventListener('change', async e => { await fn(e.target); saved(document.getElementById('set-msg'), 'Saved'); });
    onChange('p-margin', async el => { await savePrefs({ marginPct: Math.max(0, +el.value || 0) }); Object.keys(sellEdits).forEach(k => { if (sellEdits[k]) delete sellEdits[k].price; }); });
    onChange('p-os', async el => { await savePrefs({ os: el.value }); Object.keys(sellEdits).forEach(k => delete sellEdits[k]); });
    onChange('sh-osval', async el => { await savePrefs({ osValue: Math.max(0, +el.value || 0) }); });
    onChange('sh-values', async el => { await savePrefs({ sheetValues: el.checked }); });
    onChange('sh-footer', async el => { await savePrefs({ sheetFooter: el.value }); });
    const saveTpl = async text => {
      await savePrefs({ descTemplate: text });
      Object.keys(sellEdits).forEach(k => { if (sellEdits[k]) delete sellEdits[k].description; });   // redraft with the new wording
      if (!text) document.getElementById('tpl').value = LI.DEFAULT_TEMPLATE;
      saved(document.getElementById('tpl-msg'), text ? 'Saved. Descriptions now use it.' : 'Back to the default wording.');
    };
    document.getElementById('tpl-save').addEventListener('click', () => saveTpl(document.getElementById('tpl').value));
    document.getElementById('tpl-default').addEventListener('click', () => saveTpl(''));
  }

  // ---------------------------------------------------------------------------
  // Import: purchases (pasted from Chrome) and sales (eBay report files)
  // ---------------------------------------------------------------------------

  function renderImport() {
    if (pending) return renderPreview();
    if (pendingSales) return renderSalesPreview();
    view.innerHTML = `
      <h2><span>Purchases from eBay</span></h2>
      <ol class="steps">
        <li>In Chrome, open <b>My eBay → Purchases</b>.</li>
        <li>Click <b>Send to spreadsheet</b>, then <b>Copy rows for Excel</b>.</li>
        <li>Click in the box below and press <b>Ctrl+V</b>.</li>
      </ol>
      <textarea class="paste" id="paste" placeholder="Paste here (Ctrl+V)" aria-label="Paste from the Deal Checker extension"></textarea>
      <div id="paste-msg"></div>

      <h2 style="margin-top:22px"><span>Sales from eBay</span></h2>
      <ol class="steps">
        <li>In Seller Hub, download the <b>Transaction report</b> (Payments → Reports) and/or the <b>Orders</b> report (Reports → Downloads) as CSV.</li>
        <li>Choose the file(s) below. You’ll see what will be recorded first.</li>
      </ol>
      <label class="file"><input type="file" id="sales-file" accept=".csv,text/csv" multiple><span class="btn ghost">Choose report files…</span></label>
      <p class="small muted">The Transaction report includes eBay’s fees, promoted listing fees and postage labels, so profit comes out right. Sales are matched by <b>Custom label</b> (your Build or Part ID), or by title if there isn’t one.</p>
      <div id="sales-msg"></div>`;
    const ta = document.getElementById('paste');
    ta.addEventListener('paste', e => {
      const text = (e.clipboardData || window.clipboardData).getData('text');
      e.preventDefault();
      takePaste(text);
    });
    ta.addEventListener('input', () => { if (ta.value.includes(C.PREFIX)) takePaste(ta.value); });
    document.getElementById('sales-file').addEventListener('change', e => takeSalesFiles(e.target.files));
  }

  function takePaste(text) {
    try {
      const data = C.parsePayload(text);
      if (C.missingColumns(model.purchases.headers).length) throw new Error('Click “Set up the sheet” above first (one time only).');
      const plan = C.planImport(model.purchases, data.rows);
      pending = { data, plan };
      renderPreview();
    } catch (e) {
      document.getElementById('paste-msg').innerHTML = `<div class="err-box">${esc(e.message || e)}</div>`;
    }
  }

  function renderPreview() {
    const { data, plan } = pending;
    const n = plan.add.length;
    view.innerHTML = `
      <h2><span>Ready to add</span><span>${n}</span></h2>
      ${n ? `<ul class="list">${plan.preview.map(p => `<li><div class="row" style="cursor:default">
          <span><span class="name">${esc(p.partId)} ${esc(p.brand === 'No Brand' ? '' : p.brand)} ${esc(p.description)}</span>
          <span class="sub">${esc(p.category)} · bought ${nice(p.purchaseDate)} · ${p.deliveryDate ? `${p.deliveryEstimated ? 'due' : 'delivered'} ${nice(p.deliveryDate)}` : 'no delivery date'}${p.notes ? ' · ' + esc(p.notes) : ''}</span></span>
          <span class="days">${money(p.cost)}<small>${typeof p.marketValue === 'number' ? 'value ' + money(p.marketValue) : 'no value'}</small></span></div></li>`).join('')}</ul>` : ''}
      <p class="small muted">${plan.skipped ? `${plural(plan.skipped, 'row is', 'rows are')} already in the sheet and will be skipped. ` : ''}${plan.updates.length ? `${plural(plan.updates.length, 'delivery date', 'delivery dates')} will be updated. ` : ''}${data.watchlist.length ? `Your watchlist (${plural(data.watchlist.length, 'part')}) will be updated on the Watchlist sheet.` : ''}</p>
      <div class="actions">
        <button class="btn" id="go" type="button" ${n || plan.updates.length || data.watchlist.length ? '' : 'disabled'}>${n ? `Add ${plural(n, 'row')} to Purchases & Parts` : 'Update the sheet'}</button>
        <button class="link" id="cancel" type="button">Cancel</button>
      </div>
      <div id="result"></div>`;
    document.getElementById('cancel').addEventListener('click', () => { pending = null; renderImport(); });
    document.getElementById('go').addEventListener('click', doImport);
  }

  async function doImport(e) {
    const btn = e.target;
    btn.disabled = true;
    btn.innerHTML = '<span class="spinner" aria-hidden="true"></span>Adding…';
    const { data } = pending;
    try {
      let added = [];
      await Excel.run(async ctx => {
        // Re-read just before writing, in case your partner changed something.
        const fresh = await readTable(ctx, model.purchases.name, true);
        const plan = C.planImport(fresh, data.rows);
        const t = ctx.workbook.tables.getItem(fresh.name);
        if (plan.add.length) {
          await addRows(ctx, fresh, plan.add, ['Purchase Date', 'Delivery Date']);
          added = plan.preview.map(p => p.partId);
        }
        plan.updates.forEach(u => { t.getDataBodyRange().getCell(u.row, u.col).values = [[u.value]]; });
        if (data.watchlist.length) await writeWatchlist(ctx, data.watchlist);
        await ctx.sync();

        // Remember which delivery dates are only estimates.
        const est = Object.assign({}, estimatedKeys());
        data.rows.forEach(r => { if (r.deliveryEstimated) est[r.key] = 1; else delete est[r.key]; });
        Office.context.document.settings.set(SETTINGS_EST, est);
        await new Promise(res => Office.context.document.settings.saveAsync(() => res()));
        pending.done = { added, updates: plan.updates.length, watch: data.watchlist.length };
      });
      const done = pending.done;
      pending = null;
      await refresh();
      tab = 'import';
      render();
      document.getElementById('paste-msg').innerHTML = `<div class="ok-box"><strong>Done</strong><br>
        ${done.added.length ? `Added ${plural(done.added.length, 'row')} (${esc(done.added[0])}${done.added.length > 1 ? '–' + esc(done.added[done.added.length - 1]) : ''}).` : 'No new rows.'}
        ${done.updates ? ` Updated ${plural(done.updates, 'delivery date')}.` : ''}${done.watch ? ' Watchlist updated.' : ''}
        ${done.added.length ? '<br><span class="small">Check the new rows in Purchases &amp; Parts and add Allocation or notes as usual.</span>' : ''}</div>`;
    } catch (err) {
      btn.disabled = false;
      btn.textContent = 'Try again';
      document.getElementById('result').innerHTML = `<div class="err-box"><strong>Nothing was added</strong><br>${esc(err.message || err)}</div>`;
    }
  }

  async function writeWatchlist(ctx, list) {
    const { header, rows } = C.watchlistValues(list);
    let ws = ctx.workbook.worksheets.getItemOrNullObject('Watchlist');
    await ctx.sync();
    if (ws.isNullObject) ws = ctx.workbook.worksheets.add('Watchlist');
    const old = ctx.workbook.tables.getItemOrNullObject('Watchlist');
    await ctx.sync();
    if (!old.isNullObject) old.delete();
    ws.getRange('A:F').clear();
    const range = ws.getRange('A1').getResizedRange(rows.length, header.length - 1);
    range.formulas = [header].concat(rows);
    const table = ws.tables.add(range, true);
    table.name = 'Watchlist';
    if (rows.length) {
      ws.getRange('B2').getResizedRange(rows.length - 1, 0).numberFormat = rows.map(() => ['£#,##0.00']);
      ws.getRange('D2').getResizedRange(rows.length - 1, 0).numberFormat = rows.map(() => ['£#,##0']);
      ws.getRange('E2').getResizedRange(rows.length - 1, 0).numberFormat = rows.map(() => ['dd/mm/yyyy']);
    }
    range.format.autofitColumns();
  }

  // ---- sales ------------------------------------------------------------------

  async function takeSalesFiles(files) {
    const msg = document.getElementById('sales-msg');
    if (!model.sales) { msg.innerHTML = '<div class="err-box">Couldn’t find the Sales table. It needs columns called “Build / Part ID”, “Sold?” and “Sold Price”.</div>'; return; }
    try {
      let all = [];
      const kinds = [];
      for (const f of [...files]) {
        const r = SA.readReport(await f.text());
        kinds.push(r.kind);
        all = SA.merge(all, r.sales);
      }
      if (!all.length) throw new Error('No sales found in that file.');
      const plan = SA.planSales(model.sales, builds(), purchases(), all);
      plan.lines.sort((a, b) => String(b.sale.date || '').localeCompare(String(a.sale.date || '')));
      pendingSales = { plan, kinds, choose: plan.lines.map(l => ({ id: l.id || '', on: l.action !== 'skip' })) };
      renderSalesPreview();
    } catch (e) {
      msg.innerHTML = `<div class="err-box">${esc(e.message || e)}</div>`;
    }
  }

  function renderSalesPreview() {
    const { plan, kinds, choose } = pendingSales;
    const ids = [...new Set(salesRows().map(s => SA.idFrom(s['Build / Part ID'])).concat(builds().map(b => SA.idFrom(b['Build ID']))).filter(Boolean))];
    const count = () => choose.filter(c => c.on && SA.idFrom(c.id)).length;
    const how = l => (l.how === 'label' ? 'matched by custom label' : l.how === 'title' ? 'matched by title, please check' : l.how === 'order' ? 'already recorded: fills in fees' : 'choose a Build or Part ID');
    view.innerHTML = `
      <h2><span>eBay sales</span><span>${plan.lines.length}</span></h2>
      <p class="small muted" style="margin-top:0">From the ${[...new Set(kinds)].map(k => (k === 'transactions' ? 'Transaction report' : 'Orders report')).join(' and ')}. Sold price includes the postage the buyer paid.${kinds.includes('transactions') ? '' : ' Add the Transaction report too to fill in eBay fees and postage labels.'}</p>
      <ul class="list">${plan.lines.map((l, i) => `<li class="${l.action === 'skip' && !l.id ? 's-urgent' : ''}"><div class="row sale" style="cursor:default">
        <span><label class="chk"><input type="checkbox" data-on="${i}" ${choose[i].on ? 'checked' : ''}> <span class="name">${esc(l.sale.title)}</span></label>
          <span class="sub">${nice(l.sale.date)} · order ${esc(l.sale.order)}${l.sale.ebayFees ? ` · fees ${money(l.sale.ebayFees)}` : ''}${l.sale.promoFee ? ` · promoted ${money(l.sale.promoFee)}` : ''}${l.sale.shippingCost ? ` · label ${money(l.sale.shippingCost)}` : ''}</span>
          <span class="sub">→ <input class="idbox" list="ids" data-id="${i}" value="${esc(choose[i].id)}" aria-label="Build or Part ID" placeholder="B70 / P1004"> <span class="${l.how === 'title' || !l.id ? 'warn-text' : ''}">${esc(l.reason && l.action === 'skip' ? l.reason : how(l))}</span></span></span>
        <span class="days">${money(l.sale.soldPrice)}<small>${l.sale.qty > 1 ? `${l.sale.qty} sold` : 'sold'}</small></span></div></li>`).join('')}</ul>
      <datalist id="ids">${ids.map(id => `<option value="${esc(id)}">`).join('')}</datalist>
      <p class="small muted">Ticked sales are written to Sales (Sold?, Date of Sale, Sold Via, Sold Price, fees, and Profit if it’s empty). Single parts are also marked sold in Purchases &amp; Parts. The first time, two columns are added to the end of Sales: <b>eBay Fees</b> and <b>eBay Order</b>.</p>
      <div class="actions"><button class="btn" id="s-go" type="button">Record ${plural(count(), 'sale')}</button><button class="link" id="s-cancel" type="button">Cancel</button></div>
      <div id="s-result"></div>`;
    const go = document.getElementById('s-go');
    const update = () => { go.textContent = `Record ${plural(count(), 'sale')}`; go.disabled = !count(); };
    update();
    view.querySelectorAll('[data-on]').forEach(cb => cb.addEventListener('change', () => { choose[+cb.dataset.on].on = cb.checked; update(); }));
    view.querySelectorAll('[data-id]').forEach(inp => inp.addEventListener('input', () => {
      const c = choose[+inp.dataset.id];
      c.id = inp.value.trim();
      if (SA.idFrom(c.id)) { c.on = true; view.querySelector(`[data-on="${inp.dataset.id}"]`).checked = true; }
      update();
    }));
    document.getElementById('s-cancel').addEventListener('click', () => { pendingSales = null; renderImport(); });
    go.addEventListener('click', recordSales);
  }

  async function recordSales(e) {
    const btn = e.target;
    btn.disabled = true;
    btn.innerHTML = '<span class="spinner" aria-hidden="true"></span>Recording…';
    const { plan, choose } = pendingSales;
    try {
      const result = { added: [], updated: [], parts: [] };
      await Excel.run(async ctx => {
        // 1. Make sure Sales has the two extra columns.
        let st = await readTable(ctx, model.sales.name, true);
        const missing = SALES_NEW_COLS.filter(c => !st.headers.some(h => lc(h) === lc(c)));
        if (missing.length) {
          const t = ctx.workbook.tables.getItem(st.name);
          missing.forEach(name => t.columns.add(null, null, name));
          await ctx.sync();
          st = await readTable(ctx, model.sales.name, true);
          // Money format like Sold Price
          const ix0 = C.colIndex(st.headers);
          if (st.values.length && ix0('Sold Price') >= 0) {
            const cell = t.getDataBodyRange().getCell(0, ix0('Sold Price'));
            cell.load('numberFormat');
            await ctx.sync();
            const fmt = cell.numberFormat[0][0];
            if (fmt && fmt !== 'General') t.columns.getItem('eBay Fees').getDataBodyRange().numberFormat = st.values.map(() => [fmt]);
          }
        }
        // 2. Re-plan against the sheet as it is now, with your choices.
        const chosen = plan.lines.map((l, i) => ({ l, c: choose[i] })).filter(x => x.c.on && SA.idFrom(x.c.id));
        const sales = chosen.map(x => Object.assign({}, x.l.sale, { id: SA.idFrom(x.c.id), sku: SA.idFrom(x.c.id) }));
        const fresh = SA.planSales(st, builds(), purchases(), sales);
        const ix = C.colIndex(st.headers);
        const calc = C.detectCalculated(st.formulasA1, st.formulasR1C1);
        const profitCol = ix('Profit');
        const profitCalc = profitCol >= 0 && calc[profitCol];
        const objs = C.toObjects(st.headers, st.values);
        const body = ctx.workbook.tables.getItem(st.name).getDataBodyRange();
        const adds = [];
        const seenIds = new Set();
        fresh.lines.forEach(line => {
          if (!line.id || line.action === 'skip' || seenIds.has(line.id + line.sale.order)) return;
          seenIds.add(line.id + line.sale.order);
          const existing = line.row >= 0 ? objs[line.row] : null;
          const f = SA.fieldsFor(line, existing);
          if (!profitCalc && profitCol >= 0 && (!existing || existing['Profit'] === '' || existing['Profit'] == null) && f['Sold?'] === 'Y') {
            f['Profit'] = SA.profitFormula(st.name);
          }
          if (line.row >= 0) {
            Object.keys(f).forEach(k => {
              const c = ix(k);
              if (c < 0 || calc[c]) return;
              const cell = body.getCell(line.row, c);
              if (typeof f[k] === 'string' && f[k].startsWith('=')) cell.formulas = [[f[k]]]; else cell.values = [[f[k]]];
            });
            result.updated.push(line.id);
          } else {
            adds.push(st.headers.map(h => (f[h] !== undefined ? f[h] : '')));
            result.added.push(line.id);
          }
          if (/^P/.test(line.id) && line.how !== 'order') result.parts.push(line.id);
        });
        await ctx.sync();
        await addRows(ctx, st, adds, ['Listed Date', 'Listing Price', 'Date of Sale', 'Sold Price', 'Promotion Fee', 'Shipping Cost', 'eBay Fees', 'Profit']);
        // Date of Sale format on updated rows that were blank before
        const ds = ix('Date of Sale');
        if (ds >= 0 && result.updated.length) {
          const fmtCell = body.getCell(0, ix('Listed Date') >= 0 ? ix('Listed Date') : ds);
          fmtCell.load('numberFormat');
          await ctx.sync();
          const fmt = fmtCell.numberFormat[0][0] && fmtCell.numberFormat[0][0] !== 'General' ? fmtCell.numberFormat[0][0] : 'dd/mm/yyyy';
          fresh.lines.filter(l => l.row >= 0 && l.action !== 'skip').forEach(l => { body.getCell(l.row, ds).numberFormat = [[fmt]]; });
        }
        // 3. Single parts: mark sold in Purchases & Parts.
        if (result.parts.length) {
          const pt = await readTable(ctx, model.purchases.name, false);
          const px = C.colIndex(pt.headers);
          const pb = ctx.workbook.tables.getItem(pt.name).getDataBodyRange();
          result.parts.forEach(id => {
            const row = pt.values.findIndex(r => String(r[px('Part ID')]).trim() === id);
            if (row < 0) return;
            if (px('Sold?') >= 0) pb.getCell(row, px('Sold?')).values = [['Y']];
            if (px('Available?') >= 0) pb.getCell(row, px('Available?')).values = [['N']];
          });
        }
        await ctx.sync();
      });
      pendingSales = null;
      await refresh();
      tab = 'import';
      render();
      document.getElementById('sales-msg').innerHTML = `<div class="ok-box"><strong>Sales recorded</strong><br>
        ${result.updated.length ? `Updated ${esc(result.updated.join(', '))}. ` : ''}${result.added.length ? `Added ${esc(result.added.join(', '))}. ` : ''}
        ${result.parts.length ? `Marked ${esc(result.parts.join(', '))} sold in Purchases &amp; Parts.` : ''}</div>`;
    } catch (err) {
      btn.disabled = false;
      btn.textContent = 'Try again';
      document.getElementById('s-result').innerHTML = `<div class="err-box"><strong>Nothing was recorded</strong><br>${esc(err.message || err)}</div>`;
    }
  }

  // ---------------------------------------------------------------------------
  // Start
  // ---------------------------------------------------------------------------

  document.querySelectorAll('.tabs button').forEach(b => b.addEventListener('click', () => go(b.dataset.tab)));
  document.getElementById('refresh').addEventListener('click', () => { hubCache = null; refresh(); });
  document.getElementById('open-settings').addEventListener('click', () => (tab === 'settings' ? go(lastTab || 'home') : go('settings')));

  Office.onReady(info => {
    if (info.host !== Office.HostType.Excel) {
      view.innerHTML = `<p class="msg"><strong>Open this in Excel</strong><span class="muted">${APP_NAME} works with your Business Tracker workbook in Excel.</span></p>`;
      return;
    }
    refresh();
  });
})();
