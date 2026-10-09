/*
 * taskpane.js — connects the Deal Checker panel to your workbook (Office.js).
 * The decisions live in core.js, planner.js, listing.js and sales.js; this file reads
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
  const gbp = new Intl.NumberFormat('en-GB', { style: 'currency', currency: 'GBP' });
  const gbp0 = new Intl.NumberFormat('en-GB', { style: 'currency', currency: 'GBP', maximumFractionDigits: 0 });
  const money = v => (typeof v === 'number' && Number.isFinite(v) ? gbp.format(v) : '—');
  const pounds = v => (typeof v === 'number' && Number.isFinite(v) ? gbp0.format(v) : '—');
  const nice = iso => (iso ? new Date(iso + 'T12:00:00').toLocaleDateString('en-GB', { day: 'numeric', month: 'short' }) : '—');
  const esc = s => String(s == null ? '' : s).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  const plural = (n, one, many) => `${n} ${n === 1 ? one : (many || one + 's')}`;
  const lc = s => String(s == null ? '' : s).trim().toLowerCase();

  const view = document.getElementById('view');
  const setup = document.getElementById('setup');
  let tab = 'today';
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
    const url = URL.createObjectURL(new Blob([text], { type: type || 'text/csv;charset=utf-8' }));
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
      return m;
    });
  }

  async function refresh() {
    try {
      model = await loadModel();
      renderSetup();
      render();
    } catch (e) {
      view.innerHTML = `<div class="err-box"><strong>Couldn’t read the workbook</strong><br>${esc(e.message || e)}</div>`;
    }
  }

  const purchases = () => C.toObjects(model.purchases.headers, model.purchases.values);
  const builds = () => (model.builds ? C.toObjects(model.builds.headers, model.builds.values) : []);
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
    ({ today: renderToday, builds: renderBuilds, shop: renderShop, stock: renderStock, sell: renderSell, watch: renderWatch, import: renderImport }[tab] || renderToday)();
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

  function renderToday() {
    const today = C.todaySerial();
    const list = C.testDeadlines(purchases(), today, estimatedKeys());
    const by = s => list.filter(d => d.status === s);
    const urgent = by('over').length + by('urgent').length;
    const noValue = C.missingValues(purchases()).length;
    const section = (title, items) => items.length ? `<h2><span>${title}</span><span>${items.length}</span></h2><ul class="list">${items.map(deadlineRow).join('')}</ul>` : '';
    const later = by('ok');
    view.innerHTML = `
      <div class="big">
        <div class="stat ${urgent ? 'alert' : ''}"><b>${urgent}</b><span>to test within 7 days</span></div>
        <div class="stat"><b>${list.filter(d => d.status !== 'coming').length}</b><span>untested from eBay</span></div>
      </div>
      <p class="small muted" style="margin:0 0 4px">eBay gives you 30 days from delivery to report a part not as described. Click a part to find it in the sheet.</p>
      ${section('Past the 30 days', by('over'))}
      ${section('Test within 7 days', by('urgent'))}
      ${section('Within 14 days', by('soon'))}
      ${later.length ? `<details class="more" ${urgent || by('soon').length ? '' : 'open'}><summary>Later (${later.length})</summary><ul class="list">${later.map(deadlineRow).join('')}</ul></details>` : ''}
      ${section('On the way', by('coming'))}
      ${!list.length ? '<p class="msg"><strong>Nothing waiting to be tested</strong><span class="muted">eBay parts marked Tested? = N will appear here.</span></p>' : ''}
      ${noValue ? `<p class="small muted" style="margin-top:14px">${plural(noValue, 'part has', 'parts have')} no market value yet. Parts you check with the Chrome extension before buying get one automatically.</p>` : ''}`;
    view.querySelectorAll('[data-row]').forEach(b => b.addEventListener('click', () => selectRow(model.purchases.name, +b.dataset.row)));
  }

  function noBuilds() {
    view.innerHTML = '<p class="msg"><strong>No Builds table found</strong><span class="muted">It needs columns called “Build ID” and “Missing Parts”.</span></p>';
  }

  function renderBuilds() {
    if (!model.builds) return noBuilds();
    const list = C.buildsSummary(builds(), purchases());
    const order = { 'Collecting parts': 0, 'All parts in': 1, 'Ready to build': 2, 'Built': 3, 'Benchmarked': 4 };
    list.sort((a, b) => order[a.stage] - order[b.stage] || String(a.id).localeCompare(String(b.id), undefined, { numeric: true }));
    view.innerHTML = list.map(b => `<button class="build" type="button" data-row="${b.row}">
        <div class="head"><span class="id">${esc(b.id)}</span><span class="stage">${esc(b.stage)}</span></div>
        <div class="desc">${esc(b.description) || '<span class="muted">No description</span>'}</div>
        <div class="small muted">${plural(b.parts, 'part')} allocated · ${money(b.cost)}${b.untested ? ` · <span class="tag warn">${b.untested} untested</span>` : ''}${b.waitingOn ? ` · waiting on ${esc(b.waitingOn)}` : ''}</div>
        ${b.missing.length ? `<div class="chips">${b.missing.map(m => `<span class="chip">needs ${esc(m)}</span>`).join('')}</div>` : ''}
      </button>`).join('') || '<p class="muted">No builds yet.</p>';
    view.querySelectorAll('[data-row]').forEach(b => b.addEventListener('click', () => selectRow(model.builds.name, +b.dataset.row)));
  }

  function renderStock() {
    const g = C.stockSummary(purchases());
    const tot = g.reduce((a, x) => ({ count: a.count + x.count, cost: a.cost + x.cost, value: a.value + x.value }), { count: 0, cost: 0, value: 0 });
    view.innerHTML = `<p class="small muted" style="margin:0 0 8px">Parts marked available and not allocated to a build.</p>
      <table class="grid"><tr><th>Category</th><th class="n">Parts</th><th class="n">Cost</th><th class="n">Value</th></tr>
      ${g.map(x => `<tr><td>${esc(x.category)}</td><td class="n">${x.count}</td><td class="n">${pounds(x.cost)}</td><td class="n">${x.valued ? pounds(x.value) : '—'}</td></tr>`).join('')}
      <tr class="total"><td>Total</td><td class="n">${tot.count}</td><td class="n">${pounds(tot.cost)}</td><td class="n">${pounds(tot.value)}</td></tr></table>
      <p class="small muted">Value only counts parts with an Estimated Part Value.</p>`;
  }

  // ---------------------------------------------------------------------------
  // Shop: the build-aware shopping list
  // ---------------------------------------------------------------------------

  const KIND = PL.KIND_NAME;

  function renderShop() {
    if (!model.builds) return noBuilds();
    const plan = PL.plan(builds(), purchases());
    model.plan = plan;
    const stockBuilds = plan.builds.filter(b => b.needs.some(n => n.stock));
    const toBuy = plan.shopping.reduce((a, g) => a + g.count, 0);
    const watched = new Map(((hubCache && hubCache.watches) || []).map(w => [lc(w.label), w]));
    const hubOn = isHubSet();

    view.innerHTML = `
      <div class="big">
        <div class="stat"><b>${plan.fromStock}</b><span>parts in stock fit a build</span></div>
        <div class="stat"><b>${toBuy}</b><span>parts still to buy</span></div>
      </div>
      <p class="small muted" style="margin:0 0 4px">From each build’s <b>Missing Parts</b>, its CPU or board, and your unallocated stock. Builds with no CPU or board yet are planned around ${esc(PL.PLATFORMS.lga1151v2.cpu)} unless stock suggests otherwise.</p>

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
      ${!hubOn && plan.shopping.length ? '<p class="small muted">To get phone alerts for these, connect your Deal Alerts sheet on the <b>Watch</b> tab.</p>' : ''}

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
      if (!isHubSet()) { tab = 'watch'; render(); return; }
      const f = document.getElementById('wf-' + b.dataset.watch);
      f.hidden = !f.hidden;
      if (!f.hidden) f.querySelector('input').focus();
    }));
    view.querySelectorAll('[data-addwatch]').forEach(b => b.addEventListener('click', () => addShopWatch(plan.shopping[+b.dataset.addwatch], +b.dataset.addwatch)));
    view.querySelectorAll('[data-alloc]').forEach(b => b.addEventListener('click', () => allocateBuild(b.dataset.alloc, b)));
    if (hubOn && !hubCache) loadHub().then(() => { if (tab === 'shop') renderShop(); }).catch(() => {});
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
      setTimeout(() => { if (tab === 'shop') renderShop(); }, 1200);
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
      tab = 'shop';
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
    const desc = ed.description != null ? ed.description : d.description;
    const price = ed.price != null ? ed.price : d.price;
    const saleRow = sales.find(s => SA.idFrom(s['Build / Part ID']) === SA.idFrom(sellId));
    const listed = saleRow && lc(saleRow['Listed?']) === 'y';
    const specText = d.specifics.map(([k, v]) => `${k}: ${v}`).join('\n');
    const opt = b => `<option value="${esc(b['Build ID'])}" ${b['Build ID'] === sellId ? 'selected' : ''}>${esc(b['Build ID'])} · ${esc(String(b['Build Description'] || '').slice(0, 40))}</option>`;

    view.innerHTML = `
      <label class="field"><span>Build</span><select id="sell-pick">
        ${ready.length ? `<optgroup label="Built, not sold">${ready.map(opt).join('')}</optgroup>` : ''}
        ${rest.length ? `<optgroup label="Still being built">${rest.map(opt).join('')}</optgroup>` : ''}
      </select></label>
      ${d.missing.length ? `<div class="warn-box">No ${esc(d.missing.join(', '))} found for ${esc(sellId)}. Allocate the parts in Purchases &amp; Parts (or add them to the Build Description) and the draft fills in.</div>` : ''}
      ${listed ? `<div class="ok-box small">Listed on ${esc(saleRow['Listed Date'] ? nice(C.serialToIso(saleRow['Listed Date'])) : '?')} at ${money(saleRow['Listing Price'])}.</div>` : ''}

      <div class="field"><span>Title <small class="muted" id="t-count">${title.length}/80</small></span>
        <div class="with-btn"><input id="s-title" maxlength="80" value="${esc(title)}"><button class="btn ghost sm" type="button" data-copy="title">Copy</button></div></div>

      <div class="field"><span>Custom label (SKU)</span>
        <div class="with-btn"><input value="${esc(sellId)}" readonly><button class="btn ghost sm" type="button" data-copy="sku">Copy</button></div>
        <small class="muted">Put this in eBay’s <b>Custom label (SKU)</b> box. It’s how sales get matched back to ${esc(sellId)} automatically.</small></div>

      <div class="field"><span>Price</span>
        <div class="with-btn"><span class="pfx">£</span><input id="s-price" type="number" min="1" step="1" value="${price == null ? '' : price}"><button class="btn ghost sm" type="button" data-copy="price">Copy</button></div>
        <small class="muted">Cost ${money(d.cost)} + ${p.marginPct}% margin${fee ? ` + ${fee}% fees` : ''} → ${pounds(d.priceFromCost)}${d.priceFromValue ? ` · parts’ market value → ${pounds(d.priceFromValue)}${d.partsValue.missing ? ` (${d.partsValue.missing} part${d.partsValue.missing === 1 ? '' : 's'} unvalued)` : ''}` : ''}${d.similarUrl ? ` · ${ext(d.similarUrl, 'similar PCs sold')}` : ''}</small></div>

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

      <details class="more"><summary>Pricing and wording settings</summary>
        <div class="field inline"><span>Margin on cost</span><input id="p-margin" type="number" min="0" max="300" value="${p.marginPct}"> %</div>
        <div class="field inline"><span>Windows</span><select id="p-os">${['Windows 11 Home', 'Windows 11 Pro'].map(o => `<option ${o === p.os ? 'selected' : ''}>${o}</option>`).join('')}</select></div>
        <p class="small muted">${fee ? `Fees: ${fee}% of the sale price on average, from your recorded eBay sales.` : 'Fees: none allowed for yet. Once you record eBay sales (Import tab), the average fee is added automatically.'}</p>
      </details>`;

    const keep = (k, v) => { sellEdits[sellId] = Object.assign({}, sellEdits[sellId], { [k]: v }); };
    document.getElementById('sell-pick').addEventListener('change', e => { sellId = e.target.value; renderSell(); });
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
    document.getElementById('p-margin').addEventListener('change', async e => { await savePrefs({ marginPct: Math.max(0, +e.target.value || 0) }); delete (sellEdits[sellId] || {}).price; renderSell(); });
    document.getElementById('p-os').addEventListener('change', async e => { await savePrefs({ os: e.target.value }); delete sellEdits[sellId]; renderSell(); });
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
      tab = 'sell';
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
    if (!HUB_RE.test(p.hubUrl) || !p.hubKey) throw new Error('Connect your Deal Alerts sheet on the Watch tab first.');
    const res = await fetch(p.hubUrl, { method: 'POST', body: JSON.stringify(Object.assign({ key: p.hubKey }, body)) });
    const text = await res.text();
    let data;
    try { data = JSON.parse(text); } catch (e) {
      throw new Error(/<html/i.test(text) ? 'The alerts sheet asked for a Google sign-in. Redeploy its web app with “Who has access: Anyone”.' : 'Unexpected reply from the alerts sheet.');
    }
    if (!data.ok) throw new Error(data.error || 'The alerts sheet refused the request.');
    return data;
  }
  async function loadHub() { hubCache = await hubCall({ action: 'list' }); return hubCache; }

  function renderWatch() {
    const p = prefs();
    if (!isHubSet()) {
      view.innerHTML = `
        <h2><span>Deal alerts on your phone</span></h2>
        <p class="small">Your Deal Alerts Google Sheet checks eBay every 15 minutes for the parts you watch and sends a phone notification when one is listed at or under your alert price. Connect it here to add parts from the <b>Shop</b> tab. Both values are in the sheet: <b>Deal alerts → Show connection details</b>.</p>
        <label class="field"><span>Web app URL</span><input id="h-url" placeholder="https://script.google.com/macros/s/…/exec" value="${esc(p.hubUrl)}"></label>
        <label class="field"><span>Key</span><input id="h-key" type="password" value="${esc(p.hubKey)}"></label>
        <div class="actions"><button class="btn" id="h-save" type="button">Connect</button><span id="h-out" class="small"></span></div>
        ${sheetWatchHtml()}`;
      document.getElementById('h-save').addEventListener('click', async () => {
        const url = document.getElementById('h-url').value.trim();
        const key = document.getElementById('h-key').value.trim();
        const out = document.getElementById('h-out');
        if (!HUB_RE.test(url)) { out.textContent = 'The URL should look like https://script.google.com/macros/s/…/exec'; return; }
        out.textContent = 'Connecting…';
        try {
          hubCache = await hubCall({ action: 'list' }, { hubUrl: url, hubKey: key });
          await savePrefs({ hubUrl: url, hubKey: key, greatPct: hubCache.greatPct || 25 });
          renderWatch();
        } catch (e) { out.textContent = e.message || String(e); }
      });
      return;
    }
    if (!hubCache) {
      view.innerHTML = '<p class="msg"><span class="spinner" aria-hidden="true"></span>Reading your alerts sheet…</p>';
      loadHub().then(() => { if (tab === 'watch') renderWatch(); }).catch(e => {
        if (tab !== 'watch') return;
        view.innerHTML = `<div class="err-box"><strong>Couldn’t reach your alerts sheet</strong><br>${esc(e.message || e)}</div>
          <div class="actions"><button class="link" id="h-retry" type="button">Try again</button><button class="link" id="h-forget" type="button">Disconnect</button></div>`;
        document.getElementById('h-retry').addEventListener('click', () => renderWatch());
        document.getElementById('h-forget').addEventListener('click', async () => { await savePrefs({ hubUrl: '', hubKey: '' }); renderWatch(); });
      });
      return;
    }
    const h = hubCache;
    const ago = ts => (ts ? new Date(ts).toLocaleString('en-GB', { day: 'numeric', month: 'short', hour: '2-digit', minute: '2-digit' }) : '');
    view.innerHTML = `
      <div class="ok-box small" style="margin-top:0"><b>Deal alerts connected</b>${h.paused ? ' · <span class="tag warn">paused on the sheet</span>' : ''} · ${plural(h.watches.length, 'part')} watched
        · <button class="link" type="button" id="h-test">Send a test notification</button></div>
      ${h.alerts && h.alerts.length ? `<h2><span>Latest deals found</span></h2><ul class="list">${h.alerts.slice(0, 8).map(a => `<li><div class="row" style="cursor:default">
        <span><span class="name">${ext(a.url, esc(a.title))}</span><span class="sub">${esc(a.label)} · ${esc(ago(a.found))}</span></span>
        <span class="days">${money(a.total)}<small>alert ≤ ${pounds(a.alertAt)}</small></span></div></li>`).join('')}</ul>` : ''}
      <h2><span>Watching</span><span>${h.watches.length}</span></h2>
      ${h.watches.length ? `<ul class="list">${h.watches.map((w, i) => `<li><div class="row" style="cursor:default">
        <span><span class="name">${esc(w.label)}${w.active === false ? ' <span class="tag warn">off</span>' : ''}</span>
        <span class="sub">${w.median ? `median ${money(w.median)} from ${w.n || 0} sales` : 'not priced yet (the Chrome extension prices it next time you open an eBay listing)'}${w.source ? ` · from ${esc(w.source === 'excel' ? 'Excel' : w.source === 'chrome' ? 'Chrome' : w.source)}` : ''}
        · <button class="link" type="button" data-unwatch="${i}">Remove</button></span></span>
        <span class="days">${pounds(w.alertAt)}<small>${w.alertFixed ? 'your price' : 'alert at or under'}</small></span></div></li>`).join('')}</ul>`
        : '<p class="muted">Nothing watched yet. Use <b>Watch</b> on the Shop tab, or the Watch button in Chrome.</p>'}
      <p class="small muted">Change alert prices or switch parts off on the Watchlist tab of your Deal Alerts sheet. <button class="link" type="button" id="h-reload">Refresh</button> · <button class="link" type="button" id="h-forget">Disconnect</button></p>`;
    document.getElementById('h-test').addEventListener('click', async e => {
      e.target.textContent = 'Sending…';
      try { await hubCall({ action: 'test', from: 'Excel' }); e.target.textContent = 'Sent'; } catch (x) { e.target.textContent = x.message; }
    });
    document.getElementById('h-reload').addEventListener('click', () => { hubCache = null; renderWatch(); });
    document.getElementById('h-forget').addEventListener('click', async () => { await savePrefs({ hubUrl: '', hubKey: '' }); hubCache = null; renderWatch(); });
    view.querySelectorAll('[data-unwatch]').forEach(b => b.addEventListener('click', async () => {
      b.textContent = 'Removing…';
      try { hubCache = await hubCall({ action: 'remove', labels: [h.watches[+b.dataset.unwatch].label] }); renderWatch(); } catch (x) { b.textContent = x.message; }
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

  document.querySelectorAll('.tabs button').forEach(b => b.addEventListener('click', () => { tab = b.dataset.tab; render(); }));
  document.getElementById('refresh').addEventListener('click', () => { hubCache = null; refresh(); });

  Office.onReady(info => {
    if (info.host !== Office.HostType.Excel) {
      view.innerHTML = '<p class="msg"><strong>Open this in Excel</strong><span class="muted">Deal Checker works with your Business Tracker workbook in Excel.</span></p>';
      return;
    }
    refresh();
  });
})();
