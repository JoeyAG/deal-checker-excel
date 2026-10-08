/*
 * taskpane.js — connects the Deal Checker panel to your workbook (Office.js).
 * All the decisions live in core.js; this file reads and writes the tables.
 *
 * Tables are found by their column headings, not their names, so renaming a
 * table (e.g. "Table2") doesn't break anything.
 */
(function () {
  'use strict';
  const C = DealCore;
  const SETTINGS_EST = 'dealChecker.estimatedDelivery';
  const gbp = new Intl.NumberFormat('en-GB', { style: 'currency', currency: 'GBP' });
  const gbp0 = new Intl.NumberFormat('en-GB', { style: 'currency', currency: 'GBP', maximumFractionDigits: 0 });
  const money = v => (typeof v === 'number' && Number.isFinite(v) ? gbp.format(v) : '—');
  const pounds = v => (typeof v === 'number' && Number.isFinite(v) ? gbp0.format(v) : '—');
  const nice = iso => (iso ? new Date(iso + 'T12:00:00').toLocaleDateString('en-GB', { day: 'numeric', month: 'short' }) : '—');
  const esc = s => String(s == null ? '' : s).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  const plural = (n, one, many) => `${n} ${n === 1 ? one : (many || one + 's')}`;

  const view = document.getElementById('view');
  const setup = document.getElementById('setup');
  let tab = 'today';
  let model = null;
  let pending = null;   // parsed payload waiting for confirmation

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

  function purchases() { return C.toObjects(model.purchases.headers, model.purchases.values); }

  function render() {
    if (!model) return;
    document.querySelectorAll('.tabs button').forEach(b => b.setAttribute('aria-selected', String(b.dataset.tab === tab)));
    if (tab === 'today') renderToday();
    if (tab === 'builds') renderBuilds();
    if (tab === 'stock') renderStock();
    if (tab === 'watch') renderWatch();
    if (tab === 'import') renderImport();
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
    const section = (title, items, open) => items.length ? `<h2><span>${title}</span><span>${items.length}</span></h2><ul class="list">${items.map(deadlineRow).join('')}</ul>` : '';
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

  function renderBuilds() {
    if (!model.builds) { view.innerHTML = '<p class="msg"><strong>No Builds table found</strong><span class="muted">It needs columns called “Build ID” and “Missing Parts”.</span></p>'; return; }
    const list = C.buildsSummary(C.toObjects(model.builds.headers, model.builds.values), purchases());
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

  function renderWatch() {
    if (!model.watch || !model.watch.values.length || !String(model.watch.values[0][0] || '').trim()) {
      view.innerHTML = `<p class="msg"><strong>No watched parts yet</strong><span class="muted">In Chrome, click <b>Watch</b> in the Deal check panel on a listing. Your watchlist comes across with your next purchases import, or from the extension’s Watchlist page (“Copy watchlist for Excel”).</span></p>`;
      return;
    }
    const rows = C.toObjects(model.watch.headers, model.watch.values).filter(w => w['Part']);
    view.innerHTML = `<ul class="list">${rows.map(w => `<li><div class="row" style="cursor:default">
        <span><span class="name">${esc(w['Part'])}</span><span class="sub">median ${money(w['Median Sold'])} from ${esc(w['Sales'] || 0)} sales · updated ${nice(C.serialToIso(w['Updated']))}</span></span>
        <span class="days">${pounds(w['Alert At'])}<small>alert at or under</small></span></div></li>`).join('')}</ul>
      <p class="small muted">The full list, with <b>Find deals</b> links, is on the Watchlist sheet.</p>`;
  }

  function renderImport() {
    if (pending) return renderPreview();
    view.innerHTML = `
      <ol class="steps">
        <li>In Chrome, open <b>My eBay → Purchases</b>.</li>
        <li>Click <b>Send to spreadsheet</b>, then <b>Copy rows for Excel</b>.</li>
        <li>Click in the box below and press <b>Ctrl+V</b>.</li>
      </ol>
      <textarea class="paste" id="paste" placeholder="Paste here (Ctrl+V)" aria-label="Paste from the Deal Checker extension"></textarea>
      <div id="paste-msg"></div>`;
    const ta = document.getElementById('paste');
    ta.addEventListener('paste', e => {
      const text = (e.clipboardData || window.clipboardData).getData('text');
      e.preventDefault();
      takePaste(text);
    });
    ta.addEventListener('input', () => { if (ta.value.includes(C.PREFIX)) takePaste(ta.value); });
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
        const calc = C.detectCalculated(fresh.formulasA1, fresh.formulasR1C1);
        const startRow = fresh.values.length;

        if (plan.add.length) {
          t.rows.add(null, plan.add);
          await ctx.sync();
          const body = t.getDataBodyRange();
          const n = plan.add.length;
          const col = name => fresh.headers.findIndex(h => h.toLowerCase() === name.toLowerCase());
          const newCells = c => body.getCell(startRow, c).getResizedRange(n - 1, 0);
          // Formula columns: write the column's formula explicitly so they stay consistent.
          calc.forEach((cf, c) => {
            if (!cf) return;
            if (C.rowIndependent(cf.a1)) newCells(c).formulas = Array.from({ length: n }, () => [cf.a1]);
            else newCells(c).formulasR1C1 = Array.from({ length: n }, () => [cf.r1c1]);
          });
          // Dates: same format as the rows above.
          const pd = col('Purchase Date');
          if (pd >= 0 && startRow > 0) {
            const fmtCell = body.getCell(0, pd);
            fmtCell.load('numberFormat');
            await ctx.sync();
            const fmt = fmtCell.numberFormat[0][0];
            ['Purchase Date', 'Delivery Date'].map(col).filter(c => c >= 0)
              .forEach(c => { newCells(c).numberFormat = Array.from({ length: n }, () => [fmt]); });
          }
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

  // ---------------------------------------------------------------------------
  // Start
  // ---------------------------------------------------------------------------

  document.querySelectorAll('.tabs button').forEach(b => b.addEventListener('click', () => { tab = b.dataset.tab; render(); }));
  document.getElementById('refresh').addEventListener('click', refresh);

  Office.onReady(info => {
    if (info.host !== Office.HostType.Excel) {
      view.innerHTML = '<p class="msg"><strong>Open this in Excel</strong><span class="muted">Deal Checker works with your Business Tracker workbook in Excel.</span></p>';
      return;
    }
    refresh();
  });
})();
