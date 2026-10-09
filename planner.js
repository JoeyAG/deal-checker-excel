/*
 * planner.js — the build-aware shopping list.
 *
 * For every build that isn't built yet it works out:
 *   - which platform the build is on (from its CPU or motherboard, or a suggestion),
 *   - what each missing part needs to be (DDR generation, chipsets that fit, PSU size…),
 *   - which unallocated parts in Purchases & Parts already fit,
 *   - what's left to buy, grouped across builds, with a rough price from past purchases.
 *
 * Plain JavaScript with no Excel code, so it can be tested on its own. Uses DealSpecs.
 */
var DealPlanner = (function () {
  'use strict';

  const S = (typeof DealSpecs !== 'undefined') ? DealSpecs : require('./specs.js');
  const norm = s => String(s == null ? '' : s).trim().toLowerCase();

  const PR = (typeof DealPairings !== 'undefined') ? DealPairings : require('./pairings.js');
  const { PLATFORMS, DEFAULT_PLATFORM, CHIPSET_PLATFORM, cpuPlatform, chipsetsFor, looseChipset, gpuClass,
    psuWattsFor, suggestGpu, cpuFitsGaming } = PR;

  // ---------------------------------------------------------------------------
  // What each missing part has to be
  // ---------------------------------------------------------------------------

  const KIND_WORDS = [
    ['cpu', /\b(cpu|processor)\b/],
    ['mobo', /\b(motherboard|mobo|mainboard|board)\b/],
    ['gpu', /\b(gpu|graphics|video card)\b/],
    ['ram', /\b(ram|memory)\b/],
    ['psu', /\b(psu|power supply)\b/],
    ['cooler', /\b(cooler|cooling|heatsink|aio)\b/],
    ['case', /\b(case|chassis|tower)\b/],
    ['storage', /\b(ssd|hdd|nvme|m\.2|storage|hard drive|drive)\b/]
  ];
  const KIND_NAME = { cpu: 'CPU', mobo: 'Motherboard', gpu: 'Graphics card', ram: 'RAM', psu: 'PSU', storage: 'Storage', case: 'Case', cooler: 'CPU cooler' };
  const SHEET_CATEGORY = { cpu: 'CPU', mobo: 'Motherboard', gpu: 'GPU', ram: 'RAM', psu: 'PSU', storage: 'Storage', case: 'Case', cooler: 'Cooling' };
  const CATEGORY_KIND = { cpu: 'cpu', motherboard: 'mobo', gpu: 'gpu', ram: 'ram', psu: 'psu', storage: 'storage', case: 'case', cooling: 'cooler' };

  function missingKinds(text) {
    const out = [];
    String(text || '').split(/[,;/+&\n]+|\band\b/i).map(s => s.trim()).filter(Boolean).forEach(tok => {
      const t = norm(tok);
      const hit = KIND_WORDS.find(([, re]) => re.test(t));
      out.push(hit ? { kind: hit[0], text: tok } : { kind: 'other', text: tok });
    });
    return out;
  }

  function ramNeed(platform) {
    const p = PLATFORMS[platform];
    if (!p) return { cat: 'ram', gen: 4, totalGB: 16, modules: 2, perGB: 8, speed: null, form: null };
    if (p.ddr === 5) return { cat: 'ram', gen: 5, totalGB: 32, modules: 2, perGB: 16, speed: null, form: null };
    if (p.quad) return p.ddr === 3 ? { cat: 'ram', gen: 3, totalGB: 16, modules: 4, perGB: 4, speed: null, form: null }
      : { cat: 'ram', gen: 4, totalGB: 32, modules: 4, perGB: 8, speed: null, form: null };
    return { cat: 'ram', gen: p.ddr, totalGB: 16, modules: 2, perGB: 8, speed: null, form: null };
  }

  function storageNeed(platform, gpu) {
    return gpuClass(gpu) >= 650 && PLATFORMS[platform] && PLATFORMS[platform].nvme
      ? { cat: 'storage', type: 'ssd', capGB: 1000 } : { cat: 'storage', type: 'ssd', capGB: 500 };
  }

  // ---------------------------------------------------------------------------
  // Reading parts out of the sheet
  // ---------------------------------------------------------------------------

  function partText(p) {
    const b = String(p['Brand'] || '').replace(/\s*\(.*?\)\s*/g, ' ').trim();
    return `${b && !/^no brand$/i.test(b) ? b + ' ' : ''}${p['Description'] || ''}`.trim();
  }

  // Spec for a row of Purchases & Parts, using its Category as the part type.
  function partSpec(p) {
    const kind = CATEGORY_KIND[norm(p['Category'])];
    if (!kind) return null;
    const text = partText(p);
    if (kind === 'case' || kind === 'cooler') return { kind, text };
    let spec = S.parseAs(kind, text);
    if (kind === 'mobo' && !spec) {
      const c = looseChipset(text);
      if (c) spec = { cat: 'mobo', chipset: c, form: null };
    }
    if (kind === 'storage' && spec && !spec.type) spec.type = 'ssd';
    if (spec) { delete spec.index; delete spec.end; }
    return { kind, text, spec };
  }

  const isFree = p => norm(p['Available?']) === 'y' && !String(p['Allocation'] || '').trim() && norm(p['Sold?']) !== 'y';

  // ---------------------------------------------------------------------------
  // Stock that fits a need
  // ---------------------------------------------------------------------------

  const caseIsSmall = t => /\b(m-?atx|micro|mini|itx|sff)\b/i.test(t);
  const coolerLike = text => {
    const t = norm(text).replace(/\bcooler\s?master\b/g, ' ');
    if (/\b(cpu cooler|heatsink|heat sink|aio|liquid cooler|water cooler|tower cooler)\b/.test(t)) return true;
    return /\bcooler\b/.test(t) && !/\b(fans?|hub|paste|backplate|bracket)\b/.test(t);
  };

  function coolerFits(text, platform) {
    const t = norm(text);
    const amd = /\bam[2345]\+?\b|\bryzen\b|\bamd\b/.test(t);
    const intel = /\blga\s?\d{4}|\bintel\b|\b115x\b/.test(t);
    const isAmd = platform === 'am4' || platform === 'am5';
    if (amd && !intel) return isAmd ? 'fits' : 'no';
    if (intel && !amd) return isAmd ? 'no' : 'fits';
    return 'check';
  }

  // Score how well a stock part fills a need; null means it doesn't.
  function stockScore(need, ps) {
    const s = ps.spec;
    switch (need.kind) {
      case 'cpu':
        if (!s || cpuPlatform(s) !== need.platform) return null;
        if (need.gaming && !cpuFitsGaming(s)) return null;
        return 10;
      case 'mobo':
        return s && need.item.chipsets.includes(s.chipset) ? 10 : null;
      case 'gpu':
        if (!s || gpuClass(s) < 550) return null;
        return gpuClass(s);
      case 'psu':
        if (!s || s.watts < need.item.minWatts) return null;
        return 2000 - s.watts; // smallest that's big enough
      case 'storage':
        if (!s || !s.type || s.type === 'hdd' || s.ext || s.capGB < 240) return null;
        if (s.type === 'nvme' && !need.nvme) return null;
        return 1000 - Math.abs(s.capGB - need.item.capGB) / 10;
      case 'case':
        if (need.atx && caseIsSmall(ps.text)) return null;
        return 1;
      case 'cooler': {
        if (!coolerLike(ps.text)) return null;
        const f = coolerFits(ps.text, need.platform);
        return f === 'no' ? null : f === 'fits' ? 2 : 1;
      }
      default: return null;
    }
  }

  // RAM: one kit of the right size, or single sticks that add up to it.
  function pickRam(need, free) {
    const want = need.item;
    const sticks = free.filter(x => x.ps.kind === 'ram' && x.ps.spec && x.ps.spec.gen === want.gen && x.ps.spec.form !== 'sodimm' && x.ps.spec.totalGB);
    const whole = sticks.filter(x => x.ps.spec.totalGB === want.totalGB);
    if (whole.length) return [whole[0]];
    // Group single sticks by size and speed so a pair matches.
    const groups = new Map();
    sticks.filter(x => !x.ps.spec.modules || x.ps.spec.modules === 1).forEach(x => {
      const k = `${x.ps.spec.totalGB}|${x.ps.spec.speed || ''}`;
      if (!groups.has(k)) groups.set(k, []);
      groups.get(k).push(x);
    });
    const per = want.perGB;
    const n = want.modules;
    const exact = [...groups.entries()].find(([k, xs]) => +k.split('|')[0] === per && xs.length >= n);
    if (exact) return exact[1].slice(0, n);
    const any = sticks.filter(x => x.ps.spec.totalGB === per);
    return any.length >= n ? any.slice(0, n) : null;
  }

  // ---------------------------------------------------------------------------
  // Plan
  // ---------------------------------------------------------------------------

  const priority = { cpu: 0, mobo: 1, gpu: 2, ram: 3, psu: 4, storage: 5, cooler: 6, case: 7, other: 8 };

  /*
   * builds:    Builds table rows as objects (Build ID, Build Description, Missing Parts, Built?…)
   * purchases: Purchases & Parts rows as objects (with _row)
   * Returns { builds: [...], shopping: [...], fromStock }
   */
  function plan(builds, purchases) {
    const all = purchases.map(p => ({ p, ps: partSpec(p) }));
    const free = all.filter(x => x.ps && isFree(x.p));
    const taken = new Set();
    const avail = () => free.filter(x => !taken.has(x.p['Part ID']));
    const out = [];

    const active = builds.filter(b => String(b['Build ID'] || '').trim() && norm(b['Built?']) !== 'y')
      .sort((a, b) => String(a['Build ID']).localeCompare(String(b['Build ID']), undefined, { numeric: true }));

    for (const b of active) {
      const id = String(b['Build ID']).trim();
      const mine = all.filter(x => String(x.p['Allocation'] || '').trim() === id);
      const fromDesc = S.findParts([b['Build Description'] || '']);
      const have = {};
      fromDesc.forEach(s => { if (!have[s.cat]) have[s.cat] = s; });
      mine.forEach(x => { if (x.ps && x.ps.spec && !have[x.ps.kind]) have[x.ps.kind] = x.ps.spec; });

      const missing = missingKinds(b['Missing Parts']).sort((x, y) => priority[x.kind] - priority[y.kind]);
      const needKinds = new Set(missing.map(m => m.kind));
      const notes = [];

      // Platform: from the CPU, then the board, then stock, then the default.
      let platform = cpuPlatform(have.cpu);
      let platformFrom = platform ? 'cpu' : null;
      if (!platform && have.mobo) { platform = CHIPSET_PLATFORM[have.mobo.chipset] || null; if (platform) platformFrom = 'board'; }
      const gpu = have.gpu || null;
      const gaming = !!gpu || needKinds.has('gpu');
      const stockPicks = {};

      if (!platform && needKinds.has('cpu')) {
        // Look for a CPU (and ideally a board) already in stock.
        const cpus = avail().filter(x => x.ps.kind === 'cpu' && x.ps.spec && cpuPlatform(x.ps.spec) && (!gaming || cpuFitsGaming(x.ps.spec)));
        const boards = avail().filter(x => x.ps.kind === 'mobo' && x.ps.spec);
        let pair = null;
        for (const c of cpus) {
          const pf = cpuPlatform(c.ps.spec);
          const okChips = chipsetsFor(pf, c.ps.spec);
          const bd = needKinds.has('mobo') ? boards.find(x => okChips.includes(x.ps.spec.chipset)) : null;
          if (bd) { pair = { c, bd, pf }; break; }
        }
        if (pair) { platform = pair.pf; platformFrom = 'stock'; stockPicks.cpu = [pair.c]; stockPicks.mobo = [pair.bd]; }
        else if (cpus.length && needKinds.has('mobo')) { /* a lone CPU still needs a board: fine, use it */
          platform = cpuPlatform(cpus[0].ps.spec); platformFrom = 'stock'; stockPicks.cpu = [cpus[0]];
        } else if (cpus.length) {
          platform = cpuPlatform(cpus[0].ps.spec); platformFrom = 'stock'; stockPicks.cpu = [cpus[0]];
        }
      }
      if (!platform && needKinds.has('mobo') && needKinds.has('cpu')) { platform = DEFAULT_PLATFORM; platformFrom = 'suggested'; }
      if (!platform) { platform = DEFAULT_PLATFORM; platformFrom = 'guess'; }
      const P = PLATFORMS[platform];
      const cpuForChips = have.cpu || (stockPicks.cpu && stockPicks.cpu[0].ps.spec) || null;

      let effectiveGpu = gpu || (needKinds.has('gpu') ? suggestGpu(platform) : null);
      const needs = [];

      for (const m of missing) {
        // Already allocated? Then Missing Parts is probably out of date.
        if (m.kind !== 'other' && mine.some(x => x.ps && x.ps.kind === m.kind)) {
          const x = mine.find(y => y.ps && y.ps.kind === m.kind);
          notes.push(`${x.p['Part ID']} (${KIND_NAME[m.kind]}) is allocated but “${m.text}” is still in Missing Parts.`);
          continue;
        }
        const need = { kind: m.kind, text: m.text, platform, gaming };
        switch (m.kind) {
          case 'cpu': {
            const name = P.cpu;
            need.item = S.parseAs('cpu', name);
            if (need.item) { delete need.item.index; delete need.item.end; }
            need.label = `${name} (or any ${P.name.replace(/^\S+ \(|\)$/g, '')} CPU)`;
            need.short = name;
            break;
          }
          case 'mobo': {
            const chips = chipsetsFor(platform, cpuForChips);
            need.item = { cat: 'mobo', chipsets: chips };
            need.label = `${P.name} board: ${chips.map(c => c.toUpperCase()).join(', ')}`;
            need.short = `${chips.slice(0, 3).map(c => c.toUpperCase()).join('/')}${chips.length > 3 ? '…' : ''} board`;
            break;
          }
          case 'gpu':
            need.item = effectiveGpu;
            need.label = `${S.label(effectiveGpu)} (or similar)`;
            need.short = S.label(effectiveGpu);
            break;
          case 'ram':
            need.item = ramNeed(platform);
            need.label = S.label(need.item).replace(/ RAM$/, '') + ' RAM, any speed';
            need.short = S.label(need.item);
            if (P.quad) need.note = 'Quad-channel board: four sticks';
            break;
          case 'psu': {
            const w = psuWattsFor(effectiveGpu, have.cpu, platform);
            need.item = { cat: 'psu', minWatts: w };
            need.label = `${w}W or bigger PSU${effectiveGpu ? ` (for ${S.label(Object.assign({}, effectiveGpu, { vram: null }))})` : ''}`;
            need.short = `${w}W+ PSU`;
            break;
          }
          case 'storage':
            need.item = storageNeed(platform, effectiveGpu);
            need.nvme = !!P.nvme;
            need.label = `${S.label(need.item)} or bigger${P.nvme ? ' (NVMe fine)' : ''}`;
            need.short = S.label(need.item);
            break;
          case 'case': {
            const boardForm = (have.mobo && have.mobo.form) || (stockPicks.mobo && stockPicks.mobo[0].ps.spec.form) || null;
            need.atx = !boardForm || boardForm === 'atx' || boardForm === 'eatx';
            need.item = { cat: 'generic', query: need.atx ? 'atx case' : 'matx case' };
            need.label = need.atx ? 'ATX case' : 'mATX or ATX case';
            need.short = need.atx ? 'ATX case' : 'mATX case';
            break;
          }
          case 'cooler':
            need.item = { cat: 'generic', query: `cpu cooler ${P.mount}` };
            need.label = `CPU cooler for ${P.name.split(' ')[0]}`;
            need.short = `${P.mount.toUpperCase()} CPU cooler`;
            break;
          default:
            need.item = { cat: 'generic', query: m.text };
            need.label = m.text;
            need.short = m.text;
        }
        // Stock that fits.
        let picks = stockPicks[m.kind] || null;
        if (!picks && m.kind === 'ram') picks = pickRam(need, avail());
        if (!picks && m.kind !== 'other') {
          const scored = avail().filter(x => x.ps.kind === m.kind).map(x => ({ x, s: stockScore(need, x.ps) }))
            .filter(o => o.s != null).sort((a, b) => b.s - a.s);
          if (scored.length) picks = [scored[0].x];
          need.alternatives = scored.slice(1, 4).map(o => ({ partId: o.x.p['Part ID'], row: o.x.p._row, text: o.x.ps.text }));
        }
        if (picks) {
          picks.forEach(x => taken.add(x.p['Part ID']));
          if (m.kind === 'gpu' && picks[0].ps.spec) effectiveGpu = picks[0].ps.spec; // size the PSU for the card you'll use
          if (m.kind === 'storage' && picks[0].ps.spec && picks[0].ps.spec.capGB < need.item.capGB) need.smaller = true;
          need.stock = picks.map(x => ({ partId: x.p['Part ID'], row: x.p._row, text: x.ps.text,
            check: m.kind === 'cooler' && coolerFits(x.ps.text, platform) === 'check' }));
        }
        needs.push(need);
      }
      needs.sort((a, b) => priority[a.kind] - priority[b.kind]);
      out.push({
        id, row: b._row, description: b['Build Description'] || '',
        platform: { key: platform, name: P.name, from: platformFrom },
        gpu: gpu ? S.label(gpu) : null, needs, notes,
        specs: { cpu: have.cpu || null, mobo: have.mobo || null, ram: have.ram || null, gpu: gpu || null },
        cardFor: effectiveGpu || null
      });
    }

    // Shopping list: what isn't covered by stock, grouped across builds.
    const groups = new Map();
    out.forEach(b => b.needs.filter(n => !n.stock).forEach(n => {
      const key = n.kind + '|' + JSON.stringify(n.item);
      if (!groups.has(key)) groups.set(key, { key, kind: n.kind, label: n.short || n.label, detail: n.label, item: n.item, count: 0, builds: [] });
      const g = groups.get(key);
      g.count++;
      g.builds.push(b.id);
    }));
    const shopping = [...groups.values()].sort((a, b) => priority[a.kind] - priority[b.kind] || b.count - a.count);
    shopping.forEach(g => Object.assign(g, priceFromHistory(g.item, all)));
    return { builds: out, shopping, fromStock: taken.size };
  }

  // What you've paid / valued this kind of part at before, from Purchases & Parts.
  function priceFromHistory(item, all) {
    if (!item || item.cat === 'generic') return { typical: null, seen: 0 };
    const vals = [];
    all.forEach(x => {
      if (!x.ps || !x.ps.spec || SHEET_CATEGORY[x.ps.kind] == null) return;
      if (x.ps.spec.cat !== item.cat) return;
      if (!S.matchTitle(item, x.ps.text, item.cat === 'storage' ? { typeHint: 'ssd' } : undefined).ok) return;
      const v = x.p['Estimated Part Value at Purchase Date'];
      if (typeof v === 'number' && v > 0) vals.push(v);
    });
    if (!vals.length) return { typical: null, seen: 0 };
    vals.sort((a, b) => a - b);
    const mid = vals.length >> 1;
    const med = vals.length % 2 ? vals[mid] : (vals[mid - 1] + vals[mid]) / 2;
    return { typical: Math.round(med * 100) / 100, seen: vals.length };
  }

  // ---------------------------------------------------------------------------
  // eBay links and watch entries for a shopping item
  // ---------------------------------------------------------------------------

  function searchText(item) {
    if (item.cat === 'generic') return item.query;
    return S.queries(item).primary;
  }

  function ebayUrl(item, opts) {
    const o = opts || {};
    const p = new URLSearchParams();
    p.set('_nkw', searchText(item));
    const cat = item.cat === 'generic' ? (/case/.test(item.query) ? '42014' : /cooler/.test(item.query) ? '131486' : null) : S.ebayCategory(item);
    if (cat) p.set('_sacat', cat);
    p.set('LH_PrefLoc', '1');
    if (o.sold) { p.set('LH_Sold', '1'); p.set('LH_Complete', '1'); p.set('_sop', '13'); p.set('LH_ItemCondition', '3000'); }
    else { p.set('_sop', '10'); if (o.max) p.set('_udhi', String(Math.floor(o.max))); }
    return 'https://www.ebay.co.uk/sch/i.html?' + p.toString();
  }

  // A watch entry in the same shape the Chrome extension uses.
  function watchEntry(group, alertAt, greatPct) {
    const label = group.item.cat === 'generic' ? group.label : S.label(group.item);
    const median = group.typical || null;
    const pct = typeof greatPct === 'number' ? greatPct : 25;
    const at = typeof alertAt === 'number' && alertAt > 0 ? Math.floor(alertAt) : (median ? Math.floor(median * (1 - pct / 100)) : null);
    return {
      label, item: group.item, median, n: group.seen || 0, updated: median ? Date.now() : null,
      added: Date.now(), changed: Date.now(), alertAt: at, alertFixed: typeof alertAt === 'number' && alertAt > 0,
      source: 'excel', needsPrice: !median
    };
  }

  // ---------------------------------------------------------------------------
  // Automatic alerts for builds: what each unfinished build is waiting for
  // ---------------------------------------------------------------------------

  const range = ids => {
    const sorted = ids.slice().sort((a, b) => String(a).localeCompare(String(b), undefined, { numeric: true }));
    if (sorted.length <= 2) return sorted.join(', ');
    const nums = sorted.map(id => +String(id).replace(/\D/g, ''));
    const run = nums.every((n, i) => !i || n === nums[i - 1] + 1);
    return run ? `${sorted[0]}–${sorted[sorted.length - 1]}` : `${sorted.slice(0, 3).join(', ')}${sorted.length > 3 ? ` +${sorted.length - 3}` : ''}`;
  };
  const cardName = g => S.label(Object.assign({}, g, { vram: null }));

  /*
   * From a plan: the multi-choice needs the alerts sheet looks for in bundles and
   * listings (CPU + board for a card, a CPU for a board, a card for a CPU), plus
   * single-part watches for the rest. Parts covered by stock are left out.
   *   Returns { needs, parts, seeds }
   *   needs: [{ id, type: 'combo'|'cpu'|'gpu', label, builds, gpu, ddr, platform, chipsets, cpu, wantsRam, options }]
   *   parts: watch groups (same shape as plan.shopping) for RAM, PSU, storage and fixed boards
   *   seeds: price-book entries from your own purchase history
   */
  function needsForAlerts(planResult, purchases) {
    const groups = new Map();
    const add = (key, make, buildId) => {
      if (!groups.has(key)) groups.set(key, Object.assign(make(), { builds: [] }));
      groups.get(key).builds.push(buildId);
    };
    const partKeys = new Set();
    planResult.builds.forEach(b => {
      const open = new Set(b.needs.filter(n => !n.stock).map(n => n.kind));
      const fixed = b.platform.from === 'cpu' || b.platform.from === 'board' || b.platform.from === 'stock';
      const card = b.specs.gpu || (open.has('gpu') ? null : b.cardFor);
      const ddr = b.specs.ram && b.specs.ram.gen ? b.specs.ram.gen : null;
      if (open.has('cpu') && open.has('mobo') && !fixed && card) {
        add(`combo|${PR.gpuClass(card)}|${ddr || ''}|${open.has('ram') ? 'r' : ''}`, () => ({
          type: 'combo', gpu: Object.assign({}, card, { vram: null }), ddr, wantsRam: open.has('ram'),
          options: { cpus: PR.acceptableCpus(card, { ddr }).map(c => c.name) }
        }), b.id);
      } else if (open.has('cpu') && card) {
        const chips = b.needs.find(n => n.kind === 'mobo') ? null : (b.specs.mobo ? [b.specs.mobo.chipset] : null);
        add(`cpu|${PR.gpuClass(card)}|${b.platform.key}|${chips ? chips.join('/') : ''}`, () => ({
          type: 'cpu', gpu: Object.assign({}, card, { vram: null }), platform: b.platform.key, chipsets: chips,
          options: { cpus: PR.acceptableCpus(card, { platform: b.platform.key }).filter(c => !chips || chips.some(ch => PR.boardFits(c.spec, ch))).map(c => c.name) }
        }), b.id);
      }
      if (open.has('gpu') && b.specs.cpu) {
        add(`gpu|${PR.acceptableGpuClasses(b.specs.cpu).join('-')}`, () => ({
          type: 'gpu', cpu: b.specs.cpu, options: { gpus: PR.acceptableGpus(b.specs.cpu).map(g => S.label(g)) }
        }), b.id);
      }
      // Single parts that don't depend on which CPU you end up with.
      b.needs.filter(n => !n.stock && ['ram', 'psu', 'storage'].includes(n.kind) || (!n.stock && n.kind === 'mobo' && fixed))
        .forEach(n => partKeys.add(n.kind + '|' + JSON.stringify(n.item)));
    });

    const needs = [...groups.values()].map(g => {
      const who = range(g.builds);
      const label = g.type === 'combo' ? `CPU + board for ${cardName(g.gpu)}${g.ddr ? ` (DDR${g.ddr})` : ''}`
        : g.type === 'cpu' ? `CPU for ${cardName(g.gpu)} on ${PLATFORMS[g.platform].name.split(' (')[0]}${g.chipsets ? ' ' + g.chipsets.map(c => c.toUpperCase()).join('/') : ''}`
          : `Graphics card for ${S.label(g.cpu)}`;
      const id = [g.type, g.gpu ? cardName(g.gpu) : '', g.cpu ? S.label(g.cpu) : '', g.ddr || '', g.platform || '', (g.chipsets || []).join('/'), g.wantsRam ? 'ram' : ''].join('|').toLowerCase();
      return Object.assign(g, { id, label: `${label} · ${who}` });
    });
    const parts = planResult.shopping.filter(s => partKeys.has(s.kind + '|' + JSON.stringify(s.item)));

    // Price-book seeds from your own past purchases, for every option the needs could use.
    const all = purchases.map(p => ({ p, ps: partSpec(p) }));
    const seeds = new Map();
    const seed = spec => {
      const b = PR.bookKey(spec);
      if (!b || seeds.has(b.key)) return;
      const h = priceFromHistory(b.item, all);
      seeds.set(b.key, { key: b.key, item: b.item, median: h.typical, n: h.seen, updated: h.typical ? Date.now() : null, source: h.typical ? 'sheet' : null });
    };
    needs.forEach(n => {
      if (n.type === 'combo' || n.type === 'cpu') {
        PR.acceptableCpus(n.gpu, { ddr: n.ddr, platform: n.platform }).forEach(c => {
          seed(c.spec);
          PR.chipsetsFor(c.platform, c.spec).forEach(ch => seed({ cat: 'mobo', chipset: ch }));
        });
        if (n.wantsRam) seed({ cat: 'ram', gen: n.ddr || 4, totalGB: 16 });
      }
      if (n.type === 'gpu') PR.acceptableGpus(n.cpu).forEach(seed);
    });
    return { needs, parts, seeds: [...seeds.values()] };
  }

  return {
    needsForAlerts, PAIRINGS: PR,
    PLATFORMS, KIND_NAME, SHEET_CATEGORY, cpuPlatform, chipsetsFor, looseChipset, gpuClass, psuWattsFor, suggestGpu,
    missingKinds, partSpec, partText, plan, priceFromHistory, ebayUrl, searchText, watchEntry, coolerFits
  };
})();

if (typeof module !== 'undefined' && module.exports) module.exports = DealPlanner;
