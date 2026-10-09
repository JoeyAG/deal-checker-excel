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

  // ---------------------------------------------------------------------------
  // Platforms
  // ---------------------------------------------------------------------------

  const PLATFORMS = {
    lga1155: { name: 'LGA1155 (Intel 2nd/3rd gen)', ddr: 3, chipsets: ['h61', 'b75', 'h77', 'z77', 'z68', 'p67', 'h67'], cpu: 'i5-3570', mount: '1155', nvme: false },
    lga1150: { name: 'LGA1150 (Intel 4th/5th gen)', ddr: 3, chipsets: ['h81', 'b85', 'h87', 'z87', 'h97', 'z97'], cpu: 'i5-4590', mount: '1150', nvme: false },
    lga1151: { name: 'LGA1151 (Intel 6th/7th gen)', ddr: 4, chipsets: ['h110', 'b150', 'h170', 'z170', 'b250', 'h270', 'z270'], cpu: 'i5-6500', mount: '1151', nvme: true },
    lga1151v2: { name: 'LGA1151 (Intel 8th/9th gen)', ddr: 4, chipsets: ['h310', 'b360', 'h370', 'z370', 'b365', 'z390'], cpu: 'i5-9400F', mount: '1151', nvme: true },
    lga1200: { name: 'LGA1200 (Intel 10th/11th gen)', ddr: 4, chipsets: ['h410', 'b460', 'h470', 'z490', 'h510', 'b560', 'h570', 'z590'], cpu: 'i5-10400F', mount: '1200', nvme: true },
    lga1700: { name: 'LGA1700 (Intel 12th–14th gen)', ddr: 4, chipsets: ['h610', 'b660', 'h670', 'z690', 'b760', 'h770', 'z790'], cpu: 'i5-12400F', mount: '1700', nvme: true },
    lga1851: { name: 'LGA1851 (Intel Core Ultra)', ddr: 5, chipsets: ['b860', 'z890'], cpu: 'Core Ultra 5 245K', mount: '1851', nvme: true },
    lga2011: { name: 'LGA2011 (Intel X79)', ddr: 3, quad: true, chipsets: ['x79'], cpu: 'i7-4930K', mount: '2011', nvme: false },
    lga20113: { name: 'LGA2011-3 (Intel X99)', ddr: 4, quad: true, chipsets: ['x99'], cpu: 'i7-5820K', mount: '2011', nvme: true },
    lga2066: { name: 'LGA2066 (Intel X299)', ddr: 4, quad: true, chipsets: ['x299'], cpu: 'i7-7820X', mount: '2066', nvme: true },
    am4: { name: 'AM4 (Ryzen 1000–5000)', ddr: 4, chipsets: ['a320', 'b350', 'x370', 'b450', 'x470', 'a520', 'b550', 'x570'], cpu: 'Ryzen 5 3600', mount: 'am4', nvme: true },
    am5: { name: 'AM5 (Ryzen 7000+)', ddr: 5, chipsets: ['a620', 'b650', 'b650e', 'x670', 'x670e', 'b840', 'b850', 'x870', 'x870e'], cpu: 'Ryzen 5 7600', mount: 'am5', nvme: true }
  };
  const DEFAULT_PLATFORM = 'lga1151v2'; // cheap, plentiful used, and balanced with a 1080-class card

  const CHIPSET_PLATFORM = {};
  Object.keys(PLATFORMS).forEach(k => PLATFORMS[k].chipsets.forEach(c => { CHIPSET_PLATFORM[c] = k; }));

  function cpuPlatform(cpu) {
    if (!cpu) return null;
    const f = cpu.family;
    const sfx = cpu.suffix || '';
    if (/^i\d$/.test(f)) {
      const n = cpu.num;
      if (n.length === 3) return null;
      const gen = n.length === 5 ? +n.slice(0, 2) : +n[0];
      const hedtX = /^x/.test(sfx);
      if ((gen === 3 || gen === 4) && /^[34][89]\d\d$/.test(n)) return 'lga2011';
      if (gen === 5 && /^5[89]\d\d$/.test(n)) return 'lga20113';
      if (gen === 6 && /^6[89]\d\d$/.test(n)) return 'lga20113';
      if ((gen === 7 || gen === 9 || gen === 10) && hedtX) return 'lga2066';
      if (gen <= 3) return 'lga1155';
      if (gen <= 5) return 'lga1150';
      if (gen <= 7) return 'lga1151';
      if (gen <= 9) return 'lga1151v2';
      if (gen <= 11) return 'lga1200';
      if (gen <= 14) return 'lga1700';
      return null;
    }
    if (f === 'ultra') return 'lga1851';
    if (f === 'ryzen') return +cpu.num[0] >= 7 ? 'am5' : 'am4';
    if (f === 'pentium' || f === 'celeron') {
      const m = String(cpu.num).match(/^g(\d)(\d)/);
      if (!m) return null;
      if (m[1] === '3' || m[1] === '1') return 'lga1150';
      if (m[1] === '4') return +m[2] >= 9 ? 'lga1151v2' : 'lga1151';
      if (m[1] === '5') return 'lga1151v2';
      if (m[1] === '6') return 'lga1200';
      if (m[1] === '7') return 'lga1700';
    }
    return null;
  }

  // Chipsets on a platform that run this CPU without fuss.
  function chipsetsFor(platform, cpu) {
    const all = PLATFORMS[platform].chipsets;
    if (!cpu) return all;
    if (platform === 'am4' && cpu.family === 'ryzen') {
      const g = +cpu.num[0];
      if (g === 1) return ['a320', 'b350', 'x370', 'b450', 'x470'];
      if (g === 2) return ['b350', 'x370', 'b450', 'x470', 'x570'];
      if (g === 3) return ['b450', 'x470', 'b550', 'x570', 'a520'];
      return ['b550', 'x570', 'a520', 'b450', 'x470'];
    }
    if (platform === 'lga1200' && cpu.num && cpu.num.slice(0, 2) === '11') return ['h510', 'b560', 'h570', 'z590'];
    if (platform === 'lga1155' && cpu.num && cpu.num[0] === '3') return ['b75', 'h77', 'z77', 'z68', 'h61'];
    return all;
  }

  // A board name the strict parser misses, e.g. "GA-Z97P-D3".
  function looseChipset(text) {
    const t = norm(text);
    const s = S.parseAs('mobo', t);
    if (s) return s.chipset;
    const re = /(?:^|[^a-z0-9])([abhpxz]\d{2,3})(e)?/g;
    let m;
    while ((m = re.exec(t))) {
      const c = m[1] + (m[2] && CHIPSET_PLATFORM[m[1] + 'e'] ? 'e' : '');
      if (CHIPSET_PLATFORM[c]) return c;
    }
    return null;
  }

  // ---------------------------------------------------------------------------
  // How big a PSU, which card to suggest, which CPUs are worth pairing
  // ---------------------------------------------------------------------------

  const PSU_FOR_GPU = {
    450: 'gt710 gt730 gt1030 gtx750 gtx750ti gtx950 gtx1050 gtx1050ti gtx1630 gtx1650 gtx1650super rx460 rx550 rx560 rx6400 rx6500xt arca380',
    550: 'gtx960 gtx970 gtx1060 gtx1660 gtx1660super gtx1660ti rx470 rx480 rx570 rx580 rx590 rx5500xt rtx3050 arca580 rx7600 rtx4060',
    600: 'gtx980 gtx980ti gtx1070 gtx1070ti rtx2060 rtx2060super rx5600xt rx5700 rx6600 rx6600xt rx6650xt rtx3060 rtx4060ti arca750 arca770',
    650: 'gtx1080 gtx1080ti rtx2070 rtx2070super rtx2080 rtx3060ti rtx3070 rx5700xt rx6700 rx6700xt rx6750xt rx7700xt rtx4070 rtx4070super',
    750: 'rtx2080super rtx2080ti rtx3070ti rtx3080 rtx4070ti rtx4070tisuper rx6800 rx6800xt rx7800xt rx7900gre rtx5070',
    850: 'rtx3080ti rtx3090 rtx4080 rtx4080super rx6900xt rx6950xt rx7900xt rx7900xtx rtx5070ti rtx5080',
    1000: 'rtx3090ti rtx4090 rtx5090'
  };
  const GPU_WATTS = {};
  Object.keys(PSU_FOR_GPU).forEach(w => PSU_FOR_GPU[w].split(' ').forEach(k => { GPU_WATTS[k] = +w; }));
  const gpuKey = g => (g.prefix + g.num + String(g.variant || '').replace(/\s+/g, '')).toLowerCase();

  function gpuClass(gpu) {
    if (!gpu) return 0;
    if (GPU_WATTS[gpuKey(gpu)]) return GPU_WATTS[gpuKey(gpu)];
    if (gpu.prefix === 'gt') return 450;
    return 650;
  }

  function isHedt(cpu, platform) {
    return !!(platform && PLATFORMS[platform] && PLATFORMS[platform].quad) || (cpu && cpu.family === 'i9');
  }

  function psuWattsFor(gpu, cpu, platform) {
    let w = gpu ? gpuClass(gpu) : 400;
    if (isHedt(cpu, platform)) w += 100;
    return w;
  }

  // A sensible card for a CPU that has no graphics card yet.
  function suggestGpu(platform) {
    const pick = {
      lga1155: ['gtx', '1060', '', 6], lga1150: ['gtx', '1060', '', 6], lga2011: ['gtx', '1060', '', 6],
      lga1151: ['gtx', '1070', '', 8], lga1151v2: ['gtx', '1080', '', 8], am4: ['gtx', '1080', '', 8],
      lga20113: ['gtx', '1080', '', 8], lga2066: ['gtx', '1080', 'ti', 11],
      lga1200: ['rtx', '3060', '', 12], lga1700: ['rtx', '3060', '', 12], am5: ['rtx', '3060', '', 12], lga1851: ['rtx', '3060', '', 12]
    }[platform] || ['gtx', '1070', '', 8];
    return { cat: 'gpu', prefix: pick[0], num: pick[1], variant: pick[2], vram: pick[3] };
  }

  // CPUs worth putting with a gaming card (not a Pentium with a 1080 Ti).
  function cpuFitsGaming(cpu) {
    if (!cpu) return false;
    if (/^i[579]$/.test(cpu.family) || cpu.family === 'ultra') return true;
    return cpu.family === 'ryzen' && +cpu.tier >= 5;
  }

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
        gpu: gpu ? S.label(gpu) : null, needs, notes
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

  return {
    PLATFORMS, KIND_NAME, SHEET_CATEGORY, cpuPlatform, chipsetsFor, looseChipset, gpuClass, psuWattsFor, suggestGpu,
    missingKinds, partSpec, partText, plan, priceFromHistory, ebayUrl, searchText, watchEntry, coolerFits
  };
})();

if (typeof module !== 'undefined' && module.exports) module.exports = DealPlanner;
