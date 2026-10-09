/*
 * specs.js — recognises PC parts by their specifications rather than brand.
 *
 * For each part type it knows:
 *   - how to read the specs from text (a listing title, item specifics, a sold title)
 *   - which specs must match for two items to count as "the same" for pricing
 *   - how to build an eBay search that pools every brand with those specs
 *   - how to spot the part inside a bundle or whole-PC listing
 *
 *   RAM      DDR generation, total capacity, stick layout (2x8GB), speed, desktop/laptop
 *   GPU      chip model incl. Ti/Super/XT, VRAM (any brand: MSI, Asus, Zotac…)
 *   CPU      exact model incl. suffix (8700 ≠ 8700K)
 *   Board    chipset (B450, Z390…), form factor
 *   PSU      wattage, 80+ rating
 *   Storage  type (NVMe / SATA SSD / HDD), capacity (480/500/512GB pooled)
 *
 * Plain JavaScript, no Chrome APIs, so it can be tested outside the browser.
 */
var DealSpecs = (function () {
  'use strict';

  const low = s => String(s == null ? '' : s)
    .toLowerCase()
    .replace(/\u00a0/g, ' ')
    .replace(/[\u2013\u2014]/g, '-')
    .replace(/[\u201c\u201d\u2033]/g, '"')
    .replace(/\s+/g, ' ')
    .trim();

  const CATS = {
    cpu: { name: 'CPU', ebayCat: '164' },
    gpu: { name: 'Graphics card', ebayCat: '27386' },
    mobo: { name: 'Motherboard', ebayCat: '1244' },
    ram: { name: 'RAM', ebayCat: '170083' },
    storage: { name: 'Storage', ebayCat: '175669' },
    psu: { name: 'Power supply', ebayCat: '42017' },
    pc: { name: 'Whole PC', ebayCat: '179' },
    generic: { name: 'Other', ebayCat: null }
  };
  const STORAGE_EBAY_CAT = { nvme: '175669', ssd: '175669', hdd: '56083' };
  const EBAY_TO_CAT = {
    '27386': 'gpu', '164': 'cpu', '1244': 'mobo', '170083': 'ram',
    '42017': 'psu', '175669': 'storage', '56083': 'storage'
  };
  const ORDER = ['cpu', 'gpu', 'mobo', 'ram', 'storage', 'psu'];

  // Text like "fits B360 boards" or "for RTX 3060" describes compatibility, not contents.
  function isCompatibilityMention(t, index) {
    const before = t.slice(Math.max(0, index - 24), index);
    return /\b(for|fits?|supports?|compatible(?: with)?|works with|suits?|suitable for|designed for|ideal for|requires?|required|recommended|min(?:imum)?\.?|needs?)\s[\w\s/.-]{0,14}$/.test(before);
  }

  // ---------------------------------------------------------------------------
  // CPU
  // ---------------------------------------------------------------------------

  const CPU_PATTERNS = [
    [/\bcore\s?ultra\s?([3579])\s?(\d{3})([a-z]{0,2})\b/g, m => ({ family: 'ultra', tier: m[1], num: m[2], suffix: m[3] })],
    [/\b(i[3579])\s?-?\s?(\d{3,5})([a-z]{0,2})\b/g, m => ({ family: m[1], num: m[2], suffix: m[3] })],
    [/\bryzen\s?([3579])\s?(?:pro\s?)?(\d{4})\s?(x3d|xt|x|ge|g|f)?\b/g, m => ({ family: 'ryzen', tier: m[1], num: m[2], suffix: m[3] || '' })],
    [/\bryzen\s?(\d{4})\s?(x3d|xt|x|ge|g|f)?\b/g, m => ({ family: 'ryzen', tier: null, num: m[1], suffix: m[2] || '' })],
    [/\bxeon\s?(e[357])\s?-?\s?(\d{4})\s?(v\d)?\b/g, m => ({ family: 'xeon ' + m[1], num: m[2], suffix: m[3] || '' })],
    [/\b(pentium|celeron)\s?(?:gold\s?)?([gjn]\d{3,4}[a-z]?)\b/g, m => ({ family: m[1], num: m[2], suffix: '' })],
    [/\bfx\s?-?\s?(\d{4})\b/g, m => ({ family: 'fx', num: m[1], suffix: '' })]
  ];

  function cpuFindAll(t) {
    const out = [];
    for (const [re, fn] of CPU_PATTERNS) {
      re.lastIndex = 0;
      let m;
      while ((m = re.exec(t))) {
        out.push(Object.assign({ cat: 'cpu', index: m.index, end: m.index + m[0].length }, fn(m)));
      }
    }
    return out.sort((a, b) => a.index - b.index);
  }

  function cpuLabel(s) {
    const suf = (s.suffix || '').toUpperCase();
    if (/^i\d$/.test(s.family)) return `${s.family}-${s.num}${suf}`;
    if (s.family === 'ultra') return `Core Ultra ${s.tier} ${s.num}${suf}`;
    if (s.family === 'ryzen') return `Ryzen ${s.tier ? s.tier + ' ' : ''}${s.num}${suf}`;
    if (s.family === 'fx') return `FX-${s.num}`;
    if (/^xeon/.test(s.family)) return `Xeon ${s.family.split(' ')[1].toUpperCase()}-${s.num}${s.suffix ? ' ' + s.suffix : ''}`;
    return `${s.family[0].toUpperCase()}${s.family.slice(1)} ${s.num.toUpperCase()}`;
  }

  // ---------------------------------------------------------------------------
  // GPU
  // ---------------------------------------------------------------------------

  const VRAM_SIZES = [1, 2, 3, 4, 6, 8, 10, 11, 12, 16, 20, 24, 32];
  const GPU_RE = /\b(rtx|gtx|gt|rx|arc)\s?-?\s?(\d{3,4}|[ab]\d{3})\s?(ti|super|xtx|xt|gre)?(?:\s?(super))?(?![a-z0-9])/g;

  function gpuFindAll(t) {
    const out = [];
    GPU_RE.lastIndex = 0;
    let m;
    while ((m = GPU_RE.exec(t))) {
      const variant = [m[3], m[4]].filter(Boolean).join(' ');
      const end = m.index + m[0].length;
      const s = { cat: 'gpu', prefix: m[1], num: m[2], variant, vram: null, index: m.index, end };
      // VRAM: the first "NGB" just after the model, unless it's clearly system RAM.
      const after = t.slice(end, end + 22);
      const vm = after.match(/(?:^|[^\d])(\d{1,2})\s?g(?:b)?(?![a-z0-9])(?!\s?(?:ram|ddr\d|of ram|memory kit))/);
      if (vm && VRAM_SIZES.includes(+vm[1])) s.vram = +vm[1];
      out.push(s);
    }
    return out;
  }

  function gpuLabel(s) {
    const v = s.variant ? ' ' + s.variant.split(' ').map(w => (w === 'ti' ? 'Ti' : w === 'super' ? 'Super' : w.toUpperCase())).join(' ') : '';
    return `${s.prefix.toUpperCase()} ${s.num.toUpperCase()}${v}${s.vram ? ' ' + s.vram + 'GB' : ''}`;
  }

  // ---------------------------------------------------------------------------
  // RAM
  // ---------------------------------------------------------------------------

  const STD_SPEEDS = [800, 1066, 1333, 1600, 1866, 2133, 2400, 2666, 2800, 2933, 3000, 3200, 3333, 3466, 3600,
    3733, 3866, 4000, 4133, 4266, 4400, 4600, 4800, 5200, 5600, 6000, 6200, 6400, 6800, 7200, 7600, 8000];
  const PC_RATING = { 1066: 8500, 1333: 10600, 1600: 12800, 1866: 14900, 2133: 17000, 2400: 19200, 2666: 21300,
    2933: 23400, 3200: 25600, 3600: 28800, 4800: 38400, 5200: 41600, 5600: 44800, 6000: 48000, 6400: 51200 };

  function nearestSpeed(v) {
    let best = null;
    let d = Infinity;
    for (const s of STD_SPEEDS) {
      const dd = Math.abs(s - v);
      if (dd < d) { d = dd; best = s; }
    }
    return d <= v * 0.03 ? best : null;
  }

  function ramGen(t) {
    let m = t.match(/\bddr\s?([2-5])/);
    if (m) return +m[1];
    m = t.match(/\bpc([2-5])l?\s?-?\s?\d{4,5}/);
    return m ? +m[1] : null;
  }

  function ramSpeed(t) {
    let m = t.match(/(\d{3,4})\s?(?:mhz|mt\/?s)\b/);
    if (m) return nearestSpeed(+m[1]);
    m = t.match(/\bddr[2-5]l?\s?-\s?(\d{3,4})\b/) || t.match(/\bddr[2-5]l?\s(\d{4})\b/);
    if (m) return nearestSpeed(+m[1]);
    m = t.match(/\bpc[2-5]l?\s?-?\s?(\d{4,5})/);
    if (m) return nearestSpeed(+m[1] / 8) || nearestSpeed(+m[1]);
    const nums = t.match(/(?<![\d.-])\d{4}(?![\d.])/g) || [];
    for (const x of nums) if (STD_SPEEDS.includes(+x) && +x >= 1333) return +x;
    return null;
  }

  const MODULE_COUNTS = [1, 2, 3, 4, 6, 8, 12, 16];
  const STICK_SIZES = [1, 2, 4, 8, 12, 16, 24, 32, 48, 64, 96, 128];
  const okLayout = (n, per) => MODULE_COUNTS.includes(n) && STICK_SIZES.includes(per);

  // Stick layout: "2x8GB", "2 x 8 GB", "8GB x2", "4pcs 8GB", "2 sticks of 8GB", "1x16GB".
  function ramConfig(t) {
    const pats = [
      [/\b(\d{1,2})\s?[x×*]\s?(\d{1,3})\s?(?:g|gb)?\b(?!\s?(?:mhz|mt))/, m => [+m[1], +m[2]]],
      [/\b(\d{1,3})\s?gb?\s?[x×*]\s?(\d{1,2})\b/, m => [+m[2], +m[1]]],
      [/\b(\d{1,2})\s?(?:pcs|pc|pieces?|sticks?|modules?|dimms?)\s?(?:of\s|x\s?)?(\d{1,3})\s?gb\b/, m => [+m[1], +m[2]]],
      [/\b(\d{1,3})\s?gb\s?(?:sticks?|modules?|dimms?)\s?x\s?(\d{1,2})\b/, m => [+m[2], +m[1]]]
    ];
    for (const [re, fn] of pats) {
      const m = t.match(re);
      if (m) {
        const [n, per] = fn(m);
        if (okLayout(n, per)) return { modules: n, per };
      }
    }
    return null;
  }

  // Words that say how many sticks without saying their size, e.g. "single stick", "quad kit".
  function ramStickWords(t) {
    if (/\bsingle\s?(?:ram\s)?(?:stick|module|dimm)\b|\b1\s?(?:stick|module|dimm|pc)\b|\bone\s?(?:stick|module)\b/.test(t)) return { modules: 1 };
    const kitOf = t.match(/\bkit\s?of\s?([2-8])\b|\b([2-8])[\s-]?(?:pack|pk)\b|\b([2-8])[\s-]?(?:stick|module|dimm)s?\s?kit\b/);
    if (kitOf) return { modules: +(kitOf[1] || kitOf[2] || kitOf[3]) };
    if (/\b(?:dual|2)[\s-]?channel\s?kit\b|\bdual\s?kit\b|\bpair\b/.test(t)) return { modules: 2 };
    if (/\b(?:quad|4)[\s-]?channel\s?kit\b|\bquad\s?kit\b/.test(t)) return { modules: 4 };
    if (/\bocta?[\s-]?channel\s?kit\b/.test(t)) return { modules: 8 };
    if (/\bkit\b/.test(t)) return { kit: true };
    if (/\b(?:ram\s|memory\s)?stick\b(?!s)/.test(t) && !/\bmemory stick\b/.test(t)) return { modules: 1 }; // "8GB stick"
    return null;
  }

  function ramParse(t, forcedTotal) {
    const gen = ramGen(t);
    const cfg = ramConfig(t);
    let total = cfg ? cfg.modules * cfg.per : (forcedTotal || null);
    if (!total) {
      const m = t.match(/\b(\d{1,3})\s?gb\b/);
      if (m) total = +m[1];
    }
    if (!gen && !total) return null;
    let modules = cfg ? cfg.modules : null;
    let per = cfg ? cfg.per : null;
    let kit = false;
    if (cfg && total && cfg.modules * cfg.per !== total) { modules = null; per = null; }
    if (!modules) {
      const w = ramStickWords(t);
      if (w && w.modules && total && total % w.modules === 0 && okLayout(w.modules, total / w.modules)) {
        modules = w.modules;
        per = total / w.modules;
      } else if (w && (w.kit || w.modules > 1)) {
        kit = true;
      }
    }
    return {
      cat: 'ram',
      gen,
      totalGB: total,
      modules,
      perGB: per,
      kit: kit || (modules != null && modules > 1),
      speed: ramSpeed(t),
      form: /so-?dimm|\bsodimm|laptop|notebook/.test(t) ? 'sodimm' : null,
      ecc: /\b(ecc|registered|rdimm|lrdimm|server)\b/.test(t) && !/non[\s-]?ecc/.test(t)
    };
  }

  // Finds RAM inside mixed text such as "i5 8400, GTX 1060 6GB, 16GB DDR4 RAM, 500GB SSD".
  function ramFindAll(t) {
    const anchors = [];
    const re = /\bddr\s?[2-5]\b|\bram\b|\bpc[2-5]l?-\d{4,5}|\b(?:so-?)?dimms?\b|\bmemory\b/g;
    let m;
    while ((m = re.exec(t))) anchors.push({ index: m.index, word: m[0] });
    for (const a of anchors) {
      const from = Math.max(0, a.index - 26);
      const win = t.slice(from, a.index + 34);
      const genInWin = ramGen(win);
      if (a.word === 'memory' && !genInWin) continue; // "12GB memory" on a GPU isn't system RAM
      // Pick the capacity closest to the anchor that isn't storage or VRAM.
      const capRe = /\b(\d{1,3})\s?gb\b/g;
      let c;
      let best = null;
      while ((c = capRe.exec(win))) {
        const after = win.slice(c.index + c[0].length, c.index + c[0].length + 12);
        const before = win.slice(Math.max(0, c.index - 14), c.index);
        if (/^\s?(ssd|hdd|nvme|m\.2|emmc|storage|gddr|vram|hard)/.test(after)) continue;
        if (/(rtx|gtx|gt|rx|arc)\s?-?\s?[ab]?\d{3,4}\s?(ti|super|xtx|xt)?\s?$/.test(before)) continue;
        if (/\d\s?[x×*]\s?$/.test(before)) continue;   // the "8GB" in "4 x 8GB" is a stick, not the total
        const pos = from + c.index;
        const dist = Math.abs(pos - a.index);
        if (!best || dist < best.dist) best = { total: +c[1], dist, pos };
      }
      if (!best) continue;
      const cfgInWin = ramConfig(win);
      const spec = ramParse(win, cfgInWin ? cfgInWin.modules * cfgInWin.per : best.total);
      if (!spec) continue;
      spec.index = Math.min(best.pos, a.index);
      spec.end = spec.index + 20;
      return [spec];
    }
    return [];
  }

  // The spec a sale most often leaves out, for explaining "didn't state every spec".
  function detailWord(spec) {
    if (!spec) return 'every spec';
    if (spec.cat === 'ram') return 'the stick layout';
    if (spec.cat === 'gpu') return 'the VRAM';
    if (spec.cat === 'mobo') return 'the board size';
    if (spec.cat === 'psu') return 'the 80+ rating';
    if (spec.cat === 'storage') return spec.type === 'nvme' ? 'the PCIe generation' : spec.type === 'hdd' ? 'the spin speed' : 'the drive type';
    return 'every spec';
  }

  // "2×8GB", "Single 16GB stick", "kit, layout not stated", "any layout"
  function layoutText(s) {
    if (s.modules === 1) return `Single ${s.perGB}GB stick`;
    if (s.modules && s.perGB) return `${s.modules}\u00d7${s.perGB}GB`;
    return s.kit ? 'kit, layout not stated' : 'any layout';
  }

  function ramLabel(s) {
    return [
      s.gen ? `DDR${s.gen}` : '',
      s.totalGB ? `${s.totalGB}GB` : '',
      s.modules && s.perGB ? `(${s.modules}x${s.perGB}GB)` : s.kit ? 'kit' : '',
      s.speed ? `${s.speed}MHz` : '',
      s.form === 'sodimm' ? 'SODIMM' : '',
      s.ecc ? 'ECC' : '',
      'RAM'
    ].filter(Boolean).join(' ');
  }

  // ---------------------------------------------------------------------------
  // Motherboard
  // ---------------------------------------------------------------------------

  const CHIPSETS = new Set((
    'h61 h67 p67 z68 b75 h77 z77 h81 b85 h87 z87 h97 z97 x79 x99 h110 b150 h170 z170 b250 h270 z270 ' +
    'h310 b360 h370 z370 b365 z390 h410 b460 h470 z490 h510 b560 h570 z590 h610 b660 h670 z690 b760 h770 ' +
    'z790 b860 z890 x299 a320 b350 x370 b450 x470 a520 b550 x570 a620 b650 b650e x670 x670e b840 b850 x870 x870e'
  ).split(' '));
  const MOBO_RE = /\b([abhpxz]\d{2,3})(e)?([mi])?(?:[-\s]?(?:plus|pro|max|a|e|k|m|i|s))*(?![a-z0-9])/g;

  function moboForm(t, suffix) {
    if (/\be-?atx\b/.test(t)) return 'eatx';
    if (/\bmini[\s-]?itx\b|\bitx\b/.test(t) || suffix === 'i') return 'itx';
    if (/\bmicro[\s-]?atx\b|\bm-?atx\b|\bmatx\b|\bu?atx\s?m\b/.test(t) || suffix === 'm') return 'matx';
    if (/\batx\b/.test(t)) return 'atx';
    return null;
  }

  function moboFindAll(t, needAnchor) {
    const out = [];
    MOBO_RE.lastIndex = 0;
    let m;
    while ((m = MOBO_RE.exec(t))) {
      let chip = m[1];
      if (m[2] && CHIPSETS.has(chip + 'e')) chip += 'e';
      if (!CHIPSETS.has(chip)) continue;
      if (needAnchor) {
        const around = t.slice(Math.max(0, m.index - 45), m.index + 55);
        if (!/\b(motherboard|mobo|mainboard|main board|board|combo|bundle)s?\b/.test(around)) continue;
      }
      out.push({ cat: 'mobo', chipset: chip, form: moboForm(t.slice(Math.max(0, m.index - 30), m.index + 60), m[3]), index: m.index, end: m.index + m[0].length });
    }
    return out;
  }

  function moboLabel(s) {
    if (s.chipsets) return `${s.chipsets.map(c => c.toUpperCase()).join('/')} motherboard`;
    const form = { atx: 'ATX', matx: 'mATX', itx: 'ITX', eatx: 'E-ATX' }[s.form] || '';
    return `${s.chipset.toUpperCase()} motherboard${form ? ' ' + form : ''}`;
  }

  // ---------------------------------------------------------------------------
  // Power supply
  // ---------------------------------------------------------------------------

  function psuRating(t) {
    const m = t.match(/80\s?(?:\+|plus)\s?(white|bronze|silver|gold|platinum|titanium)?/) ||
      t.match(/\b(bronze|silver|gold|platinum|titanium)\b(?=[^.]{0,20}\b(rated|certified|psu|80|efficien))/);
    if (!m) return null;
    return m[1] || 'standard';
  }

  function psuWatts(t) {
    const re = /\b(\d{3,4})\s?(?:w|watts?)(?![a-z0-9])/g;
    let m;
    while ((m = re.exec(t))) {
      const w = +m[1];
      if (w >= 200 && w <= 2000) return { watts: w, index: m.index };
    }
    return null;
  }

  // Model names that carry the wattage, e.g. Corsair HX1000, RM850x, CX650M, VS550, SF750.
  function psuModelWatts(t) {
    const m = t.match(/\b(?:hx|hxi|ax|axi|rm|rmx|rme|rmi|cx|cxm|tx|txm|vs|cv|sf|gx|gf|bq|mwe)\s?-?(\d{3,4})(?:[a-z]{0,2})\b/);
    if (!m) return null;
    const w = +m[1];
    return w >= 300 && w <= 1600 ? { watts: w, index: m.index } : null;
  }

  function psuFindAll(t) {
    const re = /\b(psu|power supply|power unit|80\s?(?:\+|plus))/g;
    let m;
    while ((m = re.exec(t))) {
      const from = Math.max(0, m.index - 34);
      const win = t.slice(from, m.index + 40);
      const w = psuWatts(win);
      if (w) {
        return [{ cat: 'psu', watts: w.watts, rating: psuRating(win), index: from + w.index, end: from + w.index + 6 }];
      }
    }
    return [];
  }

  function psuLabel(s) {
    if (s.minWatts) return `${s.minWatts}W+ PSU`;
    const r = s.rating && s.rating !== 'standard' ? ` 80+ ${s.rating[0].toUpperCase()}${s.rating.slice(1)}` : s.rating ? ' 80+' : '';
    return `${s.watts}W${r} PSU`;
  }

  // ---------------------------------------------------------------------------
  // Storage: NVMe SSDs, SATA SSDs and hard drives
  // ---------------------------------------------------------------------------
  //
  //   NVMe SSD   capacity; PCIe generation if stated; small M.2 sizes (2230/2242) kept separate
  //   SATA SSD   capacity; form (2.5" / M.2 SATA / mSATA) if stated
  //   Hard drive exact capacity; 3.5" vs 2.5" and spin speed if stated; SAS/IDE kept separate
  //   External drives are never mixed with internal ones.
  //
  // Many titles never say "SSD" or "hard drive" ("Samsung 870 EVO 1TB", "ST2000DM008"),
  // so well-known model families fill the gap.

  const CAP_BUCKETS = [64, 120, 160, 250, 320, 500, 750, 1000, 1500, 2000, 3000, 4000, 5000, 6000, 8000,
    10000, 12000, 14000, 16000, 18000, 20000, 22000, 24000];

  // 128 → 120, 256 → 250, 480/512 → 500, 960/1024 → 1000, 1920/2048 → 2000
  function capBucket(gb) {
    let best = null;
    let d = Infinity;
    for (const b of CAP_BUCKETS) {
      const dd = Math.abs(Math.log(gb / b));
      if (dd < d) { d = dd; best = b; }
    }
    return d < 0.12 ? best : Math.round(gb);
  }

  // [pattern, type, PCIe generation (NVMe only)]
  const DRIVE_FAMILIES = [
    // NVMe
    [/\b980\s?pro\b/, 'nvme', 4], [/\b990\s?(?:pro|evo)(?:\s?plus)?\b/, 'nvme', 4],
    [/\b9[67]0\s?(?:evo|pro)(?:\s?plus)?\b/, 'nvme', 3], [/\bsamsung\s?(?:ssd\s?)?980\b/, 'nvme', 3],
    [/\bpm9a1\b/, 'nvme', 4], [/\bpm9(?:81|91)a?\b|\bsm9[56]1\b|\bmz-?v[5-8]/, 'nvme', 3],
    [/\bsn\s?580\b|\bsn\s?770\b|\bsn\s?8[15]0x?\b/, 'nvme', 4], [/\bsn\s?(?:5[2-7]0|7[05]0|720|730)\b/, 'nvme', 3],
    [/\bp[35]\s?plus\b|\bcrucial\s?t500\b/, 'nvme', 4], [/\bcrucial\s?t700\b/, 'nvme', 5],
    [/\b(?:crucial\s?)p[1235]\b(?!\s?plus)|\bct\d+p[1235]/, 'nvme', 3],
    [/\bnv[23]\b|\bsnv[23]s|\bkc3000\b|\bfury renegade\b/, 'nvme', 4], [/\bnv1\b|\ba2000\b|\bsnvs/, 'nvme', 3],
    [/\bplatinum\s?p41\b|\bpc801\b/, 'nvme', 4], [/\bgold\s?p31\b|\bbc[57]01\b|\bbc711\b|\bpc[67]01\b|\bpc711\b/, 'nvme', 3],
    [/\b[67][67]0p\b|\b760p\b/, 'nvme', 3], [/\bfirecuda\s?5[23]0\b/, 'nvme', 4], [/\bfirecuda\s?510\b/, 'nvme', 3],
    [/\bsx8200\b|\bmp3[34]\b|\bmp510\b|\bnm6[12]0\b|\bcs3030\b/, 'nvme', 3],
    [/\blegend\s?(?:7[05]0|8[05]0|960)\b|\bmp600\b|\bnm(?:7[01]0|790|800)\b/, 'nvme', 4],
    [/\bsabrent\s?rocket\b|\bgammix\b/, 'nvme', null],
    // SATA SSD
    [/\b8[4-7]0\s?(?:evo|qvo|pro)\b|\bmz-?7|\bpm8\d{2}\b/, 'ssd', null],
    [/\b(?:mx|bx)\s?[1-5]00\b|\bct\d+(?:mx|bx)\d/, 'ssd', null],
    [/\bsa[45]10\b|\bsa500\b|\bwd\s?(?:blue|green)\s?3d\b/, 'ssd', null],
    [/\ba400\b|\bsa400|\bkc600\b|\buv[45]00\b|\bcs900\b|\bsu[678][03]0\b|\b545s\b/, 'ssd', null],
    [/\bsandisk\s?(?:ssd\s?)?(?:plus|ultra\s?(?:3d|ii))\b/, 'ssd', null],
    // Hard drives
    [/\bwd\d{2,5}[a-z]{4}\b|\bst\d{3,5}[a-z]{2}\d{3}\b|\b(?:dt01aca|hdwd|hdwg|hdwe|hdwn|mg0\d)\w*/, 'hdd', null],
    [/\b(?:ironwolf|skyhawk|exos|deskstar|ultrastar|travelstar|barracuda)\b/, 'hdd', null],
    [/\bwd\s?(?:red|purple|gold)\b|\btoshiba\s?[pxns]300\b/, 'hdd', null]
  ];

  function familyOf(t) {
    for (const [re, type, gen] of DRIVE_FAMILIES) if (re.test(t)) return { type, gen };
    return null;
  }

  // 'nvme' | 'ssd' (SATA) | 'hdd' | null. hint comes from the eBay category when known.
  function storageType(t, hint) {
    if (/\bnvme\b|\bpci-?e\b|\bgen\s?[345]\b/.test(t)) return 'nvme';
    if (/\bsshd\b/.test(t)) return 'hdd';
    if (/\bssd\b|solid state|\bmsata\b/.test(t)) return (/\bm\.?2\b/.test(t) && !/\bsata\b/.test(t)) ? 'nvme' : 'ssd';
    if (/\bhdd\b|hard (?:drive|disk)|\d{4,5}\s?rpm\b|\brpm\b/.test(t)) return 'hdd';
    const fam = familyOf(t);
    if (fam) return fam.type;
    if (/\bm\.?2\b/.test(t)) return /\bsata\b/.test(t) ? 'ssd' : 'nvme';
    return hint || null;
  }

  function capGB(num, unit) {
    return unit === 'tb' ? Math.round(parseFloat(num) * 1000) : Math.round(parseFloat(num));
  }

  // Reads every storage characteristic from text already known to describe one drive.
  function storageDetails(t, type) {
    const d = { form: null, size: null, gen: null, rpm: null, iface: null };
    d.ext = /\b(?:external|portable|usb(?:\s?[23](?:\.\d)?|-?c)?|my passport|my book|wd elements|seagate expansion|backup plus|one touch|samsung t[579]|extreme portable)\b/.test(t);
    if (type === 'nvme') {
      const sz = t.match(/\b(2230|2242|2260|2280|22110)\b/);
      if (sz) d.size = sz[1];
      const g = t.match(/\bgen\s?([345])\b|pci-?e\s?(?:gen\s?)?([345])(?:\.0)?\b/);
      if (g) d.gen = +(g[1] || g[2]);
      else { const fam = familyOf(t); if (fam && fam.type === 'nvme' && fam.gen) d.gen = fam.gen; }
    } else if (type === 'ssd') {
      if (/\bmsata\b/.test(t)) d.form = 'msata';
      else if (/\bm\.?2\b/.test(t)) d.form = 'm.2';
      else if (/\b2\.5\b/.test(t)) d.form = '2.5';
      if (/\bsas\b|\bu\.2\b/.test(t)) d.iface = 'sas';
    } else if (type === 'hdd') {
      if (/\b3\.5\b|\bdesktop\b/.test(t)) d.form = '3.5';
      else if (/\b2\.5\b|\blaptop\b|\bnotebook\b/.test(t)) d.form = '2.5';
      const r = t.match(/\b(\d{4,5})\s?rpm\b/) || t.match(/\b(5400|5700|5900|7200|10000|10k|15000|15k)\b(?!\s?(?:mhz|mb|gb))/);
      if (r) { const v = /k$/.test(r[1]) ? 10000 : +r[1]; d.rpm = v <= 5900 ? '5400' : v < 10000 ? '7200' : '10k'; }
      if (/\bsas\b/.test(t)) d.iface = 'sas';
      else if (/\b(?:ide|pata)\b/.test(t)) d.iface = 'ide';
    }
    return d;
  }

  function makeStorage(type, gb, t) {
    return Object.assign({ cat: 'storage', type, capGB: capBucket(gb) }, storageDetails(t, type));
  }

  // Drives inside mixed text, e.g. "... 16GB RAM, 1TB NVMe, 2TB HDD ...".
  function storageFindAll(t, hint) {
    const out = [];
    const re = /\b(\d+(?:\.\d+)?)\s?(tb|gb)\b/g;
    let m;
    while ((m = re.exec(t))) {
      const gb = capGB(m[1], m[2]);
      if (gb < 60) continue;
      const after = t.slice(m.index + m[0].length, m.index + m[0].length + 34).split(/[,+|;/]|\band\b|\bwith\b/)[0];
      const before = t.slice(Math.max(0, m.index - 30), m.index).split(/[,+|;/]|\band\b|\bwith\b/).pop();
      if (/^\s?(?:ram|ddr|memory|gddr|vram|of ram)/.test(after)) continue;
      const ctx = before + ' ' + m[0] + after;
      let type = storageType(after, null) || storageType(before, null);
      if (!type) continue;
      out.push(Object.assign(makeStorage(type, gb, ctx), { index: m.index, end: m.index + m[0].length }));
    }
    return out;
  }

  // The whole text describes one drive (a listing or a sold title).
  function storageLoose(t, hint) {
    const type = storageType(t, hint);
    const re = /\b(\d+(?:\.\d+)?)\s?(tb|gb)\b/g;
    let m;
    while ((m = re.exec(t))) {
      const gb = capGB(m[1], m[2]);
      if (gb < 60) continue;
      const s = makeStorage(type, gb, t);
      if (!type) s.type = null; // caller decides what an unstated type means
      return s;
    }
    return null;
  }

  function capText(gb) {
    return gb >= 1000 ? `${+(gb / 1000).toFixed(1)}TB` : `${gb}GB`;
  }

  const RPM_TEXT = { 5400: '5400rpm', 7200: '7200rpm', '10k': '10000rpm' };

  function storageLabel(s) {
    const ext = s.ext ? 'External ' : '';
    if (s.type === 'nvme') return `${ext}${capText(s.capGB)} NVMe${s.gen ? ` Gen${s.gen}` : ''} M.2${s.size ? ' ' + s.size : ''} SSD`;
    if (s.type === 'ssd') {
      const form = { '2.5': ' 2.5"', 'm.2': ' M.2', msata: ' mSATA' }[s.form] || '';
      return `${ext}${capText(s.capGB)} ${s.form === 'msata' ? '' : 'SATA'}${form} SSD`.replace(/\s+/g, ' ');
    }
    return `${ext}${capText(s.capGB)}${s.form ? ` ${s.form}"` : ''}${s.rpm ? ' ' + RPM_TEXT[s.rpm] : ''}${s.iface === 'sas' ? ' SAS' : s.iface === 'ide' ? ' IDE' : ''} HDD`;
  }

  function storageChips(s) {
    const chips = [];
    if (s.ext) chips.push('External');
    if (s.type === 'nvme') {
      chips.push('NVMe SSD', capText(s.capGB), s.gen ? `PCIe Gen${s.gen}` : 'any PCIe gen', s.size && s.size !== '2280' ? `M.2 ${s.size} only` : 'M.2 2280');
    } else if (s.type === 'ssd') {
      chips.push('SATA SSD', capText(s.capGB), ({ '2.5': '2.5"', 'm.2': 'M.2 SATA', msata: 'mSATA' }[s.form]) || 'any form');
    } else {
      chips.push('Hard drive', capText(s.capGB), s.form ? `${s.form}"` : 'any size', s.rpm ? RPM_TEXT[s.rpm] : 'any speed');
      if (s.iface === 'sas' || s.iface === 'ide') chips.push(s.iface.toUpperCase());
    }
    chips.push('any brand');
    return chips;
  }

  // Which drive type an eBay category implies, used when a title doesn't say.
  function storageHintForCategory(catId) {
    return catId === STORAGE_EBAY_CAT.hdd ? 'hdd' : catId === STORAGE_EBAY_CAT.nvme ? 'ssd' : null;
  }

  // ---------------------------------------------------------------------------
  // Shared: find, parse, label
  // ---------------------------------------------------------------------------

  // Strict detection: things clearly present in mixed text (bundles, sold titles).
  function findAll(cat, t, opts) {
    const guard = opts && opts.guard;
    let list;
    switch (cat) {
      case 'cpu': list = cpuFindAll(t); break;
      case 'gpu': list = gpuFindAll(t); break;
      case 'ram': list = ramFindAll(t); break;
      case 'mobo': list = moboFindAll(t, true); break;
      case 'psu': list = psuFindAll(t); break;
      case 'storage': list = storageFindAll(t); break;
      default: list = [];
    }
    if (guard) list = list.filter(s => !isCompatibilityMention(t, s.index));
    return list;
  }

  // Loose parsing: the whole text is known to be about one part of this type.
  function parseAs(cat, text, opts) {
    const t = low(text);
    switch (cat) {
      case 'cpu': return cpuFindAll(t)[0] || null;
      case 'gpu': return gpuFindAll(t)[0] || null;
      case 'ram': return ramParse(t);
      case 'mobo': return moboFindAll(t, false)[0] || null;
      case 'psu': {
        const w = psuWatts(t) || psuModelWatts(t);
        if (!w) return null;
        const bare = t.match(/\b(bronze|silver|gold|platinum|titanium)\b/);
        return { cat: 'psu', watts: w.watts, rating: psuRating(t) || (bare ? bare[1] : null) };
      }
      case 'storage': return storageLoose(t, opts && opts.typeHint);
      default: return null;
    }
  }

  function strip(spec) {
    if (!spec) return spec;
    const s = Object.assign({}, spec);
    delete s.index;
    delete s.end;
    return s;
  }

  function label(spec) {
    switch (spec.cat) {
      case 'cpu': return cpuLabel(spec);
      case 'gpu': return gpuLabel(spec);
      case 'ram': return ramLabel(spec);
      case 'mobo': return moboLabel(spec);
      case 'psu': return psuLabel(spec);
      case 'storage': return storageLabel(spec);
      case 'pc': return `PC with ${cpuLabel(spec.cpu)} + ${spec.gpu ? gpuLabel(spec.gpu) : 'no graphics card'}`;
      default: return spec.query || '';
    }
  }

  // Short tags shown in the panel under "Matching".
  function chips(spec) {
    switch (spec.cat) {
      case 'ram': return [
        spec.gen ? `DDR${spec.gen}` : 'DDR type unknown',
        spec.totalGB ? `${spec.totalGB}GB` : null,
        layoutText(spec),
        spec.speed ? `${spec.speed}MHz` : 'any speed',
        spec.form === 'sodimm' ? 'Laptop (SODIMM)' : 'Desktop',
        spec.ecc ? 'ECC' : null
      ].filter(Boolean);
      case 'gpu': return [gpuLabel(Object.assign({}, spec, { vram: null })), spec.vram ? `${spec.vram}GB` : 'any VRAM', 'any brand'];
      case 'cpu': return [cpuLabel(spec)];
      case 'mobo': if (spec.chipsets) return [spec.chipsets.map(c => c.toUpperCase()).join(' / '), 'any brand'];
        return [spec.chipset.toUpperCase(), ({ atx: 'ATX', matx: 'mATX', itx: 'ITX', eatx: 'E-ATX' }[spec.form]) || 'any size', 'any brand'];
      case 'psu': if (spec.minWatts) return [`${spec.minWatts}W or more`, 'any brand'];
        return [`${spec.watts}W`, spec.rating ? (spec.rating === 'standard' ? '80+' : `80+ ${spec.rating}`) : 'any rating', 'any brand'];
      case 'storage': return storageChips(spec);
      case 'pc': return [cpuLabel(spec.cpu), spec.gpu ? gpuLabel(spec.gpu) : 'no graphics card', 'any RAM/storage'];
      default: return ['Title words'];
    }
  }

  // ---------------------------------------------------------------------------
  // Searches: pool every brand, let the spec matcher do the precise filtering
  // ---------------------------------------------------------------------------

  // eBay's search treats "(a,b,c)" as "any of these words".
  function anyOf(list) {
    const u = [...new Set(list.filter(Boolean))];
    return u.length > 1 ? '(' + u.join(',') + ')' : (u[0] || '');
  }

  // "650W or more" means the usual sizes from 650W up to about half as much again.
  const PSU_SIZES = [450, 500, 550, 600, 650, 700, 750, 800, 850, 900, 1000, 1050, 1200];
  const psuCeiling = w => Math.max(w + 150, Math.round(w * 1.55));

  function capWords(gb) {
    const near = {
      120: ['120gb', '128gb'], 250: ['240gb', '250gb', '256gb'], 500: ['480gb', '500gb', '512gb'],
      1000: ['1tb', '960gb', '1000gb'], 2000: ['2tb', '1920gb', '2000gb']
    }[gb];
    return near || [gb >= 1000 ? `${+(gb / 1000).toFixed(1)}tb` : `${gb}gb`];
  }

  function queries(spec) {
    const join = parts => parts.filter(Boolean).join(' ');
    switch (spec.cat) {
      case 'ram': {
        const g = spec.gen ? `ddr${spec.gen}` : 'ram';
        const caps = [`${spec.totalGB}gb`];
        if (spec.modules && spec.perGB) caps.push(`${spec.modules}x${spec.perGB}gb`);
        const speed = spec.speed ? anyOf([`${spec.speed}`, `${spec.speed}mhz`, PC_RATING[spec.speed] ? `${PC_RATING[spec.speed]}` : null]) : '';
        const form = spec.form === 'sodimm' ? anyOf(['sodimm', 'so-dimm', 'laptop']) : '';
        return {
          primary: join([g, spec.totalGB ? anyOf(caps) : '', speed, form]),
          fallback: join([g, spec.totalGB ? `${spec.totalGB}gb` : '', form ? 'sodimm' : ''])
        };
      }
      case 'gpu': {
        const n = spec.num;
        const p = spec.prefix;
        const words = spec.variant ? spec.variant.split(' ') : [];
        const v = words.length ? words[words.length - 1] : '';
        const joined = words.join('');
        const first = anyOf([n, p + n].concat(...words.map(w => [n + w, p + n + w]), joined && words.length > 1 ? [n + joined] : []));
        const second = v ? anyOf([v, n + v, p + n + v, words.length > 1 ? n + joined : null]) : '';
        return { primary: join([first, second]), fallback: join([p, n, v]) };
      }
      case 'cpu': {
        if (/^i\d$/.test(spec.family)) {
          return { primary: `${spec.num}${spec.suffix}`, fallback: `${spec.family} ${spec.num}${spec.suffix}` };
        }
        if (spec.family === 'ryzen') return { primary: `ryzen ${spec.num}${spec.suffix}`, fallback: `${spec.num}${spec.suffix}` };
        const l = cpuLabel(spec).toLowerCase();
        return { primary: l, fallback: spec.num };
      }
      case 'mobo': {
        if (spec.chipsets) return { primary: anyOf(spec.chipsets), fallback: `${spec.chipsets[0]} motherboard` };
        const c = spec.chipset;
        return { primary: anyOf([c, c + 'm', c + 'i']), fallback: `${c} motherboard` };
      }
      case 'psu':
        if (spec.minWatts) {
          const range = PSU_SIZES.filter(w => w >= spec.minWatts && w <= psuCeiling(spec.minWatts));
          return { primary: anyOf(range.map(w => `${w}w`)), fallback: `${spec.minWatts}w psu` };
        }
        return { primary: anyOf([`${spec.watts}w`, `${spec.watts}`]), fallback: `${spec.watts}w psu` };
      case 'storage': {
        const caps = anyOf(capWords(spec.capGB));
        const ext = spec.ext ? anyOf(['external', 'portable', 'usb']) : '';
        if (spec.type === 'nvme') {
          return { primary: join([anyOf(['nvme', 'm.2', 'pcie']), caps, spec.size && spec.size !== '2280' ? spec.size : '', ext]),
            fallback: join(['nvme', capText(spec.capGB).toLowerCase(), ext]) };
        }
        if (spec.type === 'ssd') {
          return { primary: join([anyOf(['ssd', 'sata', 'solid']), caps, ext]), fallback: join(['ssd', capText(spec.capGB).toLowerCase(), ext]) };
        }
        return { primary: join([anyOf(['hdd', 'hard', 'sata', 'rpm', '7200rpm', '5400rpm']), caps, ext]),
          fallback: join(['hdd', capText(spec.capGB).toLowerCase(), ext]) };
      }
      case 'pc': {
        const c = queries(spec.cpu);
        const g = spec.gpu ? anyOf([spec.gpu.num, spec.gpu.prefix + spec.gpu.num]) : '';
        return { primary: join([c.primary, g]), fallback: join([c.fallback, spec.gpu ? spec.gpu.num : 'pc']) };
      }
      default:
        return { primary: spec.query, fallback: spec.query };
    }
  }

  function ebayCategory(spec) {
    if (spec.cat === 'storage') return spec.ext ? null : (STORAGE_EBAY_CAT[spec.type] || null);
    if (spec.cat === 'ram' && spec.form === 'sodimm') return null; // laptop RAM lives elsewhere
    return (CATS[spec.cat] && CATS[spec.cat].ebayCat) || null;
  }

  // ---------------------------------------------------------------------------
  // Is this sold listing the same thing, by spec?
  // ---------------------------------------------------------------------------

  const cpuKey = s => `${s.family}|${s.num}|${s.suffix || ''}`;
  const gpuKey = s => `${s.prefix}|${s.num}|${s.variant || ''}`;

  const PC_WORDS = /\b(pc|computer|desktop|tower|workstation|system|sff|mini pc|optiplex|elitedesk|prodesk|thinkcentre|thinkstation|precision|build|rig)\b/;

  // A sold whole PC matches if it has the same CPU and the same graphics card (or none).
  function matchPc(target, t) {
    if (/\b(laptop|notebook|macbook|chromebook|all[\s-]?in[\s-]?one)\b/.test(t)) return { ok: false, why: 'spec' };
    if (!PC_WORDS.test(t)) return { ok: false, why: 'spec' };
    const cpus = findAll('cpu', t, { guard: true });
    if (!cpus.some(c => cpuKey(c) === cpuKey(target.cpu))) return { ok: false, why: 'spec' };
    const gpus = findAll('gpu', t, { guard: true });
    if (target.gpu) {
      const g = target.gpu;
      if (!gpus.some(x => gpuKey(x) === gpuKey(g) && (!g.vram || !x.vram || x.vram === g.vram))) return { ok: false, why: 'spec' };
    } else if (gpus.length) {
      return { ok: false, why: 'spec' };
    }
    return { ok: true };
  }

  // Returns { ok, why, exact, spec }. "exact" means the sale stated every spec that matters
  // here (e.g. the stick layout or the PCIe gen); "ok" alone means nothing contradicted it.
  // A title that lists several sizes side by side ("128GB 256GB 512GB 1TB", "8GB/16GB/32GB")
  // is a listing with options: its price is usually the cheapest option, not the one you want.
  function sizeOptions(text) {
    const t = low(text);
    const re = /(\d+(?:\.\d+)?)\s?(tb|gb)\b/g;
    const toks = [];
    let m;
    while ((m = re.exec(t))) toks.push({ gb: m[2] === 'tb' ? Math.round(parseFloat(m[1]) * 1000) : parseFloat(m[1]), index: m.index, end: m.index + m[0].length });
    for (let i = 1; i < toks.length; i++) {
      const between = t.slice(toks[i - 1].end, toks[i].index);
      if (toks[i].gb !== toks[i - 1].gb && between.length <= 5 && /^[\s/,|]*(?:or)?[\s/,|]*$/.test(between)) return true;
    }
    return false;
  }

  function matchTitle(target, title, opts) {
    const t = low(title);
    if (target.cat === 'pc') return Object.assign({ exact: true }, matchPc(target, t));
    if ((target.cat === 'storage' || target.cat === 'ram') && sizeOptions(t)) return { ok: false, why: 'spec' };

    // Another kind of part in the title: a bundle if this part is in it too, otherwise a different item.
    const own = findAll(target.cat, t, { guard: true });
    for (const cat of ORDER) {
      if (cat === target.cat) continue;
      if (findAll(cat, t, { guard: true }).length) return { ok: false, why: own.length ? 'bundle' : 'spec' };
    }
    if (target.cat === 'gpu' || target.cat === 'cpu') {
      const keys = new Set(own.map(target.cat === 'gpu' ? gpuKey : cpuKey));
      if (keys.size > 1) return { ok: false, why: 'bundle' }; // e.g. "2x different GPUs"
    }
    if (/\b(laptop|notebook|mobile|max-?q)\b/.test(t) && target.cat === 'gpu') return { ok: false, why: 'spec' };

    const c = parseAs(target.cat, t, opts);
    if (!c) return { ok: false, why: 'spec' };
    const no = { ok: false, why: 'spec' };
    const yes = exact => ({ ok: true, exact: !!exact, spec: strip(c) });

    switch (target.cat) {
      case 'cpu':
        return cpuKey(c) === cpuKey(target) ? yes(true) : no;
      case 'gpu':
        if (gpuKey(c) !== gpuKey(target)) return no;
        if (target.vram && c.vram && c.vram !== target.vram) return no;
        return yes(!target.vram || c.vram);
      case 'ram':
        if (target.gen && c.gen !== target.gen) return no;
        if (target.totalGB && c.totalGB !== target.totalGB) return no;
        if (target.speed && c.speed !== target.speed) return no;
        if (target.modules && c.modules && c.modules !== target.modules) return no;
        if (target.modules === 1 && c.kit && !c.modules) return no;       // a "kit" is never one stick
        if (target.kit && !target.modules && c.modules === 1) return no;
        if ((target.form === 'sodimm') !== (c.form === 'sodimm')) return no;
        if (c.ecc && !target.ecc) return no;
        return yes(!target.modules || c.modules);
      case 'mobo':
        if (target.chipsets ? !target.chipsets.includes(c.chipset) : c.chipset !== target.chipset) return no;
        if (target.form && c.form && c.form !== target.form) return no;
        return yes(!target.form || c.form);
      case 'psu':
        if (target.minWatts ? (c.watts < target.minWatts || c.watts > psuCeiling(target.minWatts)) : c.watts !== target.watts) return no;
        if (target.rating && c.rating && c.rating !== target.rating) return no;
        return yes(!target.rating || c.rating);
      case 'storage':
        return matchStorage(target, c, opts);
      default:
        return yes(true);
    }
  }

  function matchStorage(target, c, opts) {
    const no = { ok: false, why: 'spec' };
    const yes = exact => ({ ok: true, exact: !!exact, spec: strip(c) });
    if (!!c.ext !== !!target.ext) return no;                 // external vs internal
    if (c.capGB !== target.capGB) return no;
    // An unstated type is accepted only when the search was limited to the matching eBay
    // category (SSDs or hard drives), and then it counts as a less certain match.
    let typeStated = true;
    if (!c.type) {
      const hint = opts && opts.typeHint;
      const fits = hint === 'hdd' ? target.type === 'hdd' : hint === 'ssd' ? target.type !== 'hdd' : false;
      if (!fits) return no;
      typeStated = false;
    } else if (c.type !== target.type) {
      return no;
    }
    if (target.type === 'nvme') {
      const small = s => s === '2230' || s === '2242';
      if (small(target.size) ? c.size !== target.size : (c.size && c.size !== '2280')) return no;
      if (target.gen && c.gen && c.gen !== target.gen) return no;
      return yes(typeStated && (!target.gen || c.gen));
    }
    if (target.type === 'ssd') {
      if ((c.iface === 'sas') !== (target.iface === 'sas')) return no;
      if (target.form && c.form && c.form !== target.form) return no;
      const formKnown = c.form || target.form === '2.5' || !target.form;
      return yes(typeStated && formKnown);
    }
    // hard drives
    if ((c.iface || 'sata') !== (target.iface || 'sata')) return no;
    if (target.form && c.form && c.form !== target.form) return no;
    if (target.rpm && c.rpm && c.rpm !== target.rpm) return no;
    const formKnown = c.form || target.form === '3.5' || !target.form;
    return yes(typeStated && formKnown && (!target.rpm || c.rpm));
  }

  // ---------------------------------------------------------------------------
  // Reading a listing: single part, or a bundle of parts?
  // ---------------------------------------------------------------------------

  // Item-specifics keys that describe a part, in the words parseAs() understands.
  const SPEC_KEYS = {
    ram: ['type', 'total capacity', 'bus speed', 'form factor', 'features', 'number of modules', 'capacity per module'],
    gpu: ['chipset/gpu model', 'gpu model', 'memory size'],
    cpu: ['processor model', 'model', 'processor'],
    mobo: ['chipset', 'chipset/graphics', 'form factor', 'model'],
    psu: ['maximum power', 'power', 'wattage', 'output power', 'certification', 'efficiency'],
    storage: ['storage capacity', 'capacity', 'interface', 'form factor', 'type', 'drive type', 'hard drive type',
      'drive type(s) supported', 'rotation speed', 'compatible slot', 'storage format', 'product line', 'model', 'mpn']
  };

  function categoryOfListing(listing) {
    const s = listing.specifics || {};
    if (EBAY_TO_CAT[listing.categoryId]) return EBAY_TO_CAT[listing.categoryId];
    if (s['operating system'] || s['ram size'] || s['ssd capacity']) return null; // a whole PC
    if (s['chipset/gpu model']) return 'gpu';
    if (s['bus speed'] || s['number of modules'] || s['capacity per module']) return 'ram';
    if (s['processor model'] && !s['socket type']) return 'cpu';
    if (s['certification'] || s['maximum power']) return 'psu';
    if (s['storage capacity'] || s['rotation speed'] || /ssd|solid state|hard/.test(low(s['type']) + low(s['drive type']))) return 'storage';
    const t = low(listing.title);
    const found = ORDER.filter(cat => findAll(cat, t).length);
    return found.length === 1 ? found[0] : null;
  }

  function specFromListing(listing, cat) {
    const s = listing.specifics || {};
    const extra = (SPEC_KEYS[cat] || []).map(k => s[k]).filter(Boolean).join(' | ');
    const opts = cat === 'storage' ? { typeHint: storageHintForCategory(listing.categoryId) } : undefined;
    let spec = parseAs(cat, `${listing.title} | ${extra}`, opts);
    if (!spec) spec = parseAs(cat, extra, opts);
    if (!spec) return null;
    if (cat === 'storage' && !spec.type) return null;
    if (cat === 'ram') {
      const n = parseInt(s['number of modules'], 10);
      const per = parseInt(s['capacity per module'], 10);
      if (!spec.modules && n && per) { spec.modules = n; spec.perGB = per; }
      if (/so-?dimm/i.test(s['form factor'] || '')) spec.form = 'sodimm';
      const total = parseInt(s['total capacity'], 10);
      if (total) spec.totalGB = total;
    }
    if (cat === 'gpu' && !spec.vram) {
      const m = low(s['memory size']).match(/(\d{1,2})\s?gb/);
      if (m) spec.vram = +m[1];
    }
    return strip(spec);
  }

  // Which DDR generation goes with a CPU, used when a bundle says just "16GB RAM".
  function ramGenForCpu(cpu) {
    if (!cpu) return null;
    if (/^i\d$/.test(cpu.family)) {
      const gen = cpu.num.length === 5 ? +cpu.num.slice(0, 2) : +cpu.num[0];
      if (cpu.num.length === 3) return 3;
      if (gen <= 5) return 3;
      if (gen <= 13) return 4; // 12th–13th gen boards come in both; DDR4 is the common used-market pairing
      return 5;
    }
    if (cpu.family === 'ultra') return 5;
    if (cpu.family === 'ryzen') return +cpu.num[0] >= 7 ? 5 : 4; // Ryzen 1000–5000 = DDR4, 7000+ = DDR5
    if (cpu.family === 'fx') return 3;
    return null;
  }

  // A later mention with more detail fills in gaps, e.g. title "16GB RAM" +
  // description "16GB DDR4 2400MHz (2x8GB)". Only if they clearly describe the same part.
  function enrich(existing, extra) {
    if (!existing || !extra) return;
    const same = {
      ram: () => !existing.totalGB || !extra.totalGB || existing.totalGB === extra.totalGB,
      gpu: () => gpuKey(existing) === gpuKey(extra),
      cpu: () => cpuKey(existing) === cpuKey(extra),
      mobo: () => existing.chipset === extra.chipset,
      psu: () => existing.watts === extra.watts
    }[existing.cat];
    if (!same || !same()) return;
    for (const k of Object.keys(extra)) {
      if (existing[k] == null || existing[k] === '' || existing[k] === false) existing[k] = extra[k];
    }
  }

  // All parts found in a bundle's text. Title first, then specifics/description fill gaps.
  function findParts(texts) {
    const parts = [];
    const have = new Set();
    for (const raw of texts) {
      const t = low(raw);
      if (!t) continue;
      for (const cat of ORDER) {
        const found = findAll(cat, t, { guard: true });
        for (const f of found) {
          const s = strip(f);
          const key = cat === 'storage' ? `storage|${s.type}|${s.capGB}` : cat;
          if (cat !== 'storage' && have.has(cat)) {
            enrich(parts.find(p => p.cat === cat), s);
            break;
          }
          if (have.has(key)) continue;
          if (cat === 'storage' && [...have].filter(k => k.startsWith('storage')).length >= 3) continue;
          have.add(key);
          parts.push(s);
          if (cat !== 'storage') break;
        }
      }
    }
    const cpu = parts.find(p => p.cat === 'cpu');
    parts.forEach(p => {
      if (p.cat === 'ram' && !p.gen) {
        const g = ramGenForCpu(cpu);
        if (g) { p.gen = g; p.assumed = `DDR${g} assumed from the CPU`; }
      }
    });
    return parts.sort((a, b) => ORDER.indexOf(a.cat) - ORDER.indexOf(b.cat));
  }

  const BUNDLE_WORDS = /\b(bundle|combo|job ?lot|joblot|upgrade kit|gaming pc|gaming computer|desktop pc|pc tower|tower pc|full (?:pc|system|build)|computer|workstation|mobo combo|cpu \+|\+ ?motherboard)\b/;

  function looksLikeBundle(listing) {
    const s = listing.specifics || {};
    if (s['operating system'] || s['ram size'] || s['ssd capacity']) return true;
    const t = low(listing.title);
    const cats = new Set(ORDER.filter(cat => findAll(cat, t, { guard: true }).length));
    if (cats.size >= 2) return true;
    if (EBAY_TO_CAT[listing.categoryId]) return false; // a component category with one component in the title
    return BUNDLE_WORDS.test(t) && cats.size >= 1;
  }

  // Extra text for whole-PC listings, built from their item specifics.
  function pcSpecificsText(specs) {
    const lines = [];
    if (specs['processor model']) lines.push(specs['processor model']);
    else if (specs['processor']) lines.push(specs['processor']);
    if (specs['gpu']) lines.push(specs['gpu'] + (specs['graphics memory'] ? ' ' + specs['graphics memory'] : ''));
    if (specs['ram size']) lines.push(`${specs['ram size']} ${specs['memory type'] || ''} RAM`);
    if (specs['ssd capacity']) lines.push(`${specs['ssd capacity']} SSD`);
    if (specs['hard drive capacity']) lines.push(`${specs['hard drive capacity']} HDD`);
    return lines.join(', ');
  }

  // One line of the bundle parts list, e.g. "DDR4 16GB 3200MHz RAM" or "Case = 25".
  function parseLine(line, hintCat) {
    const raw = String(line || '').trim();
    if (!raw) return null;
    const fixed = raw.match(/^(.*?)(?:=|:)\s*£?\s*(\d+(?:\.\d{1,2})?)\s*$/);
    if (fixed) return { cat: 'fixed', label: fixed[1].trim() || 'Fixed value', value: parseFloat(fixed[2]) };

    const t = low(raw);
    if (hintCat && CATS[hintCat] && hintCat !== 'generic') {
      const s = parseAs(hintCat, t);
      if (s) return strip(s);
    }
    for (const cat of ORDER) {
      const f = findAll(cat, t)[0];
      if (f) return strip(f);
    }
    for (const cat of ORDER) {
      const s = parseAs(cat, t);
      if (s) return strip(s);
    }
    return { cat: 'generic', query: raw };
  }

  // ---------------------------------------------------------------------------
  // Whole PCs
  // ---------------------------------------------------------------------------

  const OEM_RE = /\b(dell|optiplex|vostro|inspiron desktop|alienware|hp|hewlett|elitedesk|prodesk|pavilion|omen|compaq|lenovo|thinkcentre|thinkstation|ideacentre|legion tower|acer|veriton|aspire tc|fujitsu|esprimo|celsius|medion|packard bell)\b/;
  const OEM_BRANDS = /^(dell|alienware|hp|hewlett[\s-]packard|lenovo|acer|fujitsu|medion|packard bell|compaq)\b/;

  function isWholePc(listing) {
    const s = listing.specifics || {};
    if (listing.categoryId === '179' || s['operating system'] || s['ram size'] || s['ssd capacity']) return true;
    const t = low(listing.title);
    return PC_WORDS.test(t) && !/\b(laptop|notebook)\b/.test(t) && looksLikeBundle(listing) &&
      !/\b(motherboard|mobo) (?:and|&|\+) (?:cpu|processor)\b/.test(t);
  }

  function isOem(listing) {
    const s = listing.specifics || {};
    return OEM_BRANDS.test(low(s['brand'])) || OEM_RE.test(low(listing.title));
  }

  // Working, untested or faulty, from eBay's condition and the title.
  function detectCondition(listing) {
    const cond = low(listing.condition) + ' ' + low((listing.specifics || {})['condition']);
    const t = low(listing.title);
    if (/for parts|not working/.test(cond)) return { key: 'faulty', why: 'eBay condition: for parts or not working' };
    const faulty = t.match(/\b(faulty|not working|no power|no post|no display|won'?t (?:boot|turn on|power on)|doesn'?t (?:boot|turn on|power on)|spares|repair|for parts|broken|dead)\b/);
    if (faulty && !/\b(not|no) faulty\b|\bno faults?\b/.test(t)) return { key: 'faulty', why: `title says \u201c${faulty[1]}\u201d` };
    const untested = t.match(/\b(untested|not tested|unable to test|can'?t test|sold as seen|no hdd|no ssd|no hard drive|no ram|no os|no operating system)\b/);
    if (untested) return { key: 'untested', why: `title says \u201c${untested[1]}\u201d` };
    return { key: 'working', why: 'no problems mentioned' };
  }

  // Which part types an "extras" line stands in for, e.g. "Power supply = 15" -> ['psu'].
  function extraCategories(text) {
    const t = low(text);
    const cats = [];
    if (/\b(psu|power supply|power unit)\b/.test(t)) cats.push('psu');
    if (/\b(motherboard|mobo|mainboard|board)\b/.test(t)) cats.push('mobo');
    if (/\b(graphics|gpu|video card)\b/.test(t)) cats.push('gpu');
    if (/\b(ram|memory)\b/.test(t)) cats.push('ram');
    if (/\b(ssd|hdd|storage|hard drive)\b/.test(t)) cats.push('storage');
    if (/\b(cpu|processor)\b/.test(t) && !/\bcooler|fan|heatsink\b/.test(t)) cats.push('cpu');
    return cats;
  }

  // The "PC with this CPU + this graphics card" item used to find similar whole PCs.
  function pcItem(parts) {
    const cpu = parts.find(p => p.cat === 'cpu');
    if (!cpu) return null;
    const gpu = parts.find(p => p.cat === 'gpu') || null;
    return { cat: 'pc', cpu: strip(cpu), gpu: gpu ? strip(gpu) : null };
  }

  return {
    CATS,
    ORDER,
    low,
    parseAs,
    findAll,
    label,
    chips,
    queries,
    ebayCategory,
    matchTitle,
    categoryOfListing,
    specFromListing,
    findParts,
    looksLikeBundle,
    pcSpecificsText,
    parseLine,
    ramGenForCpu,
    isWholePc,
    isOem,
    detectCondition,
    extraCategories,
    pcItem,
    storageHintForCategory,
    layoutText,
    detailWord,
    sizeOptions
  };
})();

if (typeof module !== 'undefined' && module.exports) module.exports = DealSpecs;
