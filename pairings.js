/*
 * pairings.js — which parts go together.
 *
 *   - Platforms: sockets, chipsets and RAM type for each CPU generation.
 *   - Graphics card classes (by the PSU they need, which tracks how fast they are).
 *   - A gaming score for every common desktop CPU from Intel 2nd gen and Ryzen 1000 on,
 *     and the range of scores that suits each class of card: fast enough not to hold the
 *     card back, not so fast that the CPU is wasted on it.
 *   - Price-book keys: the level of detail a part is valued at (CPU model, board chipset,
 *     card model, RAM size and type, drive size and type).
 *
 * Shared by the Excel add-in and the Deal Alerts sheet. Uses DealSpecs.
 */
var DealPairings = (function () {
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

  // Does this board (by chipset) run this CPU?
  function boardFits(cpu, chipset) {
    const pf = cpuPlatform(cpu);
    return !!(pf && chipset && chipsetsFor(pf, cpu).includes(chipset));
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
  // Graphics card classes (by the PSU they need)
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
  // CPU gaming scores (i7-7700K = 100). Rough relative gaming performance with a
  // mid-to-high-end card; close enough to decide what pairs sensibly.
  // ---------------------------------------------------------------------------

  const CPU_SCORES = `
    i3-2100 45; i5-2400 60; i5-2500 62; i5-2500K 65; i7-2600 68; i7-2600K 70; i7-2700K 72;
    i3-3220 48; i5-3470 64; i5-3570 66; i5-3570K 68; i7-3770 72; i7-3770K 74;
    i3-4160 52; i5-4460 68; i5-4570 70; i5-4590 71; i5-4670K 73; i5-4690 73; i5-4690K 75; i7-4770 78; i7-4770K 80; i7-4790 81; i7-4790K 85;
    i3-6100 60; i5-6400 72; i5-6500 76; i5-6600 78; i5-6600K 80; i7-6700 88; i7-6700K 92;
    i3-7100 63; i5-7400 76; i5-7500 80; i5-7600 83; i5-7600K 86; i7-7700 95; i7-7700K 100;
    i3-8100 72; i5-8400 92; i5-8500 95; i5-8600 97; i5-8600K 100; i7-8700 108; i7-8700K 112;
    i3-9100F 74; i5-9400 93; i5-9400F 93; i5-9500 96; i5-9600 101; i5-9600K 102; i5-9600KF 102; i7-9700 112; i7-9700F 112; i7-9700K 116; i7-9700KF 116; i9-9900 118; i9-9900K 120; i9-9900KF 120;
    i3-10100 85; i3-10100F 85; i5-10400 100; i5-10400F 100; i5-10500 103; i5-10600 106; i5-10600K 110; i5-10600KF 110; i7-10700 118; i7-10700F 118; i7-10700K 122; i7-10700KF 122; i9-10850K 124; i9-10900 122; i9-10900K 126; i9-10900KF 126;
    i5-11400 112; i5-11400F 112; i5-11500 114; i5-11600K 120; i5-11600KF 120; i7-11700 124; i7-11700F 124; i7-11700K 128; i7-11700KF 128; i9-11900K 132;
    i3-12100 112; i3-12100F 112; i5-12400 128; i5-12400F 128; i5-12500 130; i5-12600K 142; i5-12600KF 142; i7-12700 145; i7-12700F 145; i7-12700K 150; i7-12700KF 150; i9-12900K 155;
    i5-13400 135; i5-13400F 135; i5-13600K 158; i5-13600KF 158; i7-13700K 165; i7-13700KF 165; i5-14400 138; i5-14400F 138; i5-14600K 162; i5-14600KF 162; i7-14700K 168; i7-14700KF 168;
    i7-3930K 66; i7-4930K 72; i7-5820K 80; i7-5930K 82; i7-6800K 84; i7-6850K 86; i7-6900K 88;
    i7-7800X 90; i7-7820X 96; i9-7900X 98; i9-9900X 100; i9-10900X 104; i9-10920X 105; i9-10940X 106;
    Ryzen 5 1400 62; Ryzen 5 1500X 64; Ryzen 5 1600 68; Ryzen 5 1600X 70; Ryzen 7 1700 70; Ryzen 7 1700X 72; Ryzen 7 1800X 74;
    Ryzen 5 2400G 65; Ryzen 5 2600 78; Ryzen 5 2600X 80; Ryzen 7 2700 80; Ryzen 7 2700X 84;
    Ryzen 5 3400G 70; Ryzen 5 3500 85; Ryzen 5 3500X 85; Ryzen 5 3600 95; Ryzen 5 3600X 97; Ryzen 7 3700X 100; Ryzen 7 3800X 102; Ryzen 9 3900X 104;
    Ryzen 5 4500 90; Ryzen 5 4600G 92; Ryzen 5 5500 102; Ryzen 5 5600G 104; Ryzen 5 5600 120; Ryzen 5 5600X 121; Ryzen 7 5700X 122; Ryzen 7 5800X 126; Ryzen 7 5800X3D 145; Ryzen 9 5900X 128; Ryzen 9 5950X 128;
    Ryzen 5 7600 145; Ryzen 5 7600X 147; Ryzen 7 7700 148; Ryzen 7 7700X 150; Ryzen 7 7800X3D 170`;

  const cpuKey = c => (c.family === 'ryzen' ? 'r' : 'i') + c.num + String(c.suffix || '').toLowerCase();
  const CPU_LIST = [];
  const SCORE = {};
  CPU_SCORES.split(';').map(s => s.trim()).filter(Boolean).forEach(entry => {
    const m = entry.match(/^(.*\S)\s+(\d+)$/);
    const spec = S.parseAs('cpu', m[1]);
    if (!spec) return;
    delete spec.index; delete spec.end;
    SCORE[cpuKey(spec)] = +m[2];
    CPU_LIST.push({ name: m[1], spec, score: +m[2], platform: cpuPlatform(spec) });
  });

  function cpuScore(cpu) {
    if (!cpu) return null;
    const k = cpuKey(cpu);
    if (SCORE[k] != null) return SCORE[k];
    const noF = k.replace(/f$/, '');           // 9400F scores as the 9400
    if (SCORE[noF] != null) return SCORE[noF];
    const withF = k + 'f';
    return SCORE[withF] != null ? SCORE[withF] : null;
  }

  // The CPU scores that suit a class of card.
  const CPU_RANGE = { 450: [50, 105], 550: [62, 120], 600: [70, 130], 650: [80, 140], 750: [100, 160], 850: [120, 175], 1000: [140, 180] };

  function cpuRangeFor(gpu) {
    return CPU_RANGE[gpuClass(gpu)] || CPU_RANGE[650];
  }

  /*
   * CPUs that pair well with this card. opts.ddr limits to platforms with that RAM type
   * (e.g. the build already has DDR4); opts.platform limits to one platform.
   */
  function acceptableCpus(gpu, opts) {
    const o = opts || {};
    const [lo, hi] = cpuRangeFor(gpu);
    return CPU_LIST.filter(c => c.score >= lo && c.score <= hi && c.platform && (o.allowI3 || c.spec.family !== 'i3') &&
      (!o.platform || c.platform === o.platform) &&
      (!o.ddr || PLATFORMS[c.platform].ddr === o.ddr || (c.platform === 'lga1700' && o.ddr === 5)));
  }

  function cpuFitsCard(cpu, gpu) {
    const s = cpuScore(cpu);
    if (s == null) return false;
    const [lo, hi] = cpuRangeFor(gpu);
    return s >= lo && s <= hi;
  }

  // Cards that suit a CPU: the classes whose range includes its score.
  function acceptableGpuClasses(cpu) {
    const s = cpuScore(cpu);
    if (s == null) return [550, 600];
    return Object.keys(CPU_RANGE).map(Number).filter(w => s >= CPU_RANGE[w][0] && s <= CPU_RANGE[w][1] && w >= 550);
  }

  const GPU_LIST = [];
  Object.keys(PSU_FOR_GPU).forEach(w => PSU_FOR_GPU[w].split(' ').forEach(k => {
    const m = k.match(/^(gtx|rtx|gt|rx|arc)([ab]?\d{3,4})(ti|super|xtx|xt|gre)?(super)?$/);
    if (!m) return;
    GPU_LIST.push({ spec: { cat: 'gpu', prefix: m[1], num: m[2], variant: [m[3], m[4]].filter(Boolean).join(' '), vram: null }, watts: +w });
  }));

  function acceptableGpus(cpu) {
    const classes = acceptableGpuClasses(cpu);
    return GPU_LIST.filter(g => classes.includes(g.watts)).map(g => g.spec);
  }

  function gpuFitsCpu(gpu, cpu) {
    return acceptableGpuClasses(cpu).includes(gpuClass(gpu));
  }

  // ---------------------------------------------------------------------------
  // Price book: the level of detail each part is valued at
  // ---------------------------------------------------------------------------

  /*
   * Returns { key, item } or null. item is what gets priced (any brand, as in the
   * extension); key is its label, shared by every listing that maps to it.
   */
  function bookKey(spec) {
    if (!spec || !spec.cat) return null;
    let item;
    switch (spec.cat) {
      case 'cpu': item = { cat: 'cpu', family: spec.family, tier: spec.tier, num: spec.num, suffix: spec.suffix || '' }; if (!item.tier) delete item.tier; break;
      case 'gpu': item = { cat: 'gpu', prefix: spec.prefix, num: spec.num, variant: spec.variant || '', vram: null }; break;
      case 'mobo': if (!spec.chipset) return null; item = { cat: 'mobo', chipset: spec.chipset, form: null }; break;
      case 'ram': if (!spec.gen || !spec.totalGB) return null; item = { cat: 'ram', gen: spec.gen, totalGB: spec.totalGB, modules: null, perGB: null, kit: false, speed: null, form: spec.form === 'sodimm' ? 'sodimm' : null, ecc: false }; break;
      case 'storage': if (!spec.type || !spec.capGB || spec.ext) return null; item = { cat: 'storage', type: spec.type, capGB: spec.capGB, form: null, size: null, gen: null, rpm: null, iface: null, ext: false }; break;
      case 'psu': if (!spec.watts) return null; item = { cat: 'psu', watts: spec.watts, rating: null }; break;
      default: return null;
    }
    return { key: S.label(item), item };
  }

  // Is this spec already at price-book detail (so its sold median can go in the book)?
  function isBookLevel(spec) {
    const b = bookKey(spec);
    if (!b) return false;
    if (spec.cat === 'cpu' || spec.cat === 'gpu') return true;
    const strip = o => JSON.stringify(Object.keys(o).filter(k => o[k] != null && o[k] !== false && o[k] !== '' && k !== 'index' && k !== 'end').sort().map(k => [k, o[k]]));
    return strip(b.item) === strip(spec);
  }

  return {
    PLATFORMS, DEFAULT_PLATFORM, CHIPSET_PLATFORM, CPU_LIST, GPU_LIST, CPU_RANGE,
    cpuPlatform, chipsetsFor, boardFits, looseChipset, gpuClass, isHedt, psuWattsFor, suggestGpu, cpuFitsGaming,
    cpuScore, cpuRangeFor, acceptableCpus, cpuFitsCard, acceptableGpuClasses, acceptableGpus, gpuFitsCpu,
    bookKey, isBookLevel
  };
})();

if (typeof module !== 'undefined' && module.exports) module.exports = DealPairings;
