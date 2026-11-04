/*
 * specsheet.js — the spec sheet image for a build's sale listing.
 *
 *   model(build, purchases, benchmarks, games, opts) → what goes on it (plain data, testable)
 *   draw(canvas, model, photo, opts)                 → draws it, 1600px square for eBay
 *   findThumbnails(files)                            → { B69: File } from a picked folder
 *
 * Specifications come from the parts allocated to the build in Purchases & Parts (with
 * their Estimated Part Value), benchmarks from Build Benchmarking, games from
 * Build Game Testing.
 */
var DealSpecSheet = (function () {
  'use strict';

  const S = (typeof DealSpecs !== 'undefined') ? DealSpecs : require('./specs.js');
  const norm = s => String(s == null ? '' : s).trim().toLowerCase();
  const clean = s => String(s == null ? '' : s).replace(/\s+/g, ' ').trim();

  // ---------------------------------------------------------------------------
  // Specifications
  // ---------------------------------------------------------------------------

  const ROWS = ['CPU', 'GPU', 'RAM', 'Storage', 'Motherboard', 'Power Supply', 'Case', 'Cooling', 'Operating System', 'Extras'];
  const CATEGORY_ROW = {
    cpu: 'CPU', gpu: 'GPU', ram: 'RAM', storage: 'Storage', motherboard: 'Motherboard', psu: 'Power Supply',
    case: 'Case', 'cpu cooler': 'Cooling', cooling: 'Cooling', 'cables / accessories': 'Extras', other: 'Extras'
  };

  // "Corsair" + "Vengeance LPX 4GB DDR4 3000" → "Corsair Vengeance LPX 4GB DDR4 3000"
  function partName(p) {
    const brand = clean(p['Brand']);
    const desc = clean(p['Description']);
    if (!brand || /^(no brand|unbranded|unknown|n\/a|-)$/i.test(brand)) return desc;
    const b = brand.replace(/\s*\(.*?\)\s*/g, ' ').trim();
    if (b && norm(desc).startsWith(norm(b))) return desc;   // brand already in the description
    return `${b} ${desc}`;
  }

  // Identical parts counted: ["3x IONZ 120mm RGB Fan", "Antec 120mm Fan"]
  function grouped(names) {
    const order = [];
    const count = new Map();
    names.forEach(n => {
      const k = norm(n);
      if (!count.has(k)) { order.push(n); count.set(k, 0); }
      count.set(k, count.get(k) + 1);
    });
    return order.map(n => (count.get(norm(n)) > 1 ? `${count.get(norm(n))}x ${n}` : n));
  }

  // RAM: total first, then the sticks. "16GB DDR4 3000MHz (4x Corsair Vengeance LPX 4GB DDR4 3000)"
  function ramText(rows) {
    const specs = rows.map(p => S.parseAs('ram', partName(p)));
    const names = grouped(rows.map(partName));
    let total = 0;
    let gen = null;
    const speeds = [];
    let known = true;
    specs.forEach(s => {
      if (!s || !s.totalGB) { known = false; return; }
      total += s.totalGB;
      gen = gen || s.gen;
      if (s.speed) speeds.push(s.speed);
    });
    if (!known || !total) return names.join(', ');
    if (rows.length === 1 && specs[0].modules && specs[0].modules > 1) return names[0];
    const speed = speeds.length === specs.length && speeds.every(v => v === speeds[0]) ? ` ${speeds[0]}MHz` : '';
    const head = `${total}GB${gen ? ` DDR${gen}` : ''}${speed}`;
    if (rows.length === 1) return names[0].includes(`${total}GB`) ? names[0] : `${head} (${names[0]})`;
    return `${head} (${names.join(', ')})`;
  }

  const valueOf = p => {
    const v = p['Estimated Part Value at Purchase Date'];
    return typeof v === 'number' && v > 0 ? v : null;
  };

  function specRows(build, purchases, opts) {
    const o = opts || {};
    const id = String(build['Build ID'] || '').trim();
    const mine = purchases.filter(p => String(p['Allocation'] || '').trim() === id);
    const by = {};
    mine.forEach(p => {
      const row = CATEGORY_ROW[norm(p['Category'])];
      if (!row) return;                         // Shipping, Games… aren't part of the PC
      (by[row] = by[row] || []).push(p);
    });
    if (by.Cooling) by.Cooling.sort((a, b) => (norm(b['Category']) === 'cpu cooler') - (norm(a['Category']) === 'cpu cooler'));
    if (by.Storage && by.Storage.length > 1) {
      // Boot drive first: NVMe, then SATA SSD, then hard drives; bigger first within each.
      const rankOf = p => { const sp = S.parseAs('storage', partName(p)) || {}; return [{ nvme: 0, ssd: 1, hdd: 2 }[sp.type] ?? 1, -(sp.capGB || 0)]; };
      by.Storage.sort((a, b) => { const x = rankOf(a), y = rankOf(b); return x[0] - y[0] || x[1] - y[1]; });
    }

    // Anything not allocated as a row can still come from the Build Description.
    const fromDesc = {};
    S.findParts([build['Build Description'] || '']).forEach(s => {
      const row = { cpu: 'CPU', gpu: 'GPU', ram: 'RAM', storage: 'Storage', mobo: 'Motherboard', psu: 'Power Supply' }[s.cat];
      if (row && !fromDesc[row]) fromDesc[row] = S.label(s);
    });

    return ROWS.map(label => {
      if (label === 'Operating System') {
        const os = clean(o.os == null ? 'Windows 11 Home' : o.os);
        return { label, text: os ? `${os}, activated` : '', value: o.osValue > 0 ? +o.osValue : null, valued: o.osValue > 0, parts: 0 };
      }
      const rows = by[label] || [];
      let text;
      if (label === 'RAM' && rows.length) text = ramText(rows);
      else if (rows.length) text = grouped(rows.map(partName)).join(', ');
      else text = fromDesc[label] || '';
      const vals = rows.map(valueOf);
      const known = vals.filter(v => v != null);
      return {
        label, text,
        value: known.length ? Math.round(known.reduce((a, b) => a + b, 0) * 100) / 100 : null,
        valued: rows.length > 0 && known.length === rows.length,
        parts: rows.length
      };
    });
  }

  // ---------------------------------------------------------------------------
  // Benchmarks: "Cinebench R23 Single-Core" + "Multi-Core" → one line
  // ---------------------------------------------------------------------------

  const TOOL_RE = /^(cinebench\s*(?:r\d+|20\d\d)?|geekbench\s*\d*|3dmark(?:\s+(?:time spy|fire strike|port royal|steel nomad|speed way|night raid)(?:\s+extreme|\s+ultra)?)?|furmark(?:\s*2)?(?:\s*p?\d{3,4}p?)?|crystal\s*disk\s*mark|unigine\s+\w+|superposition|pcmark\s*\d*|userbenchmark|blender|cpu-?z|aida64|time spy|fire strike)\b\s*(.*)$/i;

  const num = v => {
    const n = typeof v === 'number' ? v : parseFloat(String(v).replace(/,/g, ''));
    if (!Number.isFinite(n)) return clean(v);
    return Number.isInteger(n) ? n.toLocaleString('en-GB') : String(Math.round(n * 100) / 100);
  };
  const prettyPart = s => clean(s).replace(/[-_]/g, ' ').replace(/\b(single|multi)\s?core\b/gi, (m, a) => `${a[0].toUpperCase()}${a.slice(1).toLowerCase()} Core`);
  const rank = s => (/single/i.test(s) ? 0 : /multi/i.test(s) ? 1 : 2);

  function benchmarkRows(id, benches) {
    const mine = benches.filter(b => norm(b['Build ID']) === norm(id) && clean(b['Benchmark']) && b['Score'] !== '' && b['Score'] != null);
    const tools = new Map();
    mine.forEach(b => {
      const name = clean(b['Benchmark']);
      const m = name.match(TOOL_RE);
      const tool = m ? clean(m[1]).replace(/\s+/g, ' ') : name;
      const part = m ? clean(m[2]) : '';
      const key = norm(tool);
      if (!tools.has(key)) tools.set(key, { tool: tool.replace(/^crystal\s*disk\s*mark$/i, 'CrystalDiskMark').replace(/^furmark/i, 'FurMark'), parts: [] });
      tools.get(key).parts.push({ part, score: b['Score'] });
    });
    return [...tools.values()].map(({ tool, parts }) => {
      const rw = new Map();
      const rest = [];
      parts.forEach(p => {
        const m = p.part.match(/^(.*?)\s*\b(read|write)\b\s*(.*)$/i);
        if (m) {
          const drive = clean(`${m[1]} ${m[3]}`) || 'Drive';
          if (!rw.has(drive)) rw.set(drive, {});
          rw.get(drive)[m[2].toLowerCase()] = p.score;
        } else rest.push(p);
      });
      rest.sort((a, b) => rank(a.part) - rank(b.part));
      const bits = rest.map(p => (p.part ? `${prettyPart(p.part)} ${num(p.score)}` : num(p.score)))
        .concat([...rw.entries()].map(([d, v]) => `${d} ${[v.read != null ? `${num(v.read)} MB/s read` : '', v.write != null ? `${num(v.write)} MB/s write` : ''].filter(Boolean).join(' · ')}`));
      return [tool, bits.join('  ·  ')];
    });
  }

  function gameRows(id, games) {
    return games.filter(g => norm(g['Build ID']) === norm(id) && clean(g['Game']))
      .map(g => ({ game: clean(g['Game']), settings: clean(g['Settings']), avg: g['Average FPS'] === '' ? '' : num(g['Average FPS']), low: g['1% Low FPS'] === '' || g['1% Low FPS'] == null ? '' : num(g['1% Low FPS']) }));
  }

  // "Intel Core i9-10900X", "GeForce GTX 1080 Ti 11GB"
  function headline(specs) {
    const cpuT = (specs.find(r => r.label === 'CPU') || {}).text || '';
    const gpuT = (specs.find(r => r.label === 'GPU') || {}).text || '';
    const c = S.parseAs('cpu', cpuT);
    const g = S.parseAs('gpu', gpuT);
    const cpu = c ? (/^i\d$/.test(c.family) ? 'Intel Core ' : c.family === 'ryzen' ? 'AMD ' : c.family === 'ultra' ? 'Intel ' : '') + S.label(c) : cpuT;
    const gpu = g ? (g.prefix === 'rx' ? 'Radeon ' : g.prefix === 'arc' ? 'Intel ' : 'GeForce ') + S.label(g) : gpuT;
    return { cpu, gpu, gaming: !!g };
  }

  // Short line under the headline: "16GB DDR4  ·  1TB SSD  ·  Windows 11"
  function subline(specs) {
    const ram = (specs.find(r => r.label === 'RAM') || {}).text || '';
    const st = (specs.find(r => r.label === 'Storage') || {}).text || '';
    const os = (specs.find(r => r.label === 'Operating System') || {}).text || '';
    const r = ram.match(/^(\d+GB(?: DDR\d)?)/);
    // Sizes as written ("512GB", not the pooled "500GB" used for pricing)
    const drives = st.split(/,\s*/).map(x => ({ d: S.parseAs('storage', x), m: x.match(/(\d+(?:\.\d+)?)\s?(TB|GB)\b/i) }))
      .filter(x => x.d && x.m)
      .map(({ d, m }) => `${m[1]}${m[2].toUpperCase()} ${d.type === 'nvme' ? 'NVMe SSD' : d.type === 'hdd' ? 'HDD' : 'SSD'}`);
    return [r && r[1], ...drives.slice(0, 2), os.replace(/,? ?activated$/i, '').replace(/ Home| Pro/, '')].filter(Boolean).join('  ·  ');
  }

  function model(build, purchases, benchmarks, games, opts) {
    const o = opts || {};
    const id = String(build['Build ID'] || '').trim();
    const specs = specRows(build, purchases, o);
    const counted = specs.filter(r => r.value != null);
    const totalValue = Math.round(counted.reduce((a, r) => a + r.value, 0));
    const price = typeof o.price === 'number' && o.price > 0 ? Math.round(o.price) : null;
    return {
      id, specs,
      headline: headline(specs),
      subline: subline(specs),
      benchmarks: benchmarkRows(id, benchmarks || []),
      games: gameRows(id, games || []),
      totalValue: totalValue || null,
      unvalued: specs.filter(r => r.parts > 0 && !r.valued).map(r => r.label),
      price,
      saving: price && totalValue > price ? totalValue - price : null
    };
  }

  // ---------------------------------------------------------------------------
  // Finding each build's Thumbnail.jpg in a picked folder
  // ---------------------------------------------------------------------------

  // Folder names "B69", "Build 69", "Build 69 - sold" → "B69"
  function buildIdFromFolder(name) {
    const m = String(name || '').trim().match(/^(?:b|build)\s*[-_ ]?\s*(\d{1,5})\b/i);
    return m ? 'B' + +m[1] : null;
  }

  /*
   * files: File objects from <input webkitdirectory> (each has webkitRelativePath).
   * Returns { B69: File, … } using the Thumbnail image in each build's folder.
   */
  function findThumbnails(files) {
    const out = {};
    [...files].forEach(f => {
      const path = String(f.webkitRelativePath || f.name).split('/');
      const name = path[path.length - 1];
      if (!/^thumbnail\.(jpe?g|png|webp|heic)$/i.test(name) || path.length < 2) return;
      const id = buildIdFromFolder(path[path.length - 2]);
      if (!id) return;
      if (out[id] && !/\.heic$/i.test(out[id].name)) return;   // prefer JPG/PNG over HEIC
      out[id] = f;
    });
    return out;
  }

  // ---------------------------------------------------------------------------
  // Drawing (browser only)
  // ---------------------------------------------------------------------------

  const T = {
    bg: '#0d1117', card: '#151b23', card2: '#1b222c', line: '#273140', ink: '#f2f5f8', soft: '#c9d1da',
    muted: '#8a96a6', accent: '#3ddc97', accentInk: '#06261a', strike: '#7b8796'
  };
  const FONT = 'Inter, "Segoe UI", "Helvetica Neue", Arial, sans-serif';
  const money = v => '£' + Math.round(v).toLocaleString('en-GB');

  function wrap(ctx, text, width) {
    const words = String(text || '').split(/\s+/).filter(Boolean);
    const lines = [];
    let cur = '';
    words.forEach(w => {
      const t = cur ? cur + ' ' + w : w;
      if (ctx.measureText(t).width <= width || !cur) cur = t;
      else { lines.push(cur); cur = w; }
    });
    if (cur) lines.push(cur);
    return lines.length ? lines : [''];
  }

  function roundRect(ctx, x, y, w, h, r) {
    ctx.beginPath();
    ctx.moveTo(x + r, y);
    ctx.arcTo(x + w, y, x + w, y + h, r);
    ctx.arcTo(x + w, y + h, x, y + h, r);
    ctx.arcTo(x, y + h, x, y, r);
    ctx.arcTo(x, y, x + w, y, r);
    ctx.closePath();
  }

  // Loads the bundled Inter font (once) so the image looks the same on every PC.
  let fontsReady = null;
  function loadFonts(base) {
    if (fontsReady) return fontsReady;
    if (typeof FontFace === 'undefined' || typeof document === 'undefined') return (fontsReady = Promise.resolve());
    // The font files can sit next to taskpane.html or in a fonts folder.
    const bases = base != null ? [base] : ['', 'fonts/'];
    fontsReady = Promise.all([400, 500, 600, 700, 800].map(w => {
      const src = bases.map(b => `url(${b}inter-${w}.woff2) format("woff2")`).join(', ');
      const f = new FontFace('Inter', src, { weight: String(w) });
      return f.load().then(ff => { document.fonts.add(ff); }).catch(() => null);
    })).then(() => undefined);
    return fontsReady;
  }

  /*
   * Draws the sheet. opts: { showValues: true, footer: [short points], label: 'Custom Gaming PC' }.
   * photo: a loaded Image (or null). Returns the height used.
   */
  function draw(canvas, m, photo, opts) {
    const o = Object.assign({ showValues: true, footer: null, label: null }, opts || {});
    const ctx = canvas.getContext('2d');
    const W = 1600;
    const P = 64;
    const hasPhoto = !!photo;
    const leftW = hasPhoto ? 600 : 0;
    const gap = hasPhoto ? 52 : 0;
    const colX = P + leftW + gap;
    const colW = W - colX - P;
    const osRow = m.specs.find(r => r.label === 'Operating System');
    const osName = osRow && osRow.text ? osRow.text.replace(/,? ?activated$/i, '').replace(/ (Home|Pro)$/, '') : '';
    const footer = (o.footer || ['Professionally built',
      m.benchmarks.length || m.games.length ? 'Stress-tested & benchmarked' : 'Stress-tested',
      osName ? `${osName} activated` : null]).filter(Boolean);
    const showValue = !!(o.showValues && (m.totalValue || m.price));
    const VALUE_H = 280;

    let s = 1;   // text scale, reduced if the right-hand column runs long

    const font = (size, weight) => { ctx.font = `${weight || 400} ${Math.round(size * s)}px ${FONT}`; };
    const text = (str, x, y, size, weight, color, align, paint, spacing) => {
      font(size, weight);
      if (!paint) return ctx.measureText(str).width;
      ctx.fillStyle = color || T.ink;
      ctx.textAlign = align || 'left';
      ctx.textBaseline = 'alphabetic';
      if (spacing && 'letterSpacing' in ctx) ctx.letterSpacing = `${spacing}px`;
      ctx.fillText(str, x, y);
      const w = ctx.measureText(str).width;
      if (spacing && 'letterSpacing' in ctx) ctx.letterSpacing = '0px';
      return w;
    };
    const eyebrow = (str, x, y, paint, color) => text(str.toUpperCase(), x, y, 19, 700, color || T.accent, 'left', paint, 2.5);

    // ---- header ----------------------------------------------------------------
    const header = (paint) => {
      let y = P + 22;
      eyebrow(o.label || (m.headline.gaming ? 'Custom Gaming PC' : 'Custom Desktop PC'), P, y, paint);
      y += 16;
      font(56, 800);
      const head = [m.headline.cpu, m.headline.gpu].filter(Boolean).join('  ·  ');
      wrap(ctx, head, W - P * 2).forEach(ln => { y += 66 * s; text(ln, P, y, 56, 800, T.ink, 'left', paint); });
      if (m.subline) { y += 46 * s; text(m.subline, P, y, 28, 500, T.soft, 'left', paint); }
      y += 34;
      if (paint) { ctx.fillStyle = T.line; ctx.fillRect(P, y, W - P * 2, 2); }
      return y + 38;
    };

    // ---- right column: specs and performance ------------------------------------
    const right = (y0, paint) => {
      let y = y0;
      eyebrow('Specifications', colX, y + 4, paint, T.muted);
      if (o.showValues && m.totalValue) text('PART VALUE', colX + colW, y + 4, 17, 700, T.muted, 'right', paint, 2);
      y += 22;
      const labelW = 196 * s;
      const valW = o.showValues ? 104 * s : 0;
      const textW = colW - labelW - valW - 16;
      m.specs.filter(r => clean(r.text)).forEach((r, i) => {
        font(24, 500);
        const lines = wrap(ctx, r.text, textW);
        const h = lines.length * 32 * s + 24 * s;
        if (paint && i % 2 === 0) { ctx.fillStyle = T.card; roundRect(ctx, colX - 16, y, colW + 32, h, 10); ctx.fill(); }
        text(r.label, colX, y + 36 * s, 19, 600, T.muted, 'left', paint);
        lines.forEach((ln, k) => text(ln, colX + labelW, y + 36 * s + k * 32 * s, 24, 500, T.ink, 'left', paint));
        if (o.showValues && r.value != null) text(money(r.value), colX + colW, y + 36 * s, 22, 600, T.soft, 'right', paint);
        y += h;
      });

      if (m.games.length) {
        y += 40 * s;
        eyebrow('Gaming performance', colX, y, paint, T.muted);
        y += 18 * s;
        const per = Math.min(3, m.games.length);
        const tw = (colW - (per - 1) * 16) / per;
        const th = 164 * s;
        m.games.forEach((g, i) => {
          const x = colX + (i % per) * (tw + 16);
          const ty = y + Math.floor(i / per) * (th + 16);
          if (paint) { ctx.fillStyle = T.card; roundRect(ctx, x, ty, tw, th, 14); ctx.fill(); }
          font(23, 700);
          let name = g.game;
          while (name.length > 4 && ctx.measureText(name).width > tw - 44) name = name.slice(0, -2).trim() + '…';
          text(name, x + 24, ty + 40 * s, 23, 700, T.ink, 'left', paint);
          text(g.settings, x + 24, ty + 70 * s, 19, 500, T.muted, 'left', paint);
          const big = text(String(g.avg), x + 24, ty + 134 * s, 54, 800, T.accent, 'left', paint);
          text('FPS', x + 34 + big, ty + 134 * s, 20, 700, T.soft, 'left', paint);
          if (g.low !== '') text(`1% low ${g.low}`, x + tw - 24, ty + 134 * s, 19, 500, T.muted, 'right', paint);
        });
        y += Math.ceil(m.games.length / per) * (th + 16) - 16;
      }

      if (m.benchmarks.length) {
        y += 40 * s;
        eyebrow('Benchmarks', colX, y, paint, T.muted);
        y += 12 * s;
        const tl = 230 * s;
        m.benchmarks.forEach(([tool, val], i) => {
          font(22, 500);
          const lines = wrap(ctx, val, colW - tl);
          const h = lines.length * 31 * s + 20 * s;
          if (paint && i) { ctx.fillStyle = T.line; ctx.fillRect(colX, y, colW, 1); }
          text(tool, colX, y + 35 * s, 19, 600, T.muted, 'left', paint);
          lines.forEach((ln, k) => text(ln, colX + tl, y + 35 * s + k * 31 * s, 22, 500, T.ink, 'left', paint));
          y += h;
        });
      }
      return y;
    };

    // ---- the value card ------------------------------------------------------------
    const valueCard = (x, y, w, paint) => {
      if (!showValue) return y;
      const h = VALUE_H;
      if (paint) {
        ctx.fillStyle = T.card2;
        roundRect(ctx, x, y, w, h, 18);
        ctx.fill();
        ctx.strokeStyle = T.line;
        ctx.lineWidth = 2;
        ctx.stroke();
      }
      const ix = x + 34;
      if (m.price && m.saving) {
        eyebrow('Parts bought separately', ix, y + 52, paint, T.muted);
        const vw = text(money(m.totalValue), ix, y + 104, 42, 700, T.strike, 'left', paint);
        if (paint) { ctx.fillStyle = T.strike; ctx.fillRect(ix - 3, y + 89, vw + 6, 3); }
        eyebrow('Our price', ix, y + 160, paint, T.muted);
        text(money(m.price), ix, y + 238, 76, 800, T.ink, 'left', paint);
        font(25, 800);
        const label = `You save ${money(m.saving)}`;
        const bw = ctx.measureText(label).width + 46;
        const bx = x + w - bw - 30;
        if (paint) { ctx.fillStyle = T.accent; roundRect(ctx, bx, y + 186, bw, 58, 29); ctx.fill(); }
        text(label, bx + bw / 2, y + 224, 25, 800, T.accentInk, 'center', paint);
      } else if (m.totalValue) {
        eyebrow('Value of the parts', ix, y + 52, paint, T.muted);
        text(money(m.totalValue), ix, y + 132, 76, 800, T.ink, 'left', paint);
        if (m.price) {
          eyebrow('Our price', ix, y + 192, paint, T.muted);
          text(money(m.price), ix, y + 246, 46, 800, T.accent, 'left', paint);
        } else {
          text('at current used prices', ix, y + 176, 22, 500, T.muted, 'left', paint);
        }
      } else {
        eyebrow('Our price', ix, y + 52, paint, T.muted);
        text(money(m.price), ix, y + 150, 84, 800, T.ink, 'left', paint);
      }
      return y + h;
    };

    // ---- footer points ------------------------------------------------------------
    const footerRow = (y, paint) => {
      let x = P;
      footer.forEach(f => {
        if (paint) {
          ctx.fillStyle = T.accent;
          ctx.beginPath(); ctx.arc(x + 14, y - 8, 14, 0, Math.PI * 2); ctx.fill();
          ctx.strokeStyle = T.accentInk; ctx.lineWidth = 3.5; ctx.lineCap = 'round'; ctx.lineJoin = 'round';
          ctx.beginPath(); ctx.moveTo(x + 7, y - 8); ctx.lineTo(x + 12, y - 3); ctx.lineTo(x + 21, y - 13); ctx.stroke();
        }
        const w = text(f, x + 38, y, 22, 600, T.soft, 'left', paint);
        x += 38 + w + 46;
      });
    };

    // ---- layout: shrink the text a little if the right column would run long ----------
    const FOOT = footer.length ? 104 : 30;
    const plan = () => {
      const top = header(false);
      const rightEnd = right(top, false);
      const photoH = hasPhoto ? Math.round(leftW * 4 / 3) : 0;
      const leftEnd = hasPhoto ? top + photoH + (showValue ? 28 + VALUE_H : 0) : 0;
      const bodyEnd = hasPhoto ? Math.max(rightEnd, leftEnd) : rightEnd + (showValue ? 44 + VALUE_H : 0);
      return { top, rightEnd, photoH, bodyEnd };
    };
    canvas.width = W;
    canvas.height = 10;
    let lay = plan();
    while (s > 0.76 && lay.bodyEnd + FOOT + P > W) { s -= 0.04; lay = plan(); }
    const H = hasPhoto ? Math.max(W, Math.ceil(lay.bodyEnd + FOOT + P)) : Math.ceil(lay.bodyEnd + FOOT + P);
    canvas.width = W;
    canvas.height = H;

    ctx.fillStyle = T.bg;
    ctx.fillRect(0, 0, W, H);
    const glow = ctx.createRadialGradient(W * 0.9, 0, 40, W * 0.9, 0, W);
    glow.addColorStop(0, 'rgba(61,220,151,0.10)');
    glow.addColorStop(1, 'rgba(61,220,151,0)');
    ctx.fillStyle = glow;
    ctx.fillRect(0, 0, W, H);

    const top = header(true);
    right(top, true);

    if (hasPhoto) {
      // Photo fills its frame (cropped evenly if its shape differs), with rounded corners.
      const ph = lay.photoH;
      const iw = photo.naturalWidth || photo.width;
      const ih = photo.naturalHeight || photo.height;
      const sc = Math.max(leftW / iw, ph / ih);
      ctx.save();
      roundRect(ctx, P, top, leftW, ph, 22);
      ctx.clip();
      ctx.drawImage(photo, P + (leftW - iw * sc) / 2, top + (ph - ih * sc) / 2, iw * sc, ih * sc);
      ctx.restore();
      ctx.strokeStyle = T.line;
      ctx.lineWidth = 2;
      roundRect(ctx, P, top, leftW, ph, 22);
      ctx.stroke();
      valueCard(P, top + ph + 28, leftW, true);
    } else {
      valueCard(P, lay.rightEnd + 44, W - P * 2, true);
    }
    if (footer.length) {
      ctx.fillStyle = T.line;
      ctx.fillRect(P, H - P - 62, W - P * 2, 2);
      footerRow(H - P, true);
    }
    return H;
  }

  return { model, draw, loadFonts, findThumbnails, buildIdFromFolder, partName, ramText, specRows, benchmarkRows, gameRows, headline, subline, ROWS };
})();

if (typeof module !== 'undefined' && module.exports) module.exports = DealSpecSheet;
