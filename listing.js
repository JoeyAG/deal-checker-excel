/*
 * listing.js — eBay listing drafts for finished builds.
 *
 * Reads a build's parts (Builds row + the Purchases & Parts rows allocated to it) and writes:
 * a title (80 characters max), item specifics for eBay's Desktops & All-In-Ones category,
 * a plain description, a suggested price and a Seller Hub drafts CSV.
 *
 * No Excel code; uses DealSpecs and DealPlanner.
 */
var DealListing = (function () {
  'use strict';

  const S = (typeof DealSpecs !== 'undefined') ? DealSpecs : require('./specs.js');
  const PL = (typeof DealPlanner !== 'undefined') ? DealPlanner : require('./planner.js');
  const norm = s => String(s == null ? '' : s).trim().toLowerCase();
  const EBAY_CATEGORY = '179';  // Computers/Tablets & Networking > Desktops & All-In-Ones
  const USED = '3000';

  function cpuName(c) {
    const l = S.label(c);
    if (/^i\d$/.test(c.family)) return 'Intel Core ' + l;
    if (c.family === 'ultra') return 'Intel ' + l;
    if (c.family === 'ryzen') return 'AMD ' + l;
    if (/^xeon/.test(c.family)) return 'Intel ' + l;
    return l;
  }
  function gpuName(g) {
    const l = S.label(g);
    if (g.prefix === 'rx') return 'AMD Radeon ' + l;
    if (g.prefix === 'arc') return 'Intel ' + l;
    return 'NVIDIA GeForce ' + l;
  }
  const capText = gb => (gb >= 1000 ? `${+(gb / 1000).toFixed(1)}TB` : `${gb}GB`);

  // Everything we know about the build, part by part.
  function partsOf(build, purchases) {
    const id = String(build['Build ID'] || '').trim();
    const mine = purchases.filter(p => String(p['Allocation'] || '').trim() === id);
    const desc = S.findParts([build['Build Description'] || '']);
    const out = { id, cpu: null, gpu: null, ram: null, storage: [], mobo: null, psu: null, case: null, cooler: null, wifi: false, rows: mine };

    const ramSticks = [];
    mine.forEach(p => {
      const ps = PL.partSpec(p);
      const text = PL.partText(p);
      if (/wi-?fi|\bwireless\b/i.test(text)) out.wifi = true;
      if (!ps) return;
      const s = ps.spec;
      switch (ps.kind) {
        case 'cpu': if (s && !out.cpu) out.cpu = { spec: s, text }; break;
        case 'gpu': if (s && !out.gpu) out.gpu = { spec: s, text }; break;
        case 'ram': if (s) ramSticks.push(s); break;
        case 'storage': if (s && s.type) out.storage.push({ spec: s, text }); break;
        case 'mobo': if (!out.mobo) out.mobo = { spec: s, text }; break;
        case 'psu': if (!out.psu) out.psu = { spec: s, text }; break;
        case 'case': if (!out.case) out.case = { text }; break;
        case 'cooler': if (!out.cooler && /\b(cpu cooler|heatsink|aio|liquid|cooler)\b/i.test(text.replace(/cooler\s?master/i, ''))) out.cooler = { text }; break;
        default:
      }
    });
    if (ramSticks.length) {
      const total = ramSticks.reduce((a, s) => a + (s.totalGB || 0), 0);
      const sticks = ramSticks.reduce((a, s) => a + (s.modules || 1), 0);
      const speeds = ramSticks.map(s => s.speed).filter(Boolean);
      const per = ramSticks.every(s => s.totalGB === ramSticks[0].totalGB && !s.modules) ? ramSticks[0].totalGB : null;
      out.ram = { gen: ramSticks[0].gen, totalGB: total, sticks, perGB: per, speed: speeds.length ? Math.min(...speeds) : null };
    }
    // The build description fills anything not allocated as a row.
    desc.forEach(s => {
      if (s.cat === 'cpu' && !out.cpu) out.cpu = { spec: s, text: S.label(s) };
      if (s.cat === 'gpu' && !out.gpu) out.gpu = { spec: s, text: S.label(s) };
      if (s.cat === 'ram' && !out.ram) out.ram = { gen: s.gen, totalGB: s.totalGB, sticks: s.modules || null, perGB: s.perGB || null, speed: s.speed || null };
      if (s.cat === 'psu' && !out.psu) out.psu = { spec: s, text: S.label(s) };
      if (s.cat === 'storage' && !out.storage.length) out.storage.push({ spec: s, text: S.label(s) });
      if (s.cat === 'mobo' && !out.mobo) out.mobo = { spec: s, text: S.label(s) };
    });
    out.storage.sort((a, b) => (a.spec.type === 'hdd') - (b.spec.type === 'hdd') || b.spec.capGB - a.spec.capGB);
    return out;
  }

  function ramText(r, long) {
    if (!r) return '';
    const gen = r.gen ? `DDR${r.gen}` : '';
    if (!long) return `${r.totalGB}GB ${gen}`.trim();
    const layout = r.sticks && r.perGB ? ` (${r.sticks}x${r.perGB}GB)` : '';
    return `${r.totalGB}GB ${gen}${r.speed ? ' ' + r.speed + 'MHz' : ''}${layout}`.replace(/\s+/g, ' ').trim();
  }
  function storageShort(st) {
    const s = st.spec;
    return `${capText(s.capGB)} ${s.type === 'hdd' ? 'HDD' : s.type === 'nvme' ? 'NVMe SSD' : 'SSD'}`;
  }

  // ---------------------------------------------------------------------------
  // Title (eBay allows 80 characters)
  // ---------------------------------------------------------------------------

  function title(parts, opts) {
    const o = opts || {};
    const os = o.os || 'Windows 11';
    const gpu = parts.gpu ? S.label(Object.assign({}, parts.gpu.spec)) : null;
    const cpuShort = parts.cpu ? S.label(parts.cpu.spec) : null;
    const brand = parts.cpu ? (/^i\d$|ultra|xeon/.test(parts.cpu.spec.family) ? 'Intel' : parts.cpu.spec.family === 'ryzen' ? 'AMD' : '') : '';
    const ssd = parts.storage[0] ? storageShort(parts.storage[0]).replace(' NVMe SSD', ' NVMe') : null;
    const lead = gpu ? 'Gaming PC' : 'Desktop PC';
    const variants = [
      [lead, brand, cpuShort, gpu, parts.ram && ramText(parts.ram) + ' RAM', ssd, parts.wifi && 'WiFi', os],
      [lead, brand, cpuShort, gpu, parts.ram && ramText(parts.ram), ssd, parts.wifi && 'WiFi', os],
      [lead, cpuShort, gpu, parts.ram && ramText(parts.ram), ssd, parts.wifi && 'WiFi', os],
      [lead, cpuShort, gpu, parts.ram && ramText(parts.ram), ssd, parts.wifi && 'WiFi', os.replace('Windows', 'Win')],
      [lead, cpuShort, gpu, parts.ram && `${parts.ram.totalGB}GB`, ssd, os.replace('Windows', 'Win')],
      [lead, cpuShort, gpu && gpu.replace(/ \d+GB$/, ''), parts.ram && `${parts.ram.totalGB}GB`, ssd && ssd.replace(/ (SSD|NVMe|HDD)$/, ''), os.replace('Windows', 'Win')]
    ].map(v => v.filter(Boolean).join(' ').replace(/\s+/g, ' ').trim());
    return variants.find(t => t.length <= 80) || variants[variants.length - 1].slice(0, 80);
  }

  // ---------------------------------------------------------------------------
  // Item specifics for Desktops & All-In-Ones
  // ---------------------------------------------------------------------------

  function specifics(parts, opts) {
    const o = opts || {};
    const os = o.os || 'Windows 11 Home';
    const list = [['Brand', 'Unbranded'], ['Model', 'Custom Build'], ['MPN', 'Does not apply']];
    if (parts.cpu) list.push(['Processor', cpuName(parts.cpu.spec)]);
    if (parts.gpu) {
      list.push(['GPU', gpuName(Object.assign({}, parts.gpu.spec, { vram: null }))]);
      list.push(['Graphics Processing Type', 'Dedicated Graphics']);
    }
    if (parts.ram) list.push(['RAM Size', `${parts.ram.totalGB} GB`]);
    const ssd = parts.storage.filter(s => s.spec.type !== 'hdd').reduce((a, s) => a + s.spec.capGB, 0);
    const hdd = parts.storage.filter(s => s.spec.type === 'hdd').reduce((a, s) => a + s.spec.capGB, 0);
    if (ssd) list.push(['SSD Capacity', capText(ssd).replace(/(GB|TB)$/, ' $1')]);
    if (hdd) list.push(['Hard Drive Capacity', capText(hdd).replace(/(GB|TB)$/, ' $1')]);
    if (ssd || hdd) list.push(['Storage Type', ssd && hdd ? 'SSD + HDD' : ssd ? 'SSD (Solid State Drive)' : 'HDD (Hard Disk Drive)']);
    list.push(['Operating System', os]);
    list.push(['Operating System Edition', /pro/i.test(os) ? 'Pro' : 'Home']);
    list.push(['Form Factor', 'Tower']);
    if (parts.gpu) list.push(['Most Suitable For', 'Gaming']);
    if (parts.wifi) list.push(['Features', 'Wi-Fi']);
    return list;
  }

  // ---------------------------------------------------------------------------
  // Description
  // ---------------------------------------------------------------------------

  function specLines(parts, opts) {
    const o = opts || {};
    const lines = [];
    if (parts.cpu) lines.push(['Processor', cpuName(parts.cpu.spec)]);
    if (parts.gpu) lines.push(['Graphics', gpuName(parts.gpu.spec)]);
    if (parts.ram) lines.push(['Memory', ramText(parts.ram, true)]);
    parts.storage.forEach((s, i) => lines.push([i ? 'Extra storage' : 'Storage', /\d\s?(gb|tb)\b/i.test(s.text || '') && !/^\d/.test(s.text) ? s.text : storageShort(s)]));
    if (parts.mobo) lines.push(['Motherboard', parts.mobo.text]);
    if (parts.psu) lines.push(['Power supply', parts.psu.spec && !/\d{3,4}\s?w/i.test(parts.psu.text) ? `${parts.psu.text} ${parts.psu.spec.watts}W` : parts.psu.text]);
    if (parts.cooler) lines.push(['CPU cooler', parts.cooler.text]);
    if (parts.case) lines.push(['Case', parts.case.text]);
    if (parts.wifi) lines.push(['Wireless', 'Wi-Fi']);
    lines.push(['Operating system', `${o.os || 'Windows 11 Home'}, installed and activated`]);
    return lines;
  }

  function description(parts, opts) {
    const o = opts || {};
    const lines = specLines(parts, o);
    const kind = parts.gpu ? 'gaming PC' : 'desktop PC';
    return [
      `Custom-built ${kind}, assembled, tested and ready to use.`,
      '',
      'Specification',
      ...lines.map(([k, v]) => `• ${k}: ${v}`),
      '',
      `Every part has been cleaned and tested, and the finished PC has been stress-tested${o.benchmarked ? ' and benchmarked' : ''} before listing. It comes with a fresh install of ${(o.os || 'Windows 11 Home').replace(/ (Home|Pro)$/, '')} and up-to-date drivers, so it's ready to go out of the box.`,
      '',
      "What's included: the PC and a UK power cable. Monitor, keyboard and mouse are not included.",
      '',
      'Any questions, just send a message.'
    ].join('\n');
  }

  function descriptionHtml(text) {
    const esc = s => s.replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
    const out = [];
    let list = null;
    String(text).split('\n').forEach(line => {
      if (/^•\s?/.test(line)) { list = list || []; list.push(`<li>${esc(line.replace(/^•\s?/, ''))}</li>`); return; }
      if (list) { out.push(`<ul>${list.join('')}</ul>`); list = null; }
      if (line.trim()) out.push(`<p>${esc(line)}</p>`);
    });
    if (list) out.push(`<ul>${list.join('')}</ul>`);
    return out.join('');
  }

  // ---------------------------------------------------------------------------
  // Price
  // ---------------------------------------------------------------------------

  // Cost plus your margin, allowing for selling fees, rounded to a price ending in 9.
  function suggestPrice(cost, marginPct, feePct) {
    if (typeof cost !== 'number' || !(cost > 0)) return null;
    const m = (typeof marginPct === 'number' ? marginPct : 25) / 100;
    const f = Math.min(0.5, Math.max(0, (feePct || 0) / 100));
    const raw = cost * (1 + m) / (1 - f);
    return Math.ceil((raw + 1) / 10) * 10 - 1;
  }

  const round9 = v => Math.ceil((v + 1) / 10) * 10 - 1;

  // What the allocated parts are worth on the used market (Estimated Part Value column).
  function partsValue(rows) {
    let total = 0;
    let missing = 0;
    rows.forEach(p => {
      const cat = norm(p['Category']);
      if (!['cpu', 'gpu', 'ram', 'motherboard', 'psu', 'storage', 'case', 'cooling'].includes(cat)) return;
      const v = p['Estimated Part Value at Purchase Date'];
      if (typeof v === 'number' && v > 0) total += v; else missing++;
    });
    return { total: Math.round(total * 100) / 100, missing };
  }

  // Average share of the sale price that went on fees, from Sales rows already recorded.
  function feeShare(sales) {
    const got = sales.filter(s => norm(s['Sold?']) === 'y' && typeof s['Sold Price'] === 'number' && s['Sold Price'] > 0);
    if (!got.length) return null;
    const share = got.map(s => ((+s['eBay Fees'] || 0) + (+s['Promotion Fee'] || 0)) / s['Sold Price']);
    return Math.round(share.reduce((a, b) => a + b, 0) / share.length * 1000) / 10;
  }

  // Sold whole PCs with the same CPU and graphics card, for a sense check on price.
  function similarSoldUrl(parts) {
    if (!parts.cpu) return null;
    const item = S.pcItem([parts.cpu.spec].concat(parts.gpu ? [parts.gpu.spec] : []));
    const p = new URLSearchParams({ _nkw: S.queries(item).primary, _sacat: EBAY_CATEGORY, LH_Sold: '1', LH_Complete: '1', LH_PrefLoc: '1', _sop: '13' });
    return 'https://www.ebay.co.uk/sch/i.html?' + p.toString();
  }

  // ---------------------------------------------------------------------------
  // The whole draft
  // ---------------------------------------------------------------------------

  function draft(build, purchases, opts) {
    const o = Object.assign({ os: 'Windows 11 Home', marginPct: 25, feePct: 0 }, opts || {});
    const parts = partsOf(build, purchases);
    const cost = typeof build['Build Cost'] === 'number' ? build['Build Cost'] : null;
    const value = partsValue(parts.rows);
    const fromCost = suggestPrice(cost, o.marginPct, o.feePct);
    const fromValue = value.total > 0 ? round9(value.total / (1 - Math.min(0.5, (o.feePct || 0) / 100))) : null;
    const missing = [];
    if (!parts.cpu) missing.push('CPU');
    if (!parts.ram) missing.push('RAM');
    if (!parts.storage.length) missing.push('storage');
    const d = {
      id: parts.id,
      title: title(parts, { os: o.os.replace(/ (Home|Pro)$/, '') }),
      specifics: specifics(parts, o),
      description: description(parts, { os: o.os, benchmarked: norm(build['Benchmarked?']) === 'y' }),
      cost,
      price: Math.max(fromCost || 0, fromValue || 0) || null,
      priceFromCost: fromCost,
      priceFromValue: fromValue,
      partsValue: value,
      categoryId: EBAY_CATEGORY,
      conditionId: USED,
      similarUrl: similarSoldUrl(parts),
      missing,
      parts
    };
    return d;
  }

  // ---------------------------------------------------------------------------
  // Seller Hub "Create drafts" upload file
  // ---------------------------------------------------------------------------

  const csvCell = v => {
    const s = String(v == null ? '' : v);
    return /[",\n\r]/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s;
  };

  function draftsCsv(drafts) {
    const specNames = [];
    drafts.forEach(d => d.specifics.forEach(([k]) => { if (!specNames.includes(k)) specNames.push(k); }));
    const head = ['Action(SiteID=UK|Country=GB|Currency=GBP|Version=1193|CC=UTF-8)', 'Custom label (SKU)', 'Category ID', 'Title',
      'UPC', 'Price', 'Quantity', 'Item photo URL', 'Condition ID', 'Description', 'Format'].concat(specNames.map(n => 'C:' + n));
    const rows = drafts.map(d => {
      const sp = Object.fromEntries(d.specifics);
      return ['Draft', d.id, d.categoryId, d.title, '', d.price == null ? '' : d.price, 1, '', d.conditionId,
        descriptionHtml(d.description), 'FixedPrice'].concat(specNames.map(n => sp[n] || ''));
    });
    return '﻿' + [head].concat(rows).map(r => r.map(csvCell).join(',')).join('\r\n') + '\r\n';
  }

  return { partsOf, partsValue, title, specifics, description, descriptionHtml, suggestPrice, feeShare, similarSoldUrl, draft, draftsCsv, EBAY_CATEGORY };
})();

if (typeof module !== 'undefined' && module.exports) module.exports = DealListing;
