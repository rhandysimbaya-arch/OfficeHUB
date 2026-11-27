// Reads layout/formatting from an .xlsx (column widths, row heights, merged cells, fonts, fills, borders, alignment, wrap)
// SheetJS community edition ignores all of this, so we parse the XML directly with JSZip.
const XS = (function () {
  const IDX = ['000000','FFFFFF','FF0000','00FF00','0000FF','FFFF00','FF00FF','00FFFF','000000','FFFFFF','FF0000','00FF00','0000FF','FFFF00','FF00FF','00FFFF','800000','008000','000080','808000','800080','008080','C0C0C0','808080','9999FF','993366','FFFFCC','CCFFFF','660066','FF8080','0066CC','CCCCFF'];
  const THEME_DEF = ['FFFFFF','000000','E7E6E6','44546A','4472C4','ED7D31','A5A5A5','FFC000','5B9BD5','70AD47','0563C1','954F72'];
  const kids = (el, n) => el ? Array.from(el.children).filter(x => !n || x.localName === n) : [];
  const first = (el, n) => kids(el, n)[0] || null;
  const parse = s => new DOMParser().parseFromString(s, 'application/xml');
  const find = (doc, n) => Array.from(doc.getElementsByTagName('*')).filter(x => x.localName === n);

  function tint(hex, t) {
    if (!t) return hex;
    let r = parseInt(hex.substr(0, 2), 16) / 255, g = parseInt(hex.substr(2, 2), 16) / 255, b = parseInt(hex.substr(4, 2), 16) / 255;
    const mx = Math.max(r, g, b), mn = Math.min(r, g, b); let h = 0, s = 0, l = (mx + mn) / 2;
    if (mx !== mn) { const d = mx - mn; s = l > .5 ? d / (2 - mx - mn) : d / (mx + mn); h = mx === r ? (g - b) / d + (g < b ? 6 : 0) : mx === g ? (b - r) / d + 2 : (r - g) / d + 4; h /= 6; }
    l = t < 0 ? l * (1 + t) : l * (1 - t) + t;
    const f = (p, q, x) => { if (x < 0) x += 1; if (x > 1) x -= 1; return x < 1 / 6 ? p + (q - p) * 6 * x : x < .5 ? q : x < 2 / 3 ? p + (q - p) * (2 / 3 - x) * 6 : p; };
    if (s === 0) r = g = b = l; else { const q = l < .5 ? l * (1 + s) : l + s - l * s, p = 2 * l - q; r = f(p, q, h + 1 / 3); g = f(p, q, h); b = f(p, q, h - 1 / 3); }
    return [r, g, b].map(v => Math.round(Math.max(0, Math.min(1, v)) * 255).toString(16).padStart(2, '0')).join('').toUpperCase();
  }

  async function read(buf) {
    const z = await JSZip.loadAsync(buf);
    const txt = async p => { const f = z.file(p); return f ? await f.async('string') : null; };
    const wbx = parse(await txt('xl/workbook.xml')), rel = parse(await txt('xl/_rels/workbook.xml.rels'));
    const relMap = {}; find(rel, 'Relationship').forEach(r => relMap[r.getAttribute('Id')] = r.getAttribute('Target'));
    // theme colours
    let theme = THEME_DEF.slice();
    const tx = await txt('xl/theme/theme1.xml');
    if (tx) {
      const cs = find(parse(tx), 'clrScheme')[0];
      if (cs) { const o = {}; kids(cs).forEach(e => { const c = e.firstElementChild; o[e.localName] = c ? (c.getAttribute('val') || c.getAttribute('lastClr')) : null; });
        const ord = ['lt1', 'dk1', 'lt2', 'dk2', 'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6', 'hlink', 'folHlink'];
        theme = ord.map((k, i) => (o[k] || THEME_DEF[i]).toUpperCase()); }
    }
    const colorOf = el => {
      if (!el) return null;
      if (el.getAttribute('rgb')) return '#' + el.getAttribute('rgb').slice(-6);
      if (el.getAttribute('theme') !== null) { const i = +el.getAttribute('theme'); return '#' + tint(theme[i] || '000000', +(el.getAttribute('tint') || 0)); }
      if (el.getAttribute('indexed') !== null) { const i = +el.getAttribute('indexed'); return IDX[i] ? '#' + IDX[i] : (i === 64 ? '#000000' : null); }
      if (el.getAttribute('auto')) return '#000000';
      return null;
    };
    // styles
    const sx = parse(await txt('xl/styles.xml') || '<styleSheet/>');
    const fonts = kids(first(sx.documentElement, 'fonts'), 'font').map(f => ({
      b: !!first(f, 'b') && first(f, 'b').getAttribute('val') !== '0', i: !!first(f, 'i') && first(f, 'i').getAttribute('val') !== '0',
      u: !!first(f, 'u') && first(f, 'u').getAttribute('val') !== 'none', sz: first(f, 'sz') ? +first(f, 'sz').getAttribute('val') : 11,
      name: first(f, 'name') ? first(f, 'name').getAttribute('val') : '', color: colorOf(first(f, 'color'))
    }));
    const fills = kids(first(sx.documentElement, 'fills'), 'fill').map(f => {
      const p = first(f, 'patternFill'); if (!p) return null; const t = p.getAttribute('patternType'); if (!t || t === 'none') return null;
      return colorOf(first(p, 'fgColor')) || (t === 'solid' ? null : null);
    });
    const BS = { thin: '1px solid', hair: '1px dotted', dotted: '1px dotted', dashed: '1px dashed', mediumDashed: '2px dashed', dashDot: '1px dashed', dashDotDot: '1px dashed', medium: '2px solid', mediumDashDot: '2px dashed', thick: '3px solid', double: '3px double', slantDashDot: '2px dashed', mediumDashDotDot: '2px dashed' };
    const borders = kids(first(sx.documentElement, 'borders'), 'border').map(b => {
      const g = n => { const e = first(b, n); if (!e) return null; const st = e.getAttribute('style'); if (!st) return null; return (BS[st] || '1px solid') + ' ' + (colorOf(first(e, 'color')) || '#000000'); };
      return { l: g('left'), r: g('right'), t: g('top'), b: g('bottom') };
    });
    const xfs = kids(first(sx.documentElement, 'cellXfs'), 'xf').map(x => {
      const al = first(x, 'alignment');
      return { font: +(x.getAttribute('fontId') || 0), fill: +(x.getAttribute('fillId') || 0), bd: +(x.getAttribute('borderId') || 0),
        h: al && al.getAttribute('horizontal'), v: al && al.getAttribute('vertical'), wrap: !!al && (al.getAttribute('wrapText') === '1' || al.getAttribute('wrapText') === 'true'),
        indent: al ? +(al.getAttribute('indent') || 0) : 0 };
    });
    const def = fonts[0] || { sz: 11, name: '' };
    const styleFor = i => {
      const x = xfs[i]; if (!x) return null; const f = fonts[x.font] || def, o = {};
      if (f.b) o.b = 1; if (f.i) o.i = 1; if (f.u) o.u = 1;
      if (f.color && f.color !== '#000000') o.color = f.color;
      if (f.sz !== def.sz) o.sz = Math.round(f.sz * 4 / 3 * 10) / 10;
      if (f.name && f.name !== def.name) o.font = f.name;
      const bg = fills[x.fill]; if (bg && bg.toUpperCase() !== '#FFFFFF') o.bg = bg;
      if (x.h === 'left' || x.h === 'justify' || x.h === 'fill') o.align = 'left'; else if (x.h === 'center' || x.h === 'centerContinuous') o.align = 'center'; else if (x.h === 'right') o.align = 'right';
      if (x.v === 'top') o.valign = 'top'; else if (x.v === 'center') o.valign = 'middle';
      if (x.wrap) o.wrap = 1;
      if (x.indent) o.indent = x.indent;
      const b = borders[x.bd]; if (b) { if (b.t) o.bt = b.t; if (b.b) o.bb = b.b; if (b.l) o.bl = b.l; if (b.r) o.br = b.r; }
      return o;
    };
    const out = { sheets: {} };
    for (const s of find(wbx, 'sheet')) {
      const name = s.getAttribute('name'), rid = s.getAttribute('r:id') || s.getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships', 'id');
      let tgt = relMap[rid]; if (!tgt) continue; tgt = tgt.replace(/^\//, ''); if (!tgt.startsWith('xl/')) tgt = 'xl/' + tgt;
      const xml = await txt(tgt); if (!xml) continue; const d = parse(xml);
      const info = { colW: {}, rowH: {}, merges: [], style: {}, defSz: Math.round(def.sz * 4 / 3 * 10) / 10, defFont: def.name || '', grid: true };
      const sv = find(d, 'sheetView')[0]; if (sv && sv.getAttribute('showGridLines') === '0') info.grid = false;
      const fp = find(d, 'sheetFormatPr')[0];
      const dcw = fp && fp.getAttribute('defaultColWidth') ? Math.round(+fp.getAttribute('defaultColWidth') * 7) : 64;
      const drh = fp && fp.getAttribute('defaultRowHeight') ? Math.round(+fp.getAttribute('defaultRowHeight') * 4 / 3) : 20;
      info.dcw = dcw; info.drh = drh;
      for (let c = 1; c <= 60; c++) info.colW[c] = dcw;
      find(d, 'col').forEach(c => { const mn = +c.getAttribute('min'), mx = Math.min(+c.getAttribute('max'), 60), w = c.getAttribute('hidden') === '1' ? 0 : Math.round(+c.getAttribute('width') * 7); for (let k = mn; k <= mx; k++) info.colW[k] = w; });
      const colRef = a => { let n = 0; for (const ch of a) { if (ch < 'A' || ch > 'Z') break; n = n * 26 + ch.charCodeAt(0) - 64; } return n; };
      const rowOf = a => +a.replace(/^[A-Z]+/, '');
      find(d, 'row').forEach(r => {
        const rn = +r.getAttribute('r');
        if (r.getAttribute('hidden') === '1') info.rowH[rn] = 0; else if (r.getAttribute('ht')) info.rowH[rn] = Math.round(+r.getAttribute('ht') * 4 / 3);
        kids(r, 'c').forEach(c => {
          const a = c.getAttribute('r'); const st = styleFor(+(c.getAttribute('s') || 0));
          if (st && Object.keys(st).length) info.style[rowOf(a) + ',' + colRef(a)] = st;
        });
      });
      find(d, 'mergeCell').forEach(m => info.merges.push(m.getAttribute('ref')));
      // merged cells: top/left from master, bottom/right from the far cells; bg from master
      const pa = a => ({ r: rowOf(a), c: colRef(a) });
      info.merges = info.merges.map(ref => { const [a, b] = ref.split(':'); const p = pa(a), q = pa(b || a); return { r1: p.r, c1: p.c, r2: q.r, c2: q.c }; }).filter(m => m.r2 > m.r1 || m.c2 > m.c1);
      const get = (r, c) => info.style[r + ',' + c];
      for (const m of info.merges) {
        const M = get(m.r1, m.c1) || (info.style[m.r1 + ',' + m.c1] = {});
        const B = get(m.r2, m.c1), R = get(m.r1, m.c2), BR = get(m.r2, m.c2);
        const pickB = (B && B.bb) || (BR && BR.bb), pickR = (R && R.br) || (BR && BR.br);
        if (pickB) M.bb = pickB; if (pickR) M.br = pickR;
      }
      // propagate borders to the neighbouring cell so the line is drawn whichever cell the browser prefers
      const set = (r, c, k, v) => { if (r < 1 || c < 1) return; const kk = r + ',' + c; const o = info.style[kk] || (info.style[kk] = {}); if (!o[k]) o[k] = v; };
      for (const k of Object.keys(info.style)) {
        const [r, c] = k.split(',').map(Number), s = info.style[k]; let mr = 0, mc = 0;
        const mm = info.merges.find(m => m.r1 === r && m.c1 === c); if (mm) { mr = mm.r2 - r; mc = mm.c2 - c; }
        if (s.bb) set(r + mr + 1, c, 'bt', s.bb); if (s.bt) set(r - 1, c, 'bb', s.bt);
        if (s.br) set(r, c + mc + 1, 'bl', s.br); if (s.bl) set(r, c - 1, 'br', s.bl);
      }
      out.sheets[name] = info;
    }
    return out;
  }
  return { read };
})();
