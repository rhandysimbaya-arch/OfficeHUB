/* Excel Tier-2 extras: Tables + structured references, PivotTable, conditional-format extras, sparklines,
   charts, formula auditing, cell styles/number-format colours.  Loaded after sheety.js. */
'use strict';
/* ---------- post-paint hook (sparklines, icons, number-format colours, error flags) ---------- */
const _paint2 = paint;
paint = function (r, c) { _paint2(r, c); try { postPaint(r, c); } catch (e) { } };
const POST = [];
function postPaint(r, c) { const td = tds[r] && tds[r][c]; if (!td) return; for (const f of POST) f(td, r, c); }
(function () {
  const st = document.createElement('style');
  st.textContent = `.g td.nat{position:relative}.g td.nat::before{content:'';position:absolute;left:0;top:0;border-style:solid;border-width:7px 7px 0 0;border-color:#2a9d3f transparent transparent transparent;z-index:2}
.g td.tbH{font-weight:700!important}.g td.tbT{font-weight:700!important;border-top:2px solid #444!important}
.g td.tbH.tblue{background:#4472C4!important;color:#fff!important}.g td.tbA.tblue{background:#D9E1F2!important}
.g td.tbH.tgreen{background:#70AD47!important;color:#fff!important}.g td.tbA.tgreen{background:#E2EFDA!important}
.g td.tbH.torange{background:#ED7D31!important;color:#fff!important}.g td.tbA.torange{background:#FCE4D6!important}
.g td.tbH.tgray{background:#7F7F7F!important;color:#fff!important}.g td.tbA.tgray{background:#EDEDED!important}
.g td.tbH.tdark{background:#222!important;color:#fff!important}.g td.tbA.tdark{background:#D9D9D9!important}`;
  document.head.appendChild(st);
})();

/* ---------- Tables ---------- */
function allTables() { const o = []; for (const sn of wb.order) { const T = (wb.sheets[sn].ext && wb.sheets[sn].ext.tables) || []; T.forEach(t => o.push({ sn, t })); } return o; }
function tableAt(r, c, sn) { sn = sn || cur; const T = (wb.sheets[sn].ext && wb.sheets[sn].ext.tables) || []; return T.find(t => r >= t.r1 && r <= t.r2 + (t.total ? 1 : 0) && c >= t.c1 && c <= t.c2) || null; }
const clearAst = () => { for (const sn of wb.order) for (const k in wb.sheets[sn].cells) delete wb.sheets[sn].cells[k].ast; };
EXTRA.push((r, c) => { const t = tableAt(r, c); if (!t || t.style === 'none') return ''; const st = ' t' + (t.style || 'blue'); if (t.hdr !== false && r === t.r1) return ' tbH' + st; if (t.total && r === t.r2 + 1) return ' tbT'; if (t.banded !== false && (r - t.r1) % 2 === 0) return ' tbA' + st; return ''; });
function tableHeaders(t, sn) { const o = []; for (let c = t.c1; c <= t.c2; c++) o.push(String(wb.value(sn || cur, t.r1, c) ?? '')); return o; }
function tableNameFree(n) { n = n.replace(/\s+/g, '_'); return !allTables().some(x => x.t.name.toLowerCase() === n.toLowerCase()); }
function tableCreate(R, style, name) {
  for (let c = R.c1; c <= R.c2; c++) { let h = wb.value(cur, R.r1, c); if (isBlank(h) || typeof h !== 'string') wb.setRaw(cur, R.r1, c, isBlank(h) ? 'Column' + (c - R.c1 + 1) : String(h)); }
  const seen = {}; for (let c = R.c1; c <= R.c2; c++) { let h = String(wb.value(cur, R.r1, c)); if (seen[h.toLowerCase()]) { h = h + (c - R.c1 + 1); wb.setRaw(cur, R.r1, c, h); } seen[h.toLowerCase()] = 1; }
  let i = 1; while (!tableNameFree('Table' + i)) i++; name = name && tableNameFree(name) ? name.replace(/\s+/g, '_') : 'Table' + i;
  const t = { name, r1: R.r1, c1: R.c1, r2: R.r2, c2: R.c2, style: style || 'blue', banded: true, hdr: true, total: false }; (X().tables = X().tables || []).push(t);
  if (!X().af) { X().af = { r: t.r1, c1: t.c1, c2: t.c2 }; X().afv = {}; }
  clearAst(); save(); refresh(); stMsg('Table "' + name + '" created');
}
function tableDlg() {
  const R = regionSel(); if (R.r2 === R.r1) { alert('Select the data (header row + rows) first.'); return; }
  gen('Create Table', `<div style="font-size:12px;color:#555">Range ${E.addr(R.r1, R.c1)}:${E.addr(R.r2, R.c2)} — first row becomes the header</div><label>Table name</label><input type="text" id="tbN" placeholder="Table1"><label>Style</label><select id="tbS"><option value="blue">Blue</option><option value="green">Green</option><option value="orange">Orange</option><option value="gray">Gray</option><option value="dark">Dark</option><option value="none">None</option></select>`, [['OK', () => { const n = $('tbN').value.trim(), s = $('tbS').value; closeGen(); tableCreate(R, s, n); }]]);
}
fmtTable = tableDlg;
function tableProps() {
  const t = tableAt(anchor.r, anchor.c); if (!t) { alert('Select a cell inside a table first.'); return; }
  gen('Table Properties', `<label>Name</label><input type="text" id="tpN" value="${esc(t.name)}"><label>Style</label><select id="tpS">${['blue', 'green', 'orange', 'gray', 'dark', 'none'].map(s => `<option ${s === t.style ? 'selected' : ''}>${s}</option>`).join('')}</select><label><input type="checkbox" id="tpB" ${t.banded !== false ? 'checked' : ''}> Banded rows</label><label><input type="checkbox" id="tpT" ${t.total ? 'checked' : ''}> Total row</label><label>Range</label><input type="text" id="tpR" value="${E.addr(t.r1, t.c1)}:${E.addr(t.r2, t.c2)}">`, [
    ['OK', () => { const n = $('tpN').value.trim().replace(/\s+/g, '_'); if (n && n !== t.name && tableNameFree(n)) { const old = t.name; t.name = n; renameTableRefs(old, n); } t.style = $('tpS').value; t.banded = $('tpB').checked; const rg = /^([A-Z]+)(\d+):([A-Z]+)(\d+)$/i.exec($('tpR').value.trim()); if (rg) { t.c1 = E.colToNum(rg[1].toUpperCase()); t.r1 = +rg[2]; t.c2 = E.colToNum(rg[3].toUpperCase()); t.r2 = +rg[4]; } setTotal(t, $('tpT').checked); clearAst(); save(); closeGen(); refresh(); }],
    ['Convert to range', () => { X().tables = X().tables.filter(x => x !== t); clearAst(); save(); closeGen(); refresh(); }]]);
}
function renameTableRefs(a, b) { const re = new RegExp('\\b' + a.replace(/[.*+?^${}()|[\]\\]/g, '\\$&') + '(?=\\[)', 'g'); for (const sn of wb.order) for (const k in wb.sheets[sn].cells) { const x = wb.sheets[sn].cells[k]; if (x.f) x.f = x.f.replace(re, b); } }
function setTotal(t, on) {
  if (on && !t.total) { t.total = true; const r = t.r2 + 1; wb.setRaw(cur, r, t.c1, 'Total'); const lc = t.c2; const h = tableHeaders(t)[lc - t.c1]; let numeric = false; for (let rr = t.r1 + 1; rr <= t.r2; rr++) if (typeof cellVal(rr, lc) === 'number') numeric = true; if (numeric) wb.setRaw(cur, r, lc, `=SUBTOTAL(109,${t.name}[${h}])`); }
  else if (!on && t.total) { const r = t.r2 + 1; for (let c = t.c1; c <= t.c2; c++) wb.setRaw(cur, r, c, ''); t.total = false; }
}
const _commitFx5 = commitFx;
commitFx = function (m) { const ec = editCell || { sheet: cur, r: anchor.r, c: anchor.c }; const had = $('fx').value; _commitFx5(m); if (had === '') return; const sn = ec.sheet; const T = (wb.sheets[sn] && wb.sheets[sn].ext && wb.sheets[sn].ext.tables) || []; for (const t of T) { if (!t.total && ec.r === t.r2 + 1 && ec.c >= t.c1 && ec.c <= t.c2) { t.r2++; if (X().af && X().af.r === t.r1) { } clearAst(); save(); refresh(); } } };
/* structured reference preprocessor (used by the engine before parsing) */
E.Workbook.prototype.pre = function (sheet, r, c, f) {
  const all = []; for (const sn of this.order) { const T = (this.sheets[sn].ext && this.sheets[sn].ext.tables) || []; T.forEach(t => all.push({ sn, t })); } if (!all.length) return f;
  const q = sn => (/^\w+$/.test(sn) ? sn : "'" + sn + "'") + '!';
  const A = (r, c) => '$' + E.numToCol(c) + '$' + r;
  let out = '', i = 0; const n = f.length;
  while (i < n) {
    const ch = f[i];
    if (ch === '"') { let j = i + 1; while (j < n) { if (f[j] === '"') { if (f[j + 1] === '"') { j += 2; continue; } break; } j++; } out += f.slice(i, j + 1); i = j + 1; continue; }
    const m = /^([A-Za-z_][\w.]*)?\[/.exec(f.slice(i)); if (m && (m[1] || ch === '[')) {
      let j = i + (m[1] ? m[1].length : 0), depth = 0, k = j; for (; k < n; k++) { if (f[k] === '[') depth++; else if (f[k] === ']') { depth--; if (depth === 0) break; } } if (k >= n) { out += ch; i++; continue; }
      const inner = f.slice(j + 1, k); let tb = m[1] ? all.find(x => x.t.name.toLowerCase() === m[1].toLowerCase()) : all.find(x => x.sn === sheet && r >= x.t.r1 && r <= x.t.r2 + (x.t.total ? 1 : 0) && c >= x.t.c1 && c <= x.t.c2);
      if (!tb) { out += f.slice(i, k + 1); i = k + 1; continue; } const T = tb.t, hd = []; for (let cc = T.c1; cc <= T.c2; cc++) hd.push(String(this.value(tb.sn, T.r1, cc) ?? '').toLowerCase());
      let rows = [T.r1 + 1, T.r2], cols = [T.c1, T.c2], thisRow = false; const items = inner.match(/\[[^\]]*\]|[^,\[\]]+/g) || []; const colHit = [];
      for (let it of items) { it = it.trim(); const nm = it.replace(/^\[|\]$/g, '').trim(); if (nm[0] === '@') { thisRow = true; const cn = nm.slice(1).replace(/^\[|\]$/g, '').trim(); if (cn) colHit.push(cn); continue; } const lo = nm.toLowerCase(); if (lo === '#all') rows = [T.r1, T.r2 + (T.total ? 1 : 0)]; else if (lo === '#headers') rows = [T.r1, T.r1]; else if (lo === '#data') rows = [T.r1 + 1, T.r2]; else if (lo === '#totals') rows = [T.r2 + 1, T.r2 + 1]; else if (lo === '#this row') thisRow = true; else if (nm) colHit.push(nm); }
      if (colHit.length) { const idx = []; for (const cn of colHit) { const parts = cn.split(':'); parts.forEach(p => { const ix = hd.indexOf(p.trim().replace(/^\[|\]$/g, '').toLowerCase()); if (ix >= 0) idx.push(T.c1 + ix); }); } if (idx.length) cols = [Math.min(...idx), Math.max(...idx)]; }
      if (thisRow) rows = [r, r]; const rg = rows[0] === rows[1] && cols[0] === cols[1] ? A(rows[0], cols[0]) : A(rows[0], cols[0]) + ':' + A(rows[1], cols[1]);
      out += (tb.sn !== sheet ? q(tb.sn) : '') + rg; i = k + 1; continue;
    }
    if (/[A-Za-z_]/.test(ch) && (i === 0 || !/[\w.]/.test(f[i - 1]))) { let j = i; while (j < n && /[\w.]/.test(f[j])) j++; const tok = f.slice(i, j); const tb = f[j] !== '(' && f[j] !== '[' && all.find(x => x.t.name.toLowerCase() === tok.toLowerCase()); if (tb) { const T = tb.t; out += (tb.sn !== sheet ? q(tb.sn) : '') + A(T.r1 + 1, T.c1) + ':' + A(T.r2, T.c2); } else out += tok; i = j; continue; }
    out += ch; i++;
  }
  return out;
};

/* ---------- PivotTable ---------- */
function pivotDlg() {
  const R = regionSel(); if (R.r2 - R.r1 < 1) { alert('Select your data (with a header row) first.'); return; }
  const hdr = []; for (let c = R.c1; c <= R.c2; c++) hdr.push(String(cellVal(R.r1, c) ?? 'Column ' + E.numToCol(c)));
  const opt = (none, sel0) => (none ? '<option value="-1">(none)</option>' : '') + hdr.map((h, i) => `<option value="${i}" ${i === sel0 ? 'selected' : ''}>${esc(h)}</option>`).join('');
  const aggs = ['sum', 'count', 'avg', 'min', 'max'].map(a => `<option value="${a}">${a === 'avg' ? 'Average' : a[0].toUpperCase() + a.slice(1)}</option>`).join('');
  gen('Create PivotTable', `<div style="font-size:12px;color:#555">Source ${esc(cur)}!${E.addr(R.r1, R.c1)}:${E.addr(R.r2, R.c2)}</div><label>Rows</label><select id="pvR1">${opt(false, 0)}</select><select id="pvR2">${opt(true, -1)}</select><label>Columns</label><select id="pvC">${opt(true, -1)}</select><label>Values</label><div style="display:flex;gap:4px"><select id="pvV1">${opt(false, hdr.length - 1)}</select><select id="pvA1">${aggs}</select></div><div style="display:flex;gap:4px;margin-top:4px"><select id="pvV2">${opt(true, -1)}</select><select id="pvA2">${aggs}</select></div>`, [['OK', () => {
    const def = { src: { sheet: cur, r1: R.r1, c1: R.c1, r2: R.r2, c2: R.c2 }, rows: [+$('pvR1').value, +$('pvR2').value].filter(x => x >= 0), col: +$('pvC').value, vals: [{ f: +$('pvV1').value, agg: $('pvA1').value }].concat(+$('pvV2').value >= 0 ? [{ f: +$('pvV2').value, agg: $('pvA2').value }] : []) };
    closeGen(); let i = 1; while (wb.sheets['PivotTable' + i]) i++; const name = 'PivotTable' + i; wb.addSheet(name); wb.sheets[name].ext = { pivot: def }; pivotBuild(name); save(); gotoSheet(name); renderTabs();
  }]]);
}
function pivotBuild(name) {
  const S = wb.sheets[name], D = S.ext.pivot, s = D.src; const H = []; for (let c = s.c1; c <= s.c2; c++) H.push(String(wb.value(s.sheet, s.r1, c) ?? ''));
  const rows = []; for (let r = s.r1 + 1; r <= s.r2; r++) { const row = []; for (let c = s.c1; c <= s.c2; c++) row.push(wb.value(s.sheet, r, c)); if (row.some(v => !isBlank(v))) rows.push(row); }
  const agg = (vs, a) => { const n = vs.filter(v => typeof v === 'number'); if (a === 'count') return vs.filter(v => !isBlank(v)).length; if (!n.length) return a === 'sum' ? 0 : null; if (a === 'sum') return n.reduce((x, y) => x + y, 0); if (a === 'avg') return n.reduce((x, y) => x + y, 0) / n.length; if (a === 'min') return Math.min(...n); return Math.max(...n); };
  const lab = v => isBlank(v) ? '(blank)' : v; const cmpv = (a, b) => cmpSort(a, b, false);
  const rk = row => D.rows.map(i => lab(row[i])); const keyOf = a => JSON.stringify(a);
  const rkeys = new Map(); rows.forEach(r => { const k = rk(r); rkeys.set(keyOf(k), k); }); const rl = [...rkeys.values()].sort((a, b) => { for (let i = 0; i < a.length; i++) { const d = cmpv(a[i], b[i]); if (d) return d; } return 0; });
  const cl = D.col >= 0 ? [...new Set(rows.map(r => lab(r[D.col])))].sort(cmpv) : [null];
  const cells = {}; const put = (r, c, v, st, fmt) => { const k = r + ',' + c; S.cells[k] = { v }; if (st) S.style[k] = st; if (fmt) S.fmt[k] = fmt; }; S.cells = {}; S.style = {}; S.fmt = {}; S.colW = {};
  const hs = { b: 1, bg: '#D9E1F2', bb: '1px solid #4472C4' }, ts = { b: 1, bt: '1px solid #4472C4', bg: '#F2F2F2' }; const aggName = a => a === 'avg' ? 'Average' : a[0].toUpperCase() + a.slice(1); const nf = a => a === 'count' ? '0' : '#,##0.00';
  put(1, 1, 'PivotTable — source ' + s.sheet + '!' + E.addr(s.r1, s.c1) + ':' + E.addr(s.r2, s.c2), { i: 1, color: '#666' }); const R0 = 3; let c0 = 1;
  D.rows.forEach((fi, i) => put(R0, c0 + i, H[fi], hs)); c0 += D.rows.length; const heads = [];
  cl.forEach(ci => D.vals.forEach(v => heads.push({ ci, v }))); heads.forEach((h, j) => put(R0, c0 + j, (h.ci === null ? '' : h.ci + ' — ') + aggName(h.v.agg) + ' of ' + H[h.v.f], hs));
  const gtCols = D.vals.length; D.vals.forEach((v, j) => put(R0, c0 + heads.length + j, (D.col >= 0 ? 'Grand Total — ' : '') + aggName(v.agg) + ' of ' + H[v.f], hs));
  rl.forEach((k, ri) => { const rr = R0 + 1 + ri; k.forEach((x, i) => put(rr, 1 + i, x)); const mine = rows.filter(r => keyOf(rk(r)) === keyOf(k)); heads.forEach((h, j) => { const sub = h.ci === null ? mine : mine.filter(r => lab(r[D.col]) === h.ci); put(rr, c0 + j, agg(sub.map(r => r[h.v.f]), h.v.agg), null, nf(h.v.agg)); }); D.vals.forEach((v, j) => put(rr, c0 + heads.length + j, agg(mine.map(r => r[v.f]), v.agg), { b: 1 }, nf(v.agg))); });
  const tr = R0 + 1 + rl.length; put(tr, 1, 'Grand Total', ts); for (let i = 1; i < D.rows.length; i++) put(tr, 1 + i, '', ts); heads.forEach((h, j) => { const sub = h.ci === null ? rows : rows.filter(r => lab(r[D.col]) === h.ci); put(tr, c0 + j, agg(sub.map(r => r[h.v.f]), h.v.agg), ts, nf(h.v.agg)); }); D.vals.forEach((v, j) => put(tr, c0 + heads.length + j, agg(rows.map(r => r[v.f]), v.agg), ts, nf(v.agg)));
  for (let c = 1; c <= c0 + heads.length + gtCols; c++) S.colW[c] = c <= D.rows.length ? 110 : 120;
}
function pivotRefresh(all) { let n = 0; for (const sn of wb.order) { const S = wb.sheets[sn]; if (S.ext && S.ext.pivot && (all || sn === cur)) { pivotBuild(sn); n++; } } if (!n) { stMsg('No PivotTable on this sheet'); return; } save(); build(); stMsg('Refreshed ' + n + ' PivotTable(s)'); }

/* ---------- Conditional-format extras: formula rules, icon sets, manager ---------- */
const _cfHit = cfHit;
cfHit = function (R, r, c, v) {
  if (R.type === 'formula') { try { const f = shiftFormula0 ? shiftFormula0(R.a.startsWith('=') ? R.a : '=' + R.a, r - R.r1, c - R.c1) : R.a; const x = E.evaluate(wb, cur, f.startsWith('=') ? f : '=' + f); return x === true || (typeof x === 'number' && x !== 0) || String(x).toUpperCase() === 'TRUE'; } catch (e) { return false; } }
  if (R.type.indexOf('icons') === 0) return false; return _cfHit(R, r, c, v);
};
const ICONSETS = { icons3a: ['&#9660;', '&#9654;', '&#9650;'], icons3t: ['&#128308;', '&#128993;', '&#128994;'], icons3f: ['&#9873;', '&#9873;', '&#9873;'], icons4r: ['&#9675;', '&#9684;', '&#9681;', '&#9679;'], icons5s: ['&#9734;', '&#9733;&#9734;', '&#9733;&#9733;&#9734;', '&#9733;&#9733;&#9733;&#9734;', '&#9733;&#9733;&#9733;&#9733;&#9733;'] };
ICONSETS.icons3a = ['▼', '▶', '▲']; ICONSETS.icons3t = ['🔴', '🟡', '🟢']; ICONSETS.icons3f = ['⚑', '⚑', '⚑']; ICONSETS.icons4r = ['○', '◔', '◑', '●']; ICONSETS.icons5s = ['★☆☆☆☆', '★★☆☆☆', '★★★☆☆', '★★★★☆', '★★★★★'];
const ICOLOR = { icons3a: ['#C00000', '#E6A100', '#2E8B57'], icons3f: ['#C00000', '#E6A100', '#2E8B57'] };
POST.push((td, r, c) => {
  const L = sh().cf; if (!L) return; const v = wb.value(cur, r, c); if (typeof v !== 'number') return;
  for (const R of L) { if (R.type.indexOf('icons') !== 0 || r < R.r1 || r > R.r2 || c < R.c1 || c > R.c2) continue; const set = ICONSETS[R.type]; if (!set) continue; const S = cfStats(R); const t = S.max === S.min ? 1 : (v - S.min) / (S.max - S.min); const i = Math.min(set.length - 1, Math.floor(t * set.length)); const sp = document.createElement('span'); sp.textContent = set[i] + ' '; const cl = ICOLOR[R.type]; if (cl) sp.style.color = cl[Math.min(cl.length - 1, i)]; td.insertBefore(sp, td.firstChild); break; }
});
const _cfDlg = cfDlg;
cfDlg = function () {
  _cfDlg(); const s = $('cfT'); if (!s) return; const add = (v, t) => { const o = document.createElement('option'); o.value = v; o.textContent = t; s.appendChild(o); };
  add('formula', 'Use a formula (enter in Value box, e.g. =$B2>100)'); add('icons3a', 'Icon set: 3 arrows'); add('icons3t', 'Icon set: 3 traffic lights'); add('icons3f', 'Icon set: 3 flags'); add('icons4r', 'Icon set: 4 ratings'); add('icons5s', 'Icon set: 5 stars');
  const g = $('gR'); const b = document.createElement('button'); b.textContent = 'Manage rules…'; b.onclick = () => { closeGen(); cfMgr(); }; g.appendChild(b);
};
function cfDesc(R) { const m = { gt: '> ', lt: '< ', ge: '>= ', le: '<= ', eq: '= ', ne: '<> ', between: 'between ', text: 'contains ', begins: 'begins ', top: 'top ', bottom: 'bottom ', formula: 'formula ' }[R.type] || ''; return `${R.type.indexOf('icons') === 0 ? 'Icon set' : R.type === 'scale' ? 'Color scale' : R.type === 'bar' ? 'Data bars' : m + (R.a || '') + (R.type === 'between' ? ' and ' + R.b : '') || R.type} — ${E.addr(R.r1, R.c1)}:${E.addr(R.r2, R.c2)}`; }
function cfMgr() {
  const S = sh(); S.cf = S.cf || []; const render = () => S.cf.map((R, i) => `<div style="display:flex;gap:6px;align-items:center;margin:4px 0"><span style="flex:1;font-size:13px">${i + 1}. ${esc(cfDesc(R))} ${R.bg && R.type.indexOf('icons') !== 0 && R.type !== 'scale' && R.type !== 'bar' ? `<span style="background:${R.bg};color:${R.color};padding:0 6px">Aa</span>` : ''}</span><button data-up="${i}">↑</button><button data-del="${i}">Delete</button></div>`).join('') || '<i>No rules on this sheet</i>';
  gen('Conditional Formatting Rules Manager', `<div id="cfList">${render()}</div>`, []);
  const bind = () => { document.querySelectorAll('#cfList button').forEach(b => b.onclick = () => { if (b.dataset.del !== undefined) S.cf.splice(+b.dataset.del, 1); else { const i = +b.dataset.up; if (i > 0) [S.cf[i - 1], S.cf[i]] = [S.cf[i], S.cf[i - 1]]; } CFC = {}; save(); refresh(); $('cfList').innerHTML = render(); bind(); }); }; bind();
}
(function () { const o = Object.assign; })();

/* ---------- Sparklines ---------- */
function sparkDlg(type) {
  const R = regionSel(); gen('Create Sparklines', `<label>Data range</label><input type="text" id="spR" value="${E.addr(sel.r1, sel.c1)}:${E.addr(sel.r2, sel.c2)}"><label>Location (cell or range)</label><input type="text" id="spL" placeholder="e.g. F2 or F2:F6"><label>Type</label><select id="spT"><option value="line">Line</option><option value="col">Column</option><option value="wl">Win/Loss</option></select>`, [['OK', () => {
    const pr = /^([A-Z]+)(\d+):([A-Z]+)(\d+)$/i, a = pr.exec($('spR').value.trim()), l = /^([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?$/i.exec($('spL').value.trim()); if (!a || !l) { alert('Enter ranges like B2:E6 and F2:F6'); return; }
    const r1 = +a[2], c1 = E.colToNum(a[1].toUpperCase()), r2 = +a[4], c2 = E.colToNum(a[3].toUpperCase()), lr1 = +l[2], lc = E.colToNum(l[1].toUpperCase()), lr2 = l[4] ? +l[4] : lr1; const rowsMode = (lr2 - lr1 + 1) === (r2 - r1 + 1); const sp = X().spark = X().spark || {}; const ty = $('spT').value;
    for (let i = 0; lr1 + i <= lr2; i++) sp[(lr1 + i) + ',' + lc] = rowsMode ? { t: ty, r1: r1 + i, c1, r2: r1 + i, c2 } : { t: ty, r1, c1: c1 + i, r2, c2: c1 + i }; closeGen(); save(); refresh();
  }], ['Clear', () => { const sp = X().spark; if (sp) for (let r = sel.r1; r <= sel.r2; r++)for (let c = sel.c1; c <= sel.c2; c++) delete sp[r + ',' + c]; closeGen(); save(); refresh(); }]]);
}
POST.push((td, r, c) => {
  const sp = X().spark && X().spark[r + ',' + c]; if (!sp) return; const v = []; for (let i = sp.r1; i <= sp.r2; i++) for (let j = sp.c1; j <= sp.c2; j++) { const x = wb.value(cur, i, j); v.push(typeof x === 'number' ? x : null); }
  const nums = v.filter(x => x !== null); if (!nums.length) return; const mn = Math.min(...nums, 0 * (sp.t === 'line' ? 1 : 1)), mx = Math.max(...nums); const lo = sp.t === 'line' ? Math.min(...nums) : Math.min(0, ...nums), hi = mx === lo ? lo + 1 : mx; const W = 100, H = 30, n = v.length; const X0 = i => n > 1 ? 4 + i * (W - 8) / (n - 1) : W / 2, Y = x => H - 4 - (x - lo) / (hi - lo) * (H - 8); let g = '';
  if (sp.t === 'line') { const pts = v.map((x, i) => x === null ? null : X0(i) + ',' + Y(x)).filter(Boolean); g = `<polyline fill="none" stroke="#4472C4" stroke-width="2" points="${pts.join(' ')}"/>`; const mi = v.indexOf(mx); if (mi >= 0) g += `<circle cx="${X0(mi)}" cy="${Y(mx)}" r="3" fill="#C00000"/>`; }
  else { const bw = (W - 8) / n * .7; v.forEach((x, i) => { if (x === null) return; const cx = 4 + (i + .5) * (W - 8) / n; if (sp.t === 'wl') { g += `<rect x="${cx - bw / 2}" y="${x >= 0 ? 4 : H / 2}" width="${bw}" height="${H / 2 - 4}" fill="${x >= 0 ? '#4472C4' : '#C00000'}"/>`; } else { const y0 = Y(Math.max(lo, 0)), y1 = Y(x); g += `<rect x="${cx - bw / 2}" y="${Math.min(y0, y1)}" width="${bw}" height="${Math.max(1, Math.abs(y0 - y1))}" fill="${x < 0 ? '#C00000' : '#4472C4'}"/>`; } }); }
  const svg = `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 ${W} ${H}" preserveAspectRatio="none">${g}</svg>`; td.style.backgroundImage = `url("data:image/svg+xml;utf8,${encodeURIComponent(svg)}")`; td.style.backgroundRepeat = 'no-repeat'; td.style.backgroundSize = '100% 100%';
});

/* ---------- Number-format colour tags [Red] etc. ---------- */
const NFCOL = { red: '#FF0000', blue: '#0000FF', green: '#008000', yellow: '#B8860B', magenta: '#FF00FF', cyan: '#008B8B', black: '#000000', white: '#FFFFFF' };
POST.push((td, r, c) => {
  const f = sh().fmt[r + ',' + c]; if (!f || f.indexOf('[') < 0) return; const v = wb.value(cur, r, c); if (typeof v !== 'number') return; const secs = f.split(';'); const sec = secs.length > 1 && v < 0 ? secs[1] : secs.length > 2 && v === 0 ? secs[2] : secs[0]; const m = /\[(red|blue|green|yellow|magenta|cyan|black|white)\]/i.exec(sec); if (m) td.style.color = NFCOL[m[1].toLowerCase()];
});

/* ---------- Error checking (green triangle for numbers stored as text) ---------- */
function numAsText(r, c) { const raw = wb.getRaw(cur, r, c); return typeof raw === 'string' && raw !== '' && raw[0] !== '=' && raw[0] !== "'" && /^\s*[-+]?\d+(\.\d+)?\s*$/.test(raw) && !(sh().fmt[r + ',' + c] === '@'); }
POST.push((td, r, c) => { if (numAsText(r, c)) td.classList.add('nat'); else td.classList.remove('nat'); });
function errCheck() {
  const out = []; for (const k in sh().cells) { const [r, c] = k.split(',').map(Number); const v = wb.value(cur, r, c); if (v && typeof v === 'object' && v.code) out.push([r, c, v.code]); else if (numAsText(r, c)) out.push([r, c, 'Number stored as text']); }
  gen('Error Checking', out.length ? out.slice(0, 40).map(o => `<div style="margin:4px 0"><a href="#" data-g="${o[0]},${o[1]}">${E.addr(o[0], o[1])}</a> — ${esc(o[2])}</div>`).join('') + (out.length > 40 ? '<i>…and ' + (out.length - 40) + ' more</i>' : '') : '<i>No errors found on this sheet</i>', [['Convert text numbers', () => { out.filter(o => o[2][0] === 'N').forEach(o => wb.setRaw(cur, o[0], o[1], String(+wb.getRaw(cur, o[0], o[1])))); closeGen(); save(); refresh(); }]]);
  document.querySelectorAll('#gB a[data-g]').forEach(a => a.onclick = e => { e.preventDefault(); const [r, c] = a.dataset.g.split(',').map(Number); closeGen(); goCell(r, c); });
}

/* ---------- Formula auditing ---------- */
function refsOf(f, sn) { const out = []; const re = /(?:(?:'([^']+)'|([A-Za-z_]\w*))!)?\$?([A-Z]{1,3})\$?(\d+)(?::\$?([A-Z]{1,3})\$?(\d+))?/g; let m; f = f.replace(/"[^"]*"/g, '""'); while ((m = re.exec(f))) { if (/[A-Za-z0-9_]/.test(f[m.index - 1] || '') && !m[1] && !m[2]) continue; if (f[m.index + m[0].length] === '(') continue; const s = m[1] || m[2] || sn; const c1 = E.colToNum(m[3]), r1 = +m[4], c2 = m[5] ? E.colToNum(m[5]) : c1, r2 = m[6] ? +m[6] : r1; out.push({ s, r1: Math.min(r1, r2), c1: Math.min(c1, c2), r2: Math.max(r1, r2), c2: Math.max(c1, c2) }); } return out; }
function tracePrec() { const raw = wb.getRaw(cur, anchor.r, anchor.c); if (typeof raw !== 'string' || raw[0] !== '=') { stMsg('Selected cell has no formula'); return; } const ks = []; for (const R of refsOf(wb.pre ? wb.pre(cur, anchor.r, anchor.c, raw) : raw, cur)) { if (R.s !== cur) continue; for (let r = R.r1; r <= R.r2 && r < R.r1 + 400; r++) for (let c = R.c1; c <= R.c2; c++) ks.push(r + ',' + c); } setHL(ks); stMsg(ks.length + ' precedent cell(s) highlighted'); }
function traceDep() { const ks = []; const tgt = (r, c) => anchor.r === r && anchor.c === c; for (const k in sh().cells) { const raw = sh().cells[k].f; if (!raw) continue; const [r, c] = k.split(',').map(Number); for (const R of refsOf(wb.pre ? wb.pre(cur, r, c, raw) : raw, cur)) { if (R.s === cur && anchor.r >= R.r1 && anchor.r <= R.r2 && anchor.c >= R.c1 && anchor.c <= R.c2) { ks.push(k); break; } } } setHL(ks); stMsg(ks.length + ' dependent cell(s) highlighted'); }
function clearArrows() { clearHL(); stMsg('Ready'); }
function evalDlg() {
  const raw = wb.getRaw(cur, anchor.r, anchor.c); if (typeof raw !== 'string' || raw[0] !== '=') { alert('Select a cell with a formula.'); return; }
  const parts = []; const re = /(?:[A-Za-z_]\w*!|'[^']+'!)?\$?[A-Z]{1,3}\$?\d+(?::\$?[A-Z]{1,3}\$?\d+)?/g; let m, seen = {}; while ((m = re.exec(raw))) { if (/[A-Za-z]/.test(raw[m.index - 1] || '') || raw[m.index + m[0].length] === '(') continue; if (seen[m[0]]) continue; seen[m[0]] = 1; let v; try { if (m[0].indexOf(':') > 0 && m[0].indexOf('!') < 0) { const q = refsOf(m[0], cur)[0], a = []; for (let i = q.r1; i <= q.r2; i++)for (let j = q.c1; j <= q.c2; j++)a.push(wb.value(cur, i, j)); v = '{' + a.join(',') + '}'; } else v = E.evaluate(wb, cur, '=' + m[0]); } catch (e) { v = '?'; } if (Array.isArray(v)) v = '{' + v.map(r => r.join(',')).join(';') + '}'; parts.push([m[0], v]); }
  let step = raw; parts.forEach(p => { step = step.split(p[0]).join(typeof p[1] === 'string' ? (p[1][0] === '{' ? p[1] : '"' + p[1] + '"') : (p[1] && p[1].code) || String(p[1])); });
  gen('Evaluate Formula', `<div style="font-family:monospace;font-size:13px;word-break:break-all"><b>Formula</b><br>${esc(raw)}<br><br><b>References replaced by values</b><br>${esc(step)}<br><br><b>Result</b><br>${esc(String(cellText(anchor.r, anchor.c)))}</div>`, []);
}
/* Watch window */
function watchDlg() {
  const W = X().watch = X().watch || []; const render = () => (W.map((k, i) => `<tr><td>${esc(cur)}</td><td>${E.addr(...k.split(',').map(Number))}</td><td>${esc(String(cellText(...k.split(',').map(Number))))}</td><td><button data-d="${i}">×</button></td></tr>`).join('') || '<tr><td colspan=4><i>Empty</i></td></tr>');
  gen('Watch Window', `<table style="width:100%;font-size:13px"><tr><th>Sheet</th><th>Cell</th><th>Value</th><th></th></tr><tbody id="wwB">${render()}</tbody></table>`, [['Add watch (selected cell)', () => { const k = anchor.r + ',' + anchor.c; if (!W.includes(k)) W.push(k); save(); closeGen(); watchDlg(); }]]);
  document.querySelectorAll('#wwB button').forEach(b => b.onclick = () => { W.splice(+b.dataset.d, 1); save(); closeGen(); watchDlg(); });
}

/* ---------- more cell styles + ribbon ---------- */
Object.assign(CSTYLES, { calc: { bg: '#F2F2F2', color: '#FA7D00', b: 1, bt: '1px solid #7F7F7F', bb: '1px solid #7F7F7F', bl: '1px solid #7F7F7F', br: '1px solid #7F7F7F' }, check: { bg: '#A5A5A5', color: '#fff', b: 1 }, warn: { color: '#FF0000' }, explan: { color: '#7F7F7F', i: 1 }, linked: { color: '#FA7D00', bb: '1px solid #FF8001' }, output: { bg: '#F2F2F2', color: '#3F3F3F', b: 1 }, h3: { b: 1, sz: 15, color: '#44546A', bb: '1px solid #A9BCE2' }, h4: { b: 1, color: '#44546A' }, accent3: { bg: '#A5A5A5', color: '#fff' }, accent4: { bg: '#FFC000', color: '#000' }, accent5: { bg: '#5B9BD5', color: '#fff' }, accent6: { bg: '#70AD47', color: '#fff' } });
(function () {
  const sel = [...document.querySelectorAll('select')].find(s => (s.getAttribute('onchange') || '').indexOf('cellStyle') === 0); if (sel) [['h3', 'Heading 3'], ['h4', 'Heading 4'], ['calc', 'Calculation'], ['check', 'Check Cell'], ['warn', 'Warning Text'], ['explan', 'Explanatory Text'], ['linked', 'Linked Cell'], ['output', 'Output'], ['accent3', 'Accent 3'], ['accent4', 'Accent 4'], ['accent5', 'Accent 5'], ['accent6', 'Accent 6']].forEach(([v, t]) => { if (![...sel.options].some(o => o.value === v)) { const o = document.createElement('option'); o.value = v; o.textContent = t; sel.appendChild(o); } });
  const rib = document.getElementById('rib'); if (!rib) return; const grp = (label, btns) => `<div class="grp"><div class="gb">${btns.map(b => `<button class="big" onclick="${b[1]}"><i>${b[2]}</i>${b[0]}</button>`).join('')}</div><div class="gl">${label}</div></div>`; const add = (tab, h) => { const t = rib.querySelector(`[data-t="${tab}"]`); if (t) t.insertAdjacentHTML('beforeend', h); };
  add('Insert', grp('Tables', [['PivotTable', 'pivotDlg()', '&#9638;'], ['Table', 'tableDlg()', '&#9636;']]) + grp('Sparklines', [['Line', "sparkDlg()", '&#128200;'], ['Column', "sparkDlg()", '&#9641;'], ['Win/Loss', "sparkDlg()", '&#9644;']]));
  add('Data', grp('PivotTable', [['Refresh All', 'pivotRefresh(true)', '&#8635;']]));
  add('Formulas', grp('Formula Auditing', [['Trace Precedents', 'tracePrec()', '&#8678;'], ['Trace Dependents', 'traceDep()', '&#8680;'], ['Remove Arrows', 'clearArrows()', '&#10006;'], ['Show Formulas', 'toggleShowF()', '&#402;'], ['Error Checking', 'errCheck()', '&#9888;'], ['Evaluate Formula', 'evalDlg()', '&#9654;'], ['Watch Window', 'watchDlg()', '&#128065;']]));
  add('Home', grp('Rules', [['CF Rules', 'cfMgr()', '&#9776;'], ['Table Options', 'tableProps()', '&#9881;']]));
})();
document.addEventListener('keydown', e => { if (e.key === 'F5' && e.altKey) { e.preventDefault(); pivotRefresh(false); } else if (e.key === '[' && e.ctrlKey && !e.shiftKey && !editCell) { e.preventDefault(); tracePrec(); } else if (e.key === ']' && e.ctrlKey && !e.shiftKey && !editCell) { e.preventDefault(); traceDep(); } }, true);
