/* Excel extras: undo/redo, desktop-style formatting, conditional formatting, freeze panes, filter, shortcuts.
   Loaded after the main script in sheet.html (uses its globals: wb, cur, sel, anchor, sh(), key(), E, paint(), refresh(), build() ...). */
'use strict';
/* ---------------- undo / redo ---------------- */
const H = { a: [], i: -1, busy: false };
function snap() { return JSON.stringify(wb.toJSON()); }
function pushHist() { if (H.busy) return; const s = snap(); if (H.a[H.i] === s) return; H.a = H.a.slice(0, H.i + 1); H.a.push(s); if (H.a.length > 60) H.a.shift(); H.i = H.a.length - 1; }
const _save0 = save;
save = function () { _save0(); pushHist(); };
function restoreSnap(s) { H.busy = true; try { const keep = cur; wb = E.Workbook.fromJSON(JSON.parse(s)); cur = wb.sheets[keep] ? keep : wb.order[0]; _save0(); renderTabs(); build(); } finally { H.busy = false; } }
function undo() { if (editing) { const ec = editCell; if (ec && $('fx').value === String(wb.getRaw(ec.sheet, ec.r, ec.c))) cancelEdit(); else return; } if (H.i <= 0) { $('stL').textContent = 'Nothing to undo'; return; } H.i--; restoreSnap(H.a[H.i]); $('stL').textContent = 'Undo'; }
function redo() { if (editing) { const ec = editCell; if (ec && $('fx').value === String(wb.getRaw(ec.sheet, ec.r, ec.c))) cancelEdit(); else return; } if (H.i >= H.a.length - 1) { $('stL').textContent = 'Nothing to redo'; return; } H.i++; restoreSnap(H.a[H.i]); $('stL').textContent = 'Redo'; }
pushHist();
const _refresh0 = refresh;
refresh = function () { CFC = {}; _refresh0(); };

/* ---------------- helpers ---------------- */
const PT = 4 / 3;
function selStyle(fn) { eachSel((r, c) => { const kk = key(r, c); const s = sh().style[kk] = sh().style[kk] || {}; fn(s, r, c); if (!Object.keys(s).length) delete sh().style[kk]; }); save(); refresh(); }
function resetSel(el) { if (el && el.tagName === 'SELECT') el.selectedIndex = 0; }
function szPt(pt) { sty('sz', Math.round(+pt * PT * 10) / 10); }
const SIZES = [8, 9, 10, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48, 72];
function curPt() { const st = styleOf(anchor.r, anchor.c) || {}; return Math.round((st.sz || sh().defSz || 14.67) / PT); }
function fontStep(d) { const p = curPt(); let i = SIZES.findIndex(x => x >= p); if (d > 0) { i = SIZES.findIndex(x => x > p); if (i < 0) i = SIZES.length - 1; } else { i = SIZES.length - 1; while (i > 0 && SIZES[i] >= p) i--; } szPt(SIZES[i]); }
function setValign(v) { sty('valign', v); }
function wrapToggle() { const on = !(styleOf(anchor.r, anchor.c) || {}).wrap; selStyle(s => { if (on) s.wrap = 1; else delete s.wrap; }); autoFitRow(true); }
function indentBy(d) { selStyle(s => { const n = Math.max(0, (s.indent || 0) + d); if (n) { s.indent = n; if (!s.align) s.align = 'left'; } else delete s.indent; }); }
function noFill() { selStyle(s => { delete s.bg; }); }
function rotateNote() { }

/* ---------------- merge ---------------- */
function mergeSel(mode) {
  const S = sh(); S.merges = S.merges || [];
  const hit = m => !(m.r2 < sel.r1 || m.r1 > sel.r2 || m.c2 < sel.c1 || m.c1 > sel.c2);
  if (mode === 'unmerge') { S.merges = S.merges.filter(m => !hit(m)); save(); refresh(); return; }
  if (sel.r1 === sel.r2 && sel.c1 === sel.c2) { alert('Select two or more cells to merge.'); return; }
  S.merges = S.merges.filter(m => !hit(m));
  const mk = (r1, c1, r2, c2) => {
    if (r1 === r2 && c1 === c2) return;
    for (let r = r1; r <= r2; r++) for (let c = c1; c <= c2; c++) if ((r !== r1 || c !== c1) && wb.getRaw(cur, r, c) !== '' && wb.getRaw(cur, r, c) !== null) { wb.setRaw(cur, r, c, ''); }
    S.merges.push({ r1, c1, r2, c2 });
    if (mode === 'center') { const kk = key(r1, c1); (S.style[kk] = S.style[kk] || {}).align = 'center'; }
  };
  if (mode === 'across') { for (let r = sel.r1; r <= sel.r2; r++) mk(r, sel.c1, r, sel.c2); } else mk(sel.r1, sel.c1, sel.r2, sel.c2);
  save(); refresh();
}

/* ---------------- borders ---------------- */
function bdCss() { const w = $('bdSt').value, col = $('bdCol').value || '#000000'; return w + ' ' + col; }
function setSide(r, c, side, val) {
  if (r < 1 || c < 1) return; const S = sh();
  const ap = (rr, cc, sd) => { const kk = key(rr, cc); const s = S.style[kk] = S.style[kk] || {}; if (val) s[sd] = val; else delete s[sd]; if (!Object.keys(s).length) delete S.style[kk]; };
  ap(r, c, side);
  if (side === 'bt') ap(r - 1, c, 'bb'); else if (side === 'bb') ap(r + 1, c, 'bt'); else if (side === 'bl') ap(r, c - 1, 'br'); else ap(r, c + 1, 'bl');
}
function bordersSel(type) {
  const v = bdCss(), R = sel;
  const edge = (side, rr1, cc1, rr2, cc2, val) => { for (let r = rr1; r <= rr2; r++) for (let c = cc1; c <= cc2; c++) setSide(r, c, side, val); };
  const outside = val => { edge('bt', R.r1, R.c1, R.r1, R.c2, val); edge('bb', R.r2, R.c1, R.r2, R.c2, val); edge('bl', R.r1, R.c1, R.r2, R.c1, val); edge('br', R.r1, R.c2, R.r2, R.c2, val); };
  if (type === 'none') { for (let r = R.r1 - 1; r <= R.r2 + 1; r++) for (let c = R.c1 - 1; c <= R.c2 + 1; c++) { if (r < 1 || c < 1) continue; const kk = key(r, c), s = sh().style[kk]; if (!s) continue; const inR = r >= R.r1 && r <= R.r2 && c >= R.c1 && c <= R.c2; if (inR) { delete s.bt; delete s.bb; delete s.bl; delete s.br; delete s.bd; } else { if (r < R.r1) delete s.bb; if (r > R.r2) delete s.bt; if (c < R.c1) delete s.br; if (c > R.c2) delete s.bl; } if (!Object.keys(s).length) delete sh().style[kk]; } }
  else if (type === 'all') { for (let r = R.r1; r <= R.r2; r++) for (let c = R.c1; c <= R.c2; c++) { setSide(r, c, 'bt', v); setSide(r, c, 'bb', v); setSide(r, c, 'bl', v); setSide(r, c, 'br', v); } }
  else if (type === 'outside') outside(v);
  else if (type === 'thick') { const tv = '3px solid ' + ($('bdCol').value || '#000000'); outside(tv); }
  else if (type === 'inside') { for (let r = R.r1; r <= R.r2; r++) for (let c = R.c1; c <= R.c2; c++) { if (r > R.r1) setSide(r, c, 'bt', v); if (c > R.c1) setSide(r, c, 'bl', v); } }
  else if (type === 'top') edge('bt', R.r1, R.c1, R.r1, R.c2, v);
  else if (type === 'bottom') edge('bb', R.r2, R.c1, R.r2, R.c2, v);
  else if (type === 'left') edge('bl', R.r1, R.c1, R.r2, R.c1, v);
  else if (type === 'right') edge('br', R.r1, R.c2, R.r2, R.c2, v);
  else if (type === 'dbl') edge('bb', R.r2, R.c1, R.r2, R.c2, '3px double ' + ($('bdCol').value || '#000000'));
  else if (type === 'topbottom') { edge('bt', R.r1, R.c1, R.r1, R.c2, v); edge('bb', R.r2, R.c1, R.r2, R.c2, '3px double ' + ($('bdCol').value || '#000000')); }
  save(); refresh();
}

/* ---------------- clear / paste special / painter ---------------- */
function clearMenu(m) {
  if (m === 'contents') { clearSel(); return; }
  eachSel((r, c) => { const kk = key(r, c); if (m === 'all' || m === 'formats') { delete sh().style[kk]; delete sh().fmt[kk]; } if (m === 'all') wb.setRaw(cur, r, c, ''); });
  if (m === 'all' || m === 'formats') { sh().merges = (sh().merges || []).filter(g => g.r2 < sel.r1 || g.r1 > sel.r2 || g.c2 < sel.c1 || g.c1 > sel.c2); }
  save(); refresh();
}
function pasteSpecial(mode) {
  if (!clip) { $('stL').textContent = 'Nothing copied'; return; }
  if (mode === 'all') { paste(); return; }
  for (const x of clip.cells) {
    const r = anchor.r + x.r, c = anchor.c + x.c; if (r > ROWS || c > COLS) continue; const kk = key(r, c);
    if (mode === 'values') { const v = wb.value(clip.sheet, x.src[0], x.src[1]); wb.setRaw(cur, r, c, v === null || v === undefined ? '' : (E.isErr(v) ? '' : v)); }
    else if (mode === 'formulas') { let raw = x.raw; if (typeof raw === 'string' && raw[0] === '=') raw = shiftFormula(raw, anchor.r - clip.s.r1, anchor.c - clip.s.c1); wb.setRaw(cur, r, c, raw); }
    else if (mode === 'formats') { if (x.st) sh().style[kk] = Object.assign({}, x.st); else delete sh().style[kk]; if (x.f) sh().fmt[kk] = x.f; else delete sh().fmt[kk]; }
    else if (mode === 'transpose') { const rr = anchor.r + x.c, cc = anchor.c + x.r; if (rr <= ROWS && cc <= COLS) { let raw = x.raw; if (typeof raw === 'string' && raw[0] === '=') raw = shiftFormula(raw, rr - x.src[0], cc - x.src[1]); wb.setRaw(cur, rr, cc, raw); if (x.st) sh().style[key(rr, cc)] = Object.assign({}, x.st); if (x.f) sh().fmt[key(rr, cc)] = x.f; } }
  }
  save(); refresh();
}
let PAINT = null;
function painter() {
  if (PAINT) { PAINT = null; $('stL').textContent = 'Ready'; return; }
  PAINT = { r: anchor.r, c: anchor.c, st: Object.assign({}, styleOf(anchor.r, anchor.c) || {}), f: sh().fmt[key(anchor.r, anchor.c)] };
  $('stL').textContent = 'Format Painter: select the cells to format';
}
window.addEventListener('pointerup', () => { setTimeout(() => { if (!PAINT) return; if (sel.r1 === PAINT.r && sel.c1 === PAINT.c && sel.r2 === PAINT.r && sel.c2 === PAINT.c) return; const P = PAINT; PAINT = null; selStyle(s => { for (const k in s) delete s[k]; Object.assign(s, P.st); }); if (P.f) nfmt(P.f); $('stL').textContent = 'Ready'; }, 0); });

/* ---------------- cell sizes ---------------- */
function askNum(title, def, cb) { RB.ask(title, String(def), t => { const n = parseFloat(t); if (isFinite(n) && n >= 0) cb(n); }); }
function colWidthDlg() { askNum('Column width (characters)', Math.round((colW(sel.c1) - 5) / 7 * 10) / 10, n => { for (let c = sel.c1; c <= sel.c2; c++) sh().colW[c] = Math.round(n * 7 + 5); save(); build(); }); }
function rowHeightDlg() { askNum('Row height (points)', Math.round(rowHt(sel.r1) / PT * 10) / 10, n => { for (let r = sel.r1; r <= sel.r2; r++) sh().rowH[r] = Math.round(n * PT); save(); build(); }); }
const _cx = document.createElement('canvas').getContext('2d');
function fontOf(st) { const S = sh(); return (st && st.b ? 'bold ' : '') + (st && st.i ? 'italic ' : '') + ((st && st.sz) || S.defSz || 14.67) + 'px ' + (famCss((st && st.font) || S.defFont) || 'sans-serif'); }
function autoFitCol(only) {
  const lo = typeof only === 'number' ? only : sel.c1, hi = typeof only === 'number' ? only : sel.c2;
  for (let c = lo; c <= hi; c++) { let w = 30; for (let r = 1; r <= ROWS; r++) { if (MM[key(r, c)]) continue; const t = cellText(r, c); if (!t) continue; const st = styleOf(r, c); if (st && st.wrap) continue; _cx.font = fontOf(st); w = Math.max(w, _cx.measureText(t).width + 14 + (st && st.indent ? st.indent * 9 : 0)); } sh().colW[c] = Math.ceil(w); }
  save(); build();
}
function wrapLines(txt, font, w) { _cx.font = font; let lines = 0; for (const para of String(txt).split('\n')) { let line = '', n = 1; for (const word of para.split(' ')) { const t = line ? line + ' ' + word : word; if (_cx.measureText(t).width > w && line) { n++; line = word; } else line = t; } lines += n; } return lines; }
function autoFitRow(quiet, only) {
  const r1 = typeof only === 'number' ? only : sel.r1, r2 = typeof only === 'number' ? only : sel.r2;
  for (let r = r1; r <= r2; r++) { let h = Math.round((sh().defSz || 14.67) * 1.5); for (let c = 1; c <= COLS; c++) { const mm = MM[key(r, c)]; if (mm === 'h') continue; const t = cellText(r, c); if (!t) continue; const st = styleOf(r, c) || {}; const fs = st.sz || sh().defSz || 14.67; let w = 0; for (let k = c; k < c + (mm ? mm.cs : 1); k++) w += colW(k); if (st.wrap) { const n = wrapLines(t, fontOf(st), w - 10); h = Math.max(h, Math.ceil(n * fs * 1.2 + 6)); } else h = Math.max(h, Math.ceil(fs * 1.5)); } sh().rowH[r] = h; }
  save(); build();
}
function hideRC(a) { if (a === 'r') for (let r = sel.r1; r <= sel.r2; r++) { (sh().hid = sh().hid || {}); sh().hid['r' + r] = sh().rowH[r] === undefined ? DH : sh().rowH[r]; sh().rowH[r] = 0; } else for (let c = sel.c1; c <= sel.c2; c++) { (sh().hid = sh().hid || {}); sh().hid['c' + c] = colW(c); sh().colW[c] = 0; } save(); build(); }
function unhideRC(a) {
  const S = sh(); const lo = a === 'r' ? Math.max(1, sel.r1 - 1) : Math.max(1, sel.c1 - 1), hi = a === 'r' ? sel.r2 + 1 : sel.c2 + 1;
  for (let i = lo; i <= hi; i++) { if (a === 'r' && S.rowH[i] === 0) S.rowH[i] = (S.hid && S.hid['r' + i]) || (S.defSz !== undefined ? 20 : DH); if (a === 'c' && S.colW[i] === 0) S.colW[i] = (S.hid && S.hid['c' + i]) || DW; }
  save(); build();
}
function resetSizes() { const S = sh(); for (let r = sel.r1; r <= sel.r2; r++) delete S.rowH[r]; for (let c = sel.c1; c <= sel.c2; c++) delete S.colW[c]; if (S.defSz !== undefined) for (let r = sel.r1; r <= sel.r2; r++) S.rowH[r] = 20; save(); build(); }

/* ---------------- cell styles / table ---------------- */
const CSTYLES = {
  normal: null, good: { bg: '#C6EFCE', color: '#006100' }, bad: { bg: '#FFC7CE', color: '#9C0006' }, neutral: { bg: '#FFEB9C', color: '#9C5700' },
  title: { b: 1, sz: 26.7, color: '#44546A' }, h1: { b: 1, sz: 20, color: '#44546A', bb: '3px solid #4472C4' }, h2: { b: 1, sz: 16, color: '#44546A', bb: '3px solid #A9BCE2' },
  total: { b: 1, bt: '1px solid #4472C4', bb: '3px double #4472C4' }, input: { bg: '#FFCC99', color: '#3F3F76', bt: '1px solid #7F7F7F', bb: '1px solid #7F7F7F', bl: '1px solid #7F7F7F', br: '1px solid #7F7F7F' },
  note: { bg: '#FFFFCC', bt: '1px solid #B2B2B2', bb: '1px solid #B2B2B2', bl: '1px solid #B2B2B2', br: '1px solid #B2B2B2' }, accent: { bg: '#4472C4', color: '#FFFFFF' }, accent2: { bg: '#ED7D31', color: '#FFFFFF' }, dark: { bg: '#44546A', color: '#FFFFFF', b: 1 }
};
function cellStyle(n) { const P = CSTYLES[n]; selStyle(s => { for (const k in s) delete s[k]; if (P) Object.assign(s, P); }); }
function fmtTable() {
  if (sel.r1 === sel.r2) { alert('Select the table range first (header row + data).'); return; }
  for (let r = sel.r1; r <= sel.r2; r++) for (let c = sel.c1; c <= sel.c2; c++) {
    const kk = key(r, c), s = sh().style[kk] = {}; if (r === sel.r1) { s.bg = '#4472C4'; s.color = '#FFFFFF'; s.b = 1; } else s.bg = (r - sel.r1) % 2 ? '#D9E1F2' : '#FFFFFF';
  }
  save(); refresh();
}

/* ---------------- filter ---------------- */
function filterDlg() {
  const hdr = []; for (let c = sel.c1; c <= sel.c2; c++) hdr.push(`<option value="${c}">${E.numToCol(c)}: ${String(wb.value(cur, sel.r1, c) ?? '').replace(/</g, '&lt;').slice(0, 24)}</option>`);
  gen('Filter', `<p style="font-size:12px;color:#555">First row of the selection is the header. Rows that don't match are hidden.</p><label>Column</label><select id="flC">${hdr.join('')}</select><label>Condition</label><select id="flT"><option value="contains">contains</option><option value="eq">equals</option><option value="ne">does not equal</option><option value="gt">greater than</option><option value="lt">less than</option><option value="begins">begins with</option><option value="blank">is blank</option><option value="nonblank">is not blank</option></select><label>Value</label><input type="text" id="flV">`, [['Apply', () => {
    const S = sh(), c = +$('flC').value, t = $('flT').value, val = $('flV').value.toLowerCase(), vn = parseFloat(val);
    S.flt = S.flt || {}; for (let r = sel.r1 + 1; r <= sel.r2; r++) { if (S.flt[r] === undefined) S.flt[r] = S.rowH[r] === undefined ? (S.defSz !== undefined ? 20 : DH) : S.rowH[r]; const raw = wb.value(cur, r, c), tx = String(raw ?? '').toLowerCase(), nv = typeof raw === 'number' ? raw : parseFloat(tx); let ok = true; if (t === 'contains') ok = tx.includes(val); else if (t === 'eq') ok = tx === val; else if (t === 'ne') ok = tx !== val; else if (t === 'gt') ok = nv > vn; else if (t === 'lt') ok = nv < vn; else if (t === 'begins') ok = tx.startsWith(val); else if (t === 'blank') ok = tx === ''; else if (t === 'nonblank') ok = tx !== ''; S.rowH[r] = ok ? S.flt[r] : 0; }
    save(); closeGen(); build();
  }], ['Clear filter', () => { clearFilter(); closeGen(); }]]);
}
function clearFilter() { const S = sh(); if (S.flt) for (const r in S.flt) S.rowH[r] = S.flt[r]; S.flt = {}; save(); build(); }

/* ---------------- text to columns / misc ---------------- */
function textToCols() { RB.ask('Delimiter (e.g. , ; | or tab)', ',', d => { if (d === null || d === '') return; if (/^tab$/i.test(d)) d = '\t'; for (let r = sel.r1; r <= sel.r2; r++) { const t = String(wb.getRaw(cur, r, sel.c1) ?? ''); t.split(d).forEach((p, i) => { if (sel.c1 + i <= COLS) wb.setRaw(cur, r, sel.c1 + i, p.trim()); }); } save(); refresh(); }); }
function selectAll() { const prev = Object.assign({}, sel); anchor = { r: 1, c: 1 }; sel = { r1: 1, c1: 1, r2: ROWS, c2: COLS }; paintSel(prev); }
function freezeMenu(m) {
  const S = sh(); if (m === 'none') S.frz = null; else if (m === 'row') S.frz = { r: 1, c: 0 }; else if (m === 'col') S.frz = { r: 0, c: 1 }; else S.frz = { r: Math.max(0, anchor.r - 1), c: Math.max(0, anchor.c - 1) };
  save(); refresh();
}
function xFrz(r, c) {
  const F = sh().frz; if (!F) return ''; const fr = r <= F.r, fc = c <= F.c; if (!fr && !fc) return '';
  let css = 'position:sticky;z-index:' + (fr && fc ? 5 : 2) + ';';
  if (fr) { let t = 28; for (let i = 1; i < r; i++) t += rowHt(i); css += 'top:' + t + 'px;'; }
  if (fc) { let l = 38; for (let i = 1; i < c; i++) l += colW(i); css += 'left:' + l + 'px;'; }
  return css + 'background-clip:padding-box;';
}
function toggleGrid() { const S = sh(); S.grid = S.grid === false ? true : false; save(); $('grid').classList.toggle('nogrid', S.grid === false); }

/* ---------------- conditional formatting ---------------- */
let CFC = {};
const CFCOL = { red: ['#FFC7CE', '#9C0006'], yellow: ['#FFEB9C', '#9C5700'], green: ['#C6EFCE', '#006100'], blue: ['#BDD7EE', '#1F4E78'], orange: ['#F8CBAD', '#833C0B'], gray: ['#D9D9D9', '#262626'] };
function cfStats(R) {
  if (CFC[R.id]) return CFC[R.id]; const nums = [], cnt = {};
  for (let r = R.r1; r <= R.r2; r++) for (let c = R.c1; c <= R.c2; c++) { const v = wb.value(cur, r, c); if (typeof v === 'number') nums.push(v); if (v !== null && v !== undefined && v !== '') { const k = String(v).toLowerCase(); cnt[k] = (cnt[k] || 0) + 1; } }
  nums.sort((a, b) => a - b); const st = { nums, cnt, min: nums[0], max: nums[nums.length - 1], avg: nums.length ? nums.reduce((a, b) => a + b, 0) / nums.length : 0 }; return CFC[R.id] = st;
}
function cmp(v, a) { const na = parseFloat(a); if (typeof v === 'number' && isFinite(na)) return v - na; return String(v ?? '').toLowerCase() < String(a).toLowerCase() ? -1 : String(v ?? '').toLowerCase() > String(a).toLowerCase() ? 1 : 0; }
function cfHit(R, r, c, v) {
  const t = R.type, a = R.a, b = R.b, tx = String(v ?? '').toLowerCase();
  switch (t) {
    case 'gt': return v !== null && v !== '' && cmp(v, a) > 0; case 'lt': return v !== null && v !== '' && cmp(v, a) < 0;
    case 'ge': return v !== null && v !== '' && cmp(v, a) >= 0; case 'le': return v !== null && v !== '' && cmp(v, a) <= 0;
    case 'eq': return v !== null && v !== '' && cmp(v, a) === 0; case 'ne': return !(v !== null && v !== '' && cmp(v, a) === 0);
    case 'between': return typeof v === 'number' && v >= Math.min(+a, +b) && v <= Math.max(+a, +b);
    case 'text': return tx.includes(String(a).toLowerCase()) && tx !== '';
    case 'begins': return tx.startsWith(String(a).toLowerCase()) && tx !== '';
    case 'blank': return tx === ''; case 'nonblank': return tx !== '';
    case 'dup': return tx !== '' && cfStats(R).cnt[tx] > 1; case 'uniq': return tx !== '' && cfStats(R).cnt[tx] === 1;
    case 'top': { const n = cfStats(R).nums; return typeof v === 'number' && n.length && v >= n[Math.max(0, n.length - Math.max(1, +a || 10))]; }
    case 'bottom': { const n = cfStats(R).nums; return typeof v === 'number' && n.length && v <= n[Math.min(n.length - 1, Math.max(1, +a || 10) - 1)]; }
    case 'above': return typeof v === 'number' && v > cfStats(R).avg; case 'below': return typeof v === 'number' && v < cfStats(R).avg;
  }
  return false;
}
function lerp(c1, c2, t) { const p = h => [1, 3, 5].map(i => parseInt(h.substr(i, 2), 16)); const A = p(c1), B = p(c2); return '#' + A.map((x, i) => Math.round(x + (B[i] - x) * t).toString(16).padStart(2, '0')).join(''); }
function xCF(r, c, st, v) {
  const L = sh().cf; if (!L || !L.length) return st; let out = st;
  for (const R of L) {
    if (r < R.r1 || r > R.r2 || c < R.c1 || c > R.c2) continue;
    if (R.type === 'scale') { if (typeof v !== 'number') continue; const S = cfStats(R); const t = S.max === S.min ? .5 : (v - S.min) / (S.max - S.min); const col = R.pal === 'rg' ? (t < .5 ? lerp('#63BE7B', '#FFEB84', t * 2) : lerp('#FFEB84', '#F8696B', (t - .5) * 2)) : (t < .5 ? lerp('#F8696B', '#FFEB84', t * 2) : lerp('#FFEB84', '#63BE7B', (t - .5) * 2)); out = Object.assign({}, out || {}, { bg: col }); continue; }
    if (R.type === 'bar') { if (typeof v !== 'number') continue; const S = cfStats(R); const t = S.max === S.min ? 1 : Math.max(0, (v - Math.min(0, S.min)) / (S.max - Math.min(0, S.min))); out = Object.assign({}, out || {}, { bg: 'linear-gradient(90deg,#9DC3E6 ' + Math.round(t * 100) + '%,transparent ' + Math.round(t * 100) + '%)' }); continue; }
    if (cfHit(R, r, c, v)) out = Object.assign({}, out || {}, { bg: R.bg, color: R.color });
  }
  return out;
}
function cfDlg() {
  const opt = (v, t) => `<option value="${v}">${t}</option>`;
  gen('Conditional Formatting', `<label>Rule</label><select id="cfT">${opt('gt', 'Cell value greater than')}${opt('lt', 'Cell value less than')}${opt('ge', 'Greater than or equal to')}${opt('le', 'Less than or equal to')}${opt('eq', 'Equal to')}${opt('ne', 'Not equal to')}${opt('between', 'Between')}${opt('text', 'Text contains')}${opt('begins', 'Text begins with')}${opt('dup', 'Duplicate values')}${opt('uniq', 'Unique values')}${opt('top', 'Top N values')}${opt('bottom', 'Bottom N values')}${opt('above', 'Above average')}${opt('below', 'Below average')}${opt('blank', 'Blank cells')}${opt('nonblank', 'Non-blank cells')}${opt('scale', 'Color scale (red-yellow-green)')}${opt('scale2', 'Color scale (green-yellow-red)')}${opt('bar', 'Data bars')}</select><label>Value (or N)</label><input type="text" id="cfA"><label>and (for Between)</label><input type="text" id="cfB"><label>Format with</label><select id="cfC">${opt('red', 'Light red fill, dark red text')}${opt('yellow', 'Yellow fill, dark yellow text')}${opt('green', 'Green fill, dark green text')}${opt('blue', 'Blue fill')}${opt('orange', 'Orange fill')}${opt('gray', 'Gray fill')}</select><p style="font-size:12px;color:#555">Applies to the selected range ${addrOf(sel)}.</p>`, [['Apply', () => {
    const S = sh(); S.cf = S.cf || []; let t = $('cfT').value, pal; if (t === 'scale2') { t = 'scale'; pal = 'rg'; }
    const col = CFCOL[$('cfC').value]; S.cf.push({ id: 'cf' + Date.now() + S.cf.length, type: t, pal, a: $('cfA').value, b: $('cfB').value, bg: col[0], color: col[1], r1: sel.r1, c1: sel.c1, r2: sel.r2, c2: sel.c2 }); save(); closeGen(); refresh();
  }], ['Clear in selection', () => { const S = sh(); S.cf = (S.cf || []).filter(R => R.r2 < sel.r1 || R.r1 > sel.r2 || R.c2 < sel.c1 || R.c1 > sel.c2); save(); closeGen(); refresh(); }], ['Clear all', () => { sh().cf = []; save(); closeGen(); refresh(); }]]);
}

/* ---------------- Format Cells dialog (Ctrl+1) ---------------- */
function formatCellsDlg() {
  const st = styleOf(anchor.r, anchor.c) || {}, f = sh().fmt[key(anchor.r, anchor.c)] || '';
  const sel1 = (id, opts, v) => `<select id="${id}">${opts.map(o => `<option value="${o[0]}" ${String(o[0]) === String(v) ? 'selected' : ''}>${o[1]}</option>`).join('')}</select>`;
  gen('Format Cells', `<label>Number format (code)</label><input type="text" id="fcN" value="${f.replace(/"/g, '&quot;')}" placeholder="General"><div style="font-size:11px;color:#666">e.g. 0.00 &nbsp; #,##0 &nbsp; $#,##0.00 &nbsp; 0% &nbsp; mm/dd/yyyy &nbsp; 0.00E+00 &nbsp; # ?/?</div>
  <label>Horizontal</label>${sel1('fcH', [['', 'General'], ['left', 'Left'], ['center', 'Center'], ['right', 'Right']], st.align || '')}<label>Vertical</label>${sel1('fcV', [['', 'Bottom/Default'], ['top', 'Top'], ['middle', 'Middle'], ['bottom', 'Bottom']], st.valign || '')}
  <label>Indent</label><input type="number" id="fcI" min="0" max="15" value="${st.indent || 0}"><label><input type="checkbox" id="fcW" ${st.wrap ? 'checked' : ''}> Wrap text</label> <label><input type="checkbox" id="fcM"> Merge cells</label>
  <label>Font</label>${sel1('fcF', ['Calibri', 'Arial', 'Times New Roman', 'Courier New', 'Verdana', 'Georgia', 'Tahoma', 'Candara'].map(x => [x, x]), st.font || sh().defFont || 'Calibri')}<label>Size (pt)</label><input type="number" id="fcS" min="1" max="200" value="${curPt()}">
  <label><input type="checkbox" id="fcB" ${st.b ? 'checked' : ''}> Bold</label> <label><input type="checkbox" id="fcT" ${st.i ? 'checked' : ''}> Italic</label> <label><input type="checkbox" id="fcU" ${st.u ? 'checked' : ''}> Underline</label> <label><input type="checkbox" id="fcX" ${st.s ? 'checked' : ''}> Strikethrough</label>
  <label>Font color</label><input type="color" id="fcC" value="${st.color || '#000000'}"><label>Fill color</label><input type="color" id="fcG" value="${st.bg && st.bg[0] === '#' ? st.bg : '#ffffff'}"> <label><input type="checkbox" id="fcNF" ${st.bg ? '' : 'checked'}> No fill</label>`, [['OK', () => {
    const nf = $('fcN').value.trim();
    eachSel((r, c) => { const kk = key(r, c); if (nf && nf !== 'General') sh().fmt[kk] = nf; else delete sh().fmt[kk]; });
    selStyle(s => { const set = (k, v) => { if (v) s[k] = v; else delete s[k]; }; set('align', $('fcH').value); set('valign', $('fcV').value); set('indent', +$('fcI').value); set('wrap', $('fcW').checked ? 1 : 0); set('font', $('fcF').value); s.sz = Math.round(+$('fcS').value * PT * 10) / 10; set('b', $('fcB').checked ? 1 : 0); set('i', $('fcT').checked ? 1 : 0); set('u', $('fcU').checked ? 1 : 0); set('s', $('fcX').checked ? 1 : 0); s.color = $('fcC').value === '#000000' ? undefined : $('fcC').value; if ($('fcNF').checked) delete s.bg; else s.bg = $('fcG').value; if (!s.color) delete s.color; });
    if ($('fcM').checked) mergeSel('merge'); closeGen();
  }]]);
}

/* ---------------- number format shortcuts ---------------- */
function nfSet(f) { nfmt(f); }

/* ---------------- keyboard shortcuts ---------------- */
let FC = null; // moving corner while extending a selection
function setSel(r1, c1, r2, c2, ar, ac) { const prev = Object.assign({}, sel); if (ar) anchor = { r: ar, c: ac }; sel = { r1: Math.min(r1, r2), c1: Math.min(c1, c2), r2: Math.max(r1, r2), c2: Math.max(c1, c2) }; paintSel(prev); const td = tds[(FC || anchor).r] && tds[(FC || anchor).r][(FC || anchor).c]; if (td) td.scrollIntoView({ block: 'nearest', inline: 'nearest' }); }
function goCell(r, c) { r = Math.max(1, Math.min(ROWS, r)); c = Math.max(1, Math.min(COLS, c)); FC = null; setSel(r, c, r, c, r, c); }
function nonEmpty(r, c) { const v = wb.value(cur, r, c); return v !== null && v !== undefined && v !== ''; }
function jumpEdge(r, c, dr, dc) {
  const inb = (a, b) => a >= 1 && a <= ROWS && b >= 1 && b <= COLS; let nr = r + dr, nc = c + dc; if (!inb(nr, nc)) return { r, c };
  if (nonEmpty(r, c) && nonEmpty(nr, nc)) { while (inb(nr + dr, nc + dc) && nonEmpty(nr + dr, nc + dc)) { nr += dr; nc += dc; } return { r: nr, c: nc }; }
  while (inb(nr, nc) && !nonEmpty(nr, nc)) { nr += dr; nc += dc; } if (!inb(nr, nc)) { nr -= dr; nc -= dc; } return { r: nr, c: nc };
}
function moveKey(e, dr, dc) {
  const ctrl = e.ctrlKey || e.metaKey || CTRLH(); const cur0 = (e.shiftKey && FC && ((FC.r === sel.r1 || FC.r === sel.r2) && (FC.c === sel.c1 || FC.c === sel.c2))) ? FC : (e.shiftKey ? { r: anchor.r === sel.r1 ? sel.r2 : sel.r1, c: anchor.c === sel.c1 ? sel.c2 : sel.c1 } : { r: anchor.r, c: anchor.c });
  let n = ctrl ? jumpEdge(cur0.r, cur0.c, dr, dc) : { r: Math.max(1, Math.min(ROWS, cur0.r + dr)), c: Math.max(1, Math.min(COLS, cur0.c + dc)) };
  if (e.shiftKey) { FC = n; setSel(anchor.r, anchor.c, n.r, n.c); } else goCell(n.r, n.c);
}
let enterMode = false;
function startEdit(text, enter) { const f = $('fx'); f.focus(); if (text !== undefined) f.value = text; const n = f.value.length; try { f.setSelectionRange(n, n); } catch (e) { } editing = true; editCell = { sheet: cur, r: anchor.r, c: anchor.c }; enterMode = !!enter; liveCell(); }
function cycleAbs() {
  const f = $('fx'); const v = f.value, p = f.selectionStart; const re = /\$?[A-Z]{1,3}\$?\d+/gi; let m, hit = null; while ((m = re.exec(v))) { if (p >= m.index && p <= m.index + m[0].length) { hit = m; break; } } if (!hit) return;
  const mm = /(\$?)([A-Z]{1,3})(\$?)(\d+)/i.exec(hit[0]); const st = (mm[1] ? 1 : 0) + (mm[3] ? 2 : 0); const nx = [[1, 1], [0, 1], [1, 0], [0, 0]]; const cyc = { 3: [0, 1], 2: [0, 0], 1: [1, 1], 0: [1, 0] }; // $A$1 -> A$1 -> $A1 -> A1 -> $A$1
  const order = [[1, 1], [0, 1], [1, 0], [0, 0]]; const idx = order.findIndex(o => o[0] === (mm[1] ? 1 : 0) && o[1] === (mm[3] ? 1 : 0)); const o = order[(idx + 1) % 4];
  const rep = (o[0] ? '$' : '') + mm[2] + (o[1] ? '$' : '') + mm[4]; f.value = v.slice(0, hit.index) + rep + v.slice(hit.index + hit[0].length); const np = hit.index + rep.length; f.setSelectionRange(np, np); guide();
}
function tsvOfSel() { const rows = []; for (let r = sel.r1; r <= Math.min(sel.r2, ROWS); r++) { if (visOnly && sh().rowH[r] === 0) continue; const row = []; for (let c = sel.c1; c <= Math.min(sel.c2, COLS); c++) { if (visOnly && sh().colW[c] === 0) continue; row.push(cellText(r, c)); } rows.push(row.join('\t')); } visOnly = false; return rows.join('\n'); }
function sysCopy(t) { try { const ta = document.createElement('textarea'); ta.value = t; ta.style.cssText = 'position:fixed;left:-9999px;top:0;opacity:0'; document.body.appendChild(ta); ta.select(); document.execCommand('copy'); ta.remove(); } catch (e) { } try { navigator.clipboard && navigator.clipboard.writeText(t); } catch (e) { } }
const _copy0 = copy;
copy = function (cut) { _copy0(cut); if (clip) { clip.tsv = tsvOfSel(); sysCopy(clip.tsv); } };
let PSD = false;
function pasteSpecialDlg() {
  if (!clip) { $('stL').textContent = 'Copy something first (Ctrl+C), then Paste Special'; return; }
  const it = (k, t, f) => `<div style="padding:11px 6px;border-bottom:1px solid #eee;font-size:15px" onclick="PSD=false;closeGen();pasteSpecial('${f}')"><u><b>${k}</b></u> &nbsp;${t}</div>`;
  gen('Paste Special', it('A', 'All', 'all') + it('F', 'Formulas', 'formulas') + it('V', 'Values', 'values') + it('T', 'Formats', 'formats') + it('E', 'Transpose', 'transpose') + '<p style="font-size:12px;color:#666">Press the underlined letter, or tap an option. Esc to cancel.</p>', []); PSD = true;
}
function pasteValuesNow() { if (clip) { pasteSpecial('values'); return; } try { navigator.clipboard.readText().then(t => { if (t) pasteText(t); }); } catch (e) { } }
function pasteText(t) {
  const rows = t.replace(/\r/g, '').replace(/\n$/, '').split('\n').map(l => l.split('\t')); if (!rows.length) return;
  rows.forEach((row, i) => row.forEach((x, j) => { const r = anchor.r + i, c = anchor.c + j; if (r <= ROWS && c <= COLS) wb.setRaw(cur, r, c, x); }));
  setSel(anchor.r, anchor.c, Math.min(ROWS, anchor.r + rows.length - 1), Math.min(COLS, anchor.c + Math.max(...rows.map(r => r.length)) - 1)); save(); refresh();
}
let gotPaste = false;
document.addEventListener('paste', e => {
  if (modalOpen() || /INPUT|TEXTAREA|SELECT/.test(document.activeElement.tagName)) return; gotPaste = true;
  const t = e.clipboardData && e.clipboardData.getData('text/plain'); e.preventDefault();
  if (clip && (!t || t === clip.tsv)) paste(); else if (t) pasteText(t);
});
function modalOpen() { for (const id of ['mGen', 'mUF', 'mVba', 'mRun']) { const e = $(id); if (e && e.style.display === 'flex') return true; } return false; }
function insertNow(kind) { const d = new Date(); const ser = Math.floor(E.dateToSerial(d)); if (kind === 'time') { const fr = (d.getHours() * 3600 + d.getMinutes() * 60 + d.getSeconds()) / 86400; wb.setRaw(cur, anchor.r, anchor.c, ser + fr); sh().fmt[key(anchor.r, anchor.c)] = 'h:mm AM/PM'; } else { wb.setRaw(cur, anchor.r, anchor.c, ser); sh().fmt[key(anchor.r, anchor.c)] = 'm/d/yyyy'; } save(); refresh(); }
function selectRowCol(mode) { if (mode === 'row') setSel(anchor.r, 1, anchor.r, COLS, anchor.r, anchor.c); else setSel(1, anchor.c, ROWS, anchor.c, anchor.r, anchor.c); }
function shortcutsHelp() {
  const L = [['Ctrl+C / X / V', 'Copy / Cut / Paste'], ['Ctrl+T', 'Create Table'], ['Alt+F1', 'Insert chart from selection (double-tap a chart for Chart Design)'], ['Alt+F5', 'Refresh PivotTable'], ['Ctrl+[ / Ctrl+]', 'Trace precedents / dependents'], ['Alt+Shift+→ / ←', 'Group / Ungroup rows or columns'], ['Ctrl+wheel', 'Zoom'], ['Enter / Tab in a selection', 'Move within the selected range'], ['Right-click status bar', 'Choose Sum / Average / Count / Min / Max'], ['Ctrl+P', 'Print (uses Page Setup)'], ['Ctrl+Alt+V', 'Paste Special (then V=Values, F=Formulas, T=Formats, E=Transpose)'], ['Ctrl+Shift+V', 'Paste values only'], ['Ctrl+Z / Y', 'Undo / Redo'], ['Ctrl+B / I / U', 'Bold / Italic / Underline'], ['Ctrl+5', 'Strikethrough'], ['Ctrl+A', 'Select all'], ['Ctrl+D / R', 'Fill down / right'], ['Ctrl+F / H', 'Find & Replace'], ['Ctrl+S', 'Save'], ['Ctrl+O', 'Open'], ['Ctrl+1', 'Format Cells'], ['Ctrl+Shift+1 / 2 / 3', 'Number / Time / Date format'], ['Ctrl+Shift+4 / 5', 'Currency / Percent'], ['Ctrl+; / Ctrl+Shift+;', 'Insert date / time'], ['Ctrl++ / Ctrl+-', 'Insert / delete row'], ['Ctrl+9 / Ctrl+0', 'Hide row / column'], ['Ctrl+Arrow', 'Jump to edge of data'], ['Ctrl+Home / End', 'First cell / last used cell'], ['Shift+Arrow', 'Extend selection'], ['Ctrl+Shift+Arrow', 'Extend to edge of data'], ['Ctrl+Space / Shift+Space', 'Select column / row'], ['Tab / Shift+Tab', 'Move right / left'], ['Enter / Shift+Enter', 'Move down / up'], ['Page Up / Down', 'Scroll one screen'], ['F2', 'Edit cell'], ['F4', 'Cycle $ in reference (editing)'], ['F5 / Ctrl+G', 'Go to'], ['Alt+=', 'AutoSum'], ['Ctrl+Enter', 'Fill selection with entry'], ['Delete', 'Clear contents'], ['Esc', 'Cancel'],['Ctrl+PgUp / PgDn','Previous / next sheet'],['Alt+PgUp / PgDn','Move one screen left / right'],['End, then Arrow','End mode: jump to edge of data'],['Ctrl+A (twice)','Select region, then whole sheet'],['Ctrl+Shift+Space / Ctrl+*','Select current region'],['Shift+Home, Ctrl+Shift+Home/End','Extend selection to row start / A1 / last cell'],['Alt+;','Select visible cells only (copy skips hidden)'],['Ctrl+.','Move active cell to next corner'],['Ctrl+Backspace','Scroll to active cell'],["Ctrl+' / Ctrl+\"",'Copy formula / value from cell above'],['Ctrl+E','Flash Fill (give an example first)'],['Ctrl+Shift+& / _','Outline border / remove border'],['Ctrl+Shift+L, Alt+Down','Filter'],['Ctrl+T / Ctrl+L','Format as Table'],['Ctrl+`','Show / hide formulas'],['F9 / Shift+F9','Recalculate'],['Shift+F11 / Alt+Shift+F1','New sheet'],['F11 / Alt+F1','Insert chart from selection'],['F12','Save As'],['Ctrl+P','Print'],['Ctrl+N','New workbook'],['Ctrl+W / Ctrl+F4','Close (back to menu)'],['F1','This help'],['Ctrl+F1','Show / hide ribbon'],['Alt, H/N/P/M/A/W/F','Ribbon KeyTips (switch tabs)'],['Alt,H,A,C/L/R','Center / left / right align'],['Alt,H,W','Wrap text'],['Alt,H,M,C','Merge & center'],['Alt,H,0 / 9','Increase / decrease decimals'],['Alt,H,O,I / A','AutoFit column / row'],['Alt,H,O,W / H','Column width / row height'],['Alt,H,I,R / C  or  Alt,I,R / C','Insert row / column'],['Alt,A,V,V','Data Validation'],['Alt,H,D,R / C','Delete row / column'],['Alt,H,O,R','Rename sheet'],['Alt,H,D,S','Delete sheet'],['Alt,A,S,A / D','Sort ascending / descending'],['Alt,A,M','Remove duplicates'],['Alt,A,E','Text to Columns'],['Alt,W,F,F','Freeze / unfreeze panes']];
  gen('Keyboard shortcuts', '<table style="width:100%;font-size:13px;border-collapse:collapse">' + L.map(x => `<tr><td style="padding:3px 6px;border-bottom:1px solid #eee;white-space:nowrap"><b>${x[0]}</b></td><td style="padding:3px 6px;border-bottom:1px solid #eee">${x[1]}</td></tr>`).join('') + '</table>', []);
}
/* ---------------- Excel Keyboard Shortcuts Master Reference additions ---------------- */
/* modifier tracking: some tablet/Bluetooth keyboards deliver the letter key without ctrlKey/altKey set,
   so remember Control/Alt/Meta while they are physically down (or were released a split second ago). */
let _mod = { c: false, a: false, cu: 0, au: 0 };
const CTRLH = () => _mod.c, ALTH = () => _mod.a;
window.addEventListener('keydown', e => { if (e.key === 'Control' || e.key === 'Meta') _mod.c = true; else if (e.key === 'Alt') _mod.a = true; }, true);
window.addEventListener('keyup', e => { if (e.key === 'Control' || e.key === 'Meta') { _mod.c = false; _mod.cu = Date.now(); } else if (e.key === 'Alt') { _mod.a = false; _mod.au = Date.now(); } }, true);
window.addEventListener('blur', () => { _mod.c = _mod.a = false; });
let endMode = false, tipMode = false, tipSeq = '', visOnly = false, ribHidden = false;
function stMsg(t) { try { $('stL').textContent = t; } catch (e) { } }
function regionOf(r, c) {
  let r1 = r, c1 = c, r2 = r, c2 = c, ch = true;
  while (ch) { ch = false;
    for (let rr = Math.max(1, r1 - 1); rr <= Math.min(ROWS, r2 + 1); rr++) for (let cc = Math.max(1, c1 - 1); cc <= Math.min(COLS, c2 + 1); cc++) {
      if ((rr < r1 || rr > r2 || cc < c1 || cc > c2) && nonEmpty(rr, cc)) { r1 = Math.min(r1, rr); r2 = Math.max(r2, rr); c1 = Math.min(c1, cc); c2 = Math.max(c2, cc); ch = true; }
    } }
  return { r1, c1, r2, c2 };
}
function selRegion() { const g = regionOf(anchor.r, anchor.c); setSel(g.r1, g.c1, g.r2, g.c2); }
function ctrlA() { const g = regionOf(anchor.r, anchor.c); const same = sel.r1 === g.r1 && sel.c1 === g.c1 && sel.r2 === g.r2 && sel.c2 === g.c2; const single = sel.r1 === sel.r2 && sel.c1 === sel.c2; if (same || (!single && !(g.r1 === g.r2 && g.c1 === g.c2) && sel.r2 - sel.r1 + 1 >= ROWS)) selectAll(); else if (!single && !same) selectAll(); else { if (g.r1 === g.r2 && g.c1 === g.c2) selectAll(); else selRegion(); } }
function lastUsed() { let mr = 1, mc = 1; for (const kk in sh().cells) { const [r, c] = kk.split(',').map(Number); if (r > mr) mr = r; if (c > mc) mc = c; } return { r: mr, c: mc }; }
function sheetStep(d) { const i = wb.order.indexOf(cur), n = wb.order[(i + d + wb.order.length) % wb.order.length]; if (n && n !== cur) gotoSheet(n); }
function delSheetNow() { if (wb.order.length < 2) { stMsg('A workbook needs at least one sheet'); return; } if (!confirm('Delete sheet ' + cur + '?')) return; const n = cur; wb.removeSheet(n); cur = wb.order[0]; save(); renderTabs(); build(); }
function cornerNext() { const cs = [[sel.r1, sel.c1], [sel.r1, sel.c2], [sel.r2, sel.c2], [sel.r2, sel.c1]]; let i = cs.findIndex(x => x[0] === anchor.r && x[1] === anchor.c); i = (i + 1) % 4; anchor = { r: cs[i][0], c: cs[i][1] }; FC = null; paintSel(sel); const td = tds[anchor.r] && tds[anchor.r][anchor.c]; if (td) td.scrollIntoView({ block: 'nearest', inline: 'nearest' }); }
function copyAbove(asValue) { const r = anchor.r; if (r < 2) return; for (let c = sel.c1; c <= sel.c2; c++) { for (let rr = sel.r1; rr <= sel.r2; rr++) { if (rr < 2) continue; const src = asValue ? wb.value(cur, rr - 1, c) : wb.getRaw(cur, rr - 1, c); wb.setRaw(cur, rr, c, src === null || src === undefined ? '' : src); } } save(); refresh(); }
function flashFill() {
  const c = anchor.c; let r0 = anchor.r; const toks = s => String(s).split(/[\s,;@._\-\/]+/).filter(x => x !== '');
  const cs = [x => x, x => x.toUpperCase(), x => x.toLowerCase(), x => x.charAt(0).toUpperCase() + x.slice(1).toLowerCase(), x => x.charAt(0), x => x.charAt(0).toUpperCase()];
  for (let srcC = c - 1; srcC >= Math.max(1, c - 4); srcC--) {
    const ex = []; for (let r = 1; r < r0 + 40 && r <= ROWS; r++) { const o = wb.getRaw(cur, r, c), s = wb.value(cur, r, srcC); if (o !== '' && o !== null && o !== undefined && s !== '' && s !== null && s !== undefined && r < r0 + 1) ex.push([String(s), String(o)]); }
    if (!ex.length) continue;
    const gens = []; for (let i = 0; i < 6; i++) for (let k = 0; k < cs.length; k++) { gens.push({ n: [i], k: [k] }); gens.push({ n: [-1 - i], k: [k] }); }
    for (let i = 0; i < 4; i++) for (let j = 0; j < 4; j++) if (i !== j) for (let k = 0; k < 4; k++) gens.push({ n: [i, j], k: [k, k] });
    gens.push({ n: [0, -1], k: [4, 3] }); gens.push({ n: [0, 1], k: [4, 3] });
    const run = (g, s, pre, mid, post) => { const t = toks(s); const p = g.n.map((n, ix) => { const x = n < 0 ? t[t.length + n] : t[n]; return x === undefined ? null : cs[g.k[ix]](x); }); if (p.some(x => x === null)) return null; return pre + p[0] + (p.length > 1 ? mid + p[1] : '') + post; };
    for (const g of gens) {
      const e0 = ex[ex.length - 1]; let pre = '', mid = ' ', post = '';
      const base = run(g, e0[0], '', '', ''); if (base === null) continue;
      if (g.n.length === 1) { const ix = e0[1].indexOf(base); if (ix < 0 || !base) continue; pre = e0[1].slice(0, ix); post = e0[1].slice(ix + base.length); }
      else { const t = toks(e0[0]); const a = run({ n: [g.n[0]], k: [g.k[0]] }, e0[0], '', '', ''), b = run({ n: [g.n[1]], k: [g.k[1]] }, e0[0], '', '', ''); if (a === null || b === null) continue; const ia = e0[1].indexOf(a); if (ia < 0) continue; const ib = e0[1].indexOf(b, ia + a.length); if (ib < 0) continue; pre = e0[1].slice(0, ia); mid = e0[1].slice(ia + a.length, ib); post = e0[1].slice(ib + b.length); }
      if (ex.every(x => run(g, x[0], pre, mid, post) === x[1])) {
        let n = 0; for (let r = 1; r <= ROWS; r++) { const s = wb.value(cur, r, srcC), o = wb.getRaw(cur, r, c); if (s === '' || s === null || s === undefined) continue; if (o !== '' && o !== null && o !== undefined) continue; const v = run(g, s, pre, mid, post); if (v !== null) { wb.setRaw(cur, r, c, v); n++; } }
        save(); refresh(); stMsg('Flash Fill: filled ' + n + ' cell(s)'); return true;
      }
    }
  }
  stMsg('Flash Fill: type an example first in the cell next to your data'); return false;
}
const TIPS = {
  HAC: () => sty('align', 'center'), HAL: () => sty('align', 'left'), HAR: () => sty('align', 'right'), HOI: () => autoFitCol(), HOA: () => autoFitRow(),
  HH: () => { const e = document.querySelector('[title*="Fill Color" i],[title*="Fill color" i]'); stMsg('Use Fill Color in the Home ribbon'); if (e) e.click && e.click(); }, HB: () => bordersSel('all'), HW: () => wrapToggle(), HMC: () => mergeSel('center'),
  H0: () => decs(1), H9: () => decs(-1), HIR: () => insRC('r'), HIC: () => insRC('c'), HDR: () => delRC('r'), HDC: () => delRC('c'), HOH: () => rowHeightDlg(), HOW: () => colWidthDlg(),
  IC: () => insRC('c'), IR: () => insRC('r'), AVV: () => dvDlg(), HOR: () => renSheet(cur), HDS: () => delSheetNow(), ASA: () => sortSel(1), ASD: () => sortSel(-1), ASS: () => filterDlg(), AM: () => dedupe(), AE: () => textToCols(), WFF: () => freezeMenu(sh().frz ? 'none' : 'sel'),
  H: 'Home', N: 'Insert', P: 'Page Layout', M: 'Formulas', A: 'Data', R: 'Review', W: 'View', F: 'File', Q: 'Search'
};
function tipTab(t) { if (t === 'File') { openFile(); return; } if (t === 'Review' || t === 'Search') { stMsg('No ' + t + ' tab in this app'); return; } const d = [...document.querySelectorAll('#rtabs div')].find(x => x.textContent === t); if (d) d.click(); }
function tipKey(e) {
  const k = e.key; if (k === 'Escape') { tipMode = false; tipSeq = ''; stMsg('Ready'); return true; }
  if (k.length !== 1 || !/[a-z0-9]/i.test(k) || e.ctrlKey || e.metaKey) { tipMode = false; return false; }
  tipSeq += k.toUpperCase(); const keys = Object.keys(TIPS); const ex = TIPS[tipSeq];
  const more = keys.some(x => x.length > tipSeq.length && x.indexOf(tipSeq) === 0 && typeof TIPS[x] === 'function');
  e.preventDefault(); e.stopPropagation();
  if (typeof ex === 'function' && !more) { tipMode = false; const s = tipSeq; tipSeq = ''; ex(); stMsg('Alt+' + s.split('').join(',')); return true; }
  if (tipSeq.length === 1 && typeof ex === 'string') { if (!more) { tipMode = false; tipSeq = ''; } tipTab(ex); stMsg('Alt,' + k.toUpperCase() + ' ...'); if (!more) return true; return true; }
  if (more || (typeof ex === 'string')) { stMsg('Alt,' + tipSeq.split('').join(',') + ' ...'); return true; }
  tipMode = false; tipSeq = ''; stMsg('Ready'); return true;
}
function xkeys(e, k, kl, ctrl, prevent) {
  const sh_ = e.shiftKey, alt = e.altKey || (ALTH() && k !== 'Alt');
  if (dvMenuOn && dvMenuKey(e)) return true;
  if (tipMode && tipKey(e)) return true;
  if (k === 'Alt') { if (!e.repeat) { tipMode = true; tipSeq = ''; stMsg('Alt: press H Home, N Insert, P Page Layout, M Formulas, A Data, W View, F File'); } prevent(); return true; }
  if (k === 'Shift' || k === 'Control' || k === 'Meta') return false;
  tipMode = false;
  if (alt && !ctrl) {
    if (k === 'PageDown') { prevent(); goCell(anchor.r, anchor.c + 10); return true; } if (k === 'PageUp') { prevent(); goCell(anchor.r, anchor.c - 10); return true; }
    if (k === ';') { prevent(); visOnly = true; stMsg('Visible cells only (copy skips hidden rows/columns)'); return true; }
    if (k === 'ArrowDown') { prevent(); if (dvListAt(anchor.r, anchor.c)) dvOpen(); else filterDlg(); return true; }
    if (k === 'F1' && sh_) { prevent(); addSheet(); return true; } if (k === 'F1') { prevent(); if (sel.r1 !== sel.r2 || sel.c1 !== sel.c2) addChart('bar'); else stMsg('Select the data range first'); return true; }
    if (k === 'F4') { prevent(); const b = window.__back; if (b) b(); return true; }
    if (k === 'F5') { prevent(); refresh(); return true; }
    return false;
  }
  if (ctrl && alt && /^F9$/.test(k)) { prevent(); refresh(); stMsg('Full calculation done'); return true; }
  if (ctrl && !alt) {
    if (k === 'PageDown') { prevent(); sheetStep(1); return true; } if (k === 'PageUp') { prevent(); sheetStep(-1); return true; }
    if (k === 'Backspace') { prevent(); const td = tds[anchor.r] && tds[anchor.r][anchor.c]; if (td) td.scrollIntoView({ block: 'center', inline: 'center' }); return true; }
    if (kl === 'a' && !sh_) { prevent(); ctrlA(); return true; }
    if (k === '*' || (k === ' ' && sh_)) { prevent(); selRegion(); return true; }
    if (k === '.') { prevent(); cornerNext(); return true; }
    if (k === "'") { prevent(); copyAbove(false); return true; } if (k === '"') { prevent(); copyAbove(true); return true; }
    if (kl === 'e') { prevent(); flashFill(); return true; }
    if (kl === 'n' && !sh_) { prevent(); if (confirm('Create a new blank workbook? Unsaved changes are lost.')) { wb = new E.Workbook(); cur = 'Sheet1'; book.shapes = {}; save(); renderTabs(); build(); } return true; }
    if (kl === 'p') { prevent(); window.print(); return true; }
    if (kl === 'w' || k === 'F4') { prevent(); const b = window.__back; if (b) b(); return true; }
    if (kl === 't' || (kl === 'l' && !sh_)) { prevent(); fmtTable(); return true; }
    if (kl === 'l' && sh_) { prevent(); filterDlg(); return true; }
    if (kl === 'q') { prevent(); stMsg('Quick Analysis: use Insert > Charts or Data on the ribbon'); return true; }
    if (k === '`') { prevent(); toggleShowF(); return true; }
    if (k === '&') { prevent(); bordersSel('outside'); return true; }
    if (k === '_') { prevent(); const S = sel; for (const side of ['bt', 'bb']) for (let c = S.c1; c <= S.c2; c++) setSide(side === 'bt' ? S.r1 : S.r2, c, side, null); for (const side of ['bl', 'br']) for (let r = S.r1; r <= S.r2; r++) setSide(r, side === 'bl' ? S.c1 : S.c2, side, null); save(); refresh(); return true; }
    if (k === 'F1') { prevent(); ribHidden = !ribHidden; const rb = $('rib'); if (rb) rb.style.display = ribHidden ? 'none' : ''; return true; }
    if (k === 'F3' && !sh_) { prevent(); nameMgr(); return true; }
    if (k === 'F3' && sh_) { prevent(); nameFromSel(); return true; }
    if (k === 'Home' && sh_) { prevent(); FC = { r: 1, c: 1 }; setSel(anchor.r, anchor.c, 1, 1); return true; }
    if (k === 'End' && sh_) { prevent(); const u = lastUsed(); FC = u; setSel(anchor.r, anchor.c, u.r, u.c); return true; }
    if (k === 'Tab') return false;
    return false;
  }
  if (!ctrl && !alt) {
    if (k === 'Home' && sh_) { prevent(); const cr = (FC && (FC.r === sel.r1 || FC.r === sel.r2)) ? FC.r : anchor.r; FC = { r: cr, c: 1 }; setSel(anchor.r, anchor.c, cr, 1); return true; }
    if (k === 'End') { prevent(); endMode = true; stMsg('End mode — press an arrow key'); return true; }
    if (endMode && /^Arrow/.test(k)) { prevent(); endMode = false; stMsg('Ready'); const dr = k === 'ArrowDown' ? 1 : k === 'ArrowUp' ? -1 : 0, dc = k === 'ArrowRight' ? 1 : k === 'ArrowLeft' ? -1 : 0; const ee = { shiftKey: sh_, ctrlKey: true, metaKey: false }; moveKey(ee, dr, dc); return true; }
    endMode = false;
    if ((k === 'PageDown' || k === 'PageUp') && sh_) { prevent(); const d = k === 'PageDown' ? 20 : -20; const cur0 = FC || { r: anchor.r === sel.r1 ? sel.r2 : sel.r1, c: anchor.c === sel.c1 ? sel.c2 : sel.c1 }; FC = { r: Math.max(1, Math.min(ROWS, cur0.r + d)), c: cur0.c }; setSel(anchor.r, anchor.c, FC.r, FC.c); return true; }
    if (k === 'F1') { prevent(); shortcutsHelp(); return true; }
    if (k === 'F12') { prevent(); saveX(); return true; }
    if (k === 'F9') { prevent(); refresh(); stMsg(sh_ ? 'Calculated active sheet' : 'Calculated all sheets'); return true; }
    if (k === 'F11' && sh_) { prevent(); addSheet(); return true; }
    if (k === 'F11') { prevent(); if (sel.r1 !== sel.r2 || sel.c1 !== sel.c2) addChart('bar'); else stMsg('Select the data range first'); return true; }
    if (k === 'F3' && sh_) { prevent(); fnPick(); return true; } if (k === 'F3') { prevent(); namePaste(); return true; }
    if (k === 'F4') { prevent(); stMsg('F4 cycles $ references while editing a formula'); return true; }
    if (k === 'F2' && sh_) { prevent(); noteDlg(); return true; }
    if (k === 'F7') { prevent(); stMsg('Spell check is not available'); return true; }
    if (k === 'F6') { prevent(); const f = $('nm'); f.focus(); f.select(); return true; }
  }
  return false;
}

document.addEventListener('keydown', e => {
  if (PSD) { const k0 = e.key.toLowerCase(); const m = { a: 'all', f: 'formulas', v: 'values', t: 'formats', e: 'transpose', enter: 'all' }[k0]; if (e.key === 'Escape') { PSD = false; closeGen(); e.preventDefault(); return; } if (m) { e.preventDefault(); e.stopPropagation(); PSD = false; closeGen(); pasteSpecial(m); return; } }
  if (!modalOpen()) PSD = false;
  if (modalOpen()) return;
  const ae = document.activeElement, inFx = ae === $('fx'), inField = /INPUT|TEXTAREA|SELECT/.test(ae.tagName), k = e.key, ctrl = e.ctrlKey || e.metaKey || (CTRLH() && k !== 'Control' && k !== 'Meta'), alt = e.altKey || (ALTH() && k !== 'Alt'), kl = k.length === 1 ? k.toLowerCase() : k;
  if (inFx) { if (ctrl && e.shiftKey && k === 'Enter') { e.preventDefault(); commitFx(0); return; } if (ctrl && k === 'Enter') { e.preventDefault(); const v = $('fx').value, s0 = Object.assign({}, sel); commitFx(0); for (let r = s0.r1; r <= s0.r2; r++) for (let c = s0.c1; c <= s0.c2; c++) wb.setRaw(cur, r, c, v); sel = s0; save(); refresh(); } return; }
  if (inField) return;
  const prevent = () => { e.preventDefault(); e.stopPropagation(); };
  if (xkeys(e, k, kl, ctrl, prevent)) return;
  if (ctrl && alt && kl === 'v') { prevent(); pasteSpecialDlg(); return; }
  if (ctrl && e.shiftKey && kl === 'v') { prevent(); pasteValuesNow(); return; }
  if (ctrl && !alt) {
    switch (kl) {
      case 'c': prevent(); copy(false); return; case 'x': prevent(); copy(true); return;
      case 'v': gotPaste = false; setTimeout(() => { if (!gotPaste && clip) paste(); }, 250); return;
      case 'z': prevent(); undo(); return; case 'y': prevent(); redo(); return;
      case 'a': prevent(); selectAll(); return;
      case 'b': prevent(); tog('b'); return; case 'i': prevent(); tog('i'); return; case 'u': prevent(); tog('u'); return; case '5': prevent(); tog('s'); return;
      case 'd': prevent(); fill('d'); return; case 'r': prevent(); fill('r'); return;
      case 'f': case 'h': prevent(); findRep(); return;
      case 's': prevent(); saveX(); return; case 'o': prevent(); openFile(); return;
      case 'g': prevent(); if (typeof gotoDlg === 'function') gotoDlg(); else { $('nm').focus(); $('nm').select(); } return;
      case 'k': prevent(); if (typeof linkDlg === 'function') linkDlg(); return;
      case '1': if (e.shiftKey) { } else { prevent(); formatCellsDlg(); return; } break;
      case '!': prevent(); nfmt('#,##0.00'); return; case '@': prevent(); nfmt('h:mm AM/PM'); return; case '#': prevent(); nfmt('d-mmm-yy'); return;
      case '$': prevent(); nfmt('$#,##0.00'); return; case '%': prevent(); nfmt('0%'); return; case '^': prevent(); nfmt('0.00E+00'); return; case '~': prevent(); nfmt(''); return;
      case ';': prevent(); insertNow('date'); return; case ':': prevent(); insertNow('time'); return;
      case '+': case '=': prevent(); insRC('r'); return; case '-': prevent(); delRC('r'); return;
      case '9': prevent(); hideRC('r'); return; case '0': prevent(); hideRC('c'); return; case '(': prevent(); unhideRC('r'); return; case ')': prevent(); unhideRC('c'); return;
      case ' ': prevent(); selectRowCol('col'); return;
      case 'Home': prevent(); goCell(1, 1); return;
      case 'End': { prevent(); let mr = 1, mc = 1; for (const kk in sh().cells) { const [r, c] = kk.split(',').map(Number); if (r > mr) mr = r; if (c > mc) mc = c; } goCell(mr, mc); return; }
      case 'ArrowUp': prevent(); moveKey(e, -1, 0); return; case 'ArrowDown': prevent(); moveKey(e, 1, 0); return; case 'ArrowLeft': prevent(); moveKey(e, 0, -1); return; case 'ArrowRight': prevent(); moveKey(e, 0, 1); return;
      case 'Enter': prevent(); return;
    }
    if (e.shiftKey && k === ' ') { }
    return;
  }
  if (alt) { if (k === '=') { prevent(); autosum(); } return; }
  switch (k) {
    case 'ArrowUp': prevent(); moveKey(e, -1, 0); return; case 'ArrowDown': prevent(); moveKey(e, 1, 0); return;
    case 'ArrowLeft': prevent(); moveKey(e, 0, -1); return; case 'ArrowRight': prevent(); moveKey(e, 0, 1); return;
    case 'Tab': prevent(); goCell(anchor.r, anchor.c + (e.shiftKey ? -1 : 1)); return;
    case 'Enter': prevent(); goCell(anchor.r + (e.shiftKey ? -1 : 1), anchor.c); return;
    case 'Home': prevent(); goCell(anchor.r, 1); return;
    case 'End': prevent(); return;
    case 'PageDown': prevent(); goCell(anchor.r + 20, anchor.c); return; case 'PageUp': prevent(); goCell(anchor.r - 20, anchor.c); return;
    case 'F2': prevent(); startEdit(String(wb.getRaw(cur, anchor.r, anchor.c)), false); return;
    case 'F5': prevent(); if (typeof gotoDlg === 'function') gotoDlg(); else { $('nm').focus(); $('nm').select(); } return;
    case 'Delete': case 'Backspace': prevent(); clearSel(); return;
    case 'Escape': { const k0 = clip; PAINT = null; clip = null; if (k0 && window.repaintRng) repaintRng(k0.s); }  $('stL').textContent = 'Ready'; return;
    case ' ': if (e.shiftKey) { prevent(); selectRowCol('row'); return; } break;
  }
  if (k.length === 1 && !alt) { prevent(); startEdit(k, true); }
}, true);

/* ---------------- copy/cut "marching ants" indicator ---------------- */
function antCls(r, c) {
  const k = clip; if (!k || !k.ants || k.sheet !== cur) return ''; const s = k.s;
  if (r < s.r1 || r > s.r2 || c < s.c1 || c > s.c2) return '';
  return ' ant' + (r === s.r1 ? ' aT' : '') + (r === s.r2 ? ' aB' : '') + (c === s.c1 ? ' aL' : '') + (c === s.c2 ? ' aR' : '');
}
function repaintRng(s) { for (let r = s.r1; r <= Math.min(s.r2, ROWS); r++) for (let c = s.c1; c <= Math.min(s.c2, COLS); c++) paint(r, c); }
const _copy2 = copy;
copy = function (cut) { const old = clip && clip.ants ? clip.s : null; _copy2(cut); if (clip) clip.ants = true; if (old) repaintRng(old); if (clip) repaintRng(clip.s); };
const _paste2 = paste;
paste = function () {
  const k = clip;
  if (k && !k.cut) { const h = k.s.r2 - k.s.r1 + 1, w = k.s.c2 - k.s.c1 + 1, sh_ = sel.r2 - sel.r1 + 1, sw = sel.c2 - sel.c1 + 1;
    if ((sh_ > h || sw > w) && sh_ % h === 0 && sw % w === 0 && (sh_ / h) * (sw / w) <= 400) { const a0 = Object.assign({}, anchor), s0 = Object.assign({}, sel); for (let r = s0.r1; r <= s0.r2; r += h) for (let c = s0.c1; c <= s0.c2; c += w) { anchor = { r, c }; _paste2(); } anchor = a0; sel = s0; k.ants = false; refresh(); return; } }
  _paste2(); if (k) { k.ants = false; repaintRng(k.s); }
};
const _pasteSp2 = pasteSpecial;
pasteSpecial = function (m) { const k = clip; _pasteSp2(m); if (k) { k.ants = false; repaintRng(k.s); } };

/* ---------------- double-click / drag column & row borders ---------------- */
(function () {
  const g = $('grid'); const TOL = 6;
  function edge(e) {
    const th = e.target.closest && e.target.closest('th'); if (!th) return null; const R = th.getBoundingClientRect();
    if (/^ch\d+$/.test(th.id)) { const c = +th.id.slice(2); if (R.right - e.clientX <= TOL) return { t: 'c', i: c, th }; if (e.clientX - R.left <= TOL && c > 1) return { t: 'c', i: c - 1 }; }
    else if (/^rh\d+$/.test(th.id)) { const r = +th.id.slice(2); if (R.bottom - e.clientY <= TOL) return { t: 'r', i: r, th }; if (e.clientY - R.top <= TOL && r > 1) return { t: 'r', i: r - 1 }; }
    return null;
  }
  g.addEventListener('dblclick', e => { const ed = edge(e); if (!ed) return; e.preventDefault(); e.stopPropagation(); if (ed.t === 'c') autoFitCol(ed.i); else autoFitRow(true, ed.i); stMsg(ed.t === 'c' ? 'Column ' + E.numToCol(ed.i) + ' auto-fit' : 'Row ' + ed.i + ' auto-fit'); }, true);
  g.addEventListener('mousemove', e => { if (rs) return; const ed = edge(e); g.classList.toggle('rszc', !!ed && ed.t === 'c'); g.classList.toggle('rszr', !!ed && ed.t === 'r'); });
  let rs = null;
  g.addEventListener('pointerdown', e => {
    if (e.pointerType === 'touch') return; const ed = edge(e); if (!ed) return; e.preventDefault(); e.stopPropagation();
    const thx = ed.t === 'c' ? $('ch' + ed.i) : $('rh' + ed.i); const R = thx.getBoundingClientRect();
    const base = ed.t === 'c' ? colW(ed.i) : rowHt(ed.i); const z = (ed.t === 'c' ? R.width : R.height) / Math.max(1, base);
    rs = { t: ed.t, i: ed.i, x: e.clientX, y: e.clientY, base, z, v: base };
  }, true);
  window.addEventListener('pointermove', e => { if (!rs) return; const d = rs.t === 'c' ? e.clientX - rs.x : e.clientY - rs.y; rs.v = Math.max(rs.t === 'c' ? 8 : 6, Math.round(rs.base + d / (rs.z || 1))); stMsg((rs.t === 'c' ? 'Width: ' : 'Height: ') + rs.v + ' px'); });
  window.addEventListener('pointerup', () => { if (!rs) return; const o = rs; rs = null; if (o.v === o.base) return; if (o.t === 'c') sh().colW[o.i] = o.v; else sh().rowH[o.i] = o.v; save(); build(); });
})();

/* ---------------- Data Validation ---------------- */
function dvRuleAt(sheet, r, c) { const S = wb.sheets[sheet]; const L = (S && S.dv) || []; for (let i = L.length - 1; i >= 0; i--) { const q = L[i]; if (r >= q.r1 && r <= q.r2 && c >= q.c1 && c <= q.c2) return q; } return null; }
function dvListAt(r, c) { const q = dvRuleAt(cur, r, c); return q && q.type === 'list' && q.drop !== false ? q : null; }
function dvListItems(q, sheet) {
  const s = String(q.v1 === undefined ? '' : q.v1).trim(); if (!s) return [];
  if (s[0] === '=') {
    let t = s.slice(1).trim().replace(/\$/g, ''); let sn = sheet; const m0 = /^(?:'([^']+)'|([^'!]+))!(.*)$/.exec(t); if (m0) { sn = m0[1] || m0[2]; t = m0[3]; }
    const nm = book.names && book.names[t.toUpperCase()]; if (nm) { const m1 = /^(?:'([^']+)'|([^'!]+))!(.*)$/.exec(String(nm).replace(/\$/g, '')); if (m1) { sn = m1[1] || m1[2]; t = m1[3]; } else t = String(nm).replace(/\$/g, ''); }
    const m = /^([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?$/i.exec(t); if (!m || !wb.sheets[sn]) return [];
    const c1 = E.colToNum(m[1].toUpperCase()), r1 = +m[2], c2 = m[3] ? E.colToNum(m[3].toUpperCase()) : c1, r2 = m[4] ? +m[4] : r1; const out = [];
    for (let r = r1; r <= r2; r++) for (let c = c1; c <= c2; c++) { const v = wb.value(sn, r, c); if (v !== null && v !== undefined && v !== '') out.push(String(E.isErr(v) ? '' : v)); } return out;
  }
  return s.split(',').map(x => x.trim()).filter(x => x !== '');
}
function dvBound(sheet, s, type) {
  if (s === undefined || s === null || String(s).trim() === '') return null; s = String(s).trim();
  if (s[0] === '=') { try { const v = E.evaluate(wb, sheet, s); return typeof v === 'number' ? v : (v && v.v !== undefined ? v.v : v); } catch (e) { return null; } }
  if (type === 'time') { const m = /^(\d{1,2}):(\d{2})(?::(\d{2}))?\s*(AM|PM)?$/i.exec(s); if (m) { let h = +m[1]; if (m[4]) { h = h % 12 + (/pm/i.test(m[4]) ? 12 : 0); } return (h * 3600 + +m[2] * 60 + +(m[3] || 0)) / 86400; } }
  const p = wb.parseInput(s); return p.v;
}
function dvCmp(op, x, a, b) {
  switch (op) { case 'between': return x >= a && x <= b; case 'notBetween': return !(x >= a && x <= b); case 'equal': return x === a; case 'notEqual': return x !== a;
    case 'gt': return x > a; case 'lt': return x < a; case 'ge': return x >= a; case 'le': return x <= a; } return true;
}
function dvCheck(sheet, r, c, raw) {
  const q = dvRuleAt(sheet, r, c); if (!q || q.type === 'any') return { ok: true };
  if (raw === '' || raw === null || raw === undefined) return { ok: q.blank !== false, q };
  raw = String(raw); const isF = raw[0] === '=' && raw.length > 1;
  if (isF && q.type !== 'custom') return { ok: true, q };
  let ok = true;
  if (q.type === 'list') { const items = dvListItems(q, sheet); ok = items.some(x => x.toLowerCase() === raw.trim().toLowerCase() || (x !== '' && !isNaN(+x) && !isNaN(+raw) && +x === +raw)); }
  else if (q.type === 'length') { const a = dvBound(sheet, q.v1, 'n'), b = dvBound(sheet, q.v2, 'n'); ok = dvCmp(q.op, raw.length, a, b); }
  else if (q.type === 'custom') {
    const S = wb.sheets[sheet], old = wb.getRaw(sheet, r, c); let res = false;
    try { wb.setRaw(sheet, r, c, raw); wb.recalc(); const f = shiftFormula(String(q.v1 || ''), r - q.r1, c - q.c1); const v = E.evaluate(wb, sheet, f); res = !(v === false || v === 0 || v === '' || v === null || (v && v.err) || E.isErr(v)); } catch (e) { res = false; }
    wb.setRaw(sheet, r, c, old); wb.recalc(); ok = res;
  }
  else {
    const p = wb.parseInput(raw); let v = p.v; if (q.type === 'time') { const t = dvBound(sheet, raw, 'time'); if (typeof t === 'number') v = t; }
    if (typeof v !== 'number') ok = false; else {
      if (q.type === 'whole' && !Number.isInteger(v)) ok = false; else if (q.type === 'date' || q.type === 'time') { const a = dvBound(sheet, q.v1, q.type), b = dvBound(sheet, q.v2, q.type); ok = dvCmp(q.op, q.type === 'date' ? Math.floor(v) : v, q.type === 'date' && a !== null ? Math.floor(a) : a, q.type === 'date' && b !== null ? Math.floor(b) : b); }
      else { const a = dvBound(sheet, q.v1, 'n'), b = dvBound(sheet, q.v2, 'n'); ok = dvCmp(q.op, v, a, b); }
    }
  }
  return { ok, q };
}
const _commitFx = commitFx;
commitFx = function (move) {
  if (editing || editCell) {
    const ec = editCell || { sheet: cur, r: anchor.r, c: anchor.c }; const v = $('fx').value; const res = dvCheck(ec.sheet, ec.r, ec.c, v);
    if (!res.ok && res.q) {
      const er = res.q.er || {}, msg = (er.t ? er.t + '\n\n' : '') + (er.m || "This value doesn't match the data validation restrictions defined for this cell."); const st = er.on === false ? 'none' : (er.st || 'stop');
      if (st === 'stop') { alert(msg); try { $('fx').focus(); } catch (e) { } return; }
      if (st === 'warning') { if (!confirm(msg + '\n\nDo you want to continue?')) { try { $('fx').focus(); } catch (e) { } return; } }
      else if (st === 'information') { if (!confirm(msg + '\n\nOK = keep the value, Cancel = reject it.')) { try { $('fx').focus(); } catch (e) { } return; } }
    }
  }
  _commitFx(move);
};
/* dropdown button, input message, list menu */
let dvMenuOn = false, dvSel = 0, dvItems = [];
function dvEl(id) { let e = document.getElementById(id); if (!e) { e = document.createElement('div'); e.id = id; document.body.appendChild(e); } return e; }
function dvUpdate() {
  const btn = dvEl('dvBtn'), tip = dvEl('dvTip'); btn.style.display = tip.style.display = 'none'; if (editing) return;
  const q = dvRuleAt(cur, anchor.r, anchor.c); const td = tds[anchor.r] && tds[anchor.r][anchor.c]; if (!q || !td) return;
  const R = td.getBoundingClientRect(), W = $('wrap').getBoundingClientRect(); if (R.bottom < W.top + 10 || R.top > W.bottom || R.right < W.left + 10 || R.left > W.right) return;
  if (q.type === 'list' && q.drop !== false) { btn.textContent = '▼'; btn.style.cssText = `display:block;left:${R.right - 18}px;top:${R.top}px;height:${Math.max(16, R.height)}px;line-height:${Math.max(16, R.height)}px`; btn.onclick = ev => { ev.stopPropagation(); dvOpen(); }; }
  if (q.im && q.im.on && (q.im.t || q.im.m)) { tip.innerHTML = (q.im.t ? '<b>' + q.im.t.replace(/</g, '&lt;') + '</b><br>' : '') + (q.im.m || '').replace(/</g, '&lt;'); tip.style.display = 'block'; tip.style.left = Math.min(R.left, innerWidth - 250) + 'px'; tip.style.top = (R.bottom + 3) + 'px'; }
}
const _paintSel2 = paintSel;
paintSel = function (p) { _paintSel2(p); dvUpdate(); };
window.addEventListener('load', () => { $('wrap').addEventListener('scroll', () => { dvUpdate(); dvClose(); }); setTimeout(dvUpdate, 300); });
function dvClose() { dvMenuOn = false; const m = document.getElementById('dvMenu'); if (m) m.style.display = 'none'; }
function dvPick(i) { const v = dvItems[i]; dvClose(); if (v === undefined) return; setCell(cur, anchor.r, anchor.c, v); dvUpdate(); }
function dvOpen() {
  const q = dvListAt(anchor.r, anchor.c); if (!q) return; dvItems = dvListItems(q, cur); const m = dvEl('dvMenu'); const td = tds[anchor.r][anchor.c]; const R = td.getBoundingClientRect();
  const curv = String(wb.getRaw(cur, anchor.r, anchor.c)); dvSel = Math.max(0, dvItems.findIndex(x => x === curv));
  m.innerHTML = dvItems.length ? dvItems.map((x, i) => `<div data-i="${i}" class="${i === dvSel ? 'on' : ''}">${String(x).replace(/</g, '&lt;')}</div>`).join('') : '<div style="color:#888">(no items)</div>';
  m.style.display = 'block'; m.style.left = Math.min(R.left, innerWidth - 140) + 'px'; m.style.top = (R.bottom) + 'px'; m.onclick = ev => { const d = ev.target.closest('div[data-i]'); if (d) dvPick(+d.dataset.i); }; dvMenuOn = true;
}
function dvMenuKey(e) {
  const k = e.key; const m = document.getElementById('dvMenu'); const mv = d => { dvSel = Math.max(0, Math.min(dvItems.length - 1, dvSel + d)); [...m.children].forEach((x, i) => x.classList.toggle('on', i === dvSel)); const x = m.children[dvSel]; if (x) x.scrollIntoView({ block: 'nearest' }); };
  if (k === 'ArrowDown') { e.preventDefault(); e.stopPropagation(); mv(1); return true; } if (k === 'ArrowUp') { e.preventDefault(); e.stopPropagation(); mv(-1); return true; }
  if (k === 'Enter' || k === 'Tab') { e.preventDefault(); e.stopPropagation(); dvPick(dvSel); return true; } if (k === 'Escape') { e.preventDefault(); e.stopPropagation(); dvClose(); return true; }
  dvClose(); return false;
}
document.addEventListener('pointerdown', e => { if (dvMenuOn && !(e.target.closest && (e.target.closest('#dvMenu') || e.target.closest('#dvBtn')))) dvClose(); }, true);
/* dialog */
function dvDlg() {
  const q = dvRuleAt(cur, anchor.r, anchor.c) || {}; const im = q.im || {}, er = q.er || { on: true, st: 'stop' };
  const sel1 = (id, opts, v) => `<select id="${id}" onchange="dvSync()">${opts.map(o => `<option value="${o[0]}" ${o[0] === v ? 'selected' : ''}>${o[1]}</option>`).join('')}</select>`;
  const esc = x => String(x === undefined ? '' : x).replace(/"/g, '&quot;');
  const h = `<label>Allow</label>${sel1('dvA', [['any', 'Any value'], ['whole', 'Whole number'], ['decimal', 'Decimal'], ['list', 'List'], ['date', 'Date'], ['time', 'Time'], ['length', 'Text length'], ['custom', 'Custom (formula)']], q.type || 'any')}
  <div id="dvOpW"><label>Data</label>${sel1('dvO', [['between', 'between'], ['notBetween', 'not between'], ['equal', 'equal to'], ['notEqual', 'not equal to'], ['gt', 'greater than'], ['lt', 'less than'], ['ge', 'greater than or equal to'], ['le', 'less than or equal to']], q.op || 'between')}</div>
  <div id="dvV1W"><label id="dvL1">Minimum</label><input type="text" id="dvV1" value="${esc(q.v1)}"></div>
  <div id="dvV2W"><label>Maximum</label><input type="text" id="dvV2" value="${esc(q.v2)}"></div>
  <div id="dvLsW" style="font-size:11px;color:#666">List: type items separated by commas (Yes,No,Maybe) or a range such as =$A$1:$A$10 or =Sheet2!A1:A8. Custom: a formula for the first cell, e.g. =A1&gt;10.</div>
  <label><input type="checkbox" id="dvBl" ${q.blank === false ? '' : 'checked'}> Ignore blank</label>
  <label id="dvDrW"><input type="checkbox" id="dvDr" ${q.drop === false ? '' : 'checked'}> In-cell dropdown</label>
  <hr><label><input type="checkbox" id="dvIm" ${im.on ? 'checked' : ''}> Show input message when cell is selected</label>
  <input type="text" id="dvIT" placeholder="Title" value="${esc(im.t)}"><input type="text" id="dvIM" placeholder="Input message" value="${esc(im.m)}">
  <hr><label><input type="checkbox" id="dvEr" ${er.on === false ? '' : 'checked'}> Show error alert after invalid data is entered</label>
  ${sel1('dvEs', [['stop', 'Stop'], ['warning', 'Warning'], ['information', 'Information']], er.st || 'stop')}
  <input type="text" id="dvET" placeholder="Error title" value="${esc(er.t)}"><input type="text" id="dvEM" placeholder="Error message" value="${esc(er.m)}">`;
  gen('Data Validation', h, [['OK', dvApply], ['Clear in selection', dvClear], ['Clear all', () => { sh().dv = []; save(); closeGen(); refresh(); dvUpdate(); }]]);
  dvSync();
}
function dvSync() {
  const t = $('dvA').value, op = $('dvO').value; const show = (id, on) => { $(id).style.display = on ? '' : 'none'; };
  const num = ['whole', 'decimal', 'date', 'time', 'length'].includes(t); show('dvOpW', num); show('dvV1W', t !== 'any'); show('dvV2W', num && (op === 'between' || op === 'notBetween')); show('dvLsW', t === 'list' || t === 'custom'); show('dvDrW', t === 'list');
  $('dvL1').textContent = t === 'list' ? 'Source' : t === 'custom' ? 'Formula' : (op === 'between' || op === 'notBetween') ? 'Minimum' : 'Value';
}
function dvApply() {
  const S = sh(); S.dv = S.dv || []; const t = $('dvA').value;
  const covered = q => q.r1 >= sel.r1 && q.r2 <= sel.r2 && q.c1 >= sel.c1 && q.c2 <= sel.c2; S.dv = S.dv.filter(q => !covered(q));
  const rule = { r1: sel.r1, c1: sel.c1, r2: sel.r2, c2: sel.c2, type: t, op: $('dvO').value, v1: $('dvV1').value, v2: $('dvV2').value, blank: $('dvBl').checked, drop: $('dvDr').checked,
    im: { on: $('dvIm').checked, t: $('dvIT').value, m: $('dvIM').value }, er: { on: $('dvEr').checked, st: $('dvEs').value, t: $('dvET').value, m: $('dvEM').value } };
  if (t !== 'any' && String(rule.v1).trim() === '') { alert('Enter the ' + ($('dvL1').textContent.toLowerCase()) + ' value first.'); return; }
  if (t !== 'any' || rule.im.on) S.dv.push(rule); save(); closeGen(); refresh(); dvUpdate();
}
function dvClear() {
  const S = sh(); S.dv = S.dv || []; const inter = q => !(q.r2 < sel.r1 || q.r1 > sel.r2 || q.c2 < sel.c1 || q.c1 > sel.c2); const cov = q => q.r1 >= sel.r1 && q.r2 <= sel.r2 && q.c1 >= sel.c1 && q.c2 <= sel.c2;
  const part = S.dv.some(q => inter(q) && !cov(q)); S.dv = S.dv.filter(q => !cov(q)); if (part) S.dv.push({ r1: sel.r1, c1: sel.c1, r2: sel.r2, c2: sel.c2, type: 'any', op: 'between', v1: '', v2: '', blank: true });
  save(); closeGen(); refresh(); dvUpdate();
}

/* ---------------- live typing in the cell, arrow-to-commit, reference shifting, auto-uppercase ---------------- */
function liveCell() {
  if (!editing || !editCell || editCell.sheet !== cur) return; const td = tds[editCell.r] && tds[editCell.r][editCell.c]; if (!td) return;
  td.textContent = $('fx').value; td.classList.add('typing');
}
const _fxInput2 = fxInput; fxInput = function () { _fxInput2(); liveCell(); };
const _pointTo2 = pointTo; pointTo = function (r, c, ext) { _pointTo2(r, c, ext); liveCell(); };
const _cancelEdit2 = cancelEdit; cancelEdit = function () { const ec = editCell; enterMode = false; _cancelEdit2(); if (ec && ec.sheet === cur) paint(ec.r, ec.c); };
const _fxFocus2 = fxFocus; fxFocus = function () { if (!editing) enterMode = false; _fxFocus2(); };
const _fxKey2 = fxKey;
fxKey = function (e) {
  if (editing && enterMode && /^Arrow(Up|Down|Left|Right)$/.test(e.key) && !canPoint()) {
    e.preventDefault(); const dr = e.key === 'ArrowDown' ? 1 : e.key === 'ArrowUp' ? -1 : 0, dc = e.key === 'ArrowRight' ? 1 : e.key === 'ArrowLeft' ? -1 : 0;
    enterMode = false; commitFx(0); goCell(anchor.r + dr, anchor.c + dc); return;
  }
  _fxKey2(e);
};
function shiftFormula0(f, dr, dc) {
  let out = '', i = 0; const n = f.length;
  while (i < n) {
    const ch = f[i];
    if (ch === '"') { let j = i + 1; while (j < n) { if (f[j] === '"') { if (f[j + 1] === '"') { j += 2; continue; } break; } j++; } out += f.slice(i, j + 1); i = j + 1; continue; }
    if (ch === "'") { let j = f.indexOf("'", i + 1); if (j < 0) j = n - 1; out += f.slice(i, j + 1); i = j + 1; continue; }
    if (/[A-Za-z_$]/.test(ch)) {
      let j = i; while (j < n && /[A-Za-z0-9_$.]/.test(f[j])) j++; const tok = f.slice(i, j); const m = /^(\$?)([A-Za-z]{1,3})(\$?)(\d+)$/.exec(tok);
      if (m && f[j] !== '(' && f[j] !== '!') { const col = m[1] ? m[2].toUpperCase() : E.numToCol(Math.max(1, E.colToNum(m[2]) + dc)); const row = m[3] ? m[4] : String(Math.max(1, +m[4] + dr)); out += m[1] + col + m[3] + row; }
      else out += tok; i = j; continue;
    }
    out += ch; i++;
  }
  return out;
}
shiftFormula = shiftFormula0;
function normFormula(f) {
  if (typeof f !== 'string' || f[0] !== '=') return f; let out = '', i = 0; const n = f.length;
  while (i < n) {
    const ch = f[i];
    if (ch === '"') { let j = i + 1; while (j < n) { if (f[j] === '"') { if (f[j + 1] === '"') { j += 2; continue; } break; } j++; } out += f.slice(i, j + 1); i = j + 1; continue; }
    if (ch === "'") { let j = f.indexOf("'", i + 1); if (j < 0) j = n - 1; out += f.slice(i, j + 1); i = j + 1; continue; }
    if (/[A-Za-z_$]/.test(ch)) {
      let j = i; while (j < n && /[A-Za-z0-9_$.]/.test(f[j])) j++; const tok = f.slice(i, j);
      if (f[j] === '(' && /^[A-Za-z][A-Za-z0-9.]*$/.test(tok) && E.FN && E.FN[tok.toUpperCase()]) out += tok.toUpperCase();
      else if (f[j] !== '(' && f[j] !== '!' && /^\$?[A-Za-z]{1,3}\$?\d+$/.test(tok)) out += tok.toUpperCase();
      else if (/^(true|false)$/i.test(tok) && f[j] !== '(') out += tok.toUpperCase();
      else out += tok; i = j; continue;
    }
    out += ch; i++;
  }
  return out;
}
/* uppercase refs/functions when an entry is finalised (wrapped around the validation wrapper) */
const _commitFx3 = commitFx;
commitFx = function (move) { if (editing || editCell) { const f = $('fx'); const v = f.value; const nv = normFormula(v); if (nv !== v) f.value = nv; } _commitFx3(move); };
