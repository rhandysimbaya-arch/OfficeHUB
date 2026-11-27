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
function undo() { if (editing) return; if (H.i <= 0) { $('stL').textContent = 'Nothing to undo'; return; } H.i--; restoreSnap(H.a[H.i]); $('stL').textContent = 'Undo'; }
function redo() { if (editing) return; if (H.i >= H.a.length - 1) { $('stL').textContent = 'Nothing to redo'; return; } H.i++; restoreSnap(H.a[H.i]); $('stL').textContent = 'Redo'; }
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
function autoFitCol() {
  for (let c = sel.c1; c <= sel.c2; c++) { let w = 30; for (let r = 1; r <= ROWS; r++) { if (MM[key(r, c)]) continue; const t = cellText(r, c); if (!t) continue; const st = styleOf(r, c); if (st && st.wrap) continue; _cx.font = fontOf(st); w = Math.max(w, _cx.measureText(t).width + 14 + (st && st.indent ? st.indent * 9 : 0)); } sh().colW[c] = Math.ceil(w); }
  save(); build();
}
function wrapLines(txt, font, w) { _cx.font = font; let lines = 0; for (const para of String(txt).split('\n')) { let line = '', n = 1; for (const word of para.split(' ')) { const t = line ? line + ' ' + word : word; if (_cx.measureText(t).width > w && line) { n++; line = word; } else line = t; } lines += n; } return lines; }
function autoFitRow(quiet) {
  const r1 = quiet ? sel.r1 : sel.r1, r2 = sel.r2;
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
  const ctrl = e.ctrlKey || e.metaKey; const cur0 = (e.shiftKey && FC && ((FC.r === sel.r1 || FC.r === sel.r2) && (FC.c === sel.c1 || FC.c === sel.c2))) ? FC : (e.shiftKey ? { r: anchor.r === sel.r1 ? sel.r2 : sel.r1, c: anchor.c === sel.c1 ? sel.c2 : sel.c1 } : { r: anchor.r, c: anchor.c });
  let n = ctrl ? jumpEdge(cur0.r, cur0.c, dr, dc) : { r: Math.max(1, Math.min(ROWS, cur0.r + dr)), c: Math.max(1, Math.min(COLS, cur0.c + dc)) };
  if (e.shiftKey) { FC = n; setSel(anchor.r, anchor.c, n.r, n.c); } else goCell(n.r, n.c);
}
function startEdit(text) { const f = $('fx'); f.focus(); if (text !== undefined) f.value = text; const n = f.value.length; try { f.setSelectionRange(n, n); } catch (e) { } editing = true; editCell = { sheet: cur, r: anchor.r, c: anchor.c }; }
function cycleAbs() {
  const f = $('fx'); const v = f.value, p = f.selectionStart; const re = /\$?[A-Z]{1,3}\$?\d+/gi; let m, hit = null; while ((m = re.exec(v))) { if (p >= m.index && p <= m.index + m[0].length) { hit = m; break; } } if (!hit) return;
  const mm = /(\$?)([A-Z]{1,3})(\$?)(\d+)/i.exec(hit[0]); const st = (mm[1] ? 1 : 0) + (mm[3] ? 2 : 0); const nx = [[1, 1], [0, 1], [1, 0], [0, 0]]; const cyc = { 3: [0, 1], 2: [0, 0], 1: [1, 1], 0: [1, 0] }; // $A$1 -> A$1 -> $A1 -> A1 -> $A$1
  const order = [[1, 1], [0, 1], [1, 0], [0, 0]]; const idx = order.findIndex(o => o[0] === (mm[1] ? 1 : 0) && o[1] === (mm[3] ? 1 : 0)); const o = order[(idx + 1) % 4];
  const rep = (o[0] ? '$' : '') + mm[2] + (o[1] ? '$' : '') + mm[4]; f.value = v.slice(0, hit.index) + rep + v.slice(hit.index + hit[0].length); const np = hit.index + rep.length; f.setSelectionRange(np, np); guide();
}
function tsvOfSel() { const rows = []; for (let r = sel.r1; r <= Math.min(sel.r2, ROWS); r++) { const row = []; for (let c = sel.c1; c <= Math.min(sel.c2, COLS); c++) row.push(cellText(r, c)); rows.push(row.join('\t')); } return rows.join('\n'); }
function sysCopy(t) { try { const ta = document.createElement('textarea'); ta.value = t; ta.style.cssText = 'position:fixed;left:-9999px;top:0;opacity:0'; document.body.appendChild(ta); ta.select(); document.execCommand('copy'); ta.remove(); } catch (e) { } try { navigator.clipboard && navigator.clipboard.writeText(t); } catch (e) { } }
const _copy0 = copy;
copy = function (cut) { _copy0(cut); if (clip) { clip.tsv = tsvOfSel(); sysCopy(clip.tsv); } };
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
  const L = [['Ctrl+C / X / V', 'Copy / Cut / Paste'], ['Ctrl+Z / Y', 'Undo / Redo'], ['Ctrl+B / I / U', 'Bold / Italic / Underline'], ['Ctrl+5', 'Strikethrough'], ['Ctrl+A', 'Select all'], ['Ctrl+D / R', 'Fill down / right'], ['Ctrl+F / H', 'Find & Replace'], ['Ctrl+S', 'Save'], ['Ctrl+O', 'Open'], ['Ctrl+1', 'Format Cells'], ['Ctrl+Shift+1 / 2 / 3', 'Number / Time / Date format'], ['Ctrl+Shift+4 / 5', 'Currency / Percent'], ['Ctrl+; / Ctrl+Shift+;', 'Insert date / time'], ['Ctrl++ / Ctrl+-', 'Insert / delete row'], ['Ctrl+9 / Ctrl+0', 'Hide row / column'], ['Ctrl+Arrow', 'Jump to edge of data'], ['Ctrl+Home / End', 'First cell / last used cell'], ['Shift+Arrow', 'Extend selection'], ['Ctrl+Shift+Arrow', 'Extend to edge of data'], ['Ctrl+Space / Shift+Space', 'Select column / row'], ['Tab / Shift+Tab', 'Move right / left'], ['Enter / Shift+Enter', 'Move down / up'], ['Page Up / Down', 'Scroll one screen'], ['F2', 'Edit cell'], ['F4', 'Cycle $ in reference (editing)'], ['F5 / Ctrl+G', 'Go to'], ['Alt+=', 'AutoSum'], ['Ctrl+Enter', 'Fill selection with entry'], ['Delete', 'Clear contents'], ['Esc', 'Cancel']];
  gen('Keyboard shortcuts', '<table style="width:100%;font-size:13px;border-collapse:collapse">' + L.map(x => `<tr><td style="padding:3px 6px;border-bottom:1px solid #eee;white-space:nowrap"><b>${x[0]}</b></td><td style="padding:3px 6px;border-bottom:1px solid #eee">${x[1]}</td></tr>`).join('') + '</table>', []);
}
document.addEventListener('keydown', e => {
  if (modalOpen()) return;
  const ae = document.activeElement, inFx = ae === $('fx'), inField = /INPUT|TEXTAREA|SELECT/.test(ae.tagName), k = e.key, ctrl = e.ctrlKey || e.metaKey, kl = k.length === 1 ? k.toLowerCase() : k;
  if (inFx) { if (ctrl && k === 'Enter') { e.preventDefault(); const v = $('fx').value, s0 = Object.assign({}, sel); commitFx(0); for (let r = s0.r1; r <= s0.r2; r++) for (let c = s0.c1; c <= s0.c2; c++) wb.setRaw(cur, r, c, v); sel = s0; save(); refresh(); } return; }
  if (inField) return;
  const prevent = () => { e.preventDefault(); e.stopPropagation(); };
  if (ctrl && !e.altKey) {
    switch (kl) {
      case 'c': prevent(); copy(false); return; case 'x': prevent(); copy(true); return;
      case 'v': gotPaste = false; setTimeout(() => { if (!gotPaste && clip) paste(); }, 250); return;
      case 'z': prevent(); undo(); return; case 'y': prevent(); redo(); return;
      case 'a': prevent(); selectAll(); return;
      case 'b': prevent(); tog('b'); return; case 'i': prevent(); tog('i'); return; case 'u': prevent(); tog('u'); return; case '5': prevent(); tog('s'); return;
      case 'd': prevent(); fill('d'); return; case 'r': prevent(); fill('r'); return;
      case 'f': case 'h': prevent(); findRep(); return;
      case 's': prevent(); saveX(); return; case 'o': prevent(); openFile(); return;
      case 'g': prevent(); $('nm').focus(); $('nm').select(); return;
      case 'k': prevent(); return;
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
  if (e.altKey) { if (k === '=') { prevent(); autosum(); } return; }
  switch (k) {
    case 'ArrowUp': prevent(); moveKey(e, -1, 0); return; case 'ArrowDown': prevent(); moveKey(e, 1, 0); return;
    case 'ArrowLeft': prevent(); moveKey(e, 0, -1); return; case 'ArrowRight': prevent(); moveKey(e, 0, 1); return;
    case 'Tab': prevent(); goCell(anchor.r, anchor.c + (e.shiftKey ? -1 : 1)); return;
    case 'Enter': prevent(); goCell(anchor.r + (e.shiftKey ? -1 : 1), anchor.c); return;
    case 'Home': prevent(); goCell(anchor.r, 1); return;
    case 'End': prevent(); return;
    case 'PageDown': prevent(); goCell(anchor.r + 20, anchor.c); return; case 'PageUp': prevent(); goCell(anchor.r - 20, anchor.c); return;
    case 'F2': prevent(); startEdit(String(wb.getRaw(cur, anchor.r, anchor.c))); return;
    case 'F5': prevent(); $('nm').focus(); $('nm').select(); return;
    case 'Delete': case 'Backspace': prevent(); clearSel(); return;
    case 'Escape': PAINT = null; clip = null; $('stL').textContent = 'Ready'; return;
    case ' ': if (e.shiftKey) { prevent(); selectRowCol('row'); return; } break;
  }
  if (k.length === 1 && !e.altKey) { prevent(); startEdit(k); }
}, true);
