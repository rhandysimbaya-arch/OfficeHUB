/* Excel Tier-1 extras: fill handle, context menu, filter arrows, sort dialog, Go To / Find, names, Insert Function,
   hyperlinks, notes, protection.  Loaded after sheetx.js. Per-sheet extras live in sh().ext (persisted by engine). */
'use strict';
const X = () => { const S = sh(); return S.ext || (S.ext = {}); };
const eachCell = (fn, s) => { s = s || sel; for (let r = s.r1; r <= Math.min(s.r2, ROWS); r++) for (let c = s.c1; c <= Math.min(s.c2, COLS); c++) fn(r, c); };
const absRef = (s, sheet) => { s = s || sel; const a = (r, c) => '$' + E.numToCol(c) + '$' + r; const q = /^\w+$/.test(sheet || cur) ? (sheet || cur) : "'" + (sheet || cur) + "'"; return q + '!' + (s.r1 === s.r2 && s.c1 === s.c2 ? a(s.r1, s.c1) : a(s.r1, s.c1) + ':' + a(s.r2, s.c2)); };
const esc = x => String(x === undefined || x === null ? '' : x).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/"/g, '&quot;');
function cellVal(r, c) { const v = wb.value(cur, r, c); return v === undefined ? null : v; }
function isBlank(v) { return v === null || v === undefined || v === ''; }
function regionSel() { if (sel.r1 === sel.r2 && sel.c1 === sel.c2) return regionOf(anchor.r, anchor.c); return Object.assign({}, sel); }
function lastRowOf(c1, c2, from) { let m = from; for (let r = from; r <= ROWS; r++) for (let c = c1; c <= c2; c++) if (!isBlank(cellVal(r, c))) { m = r; } return m; }
/* extra cell classes (filter arrows, links, notes, highlights) composed onto the marching-ants hook */
const EXTRA = [];
const _ant0 = antCls;
antCls = function (r, c) { let s = _ant0(r, c); for (const f of EXTRA) s += f(r, c); return s; };
const HL = { set: null };
EXTRA.push((r, c) => HL.set && HL.set.has(r + ',' + c) ? ' hlc' : '');
(function () {
  const st = document.createElement('style');
  st.textContent = `.g td.hlc{background:#ffe08a!important}.g td.afh{position:relative}.g td.afh::after{content:'\\25BC';position:absolute;right:1px;bottom:1px;width:15px;height:15px;line-height:15px;font-size:9px;text-align:center;background:#f3f3f3;border:1px solid #888;color:#333;z-index:2;box-sizing:border-box}.g td.afh.afa::after{background:#cfe8d9;color:#217346;border-color:#217346}
.g td.lnk{color:#0563c1!important;text-decoration:underline}.g td.nte{position:relative}.g td.nte::before{content:'';position:absolute;right:0;top:0;border-style:solid;border-width:0 8px 8px 0;border-color:transparent #d00 transparent transparent;z-index:2}
#fh{position:fixed;width:22px;height:22px;display:none;z-index:55;cursor:crosshair;touch-action:none}#fh i{position:absolute;right:5px;bottom:5px;width:8px;height:8px;background:#217346;border:1px solid #fff;box-sizing:content-box}
#ctxm,#afm{position:fixed;z-index:80;display:none;min-width:170px;max-height:70vh;overflow:auto;background:#fff;border:1px solid #888;box-shadow:2px 3px 8px rgba(0,0,0,.35);font-size:13px;padding:2px 0}
#ctxm div{padding:7px 16px;cursor:pointer;white-space:nowrap}#ctxm div:hover,#ctxm div:active{background:#cfe8d9}#ctxm hr{border:0;border-top:1px solid #ddd;margin:2px 0}
#afm{padding:6px;width:230px}#afm .row{padding:5px 6px;cursor:pointer;border-radius:3px}#afm .row:hover{background:#eef6f1}#afm input[type=text]{width:100%;box-sizing:border-box;margin:4px 0}#afm .lst{max-height:200px;overflow:auto;border:1px solid #ccc;padding:2px}#afm label{display:flex;gap:6px;align-items:center;padding:2px 4px}
#ntTip{position:fixed;z-index:60;display:none;max-width:260px;background:#ffffe1;border:1px solid #777;padding:5px 8px;font-size:12px;white-space:pre-wrap;box-shadow:1px 1px 3px rgba(0,0,0,.3)}`;
  document.head.appendChild(st);
})();
function menuEl(id) { let e = document.getElementById(id); if (!e) { e = document.createElement('div'); e.id = id; document.body.appendChild(e); } return e; }
function showAt(m, x, y) { m.style.display = 'block'; const W = innerWidth, H = innerHeight; m.style.left = Math.max(4, Math.min(x, W - m.offsetWidth - 4)) + 'px'; m.style.top = Math.max(4, Math.min(y, H - m.offsetHeight - 4)) + 'px'; }
function closeMenus() { for (const id of ['ctxm', 'afm']) { const m = document.getElementById(id); if (m) m.style.display = 'none'; } }
document.addEventListener('pointerdown', e => { if (!(e.target.closest && (e.target.closest('#ctxm') || e.target.closest('#afm')))) closeMenus(); if (HL.set && !(e.target.closest && e.target.closest('#mGen'))) clearHL(); }, true);
function clearHL() { if (!HL.set) return; const old = HL.set; HL.set = null; for (const k of old) { const [r, c] = k.split(',').map(Number); paint(r, c); } }
function setHL(keys) { clearHL(); HL.set = new Set(keys); for (const k of HL.set) { const [r, c] = k.split(',').map(Number); paint(r, c); } }

/* ---------------- unified editing guard: protection ---------------- */
function protOn(sheet) { const S = wb.sheets[sheet || cur]; return !!(S && S.ext && S.ext.prot && S.ext.prot.on); }
function isLocked(sheet, r, c) { const S = wb.sheets[sheet]; return !(S && S.ext && S.ext.unl && S.ext.unl[r + ',' + c]); }
const _setRaw0 = E.Workbook.prototype.setRaw;
E.Workbook.prototype.setRaw = function (sheet, r, c, raw) { if (!PROT_OFF && protOn(sheet) && isLocked(sheet, r, c)) { PROT_HIT = true; return; } return _setRaw0.call(this, sheet, r, c, raw); };
let PROT_OFF = false, PROT_HIT = false;
function protMsg() { stMsg('The cell is protected. Review > Unprotect Sheet to edit.'); try { alert('The cell or chart you are trying to change is on a protected sheet.\nTo make a change, unprotect the sheet (Review > Unprotect Sheet).'); } catch (e) { } }
const _commitFx4 = commitFx;
commitFx = function (m) { if ((editing || editCell)) { const ec = editCell || { sheet: cur, r: anchor.r, c: anchor.c }; if (protOn(ec.sheet) && isLocked(ec.sheet, ec.r, ec.c)) { protMsg(); cancelEdit(); return; } } _commitFx4(m); };
const _clearSel = clearSel; clearSel = function () { PROT_HIT = false; _clearSel(); if (PROT_HIT) protMsg(); };
function protectDlg() {
  if (protOn()) { const pw = X().prot.pw; const go = () => { X().prot = { on: false }; save(); closeGen(); stMsg('Sheet unprotected'); }; if (!pw) { go(); return; } gen('Unprotect Sheet', '<label>Password</label><input type="password" id="ppw">', [['OK', () => { if (hashPw($('ppw').value) === pw) go(); else alert('The password you supplied is not correct.'); }]]); return; }
  gen('Protect Sheet', '<label>Password (optional)</label><input type="password" id="ppw"><p style="font-size:12px;color:#555">Locked cells cannot be edited, cleared, pasted over, filled or sorted. Use Review > Lock Cell to unlock the cells people may edit (cells are locked by default).</p>', [['Protect', () => { const p = $('ppw').value; X().prot = { on: true, pw: p ? hashPw(p) : '' }; save(); closeGen(); stMsg('Sheet protected'); }]]);
}
function hashPw(s) { let h = 5381; for (const ch of String(s)) h = ((h << 5) + h + ch.charCodeAt(0)) | 0; return String(h); }
function lockCells() { const U = X().unl = X().unl || {}; const allLocked = (() => { let a = true; eachCell((r, c) => { if (U[r + ',' + c]) a = false; }); return a; })(); eachCell((r, c) => { if (allLocked) U[r + ',' + c] = 1; else delete U[r + ',' + c]; }); save(); stMsg(allLocked ? 'Selected cells unlocked' : 'Selected cells locked'); }

/* ---------------- hyperlinks ---------------- */
EXTRA.push((r, c) => { const L = X().links; if (L && L[r + ',' + c]) return ' lnk'; const raw = wb.getRaw(cur, r, c); return typeof raw === 'string' && /^=HYPERLINK\(/i.test(raw) ? ' lnk' : ''; });
function linkDlg() {
  const k = anchor.r + ',' + anchor.c, L = (X().links || {})[k] || {};
  gen('Insert Hyperlink', `<label>Text to display</label><input type="text" id="hlT" value="${esc(L.text || (wb.getRaw(cur, anchor.r, anchor.c) + ''))}"><label>Address (https://… , mailto:… or a place like Sheet2!A1)</label><input type="text" id="hlA" value="${esc(L.url || '')}">`, [['OK', () => { const u = $('hlA').value.trim(), t = $('hlT').value; if (!u) return; const Lk = X().links = X().links || {}; Lk[k] = { url: u, text: t }; if (t) wb.setRaw(cur, anchor.r, anchor.c, t); save(); closeGen(); refresh(); }], ['Remove link', () => { if (X().links) delete X().links[k]; save(); closeGen(); refresh(); }]]);
}
function openLink(r, c) {
  const k = r + ',' + c; let u = ((X().links || {})[k] || {}).url; if (!u) { const raw = wb.getRaw(cur, r, c); const m = typeof raw === 'string' && /^=HYPERLINK\(\s*"([^"]*)"/i.exec(raw); if (m) u = m[1]; }
  if (!u) return false;
  const pl = /^#?(?:'([^']+)'|([^'!]+))?!?([A-Za-z]+\d+(?::[A-Za-z]+\d+)?)$/.exec(u.replace(/^#/, ''));
  if (!/^[a-z]+:/i.test(u) && pl) { const sn = pl[1] || pl[2]; if (u.includes('!') && sn && wb.sheets[sn] && sn !== cur) gotoSheet(sn); setTimeout(() => gotoName(pl[3]), 30); return true; }
  try { if (window.RB && RB.openUrl) RB.openUrl(u); else window.open(/^[a-z]+:/i.test(u) ? u : 'https://' + u, '_blank'); } catch (e) { } return true;
}
let _pdAnchor = null;
$('grid').addEventListener('pointerdown', e => { _pdAnchor = { r: anchor.r, c: anchor.c }; }, true);
$('grid').addEventListener('click', e => { const td = e.target.closest && e.target.closest('td'); if (!td || design) return; const r = +td.dataset.r, c = +td.dataset.c; if (!(EXTRA[2](r, c))) return; if (e.ctrlKey || e.metaKey || (_pdAnchor && _pdAnchor.r === r && _pdAnchor.c === c)) openLink(r, c); });

/* ---------------- notes ---------------- */
EXTRA.push((r, c) => (X().notes && X().notes[r + ',' + c]) ? ' nte' : '');
function noteDlg() {
  const k = anchor.r + ',' + anchor.c, N = X().notes = X().notes || {};
  gen(N[k] ? 'Edit Note' : 'New Note', `<textarea id="ntX" rows="5" style="width:100%;box-sizing:border-box">${esc(N[k] || '')}</textarea>`, [['OK', () => { const v = $('ntX').value; if (v) N[k] = v; else delete N[k]; save(); closeGen(); refresh(); ntUpdate(); }], ['Delete note', () => { delete N[k]; save(); closeGen(); refresh(); ntUpdate(); }]]);
}
function ntUpdate() {
  const t = menuEl('ntTip'); t.style.display = 'none'; if (editing) return; const n = X().notes && X().notes[anchor.r + ',' + anchor.c]; if (!n) return;
  const td = tds[anchor.r] && tds[anchor.r][anchor.c]; if (!td) return; const R = td.getBoundingClientRect(); t.id = 'ntTip'; t.textContent = n; t.style.display = 'block'; t.style.left = Math.min(R.right + 4, innerWidth - 270) + 'px'; t.style.top = R.top + 'px';
}
const _paintSel3 = paintSel; paintSel = function (p) { _paintSel3(p); ntUpdate(); fhUpdate(); };

/* ---------------- fill handle / AutoFill ---------------- */
const DAYS_S = ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'], DAYS_L = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];
const MON_S = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'], MON_L = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
const LISTS = [DAYS_S, DAYS_L, MON_S, MON_L];
function caseLike(sample, s) { if (sample === sample.toUpperCase()) return s.toUpperCase(); if (sample === sample.toLowerCase()) return s.toLowerCase(); return s; }
function isDateFmt(f) { return !!f && /[dmy]/i.test(String(f).replace(/"[^"]*"/g, '').replace(/\[[^\]]*\]/g, '')) && !/^[#0,.\s%$-]*$/.test(f) && !/0\.0|#,#|E\+/.test(f); }
function seriesGen(items) { // items: [{raw, v, fmt}] in fill order; returns fn(k) -> raw for k>=1 ; or null for plain copy
  const n = items.length; const nums = items.every(x => typeof x.v === 'number' && !(typeof x.raw === 'string' && x.raw[0] === '='));
  if (nums) {
    if (n === 1) { if (isDateFmt(items[0].fmt)) return k => items[0].v + k; return null; }
    const step = (items[n - 1].v - items[0].v) / (n - 1); return k => +(items[n - 1].v + step * k).toPrecision(14);
  }
  if (items.some(x => typeof x.raw === 'string' && x.raw[0] === '=')) return null;
  const strs = items.map(x => String(x.raw));
  for (const L of LISTS) { const idx = strs.map(s => L.findIndex(x => x.toLowerCase() === s.toLowerCase())); if (idx.every(i => i >= 0)) { const step = n > 1 ? (idx[n - 1] - idx[0]) || 1 : 1; return k => caseLike(strs[n - 1], L[(((idx[n - 1] + step * k) % L.length) + L.length) % L.length]); } }
  const ms = strs.map(s => /^(.*?)(\d+)$/.exec(s)); if (ms.every(m => m) && ms.every(m => m[1] === ms[0][1])) { const nn = ms.map(m => +m[2]); const step = n > 1 ? (nn[n - 1] - nn[0]) / (n - 1) : 1; const w = ms[n - 1][2].length; return k => ms[0][1] + String(Math.max(0, Math.round(nn[n - 1] + step * k))).padStart(/^0/.test(ms[n - 1][2]) ? w : 0, '0'); }
  return null;
}
function autoFill(src, dir, count) {
  if (count < 1) return; const vert = dir === 'down' || dir === 'up', fwd = dir === 'down' || dir === 'right';
  const lines = vert ? src.c2 - src.c1 + 1 : src.r2 - src.r1 + 1, len = vert ? src.r2 - src.r1 + 1 : src.c2 - src.c1 + 1;
  for (let li = 0; li < lines; li++) {
    const pos = i => vert ? { r: fwd ? src.r1 + i : src.r2 - i, c: src.c1 + li } : { r: src.r1 + li, c: fwd ? src.c1 + i : src.c2 - i };
    const items = []; for (let i = 0; i < len; i++) { const p = pos(i); items.push({ raw: wb.getRaw(cur, p.r, p.c), v: cellVal(p.r, p.c), fmt: sh().fmt[key(p.r, p.c)], st: styleOf(p.r, p.c), p }); }
    const gen_ = seriesGen(items);
    for (let k = 1; k <= count; k++) {
      const i = (k - 1) % len, it = items[i]; const tp = vert ? { r: fwd ? src.r2 + k : src.r1 - k, c: it.p.c } : { r: it.p.r, c: fwd ? src.c2 + k : src.c1 - k };
      if (tp.r < 1 || tp.c < 1 || tp.r > ROWS || tp.c > COLS) continue; let raw;
      if (gen_) raw = gen_(k); else if (typeof it.raw === 'string' && it.raw[0] === '=') { const off = (Math.floor((k - 1) / len) + 1) * len * (fwd ? 1 : -1); raw = shiftFormula(it.raw, vert ? off : 0, vert ? 0 : off); } else raw = it.raw;
      wb.setRaw(cur, tp.r, tp.c, raw); const kk = key(tp.r, tp.c); if (it.st) sh().style[kk] = Object.assign({}, it.st); else delete sh().style[kk]; if (it.fmt) sh().fmt[kk] = it.fmt; else if (typeof raw === 'number' && gen_ && it.fmt) sh().fmt[kk] = it.fmt; else delete sh().fmt[kk];
    }
  }
  save(); refresh();
}
function fhEl() { let e = document.getElementById('fh'); if (!e) { e = document.createElement('div'); e.id = 'fh'; e.innerHTML = '<i></i>'; document.body.appendChild(e); } return e; }
function fhUpdate() {
  const h = fhEl(); h.style.display = 'none'; if (editing || design || FHD) return; const td = tds[Math.min(sel.r2, ROWS)] && tds[Math.min(sel.r2, ROWS)][Math.min(sel.c2, COLS)]; if (!td) return;
  const R = td.getBoundingClientRect(), W = $('wrap').getBoundingClientRect(); if (R.bottom < W.top + 24 || R.bottom > W.bottom + 2 || R.right < W.left + 40 || R.right > W.right + 2) return;
  h.style.display = 'block'; h.style.left = (R.right - 17) + 'px'; h.style.top = (R.bottom - 17) + 'px';
}
let FHD = null;
fhEl().addEventListener('pointerdown', e => { e.preventDefault(); e.stopPropagation(); FHD = { src: Object.assign({}, sel), cur: null }; try { fhEl().setPointerCapture(e.pointerId); } catch (x) { } });
fhEl().addEventListener('dblclick', e => { // fill down alongside neighbouring data
  const s = sel; let nc = s.c1 - 1; if (nc < 1 || isBlank(cellVal(s.r2, nc))) nc = s.c2 + 1; if (nc > COLS || isBlank(cellVal(s.r2, nc))) return; let r = s.r2; while (r + 1 <= ROWS && !isBlank(cellVal(r + 1, nc))) r++; if (r > s.r2) { const src = Object.assign({}, s); autoFill(src, 'down', r - s.r2); setSel(src.r1, src.c1, r, src.c2, src.r1, src.c1); }
});
window.addEventListener('pointermove', e => {
  if (!FHD) return; const el = document.elementFromPoint(e.clientX, e.clientY); const td = el && el.closest && el.closest('td'); if (!td) return; const r = +td.dataset.r, c = +td.dataset.c, s = FHD.src;
  const dr = r > s.r2 ? r - s.r2 : r < s.r1 ? r - s.r1 : 0, dc = c > s.c2 ? c - s.c2 : c < s.c1 ? c - s.c1 : 0; let dir = null, n = 0;
  if (Math.abs(dr) >= Math.abs(dc) && dr !== 0) { dir = dr > 0 ? 'down' : 'up'; n = Math.abs(dr); } else if (dc !== 0) { dir = dc > 0 ? 'right' : 'left'; n = Math.abs(dc); }
  FHD.dir = dir; FHD.n = n; const u = Object.assign({}, s); if (dir === 'down') u.r2 += n; else if (dir === 'up') u.r1 -= n; else if (dir === 'right') u.c2 += n; else if (dir === 'left') u.c1 -= n; FHD.u = u;
  const prev = Object.assign({}, sel); sel = { r1: u.r1, c1: u.c1, r2: u.r2, c2: u.c2 }; paintSel(prev); stMsg(dir ? 'Fill ' + dir : 'Ready');
});
window.addEventListener('pointerup', () => { if (!FHD) return; const f = FHD; FHD = null; if (f.dir && f.n) { sel = Object.assign({}, f.src); autoFill(f.src, f.dir, f.n); const u = f.u; sel = { r1: u.r1, c1: u.c1, r2: u.r2, c2: u.c2 }; refresh(); } fhUpdate(); });
$('wrap').addEventListener('scroll', () => { fhUpdate(); ntUpdate(); closeMenus(); });
window.addEventListener('resize', () => { fhUpdate(); });
setTimeout(fhUpdate, 400);

/* ---------------- context menu (right-click / long-press) ---------------- */
function ctxItems(kind) {
  const sep = ['-'];
  if (kind === 'col') return [['Cut', () => copy(true)], ['Copy', () => copy(false)], ['Paste', () => paste()], ['Paste Special…', () => pasteSpecialDlg()], sep, ['Insert column(s)', () => insRC('c')], ['Delete column(s)', () => delRC('c')], ['Clear contents', () => clearSel()], sep, ['Column width…', () => colWidthDlg()], ['AutoFit column width', () => autoFitCol()], ['Hide', () => hideRC('c')], ['Unhide', () => unhideRC('c')]];
  if (kind === 'row') return [['Cut', () => copy(true)], ['Copy', () => copy(false)], ['Paste', () => paste()], ['Paste Special…', () => pasteSpecialDlg()], sep, ['Insert row(s)', () => insRC('r')], ['Delete row(s)', () => delRC('r')], ['Clear contents', () => clearSel()], sep, ['Row height…', () => rowHeightDlg()], ['AutoFit row height', () => autoFitRow()], ['Hide', () => hideRC('r')], ['Unhide', () => unhideRC('r')]];
  const k = anchor.r + ',' + anchor.c, hasNote = !!(X().notes && X().notes[k]), hasLink = !!(X().links && X().links[k]);
  return [['Cut', () => copy(true)], ['Copy', () => copy(false)], ['Paste', () => paste()], ['Paste Special…', () => pasteSpecialDlg()], sep, ['Insert row above', () => insRC('r')], ['Insert column left', () => insRC('c')], ['Delete row(s)', () => delRC('r')], ['Delete column(s)', () => delRC('c')], ['Clear contents', () => clearSel()], sep,
    ['Sort A to Z', () => sortSel(1)], ['Sort Z to A', () => sortSel(-1)], ['Sort…', () => sortDlg()], ['Filter', () => filterToggle()], sep, ['Format Cells…', () => formatCellsDlg()], ['Data Validation…', () => dvDlg()], ['Define Name…', () => nameNew()],
    [hasLink ? 'Edit Hyperlink…' : 'Hyperlink…', () => linkDlg()], hasLink ? ['Open Hyperlink', () => openLink(anchor.r, anchor.c)] : null, [hasNote ? 'Edit Note…' : 'New Note…', () => noteDlg()], hasNote ? ['Delete Note', () => { delete X().notes[k]; save(); refresh(); ntUpdate(); }] : null,
    sep, ['Row height…', () => rowHeightDlg()], ['Column width…', () => colWidthDlg()]].filter(Boolean);
}
function showCtx(x, y, kind) {
  const m = menuEl('ctxm'); m.innerHTML = ''; for (const it of ctxItems(kind)) { if (it[0] === '-') { m.appendChild(document.createElement('hr')); continue; } const d = document.createElement('div'); d.textContent = it[0]; d.onclick = () => { closeMenus(); try { it[1](); } catch (e) { console.log(e); } }; m.appendChild(d); }
  showAt(m, x, y);
}
$('grid').addEventListener('contextmenu', e => {
  const th = e.target.closest && e.target.closest('th'), td = e.target.closest && e.target.closest('td'); e.preventDefault();
  if (td) { const r = +td.dataset.r, c = +td.dataset.c; if (!(r >= sel.r1 && r <= sel.r2 && c >= sel.c1 && c <= sel.c2)) goCell(r, c); showCtx(e.clientX, e.clientY, 'cell'); }
  else if (th && /^ch\d+$/.test(th.id)) { const c = +th.id.slice(2); if (!(c >= sel.c1 && c <= sel.c2 && sel.r1 === 1 && sel.r2 >= ROWS)) setSel(1, c, ROWS, c, 1, c); showCtx(e.clientX, e.clientY, 'col'); }
  else if (th && /^rh\d+$/.test(th.id)) { const r = +th.id.slice(2); if (!(r >= sel.r1 && r <= sel.r2 && sel.c1 === 1 && sel.c2 >= COLS)) setSel(r, 1, r, COLS, r, 1); showCtx(e.clientX, e.clientY, 'row'); }
});
(function () { // long-press for touch
  let t = null, sx = 0, sy = 0, tgt = null;
  $('grid').addEventListener('pointerdown', e => { if (e.pointerType !== 'touch') return; sx = e.clientX; sy = e.clientY; tgt = e.target; clearTimeout(t); t = setTimeout(() => { t = null; const td = tgt.closest && tgt.closest('td'); if (!td) return; const r = +td.dataset.r, c = +td.dataset.c; if (!(r >= sel.r1 && r <= sel.r2 && c >= sel.c1 && c <= sel.c2)) goCell(r, c); showCtx(sx, sy, 'cell'); }, 550); }, true);
  const cancel = e => { if (t && (e.type !== 'pointermove' || Math.hypot(e.clientX - sx, e.clientY - sy) > 8)) { clearTimeout(t); t = null; } };
  window.addEventListener('pointermove', cancel, true); window.addEventListener('pointerup', cancel, true); window.addEventListener('pointercancel', cancel, true);
})();
document.addEventListener('keydown', e => { if (e.key === 'F10' && e.shiftKey) { e.preventDefault(); const td = tds[anchor.r][anchor.c], R = td.getBoundingClientRect(); showCtx(R.left + 20, R.bottom, 'cell'); } else if (e.key === 'ContextMenu') { e.preventDefault(); const td = tds[anchor.r][anchor.c], R = td.getBoundingClientRect(); showCtx(R.left + 20, R.bottom, 'cell'); } else if (e.key === 'Escape') closeMenus(); }, true);

/* ---------------- sort (generic) & Sort dialog ---------------- */
function cmpSort(a, b, desc) {
  const rank = v => isBlank(v) ? 9 : typeof v === 'number' ? 0 : typeof v === 'string' ? 1 : typeof v === 'boolean' ? 2 : 3;
  const ra = rank(a), rb = rank(b); if (ra === 9 || rb === 9) return ra === rb ? 0 : ra === 9 ? 1 : -1; // blanks always last
  let r; if (ra !== rb) r = ra - rb; else if (ra === 0) r = a - b; else if (ra === 1) r = a.toLowerCase().localeCompare(b.toLowerCase()); else r = String(a).localeCompare(String(b)); return desc ? -r : r;
}
function sortRange(r1, c1, r2, c2, levels) {
  if (r2 < r1) return; const rows = [];
  for (let r = r1; r <= r2; r++) { const cells = []; for (let c = c1; c <= c2; c++) cells.push({ raw: wb.getRaw(cur, r, c), st: styleOf(r, c), f: sh().fmt[key(r, c)] }); rows.push({ r, cells, keys: levels.map(l => cellVal(r, l.col)) }); }
  rows.sort((A, B) => { for (let i = 0; i < levels.length; i++) { const d = cmpSort(A.keys[i], B.keys[i], levels[i].desc); if (d) return d; } return A.r - B.r; });
  rows.forEach((row, i) => { const nr = r1 + i; row.cells.forEach((x, j) => { const c = c1 + j, kk = key(nr, c); let raw = x.raw; if (typeof raw === 'string' && raw[0] === '=') raw = shiftFormula(raw, nr - row.r, 0); wb.setRaw(cur, nr, c, raw); if (x.st) sh().style[kk] = Object.assign({}, x.st); else delete sh().style[kk]; if (x.f) sh().fmt[kk] = x.f; else delete sh().fmt[kk]; }); });
  save(); refresh();
}
function guessHeader(R) { if (R.r2 <= R.r1) return false; let allText = true; for (let c = R.c1; c <= R.c2; c++) { const v = cellVal(R.r1, c); if (typeof v !== 'string' || v === '') allText = false; } if (!allText) return false; for (let c = R.c1; c <= R.c2; c++) { const v = cellVal(R.r1 + 1, c); if (typeof v !== 'string' && !isBlank(v)) return true; } return false; }
sortSel = function (dir) {
  const R = regionSel(); if (R.r2 === R.r1) return; const hdr = guessHeader(R); const col = (anchor.c >= R.c1 && anchor.c <= R.c2) ? anchor.c : R.c1;
  sortRange(R.r1 + (hdr ? 1 : 0), R.c1, R.r2, R.c2, [{ col, desc: dir < 0 }]);
};
let SD = null;
function sortDlg() {
  const R = regionSel(); SD = { R, hdr: guessHeader(R), lv: [{ col: (anchor.c >= R.c1 && anchor.c <= R.c2) ? anchor.c : R.c1, desc: false }] }; sortRender();
}
function sortRender() {
  const R = SD.R; const opts = v => { let h = ''; for (let c = R.c1; c <= R.c2; c++) { const t = SD.hdr ? String(cellVal(R.r1, c) ?? '') : ''; h += `<option value="${c}" ${c === v ? 'selected' : ''}>${SD.hdr && t ? esc(t) : 'Column ' + E.numToCol(c)}</option>`; } return h; };
  const rows = SD.lv.map((l, i) => `<div style="display:flex;gap:4px;margin:5px 0;align-items:center"><span style="width:52px;font-size:12px">${i ? 'Then by' : 'Sort by'}</span><select onchange="SD.lv[${i}].col=+this.value">${opts(l.col)}</select><select onchange="SD.lv[${i}].desc=this.value==='d'"><option value="a" ${l.desc ? '' : 'selected'}>A → Z / smallest</option><option value="d" ${l.desc ? 'selected' : ''}>Z → A / largest</option></select>${SD.lv.length > 1 ? `<button onclick="SD.lv.splice(${i},1);sortRender()">✕</button>` : ''}</div>`).join('');
  gen('Sort', `<div style="font-size:12px;color:#555">Range ${E.addr(R.r1, R.c1)}:${E.addr(R.r2, R.c2)}</div><label><input type="checkbox" id="sdH" ${SD.hdr ? 'checked' : ''} onchange="SD.hdr=this.checked;sortRender()"> My data has headers</label>${rows}<button onclick="SD.lv.push({col:SD.R.c1,desc:false});sortRender()">+ Add level</button>`, [['OK', () => { sortRange(R.r1 + (SD.hdr ? 1 : 0), R.c1, R.r2, R.c2, SD.lv); closeGen(); }]]);
}

/* ---------------- AutoFilter with header arrows ---------------- */
const filterCond = filterDlg;
EXTRA.push((r, c) => { const a = X().af; if (!a || r !== a.r || c < a.c1 || c > a.c2) return ''; return ' afh' + (X().afv && X().afv[c] ? ' afa' : ''); });
function filterToggle() {
  if (X().af) { _clearFilter0(); X().af = null; X().afv = {}; save(); refresh(); stMsg('Filter off'); return; }
  const R = regionSel(); X().af = { r: R.r1, c1: R.c1, c2: R.c2 }; X().afv = {}; save(); refresh(); stMsg('Filter on: click a ▼ in the header row');
}
filterDlg = filterToggle;
function applyAf() {
  const S = sh(), a = X().af; if (!a) return; _clearFilter0(); const last = lastRowOf(a.c1, a.c2, a.r + 1); const act = Object.keys(X().afv || {}).filter(c => X().afv[c]);
  S.flt = S.flt || {}; for (let r = a.r + 1; r <= last; r++) { let ok = true; for (const c of act) { if (!new Set(X().afv[c]).has(cellText(r, +c))) { ok = false; break; } } if (!ok) { S.flt[r] = S.rowH[r] === undefined ? (S.defSz !== undefined ? 20 : DH) : S.rowH[r]; S.rowH[r] = 0; } }
  save(); build();
}
$('grid').addEventListener('pointerdown', e => {
  const td = e.target.closest && e.target.closest('td'); if (!td || !td.classList.contains('afh')) return; const R = td.getBoundingClientRect(); if (R.right - e.clientX > 20 || R.bottom - e.clientY > 20) return;
  e.preventDefault(); e.stopPropagation(); afMenu(+td.dataset.c, R.left, R.bottom);
}, true);
function afMenu(c, x, y) {
  const a = X().af, last = lastRowOf(a.c1, a.c2, a.r + 1); const vals = []; const seen = new Set(); for (let r = a.r + 1; r <= last; r++) { const t = cellText(r, c); if (!seen.has(t)) { seen.add(t); vals.push(t); } }
  vals.sort((p, q) => { const n1 = parseFloat(p), n2 = parseFloat(q); if (!isNaN(n1) && !isNaN(n2)) return n1 - n2; return p.toLowerCase().localeCompare(q.toLowerCase()); });
  const cur0 = X().afv && X().afv[c] ? new Set(X().afv[c]) : null; const m = menuEl('afm'); m.id = 'afm';
  const hdrName = String(cellVal(a.r, c) ?? E.numToCol(c));
  m.innerHTML = `<div class="row" id="afA">Sort A to Z</div><div class="row" id="afD">Sort Z to A</div><div class="row" id="afC">Clear filter from "${esc(hdrName.slice(0, 18))}"</div><input type="text" id="afQ" placeholder="Search"><div class="lst" id="afL"></div><div style="display:flex;gap:6px;justify-content:flex-end;margin-top:6px"><button id="afX">Cancel</button><button id="afO" style="background:#217346;color:#fff">OK</button></div>`;
  const st = new Set(cur0 || vals); const L = m.querySelector('#afL');
  const draw = () => { const q = m.querySelector('#afQ').value.toLowerCase(); const shown = vals.filter(v => v.toLowerCase().includes(q)); L.innerHTML = `<label><input type="checkbox" id="afAll" ${shown.every(v => st.has(v)) ? 'checked' : ''}> (Select All)</label>` + shown.slice(0, 400).map((v, i) => `<label><input type="checkbox" data-v="${i}" ${st.has(v) ? 'checked' : ''}> ${esc(v === '' ? '(Blanks)' : v)}</label>`).join(''); L.querySelector('#afAll').onchange = e2 => { shown.forEach(v => e2.target.checked ? st.add(v) : st.delete(v)); draw(); }; L.querySelectorAll('input[data-v]').forEach(i => i.onchange = () => { const v = shown[+i.dataset.v]; i.checked ? st.add(v) : st.delete(v); }); };
  draw(); m.querySelector('#afQ').oninput = draw;
  m.querySelector('#afA').onclick = () => { closeMenus(); sortRange(a.r + 1, a.c1, last, a.c2, [{ col: c, desc: false }]); };
  m.querySelector('#afD').onclick = () => { closeMenus(); sortRange(a.r + 1, a.c1, last, a.c2, [{ col: c, desc: true }]); };
  m.querySelector('#afC').onclick = () => { closeMenus(); delete X().afv[c]; applyAf(); refresh(); };
  m.querySelector('#afX').onclick = closeMenus;
  m.querySelector('#afO').onclick = () => { closeMenus(); if (vals.every(v => st.has(v))) delete X().afv[c]; else X().afv[c] = [...st]; applyAf(); refresh(); };
  showAt(m, Math.min(x, innerWidth - 240), y);
}
const _clearFilter0 = clearFilter;

/* ---------------- Go To / Go To Special / Find & Replace ---------------- */
function gotoDlg() {
  const sp = (id, t) => `<label><input type="radio" name="gsp" value="${id}"> ${t}</label>`;
  gen('Go To', `<label>Reference (cell, range or name)</label><input type="text" id="gtR" placeholder="e.g. B5 or A1:C9 or MyName"><div style="margin-top:8px;font-weight:600">Special</div>` + sp('blank', 'Blanks') + sp('const', 'Constants') + sp('formula', 'Formulas') + sp('error', 'Formulas returning errors') + sp('num', 'Numbers') + sp('text', 'Text') + sp('region', 'Current region') + sp('last', 'Last cell') + sp('dv', 'Data validation') + sp('cf', 'Conditional formats') + sp('note', 'Notes') + sp('vis', 'Visible cells only'), [['Go', () => { const v = $('gtR').value.trim(), sp_ = (document.querySelector('input[name=gsp]:checked') || {}).value; closeGen(); if (sp_) goSpecial(sp_); else if (v) gotoName(v); }]]);
  setTimeout(() => { try { $('gtR').focus(); } catch (e) { } }, 50);
}
function goSpecial(kind) {
  const U = lastUsed(); const R = (sel.r1 === sel.r2 && sel.c1 === sel.c2) ? { r1: 1, c1: 1, r2: U.r, c2: U.c } : sel; const hits = [];
  if (kind === 'region') { selRegion(); return; } if (kind === 'last') { goCell(U.r, U.c); return; }
  eachCell((r, c) => {
    const raw = wb.getRaw(cur, r, c), v = cellVal(r, c), isF = typeof raw === 'string' && raw[0] === '=';
    const ok = kind === 'blank' ? isBlank(raw) : kind === 'const' ? (!isBlank(raw) && !isF) : kind === 'formula' ? isF : kind === 'error' ? (isF && E.isErr(v)) : kind === 'num' ? (typeof v === 'number') : kind === 'text' ? (typeof v === 'string' && v !== '') : kind === 'dv' ? !!dvRuleAt(cur, r, c) : kind === 'note' ? !!(X().notes && X().notes[r + ',' + c]) : kind === 'cf' ? (sh().cf || []).some(q => r >= q.r1 && r <= q.r2 && c >= q.c1 && c <= q.c2) : kind === 'vis' ? (sh().rowH[r] !== 0 && sh().colW[c] !== 0) : false;
    if (ok) hits.push(r + ',' + c);
  }, R);
  if (!hits.length) { stMsg('No cells found'); alert('No cells were found.'); return; } const [fr, fc] = hits[0].split(',').map(Number); goCell(fr, fc); setHL(hits); stMsg(hits.length + ' cell(s) found (highlighted)');
}
let FIND = { q: '', rp: '', book: false, mc: false, whole: false, form: false, res: [] };
function findMatches() {
  const out = []; const q = FIND.mc ? FIND.q : FIND.q.toLowerCase(); if (!q) return out;
  for (const sn of (FIND.book ? wb.order : [cur])) { const S = wb.sheets[sn]; for (const k in S.cells) { const [r, c] = k.split(',').map(Number); const raw = wb.getRaw(sn, r, c); const v = FIND.form ? String(raw) : (sn === cur ? cellText(r, c) : String(wb.value(sn, r, c) ?? '')); const t = FIND.mc ? v : v.toLowerCase(); if (FIND.whole ? t === q : t.includes(q)) out.push({ sn, r, c, t: v }); } }
  out.sort((a, b) => wb.order.indexOf(a.sn) - wb.order.indexOf(b.sn) || a.r - b.r || a.c - b.c); return out;
}
function findDlg() {
  const ck = (id, t, v) => `<label style="display:inline-flex;gap:4px;margin-right:10px"><input type="checkbox" id="${id}" ${v ? 'checked' : ''}> ${t}</label>`;
  const rd = () => { FIND.q = $('fF').value; FIND.rp = $('fR').value; FIND.mc = $('fMC').checked; FIND.whole = $('fWh').checked; FIND.book = $('fWb').checked; FIND.form = $('fFm').checked; };
  gen('Find and Replace', `<label>Find what</label><input type="text" id="fF" value="${esc(FIND.q)}"><label>Replace with</label><input type="text" id="fR" value="${esc(FIND.rp)}"><div style="margin:6px 0">${ck('fMC', 'Match case', FIND.mc)}${ck('fWh', 'Match entire cell', FIND.whole)}</div><div>${ck('fWb', 'Whole workbook', FIND.book)}${ck('fFm', 'Look in formulas', FIND.form)}</div><div id="fRes" style="max-height:30vh;overflow:auto;margin-top:6px;font-size:12px"></div>`, [
    ['Find Next', () => { rd(); findNext(); }], ['Find All', () => { rd(); const m = findMatches(); FIND.res = m; $('fRes').innerHTML = m.length ? m.map((x, i) => `<div style="padding:5px;border-bottom:1px solid #eee" onclick="findGo(${i})">${esc(x.sn)}!${E.addr(x.r, x.c)} &nbsp; ${esc(String(x.t).slice(0, 40))}</div>`).join('') : 'No matches'; }],
    ['Replace', () => { rd(); const m = findMatches(); const t = m.find(x => x.sn === cur && x.r === anchor.r && x.c === anchor.c); if (t) doReplace(t); findNext(); }], ['Replace All', () => { rd(); let n = 0; for (const x of findMatches()) { doReplace(x); n++; } closeGen(); alert(n + ' replacement(s) made.'); }]]);
}
function doReplace(x) { const raw = String(wb.getRaw(x.sn, x.r, x.c)); let nv; if (FIND.whole) nv = FIND.rp; else { const re = new RegExp(FIND.q.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), FIND.mc ? 'g' : 'gi'); nv = raw.replace(re, () => FIND.rp); } wb.setRaw(x.sn, x.r, x.c, nv); save(); refresh(); }
function findGo(i) { const x = FIND.res[i]; if (!x) return; if (x.sn !== cur) gotoSheet(x.sn); goCell(x.r, x.c); closeGen(); }
function findNext() {
  const m = findMatches(); if (!m.length) { stMsg('Nothing found'); alert("We couldn't find what you were searching for."); return; }
  const idx = m.findIndex(x => (wb.order.indexOf(x.sn) > wb.order.indexOf(cur)) || (x.sn === cur && (x.r > anchor.r || (x.r === anchor.r && x.c > anchor.c)))); const t = m[idx < 0 ? 0 : idx]; if (t.sn !== cur) gotoSheet(t.sn); goCell(t.r, t.c); closeGen();
}
findRep = findDlg;

/* ---------------- defined names ---------------- */
function syncNames() { wb.names = book.names; }
const _refresh4 = refresh; refresh = function () { syncNames(); _refresh4(); }; syncNames();
const NAME_OK = u => /^[A-Z_][A-Z0-9_.]*$/.test(u) && !/^[A-Z]{1,3}\d+$/.test(u) && !/^R\d*C\d*$/.test(u);
function nameSet(n, ref) { n = String(n).trim().toUpperCase(); if (!NAME_OK(n)) { alert('Invalid name. Use letters, numbers, _ and . and do not look like a cell address.'); return false; } book.names[n] = ref; syncNames(); save(); return true; }
function nameNew() { gen('New Name', `<label>Name</label><input type="text" id="nnN"><label>Refers to</label><input type="text" id="nnR" value="${esc(absRef())}">`, [['OK', () => { if (nameSet($('nnN').value, $('nnR').value.replace(/^=/, ''))) { closeGen(); refresh(); } }]]); setTimeout(() => { try { $('nnN').focus(); } catch (e) { } }, 50); }
function nameMgr() {
  const L = Object.keys(book.names); gen('Name Manager', L.length ? L.map(n => `<div style="display:flex;gap:6px;align-items:center;padding:6px 0;border-bottom:1px solid #eee"><b style="min-width:80px">${esc(n)}</b><span style="flex:1;font-size:12px;word-break:break-all">${esc(book.names[n])}</span><button onclick="nameEdit('${n}')">Edit</button><button onclick="nameDel('${n}')">✕</button></div>`).join('') : '<p>No names yet. Type a name in the Name Box (left of the formula bar) while a range is selected, or press New.</p>', [['New…', () => { closeGen(); nameNew(); }]]);
}
function nameEdit(n) { gen('Edit Name', `<b>${esc(n)}</b><label>Refers to</label><input type="text" id="nnR" value="${esc(book.names[n])}">`, [['OK', () => { book.names[n] = $('nnR').value.replace(/^=/, ''); syncNames(); save(); closeGen(); refresh(); }]]); }
function nameDel(n) { delete book.names[n]; syncNames(); save(); nameMgr(); refresh(); }
function namePaste() { const L = Object.keys(book.names); gen('Paste Name', L.length ? L.map(n => `<div style="padding:9px;border-bottom:1px solid #eee" onclick="closeGen();insertFn('${n}')">${esc(n)} <span style="color:#777;font-size:11px">${esc(book.names[n])}</span></div>`).join('') : 'No names defined.', []); }
function nameFromSel() {
  gen('Create Names from Selection', `<label><input type="checkbox" id="nfT" checked> Top row</label><label><input type="checkbox" id="nfL"> Left column</label>`, [['OK', () => {
    const top = $('nfT').checked, left = $('nfL').checked; closeGen();
    if (top) for (let c = sel.c1 + (left ? 1 : 0); c <= sel.c2; c++) { const n = String(cellVal(sel.r1, c) ?? '').trim().replace(/\s+/g, '_').toUpperCase(); if (n && NAME_OK(n)) book.names[n] = absRef({ r1: sel.r1 + 1, c1: c, r2: sel.r2, c2: c }); }
    if (left) for (let r = sel.r1 + (top ? 1 : 0); r <= sel.r2; r++) { const n = String(cellVal(r, sel.c1) ?? '').trim().replace(/\s+/g, '_').toUpperCase(); if (n && NAME_OK(n)) book.names[n] = absRef({ r1: r, c1: sel.c1 + 1, r2: r, c2: sel.c2 }); }
    syncNames(); save(); refresh();
  }]]);
}
const _gotoName0 = gotoName;
gotoName = function (t) { const u = String(t).trim().toUpperCase(); if (NAME_OK(u) && !book.names[u]) { book.names[u] = absRef(); syncNames(); save(); stMsg('Name "' + u + '" defined for ' + addrOf(sel)); return; } _gotoName0(t); };

/* ---------------- Insert Function dialog ---------------- */
const CATS = {
  Math: 'SUM SUMIF SUMIFS SUMPRODUCT PRODUCT ROUND ROUNDUP ROUNDDOWN INT MOD ABS SQRT POWER CEILING FLOOR RAND RANDBETWEEN SEQUENCE SUBTOTAL AGGREGATE EXP LN LOG LOG10 PI SIGN TRUNC SUMSQ',
  Statistical: 'AVERAGE AVERAGEIF AVERAGEIFS COUNT COUNTA COUNTBLANK COUNTIF COUNTIFS MAX MAXIFS MIN MINIFS MEDIAN MODE LARGE SMALL RANK PERCENTILE STDEV STDEV.S STDEV.P VAR CORREL SLOPE INTERCEPT FORECAST',
  Logical: 'IF IFS IFERROR IFNA AND OR NOT XOR SWITCH TRUE FALSE LET',
  Text: 'CONCAT CONCATENATE TEXTJOIN LEFT RIGHT MID LEN FIND SEARCH SUBSTITUTE REPLACE UPPER LOWER PROPER TRIM CLEAN TEXT VALUE REPT EXACT CHAR CODE TEXTSPLIT TEXTBEFORE TEXTAFTER',
  'Lookup & Reference': 'VLOOKUP HLOOKUP XLOOKUP XMATCH LOOKUP INDEX MATCH CHOOSE OFFSET INDIRECT ROW COLUMN ROWS COLUMNS TRANSPOSE ADDRESS',
  'Date & Time': 'TODAY NOW DATE TIME YEAR MONTH DAY HOUR MINUTE SECOND WEEKDAY WEEKNUM EOMONTH EDATE DATEDIF DATEVALUE TIMEVALUE NETWORKDAYS WORKDAY DAYS',
  Financial: 'PMT PV FV NPV IRR RATE NPER SLN DB',
  Information: 'ISNUMBER ISTEXT ISBLANK ISERROR ISNA ISLOGICAL ISFORMULA ISEVEN ISODD NA N T TYPE',
  'Dynamic arrays': 'FILTER SORT SORTBY UNIQUE SEQUENCE RANDARRAY TAKE DROP VSTACK HSTACK TOCOL TOROW WRAPROWS CHOOSECOLS CHOOSEROWS'
};
const CATOF = {}; for (const k in CATS) CATS[k].split(' ').forEach(n => { CATOF[n] = CATOF[n] || k; });
let FNCAT = 'All';
function fnPick() {
  gen('Insert Function', `<input type="text" id="fq" placeholder="Search functions…" oninput="fnFilter()"><select id="fcat" onchange="FNCAT=this.value;fnFilter()" style="margin-top:4px">${['All', ...Object.keys(CATS), 'Other'].map(c => `<option ${c === FNCAT ? 'selected' : ''}>${c}</option>`).join('')}</select><div id="fl" style="max-height:44vh;overflow:auto;margin-top:6px"></div>`, []); fnFilter(); setTimeout(() => { try { $('fq').focus(); } catch (e) { } }, 50);
}
function fnFilter() {
  const q = $('fq').value.toUpperCase(); const names = Object.keys(E.FN).filter(n => (!q || n.includes(q)) && (FNCAT === 'All' || (CATOF[n] || 'Other') === FNCAT)).sort();
  $('fl').innerHTML = names.slice(0, 150).map(n => `<div style="padding:8px;border-bottom:1px solid #eee" onclick="closeGen();insertFn('${n}(')"><b>${n}</b> <span style="color:#666;font-size:12px">(${esc((SIG[n] || ''))})</span> <span style="float:right;color:#999;font-size:11px">${CATOF[n] || 'Other'}</span></div>`).join('') || 'No match';
}

/* ---------------- ribbon additions: Review tab, Insert/Data/Formulas extras ---------------- */
(function () {
  const rib = document.getElementById('rib'); if (!rib) return;
  const grp = (label, btns) => `<div class="grp"><div class="gb">${btns.map(b => `<button class="big" onclick="${b[1]}"><i>${b[2]}</i>${b[0]}</button>`).join('')}</div><div class="gl">${label}</div></div>`;
  if (!rib.querySelector('[data-t="Review"]')) { const d = document.createElement('div'); d.className = 'tb'; d.dataset.t = 'Review';
    d.innerHTML = grp('Notes', [['New Note', 'noteDlg()', '&#128221;'], ['Delete Note', 'noteDel()', '&#128465;']]) + grp('Links', [['Hyperlink', 'linkDlg()', '&#128279;']]) + grp('Protect', [['Protect Sheet', 'protectDlg()', '&#128274;'], ['Lock Cell', 'lockCells()', '&#128273;']]) + grp('Changes', [['Show Formulas', 'toggleShowF()', '&#402;']]); rib.appendChild(d); }
  const add = (tab, html) => { const t = rib.querySelector(`[data-t="${tab}"]`); if (t) t.insertAdjacentHTML('beforeend', html); };
  add('Data', grp('Sort', [['Sort…', 'sortDlg()', '&#8645;']]));
  add('Formulas', grp('Defined Names', [['Name Manager', 'nameMgr()', '&#127991;'], ['Define Name', 'nameNew()', '&#10133;'], ['From Selection', 'nameFromSel()', '&#9638;'], ['Use in Formula', 'namePaste()', '&#8618;']]) + grp('Function', [['Insert Function', 'fnPick()', '&#402;']]));
  add('Insert', grp('Links', [['Hyperlink', 'linkDlg()', '&#128279;'], ['Note', 'noteDlg()', '&#128221;']]));
})();
function noteDel() { const k = anchor.r + ',' + anchor.c; if (X().notes) delete X().notes[k]; save(); refresh(); ntUpdate(); }

/* ---------------- keyboard hooks ---------------- */
document.addEventListener('keydown', e => {
  if (modalOpen() || e.defaultPrevented) return; const ae = document.activeElement; if (/INPUT|TEXTAREA|SELECT/.test(ae.tagName)) return;
  const ctrl = e.ctrlKey || e.metaKey, k = e.key;
  if (k === 'F5' && !e.altKey) { e.preventDefault(); gotoDlg(); }
}, false);
