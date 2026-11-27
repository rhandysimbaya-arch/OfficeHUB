/* Excel-style charts: many types, title/legend/labels/axis titles, palettes, selectable+draggable, Chart Design dialog */
var CH_TYPES = [
  ['col', 'Clustered Column', '&#9641;'], ['colS', 'Stacked Column', '&#9636;'], ['colP', '100% Stacked Column', '&#9638;'],
  ['barH', 'Clustered Bar', '&#9644;'], ['barHS', 'Stacked Bar', '&#9645;'],
  ['line', 'Line', '&#128200;'], ['lineM', 'Line with Markers', '&#9679;'], ['area', 'Area', '&#9650;'], ['areaS', 'Stacked Area', '&#9651;'],
  ['pie', 'Pie', '&#129383;'], ['donut', 'Doughnut', '&#9678;'], ['scatter', 'Scatter (XY)', '&#8759;'], ['radar', 'Radar', '&#10023;'], ['combo', 'Combo (Column + Line)', '&#9881;']];
var CH_PAL = { Colorful: ['#4472C4', '#ED7D31', '#A5A5A5', '#FFC000', '#5B9BD5', '#70AD47', '#264478', '#9E480E'], Blue: ['#1F4E79', '#2E75B6', '#5B9BD5', '#9DC3E6', '#BDD7EE', '#0B2A4A', '#3B7DC4', '#7FB2E5'], Green: ['#375623', '#548235', '#70AD47', '#A9D18E', '#C5E0B4', '#1E3A12', '#6AA84F', '#93C47D'], Orange: ['#843C0C', '#C55A11', '#ED7D31', '#F4B183', '#F8CBAD', '#5B2A08', '#E69138', '#F6B26B'], Gray: ['#404040', '#7F7F7F', '#A5A5A5', '#BFBFBF', '#D9D9D9', '#262626', '#595959', '#8C8C8C'], Warm: ['#C00000', '#ED7D31', '#FFC000', '#7F6000', '#FF6666', '#843C0C', '#F4B183', '#BF9000'] };
function chOpt(s) { s.o = s.o || {}; const o = s.o; if (!o.kind0) { o.kind0 = 1; if (s.kind === 'bar') s.kind = 'col'; } return o; }
function chKind(s) { return s.kind === 'bar' ? 'col' : s.kind; }
function chSeries(s) {
  let d = chartData(s); const o = chOpt(s); if (o.sw) d = d[0].map((_, j) => d.map(r => r[j]));
  const names = d[0].slice(1).map(x => String(x ?? '')), rows = d.slice(1), cats = rows.map(r => r[0]);
  const vals = names.map((_, j) => rows.map(r => { const v = r[j + 1]; return typeof v === 'number' ? v : (v === '' || v == null || isNaN(+v) ? 0 : +v); }));
  return { names, cats, vals, rows };
}
function chNice(mn, mx, n) { if (mx === mn) { mx = mn + 1; } const rg = mx - mn, rough = rg / (n || 5), pw = Math.pow(10, Math.floor(Math.log10(rough))), f = rough / pw, st = (f <= 1 ? 1 : f <= 2 ? 2 : f <= 5 ? 5 : 10) * pw; const lo = Math.floor(mn / st) * st, hi = Math.ceil(mx / st) * st; const t = []; for (let v = lo; v <= hi + st / 2; v += st)t.push(+v.toFixed(10)); return { lo, hi, t }; }
function chFmt(v) { const a = Math.abs(v); if (a >= 1e9) return +(v / 1e9).toFixed(2) + 'B'; if (a >= 1e6) return +(v / 1e6).toFixed(2) + 'M'; if (a >= 1e4) return +(v / 1e3).toFixed(1) + 'K'; return String(+v.toFixed(2)); }
function drawChart(cv, s) {
  const g = cv.getContext('2d'), W = cv.width, H = cv.height, o = chOpt(s), K = chKind(s); g.fillStyle = '#fff'; g.fillRect(0, 0, W, H);
  let D; try { D = chSeries(s); } catch (e) { return; } if (!D.cats.length || !D.names.length) { g.fillStyle = '#888'; g.font = '22px sans-serif'; g.fillText('Select data with a header row', 20, H / 2); return; }
  const pal = CH_PAL[o.pal || 'Colorful'] || CH_PAL.Colorful, col = i => pal[i % pal.length], FS = 22, ff = 'px Segoe UI,Roboto,sans-serif';
  g.font = FS + ff; g.textBaseline = 'alphabetic';
  const pie = K === 'pie' || K === 'donut'; const ns = pie ? 1 : D.names.length;
  let top = 12, bot = 12, lef = 12, rig = 12;
  if (o.title) { g.font = 'bold ' + (FS + 4) + ff; g.fillStyle = '#333'; g.textAlign = 'center'; g.fillText(o.title, W / 2, 34); top = 52; }
  const legNames = pie ? D.cats.map(String) : D.names; const lg = o.legend === undefined ? (D.names.length > 1 || pie ? 'right' : 'none') : o.legend;
  g.font = FS + ff; let legW = 0; if (lg !== 'none') { legNames.forEach(n => legW = Math.max(legW, g.measureText(n.slice(0, 14)).width)); legW += 40; }
  const legBox = { };
  if (lg === 'right') rig += legW; else if (lg === 'left') lef += legW; else if (lg === 'bottom') bot += 34; else if (lg === 'top') top += 34;
  const drawLegend = () => {
    if (lg === 'none') return; g.font = FS + ff; g.textAlign = 'left'; g.fillStyle = '#333';
    if (lg === 'right' || lg === 'left') { let x = lg === 'right' ? W - legW - 4 : 6, y = Math.max(top + 8, (H - legNames.length * 32) / 2); legNames.forEach((n, i) => { g.fillStyle = col(i); g.fillRect(x, y - 14, 18, 18); g.fillStyle = '#333'; g.fillText(n.slice(0, 14), x + 26, y); y += 32; }); }
    else { const tot = legNames.reduce((a, n) => a + g.measureText(n.slice(0, 14)).width + 50, 0); let x = Math.max(8, (W - tot) / 2); const y = lg === 'top' ? top - 12 : H - 14; legNames.forEach((n, i) => { g.fillStyle = col(i); g.fillRect(x, y - 14, 18, 18); g.fillStyle = '#333'; g.fillText(n.slice(0, 14), x + 24, y); x += g.measureText(n.slice(0, 14)).width + 50; }); }
  };
  if (pie) {
    const v = D.vals[0].map(x => Math.max(0, x)), tot = v.reduce((a, b) => a + b, 0) || 1; const cx = lef + (W - lef - rig) / 2, cy = top + (H - top - bot) / 2, R = Math.min(W - lef - rig, H - top - bot) / 2 - 10; let a0 = -Math.PI / 2;
    v.forEach((x, i) => { const a1 = a0 + x / tot * 2 * Math.PI; g.beginPath(); g.moveTo(cx, cy); g.arc(cx, cy, R, a0, a1); g.closePath(); g.fillStyle = col(i); g.fill(); g.strokeStyle = '#fff'; g.lineWidth = 2; g.stroke(); if (o.labels && x > 0) { const m = (a0 + a1) / 2; g.fillStyle = '#fff'; g.font = 'bold ' + (FS - 2) + ff; g.textAlign = 'center'; g.fillText(Math.round(x / tot * 100) + '%', cx + Math.cos(m) * R * (K === 'donut' ? .78 : .65), cy + Math.sin(m) * R * (K === 'donut' ? .78 : .65) + 6); } a0 = a1; });
    if (K === 'donut') { g.beginPath(); g.arc(cx, cy, R * .55, 0, 7); g.fillStyle = '#fff'; g.fill(); }
    drawLegend(); return;
  }
  if (K === 'radar') {
    const n = D.cats.length, cx = lef + (W - lef - rig) / 2, cy = top + (H - top - bot) / 2, R = Math.min(W - lef - rig, H - top - bot) / 2 - 36; let mx = 0; D.vals.forEach(a => a.forEach(v => mx = Math.max(mx, v))); const sc = chNice(0, mx || 1, 4);
    g.strokeStyle = '#ccc'; g.lineWidth = 1; sc.t.forEach(t => { g.beginPath(); for (let i = 0; i < n; i++) { const a = -Math.PI / 2 + i / n * 2 * Math.PI, r = R * t / sc.hi; i ? g.lineTo(cx + Math.cos(a) * r, cy + Math.sin(a) * r) : g.moveTo(cx + Math.cos(a) * r, cy + Math.sin(a) * r); } g.closePath(); g.stroke(); });
    g.fillStyle = '#333'; g.textAlign = 'center'; for (let i = 0; i < n; i++) { const a = -Math.PI / 2 + i / n * 2 * Math.PI; g.beginPath(); g.moveTo(cx, cy); g.lineTo(cx + Math.cos(a) * R, cy + Math.sin(a) * R); g.stroke(); g.fillText(String(D.cats[i]).slice(0, 8), cx + Math.cos(a) * (R + 24), cy + Math.sin(a) * (R + 24) + 6); }
    D.vals.forEach((arr, j) => { g.beginPath(); arr.forEach((v, i) => { const a = -Math.PI / 2 + i / n * 2 * Math.PI, r = R * v / sc.hi; i ? g.lineTo(cx + Math.cos(a) * r, cy + Math.sin(a) * r) : g.moveTo(cx + Math.cos(a) * r, cy + Math.sin(a) * r); }); g.closePath(); g.strokeStyle = col(j); g.lineWidth = 4; g.stroke(); }); drawLegend(); return;
  }
  /* cartesian */
  const horiz = K === 'barH' || K === 'barHS', stacked = K === 'colS' || K === 'colP' || K === 'barHS' || K === 'areaS', pct = K === 'colP';
  const isScatter = K === 'scatter', cats = D.cats, n = cats.length;
  let mn = 0, mx = 0, xs = null, xmn = 0, xmx = 1;
  if (isScatter) { xs = D.rows.map((r, i) => typeof r[0] === 'number' ? r[0] : i + 1); xmn = Math.min(...xs); xmx = Math.max(...xs); if (xmn === xmx) xmx = xmn + 1; mn = Infinity; mx = -Infinity; D.vals.forEach(a => a.forEach(v => { mn = Math.min(mn, v); mx = Math.max(mx, v); })); if (mn > 0) mn = Math.min(mn, 0) === 0 && mn < (mx - mn) ? 0 : mn; }
  else if (stacked) { for (let i = 0; i < n; i++) { let p = 0, q = 0; D.vals.forEach(a => { if (a[i] >= 0) p += a[i]; else q += a[i]; }); mx = Math.max(mx, p); mn = Math.min(mn, q); } if (pct) { mx = 1; mn = 0; } }
  else D.vals.forEach(a => a.forEach(v => { mx = Math.max(mx, v); mn = Math.min(mn, v); }));
  const yt = o.yT, xt = o.xT; const sc = pct ? { lo: 0, hi: 1, t: [0, .25, .5, .75, 1] } : chNice(mn, mx, 5), xsc = isScatter ? chNice(xmn, xmx, 5) : null;
  g.font = FS + ff; let yl = 0; (horiz ? cats.map(c => String(c).slice(0, 8)) : sc.t.map(t => pct ? t * 100 + '%' : chFmt(t))).forEach(t => yl = Math.max(yl, g.measureText(t).width)); lef += yl + 14 + (yt ? 30 : 0); bot += 34 + (xt ? 30 : 0);
  const x0 = lef, x1 = W - rig, y0 = top + 6, y1 = H - bot, pw = x1 - x0, ph = y1 - y0; if (pw < 40 || ph < 40) return;
  const V = v => horiz ? x0 + (v - sc.lo) / (sc.hi - sc.lo) * pw : y1 - (v - sc.lo) / (sc.hi - sc.lo) * ph;
  const C = i => horiz ? y0 + (i + .5) / n * ph : x0 + (i + .5) / n * pw;
  /* grid + value axis */
  g.lineWidth = 1; g.font = FS + ff; g.fillStyle = '#555';
  sc.t.forEach(t => { const p = V(pct ? t : t); if (o.grid !== false) { g.strokeStyle = '#e1e1e1'; g.beginPath(); if (horiz) { g.moveTo(p, y0); g.lineTo(p, y1); } else { g.moveTo(x0, p); g.lineTo(x1, p); } g.stroke(); } g.textAlign = horiz ? 'center' : 'right'; const lab = pct ? t * 100 + '%' : chFmt(t); if (horiz) g.fillText(lab, p, y1 + 26); else g.fillText(lab, x0 - 8, p + 7); });
  if (isScatter) xsc.t.forEach(t => { const p = x0 + (t - xsc.lo) / (xsc.hi - xsc.lo) * pw; if (o.grid !== false) { g.strokeStyle = '#eee'; g.beginPath(); g.moveTo(p, y0); g.lineTo(p, y1); g.stroke(); } g.textAlign = 'center'; g.fillText(chFmt(t), p, y1 + 26); });
  g.strokeStyle = '#888'; g.beginPath(); g.moveTo(x0, y0); g.lineTo(x0, y1); g.lineTo(x1, y1); g.stroke();
  const zero = V(0); g.strokeStyle = '#777'; g.beginPath(); if (horiz) { g.moveTo(zero, y0); g.lineTo(zero, y1); } else { g.moveTo(x0, zero); g.lineTo(x1, zero); } g.stroke();
  /* category labels */
  if (!isScatter) { g.fillStyle = '#555'; const maxw = (horiz ? ph : pw) / n; D.cats.forEach((c, i) => { let t = String(c); const lim = horiz ? 99 : Math.max(2, Math.floor(maxw / 11)); if (t.length > lim) t = t.slice(0, lim - 1) + '…'; if (horiz) { g.textAlign = 'right'; g.fillText(t, x0 - 8, C(i) + 7); } else { g.textAlign = 'center'; g.fillText(t, C(i), y1 + 26); } }); }
  /* series */
  const nS = D.names.length; const grpW = (horiz ? ph : pw) / n;
  const pos = new Array(nS).fill(0).map(() => new Array(n).fill(0)); const neg = new Array(nS).fill(0).map(() => new Array(n).fill(0));
  const drawLabels = (txt, x, y, al) => { g.fillStyle = '#333'; g.font = (FS - 3) + ff; g.textAlign = al || 'center'; g.fillText(txt, x, y); g.font = FS + ff; };
  const asBar = j => (K === 'col' || K === 'colS' || K === 'colP' || K === 'barH' || K === 'barHS' || (K === 'combo' && j === 0));
  const asLine = j => (K === 'line' || K === 'lineM' || K === 'radar' || (K === 'combo' && j > 0));
  const asArea = () => K === 'area' || K === 'areaS';
  const barSeries = [...Array(nS).keys()].filter(asBar), nb = (stacked ? 1 : barSeries.length) || 1;
  if (asArea()) {
    const base = new Array(n).fill(0); for (let j = 0; j < nS; j++) { const top = D.vals[j].map((v, i) => K === 'areaS' ? base[i] + v : v); g.beginPath(); top.forEach((v, i) => { const x = x0 + (n > 1 ? i / (n - 1) : .5) * pw; i ? g.lineTo(x, V(v)) : g.moveTo(x, V(v)); }); for (let i = n - 1; i >= 0; i--) { const x = x0 + (n > 1 ? i / (n - 1) : .5) * pw; g.lineTo(x, V(K === 'areaS' ? base[i] : 0)); } g.closePath(); g.globalAlpha = .75; g.fillStyle = col(j); g.fill(); g.globalAlpha = 1; if (K === 'areaS') top.forEach((v, i) => base[i] = v); }
  }
  D.vals.forEach((arr, j) => {
    if (asBar(j)) arr.forEach((v, i) => {
      const bw = grpW * .7 / nb, off = grpW * .15 + (stacked ? 0 : barSeries.indexOf(j) * bw); let a, b;
      if (stacked) { const val = pct ? v / (D.vals.reduce((t, q) => t + Math.abs(q[i]), 0) || 1) : v; if (val >= 0) { a = pos[j][i] = (j ? pos[j - 1][i] : 0) + 0; } a = 0; const prevP = D.vals.slice(0, j).reduce((t, q) => t + (pct ? Math.max(0, q[i]) / (D.vals.reduce((t2, q2) => t2 + Math.abs(q2[i]), 0) || 1) : Math.max(0, q[i])), 0), prevN = D.vals.slice(0, j).reduce((t, q) => t + (pct ? 0 : Math.min(0, q[i])), 0); a = val >= 0 ? prevP : prevN; b = a + val; }
      else { a = 0; b = v; }
      g.fillStyle = col(j); if (horiz) g.fillRect(Math.min(V(a), V(b)), C(i) - grpW / 2 + off, Math.abs(V(b) - V(a)), bw); else g.fillRect(C(i) - grpW / 2 + off, Math.min(V(a), V(b)), bw, Math.abs(V(b) - V(a)));
      if (o.labels) { const t = pct ? Math.round(Math.abs(b - a) * 100) + '%' : chFmt(v); if (horiz) drawLabels(t, V(Math.max(a, b)) + 4, C(i) - grpW / 2 + off + bw / 2 + 6, 'left'); else drawLabels(t, C(i) - grpW / 2 + off + bw / 2, (stacked ? (Math.min(V(a), V(b)) + Math.max(V(a), V(b))) / 2 + 6 : Math.min(V(a), V(b)) - 6)); }
    });
    if (asLine(j) && K !== 'radar') {
      const px = i => x0 + (K === 'combo' ? (i + .5) / n * pw : (n > 1 ? i / (n - 1) : .5) * pw) - 0; const pxx = i => (K === 'combo' ? C(i) : x0 + (i + .5) / n * pw);
      g.strokeStyle = col(j); g.lineWidth = 4; g.lineJoin = 'round'; g.beginPath(); arr.forEach((v, i) => { const x = pxx(i), y = V(v); if (i === 0) g.moveTo(x, y); else if (o.smooth) { const xp = pxx(i - 1), yp = V(arr[i - 1]); g.bezierCurveTo((xp + x) / 2, yp, (xp + x) / 2, y, x, y); } else g.lineTo(x, y); }); g.stroke();
      if (K === 'lineM' || K === 'combo' || o.markers) arr.forEach((v, i) => { g.beginPath(); g.arc(pxx(i), V(v), 6, 0, 7); g.fillStyle = '#fff'; g.fill(); g.strokeStyle = col(j); g.lineWidth = 3; g.stroke(); });
      if (o.labels) arr.forEach((v, i) => drawLabels(chFmt(v), pxx(i), V(v) - 12));
    }
    if (isScatter) arr.forEach((v, i) => { const x = x0 + (xs[i] - xsc.lo) / (xsc.hi - xsc.lo) * pw; g.beginPath(); g.arc(x, V(v), 8, 0, 7); g.fillStyle = col(j); g.fill(); if (o.labels) drawLabels(chFmt(v), x, V(v) - 14); });
  });
  if (asArea() && o.labels) { }
  /* axis titles */
  g.fillStyle = '#444'; g.font = 'bold ' + (FS - 2) + ff; g.textAlign = 'center';
  if (xt) g.fillText(xt, x0 + pw / 2, H - bot + 62 > H - 6 ? H - 6 : y1 + 58);
  if (yt) { g.save(); g.translate(18, y0 + ph / 2); g.rotate(-Math.PI / 2); g.fillText(yt, 0, 0); g.restore(); }
  drawLegend();
}
/* ---- selection, drag, redraw ---- */
var chSel = null;
function chCharts() { return shapes().filter(s => s.type === 'chart'); }
function chRedraw() { const L = document.getElementById('shapes'); if (!L) return; [...L.children].forEach(d => { const s = shapes().find(x => x.id === d.dataset.id); if (!s || s.type !== 'chart') return; let cv = d.querySelector('canvas'); if (!cv) { cv = document.createElement('canvas'); d.insertBefore(cv, d.firstChild); } cv.width = s.w * 2; cv.height = s.h * 2; drawChart(cv, s); d.classList.toggle('chsel', chSel === s.id); }); }
(function () {
  const st = document.createElement('style'); st.textContent = '.shp.chart{cursor:move;touch-action:none;box-shadow:0 1px 4px #0003}.shp.chsel{outline:2px solid #217346}.shp.chsel .rz{display:block}.shp.chart .rz{display:none}.chgal{display:grid;grid-template-columns:repeat(2,1fr);gap:6px}.chgal button{padding:8px;border:1px solid #bbb;border-radius:6px;background:#fff;text-align:left;font-size:13px}.chgal button.on{border-color:#217346;background:#e6f2ea}';
  document.head.appendChild(st);
})();
var _rs3 = renderShapes;
renderShapes = function () {
  _rs3(); const L = document.getElementById('shapes'); chRedraw();
  [...L.children].forEach(d => { const s = shapes().find(x => x.id === d.dataset.id); if (!s || s.type !== 'chart') return; d.classList.add('chart');
    let pd = null, moved = false;
    d.addEventListener('pointerdown', e => { if (design) return; e.stopPropagation(); chSel = s.id; selShape = s.id; chRedraw(); chRibbon(true); moved = false; pd = { x: e.clientX, y: e.clientY, sx: s.x, sy: s.y, sw: s.w, sh: s.h, rz: e.target.classList.contains('rz') }; d.setPointerCapture(e.pointerId); });
    d.addEventListener('pointermove', e => { if (design || !pd) return; const dx = e.clientX - pd.x, dy = e.clientY - pd.y; if (Math.abs(dx) + Math.abs(dy) > 3) moved = true; if (pd.rz) { s.w = Math.max(120, pd.sw + dx); s.h = Math.max(90, pd.sh + dy); } else { s.x = pd.sx + dx; s.y = pd.sy + dy; } d.style.left = s.x + 'px'; d.style.top = s.y + 'px'; d.style.width = s.w + 'px'; d.style.height = s.h + 'px'; });
    d.addEventListener('pointerup', () => { if (design || !pd) return; const was = pd; pd = null; if (moved) { save(); chRedraw(); } });
    d.addEventListener('dblclick', e => { if (!design) { e.stopPropagation(); chartDlg(s); } });
    d.addEventListener('contextmenu', e => { e.preventDefault(); e.stopPropagation(); chSel = s.id; chRedraw(); chartDlg(s); });
  });
};
document.addEventListener('pointerdown', e => { if (chSel && !e.target.closest('.shp') && !e.target.closest('#mGen') && !e.target.closest('#rib')) { chSel = null; chRedraw(); chRibbon(false); } }, true);
document.addEventListener('keydown', e => { if (chSel && (e.key === 'Delete' || e.key === 'Backspace') && !/INPUT|TEXTAREA|SELECT/.test((e.target || {}).tagName || '') && !editCell) { const a = shapes(), i = a.findIndex(x => x.id === chSel); if (i >= 0) { a.splice(i, 1); chSel = null; save(); renderShapes(); chRibbon(false); e.preventDefault(); e.stopPropagation(); } } else if (e.altKey && e.key === 'F1') { e.preventDefault(); chQuick(); } else if (e.key === 'Escape' && chSel) { chSel = null; chRedraw(); chRibbon(false); } }, true);
var _rf3 = refresh, _chT = 0; refresh = function () { _rf3(); clearTimeout(_chT); _chT = setTimeout(chRedraw, 30); };
function chQuick() { if (sel.r1 === sel.r2 && sel.c1 === sel.c2) { const R = regionSel(); if (R.r2 > R.r1) setSel(R.r1, R.c1, R.r2, R.c2, R.r1, R.c1); } chAdd('col'); }
function chAdd(kind) { if (sel.r1 === sel.r2 && sel.c1 === sel.c2) { alert('Select the data range first (include headers).'); return; } const s = { id: nextId(), type: 'chart', kind, sheet: cur, r1: sel.r1, c1: sel.c1, r2: sel.r2, c2: sel.c2, x: wrapX() + 50, y: wrapY() + 50, w: 360, h: 240, o: { kind0: 1, title: 'Chart Title' } }; shapes().push(s); chSel = s.id; save(); renderShapes(); chRibbon(true); }
addChart = function (kind) { chAdd(kind === 'bar' ? 'col' : kind); };
function chPicker() {
  if (sel.r1 === sel.r2 && sel.c1 === sel.c2) { const R = regionSel(); if (R.r2 > R.r1) setSel(R.r1, R.c1, R.r2, R.c2, R.r1, R.c1); else { alert('Select the data range first (include headers).'); return; } }
  gen('Insert Chart', '<div class="chgal">' + CH_TYPES.map(t => `<button data-k="${t[0]}"><span style="font-size:18px">${t[2]}</span> ${t[1]}</button>`).join('') + '</div>', []);
  document.querySelectorAll('#gB .chgal button').forEach(b => b.onclick = () => { closeGen(); chAdd(b.dataset.k); });
}
function chartDlg(s) {
  s = s || shapes().find(x => x.id === chSel); if (!s) { alert('Tap a chart first.'); return; } const o = chOpt(s);
  const sl = (id, v, a) => `<select id="${id}">${a.map(x => `<option value="${x[0]}" ${String(x[0]) === String(v) ? 'selected' : ''}>${x[1]}</option>`).join('')}</select>`;
  gen('Chart Design', `<label>Chart type</label>${sl('cdK', chKind(s), CH_TYPES.map(t => [t[0], t[1]]))}<label>Chart title</label><input type="text" id="cdT" value="${esc(o.title || '')}"><label>Data range</label><input type="text" id="cdR" value="${esc(s.sheet && s.sheet !== cur ? s.sheet + '!' : '')}${E.addr(s.r1, s.c1)}:${E.addr(s.r2, s.c2)}"><label>Legend</label>${sl('cdL', o.legend || 'auto', [['auto', 'Automatic'], ['right', 'Right'], ['bottom', 'Bottom'], ['top', 'Top'], ['left', 'Left'], ['none', 'None']])}<label>Colors</label>${sl('cdP', o.pal || 'Colorful', Object.keys(CH_PAL).map(k => [k, k]))}<label>X-axis title</label><input type="text" id="cdX" value="${esc(o.xT || '')}"><label>Y-axis title</label><input type="text" id="cdY" value="${esc(o.yT || '')}"><label><input type="checkbox" id="cdD" ${o.labels ? 'checked' : ''}> Data labels</label><label><input type="checkbox" id="cdG" ${o.grid !== false ? 'checked' : ''}> Gridlines</label><label><input type="checkbox" id="cdS" ${o.smooth ? 'checked' : ''}> Smooth lines</label><label><input type="checkbox" id="cdW" ${o.sw ? 'checked' : ''}> Switch row/column</label>`, [
    ['OK', () => { s.kind = $('cdK').value; o.title = $('cdT').value; const lv = $('cdL').value; if (lv === 'auto') delete o.legend; else o.legend = lv; o.pal = $('cdP').value; o.xT = $('cdX').value; o.yT = $('cdY').value; o.labels = $('cdD').checked; o.grid = $('cdG').checked; o.smooth = $('cdS').checked; o.sw = $('cdW').checked; const m = /^(?:'?([^'!]+)'?!)?([A-Z]+)(\d+):([A-Z]+)(\d+)$/i.exec($('cdR').value.trim()); if (m) { if (m[1]) s.sheet = m[1]; s.c1 = E.colToNum(m[2].toUpperCase()); s.r1 = +m[3]; s.c2 = E.colToNum(m[4].toUpperCase()); s.r2 = +m[5]; } closeGen(); save(); renderShapes(); }],
    ['Delete chart', () => { const a = shapes(); a.splice(a.indexOf(s), 1); chSel = null; closeGen(); save(); renderShapes(); chRibbon(false); }]]);
}
/* contextual Chart Design ribbon tab + Insert group */
function chRibbon(on) { const rib = document.getElementById('rib'); if (!rib) return; let t = rib.querySelector('[data-t="Chart Design"]'); const bar = document.querySelector('.rib .tabs') || document.getElementById('tabs'); const btn = [...document.querySelectorAll('.rib .tab, #tabs .tab, .tabs button')].find(b => b.textContent.trim() === 'Chart Design'); if (btn) btn.style.display = on ? '' : 'none'; if (!on && t && t.classList.contains('on')) { const home = [...document.querySelectorAll('.rib .tab, #tabs .tab, .tabs button')].find(b => b.textContent.trim() === 'Home'); if (home) home.click(); } else if (on && btn && !chRibbon.shown) { } chRibbon.shown = on; }
(function () {
  const rib = document.getElementById('rib'); if (!rib) return;
  const grp = (label, btns) => `<div class="grp"><div class="gb">${btns.map(b => `<button class="big" onclick="${b[1]}"><i>${b[2]}</i>${b[0]}</button>`).join('')}</div><div class="gl">${label}</div></div>`;
  const ins = rib.querySelector('[data-t="Insert"]'); if (ins) ins.insertAdjacentHTML('beforeend', grp('More Charts', [['All Charts…', 'chPicker()', '&#128202;'], ['Chart Design', 'chartDlg()', '&#9881;'], ['Area', "chAdd('area')", '&#9650;'], ['Scatter', "chAdd('scatter')", '&#8759;'], ['Doughnut', "chAdd('donut')", '&#9678;'], ['Bar', "chAdd('barH')", '&#9644;']]));
})();
