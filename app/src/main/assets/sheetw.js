/* Tier 3: high-fidelity .xlsx export (styles, validation, freeze, names, links, notes, print setup) */
var NFB = { 'General': 0, '0': 1, '0.00': 2, '#,##0': 3, '#,##0.00': 4, '0%': 9, '0.00%': 10, 'm/d/yyyy': 14, 'mm-dd-yy': 14 };
function xlEsc(s) { return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;'); }
function xlColor(c) { c = String(c || '').trim(); let m = /^#([0-9a-f]{6})$/i.exec(c); if (m) return 'FF' + m[1].toUpperCase(); m = /^#([0-9a-f]{3})$/i.exec(c); if (m) return 'FF' + m[1].split('').map(x => x + x).join('').toUpperCase(); m = /rgb\((\d+),\s*(\d+),\s*(\d+)/.exec(c); if (m) return 'FF' + [m[1], m[2], m[3]].map(x => (+x).toString(16).padStart(2, '0')).join('').toUpperCase(); return null; }
function xlBorder(v) { if (!v) return ''; const w = parseFloat(v) || 1; const dbl = /double/.test(v), dash = /dash/.test(v), dot = /dotted/.test(v); const st = dbl ? 'double' : dot ? 'dotted' : dash ? (w >= 2 ? 'mediumDashed' : 'dashed') : (w >= 3 ? 'thick' : w >= 2 ? 'medium' : 'thin'); const col = (/#[0-9a-f]{3,6}|rgb\([^)]*\)/i.exec(v) || [])[0]; return [st, xlColor(col) || 'FF000000']; }
function buildStyles(styleMap) {
  const fonts = ['<font><sz val="11"/><name val="Calibri"/><family val="2"/></font>'], fills = ['<fill><patternFill patternType="none"/></fill>', '<fill><patternFill patternType="gray125"/></fill>'], borders = ['<border><left/><right/><top/><bottom/><diagonal/></border>'], nf = {}, xfs = ['<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>']; const idx = {}; let nid = 164; const pool = (arr, x) => { let i = arr.indexOf(x); if (i < 0) { arr.push(x); i = arr.length - 1; } return i; };
  const get = (st, fmt) => {
    const key = JSON.stringify([st || {}, fmt || '']); if (idx[key] !== undefined) return idx[key]; st = st || {};
    let f = '<font>' + (st.b ? '<b/>' : '') + (st.i ? '<i/>' : '') + (st.u ? '<u/>' : '') + (st.s || st.strike ? '<strike/>' : '') + `<sz val="${st.sz ? +(st.sz * 0.75).toFixed(1) : 11}"/>` + (st.color && xlColor(st.color) ? `<color rgb="${xlColor(st.color)}"/>` : '') + `<name val="${xlEsc(st.font || 'Calibri')}"/><family val="2"/></font>`;
    const fo = pool(fonts, f); const fl = st.bg && xlColor(st.bg) ? pool(fills, `<fill><patternFill patternType="solid"><fgColor rgb="${xlColor(st.bg)}"/><bgColor indexed="64"/></patternFill></fill>`) : 0;
    const bd = ['bl', 'br', 'bt', 'bb'].map(k => xlBorder(st[k])); const bo = bd.some(x => x) ? pool(borders, '<border>' + [['left', bd[0]], ['right', bd[1]], ['top', bd[2]], ['bottom', bd[3]]].map(([n, v]) => v ? `<${n} style="${v[0]}"><color rgb="${v[1]}"/></${n}>` : `<${n}/>`).join('') + '<diagonal/></border>') : 0;
    let nm = 0; if (fmt) { if (NFB[fmt] !== undefined) nm = NFB[fmt]; else { if (!nf[fmt]) nf[fmt] = nid++; nm = nf[fmt]; } }
    const al = (st.align || st.valign || st.wrap || st.indent) ? `<alignment${st.align ? ` horizontal="${st.align}"` : ''}${st.valign ? ` vertical="${st.valign === 'middle' ? 'center' : st.valign}"` : ''}${st.wrap ? ' wrapText="1"' : ''}${st.indent ? ` indent="${st.indent}"` : ''}/>` : '';
    xfs.push(`<xf numFmtId="${nm}" fontId="${fo}" fillId="${fl}" borderId="${bo}" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyNumberFormat="1"${al ? ' applyAlignment="1">' + al + '</xf>' : '/>'}`); return idx[key] = xfs.length - 1;
  };
  const xml = () => `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">${Object.keys(nf).length ? `<numFmts count="${Object.keys(nf).length}">${Object.entries(nf).map(([c, i]) => `<numFmt numFmtId="${i}" formatCode="${xlEsc(c)}"/>`).join('')}</numFmts>` : ''}<fonts count="${fonts.length}">${fonts.join('')}</fonts><fills count="${fills.length}">${fills.join('')}</fills><borders count="${borders.length}">${borders.join('')}</borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="${xfs.length}">${xfs.join('')}</cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>`;
  return { get, xml };
}
function dvXml(S) {
  const L = (S.dv || []).filter(q => q.type && q.type !== 'any'); if (!L.length) return ''; const tp = { whole: 'whole', decimal: 'decimal', list: 'list', date: 'date', time: 'time', length: 'textLength', custom: 'custom' };
  return `<dataValidations count="${L.length}">` + L.map(q => { const t = tp[q.type]; if (!t) return ''; const list = q.type === 'list'; const v1 = String(q.v1 || ''); const f1 = list ? (/^=/.test(v1) ? v1.slice(1) : '"' + v1.replace(/"/g, '""') + '"') : (q.type === 'custom' ? v1.replace(/^=/, '') : v1.replace(/^=/, '')); const op = (q.op && q.op !== 'between' && !list && q.type !== 'custom') ? ` operator="${q.op}"` : (q.op === 'notBetween' ? ' operator="notBetween"' : '');
    return `<dataValidation type="${t}"${op}${q.er && q.er.st && q.er.st !== 'stop' ? ` errorStyle="${q.er.st}"` : ''}${q.blank ? ' allowBlank="1"' : ''}${list && q.drop === false ? '' : ''} showInputMessage="${q.im && q.im.on ? 1 : 0}" showErrorMessage="${q.er && q.er.on === false ? 0 : 1}"${q.er && q.er.t ? ` errorTitle="${xlEsc(q.er.t)}"` : ''}${q.er && q.er.m ? ` error="${xlEsc(q.er.m)}"` : ''}${q.im && q.im.t ? ` promptTitle="${xlEsc(q.im.t)}"` : ''}${q.im && q.im.m ? ` prompt="${xlEsc(q.im.m)}"` : ''} sqref="${E.addr(q.r1, q.c1)}${q.r1 === q.r2 && q.c1 === q.c2 ? '' : ':' + E.addr(q.r2, q.c2)}"><formula1>${xlEsc(f1)}</formula1>${(q.op === 'between' || q.op === 'notBetween') && !list && q.type !== 'custom' ? `<formula2>${xlEsc(String(q.v2 || '').replace(/^=/, ''))}</formula2>` : ''}</dataValidation>`; }).join('') + '</dataValidations>';
}
function pageXml(S) { const P = (S.ext && S.ext.page) || null; if (!P) return { opt: '', mar: '', setup: '', hf: '' }; const m = P.margins || [0.7, 0.7, 0.75, 0.75]; const ps = { Letter: 1, Legal: 5, A4: 9, A3: 8, A5: 11, Folio: 14 }[P.paper || 'Letter'] || 1;
  return { opt: P.gridlines || P.center ? `<printOptions${P.gridlines ? ' gridLines="1"' : ''}${P.center ? ' horizontalCentered="1"' : ''}/>` : '', mar: `<pageMargins left="${m[0]}" right="${m[1]}" top="${m[2]}" bottom="${m[3]}" header="0.3" footer="0.3"/>`, setup: `<pageSetup paperSize="${ps}" orientation="${P.orient === 'landscape' ? 'landscape' : 'portrait'}"${P.fit ? ' fitToWidth="1" fitToHeight="0"' : P.scale ? ` scale="${P.scale}"` : ''}/>`, hf: (P.header || P.footer) ? `<headerFooter>${P.header ? `<oddHeader>${xlEsc(P.header)}</oddHeader>` : ''}${P.footer ? `<oddFooter>${xlEsc(P.footer)}</oddFooter>` : ''}</headerFooter>` : '' }; }
async function saveX() {
  const w = XLSX.utils.book_new(); const styleSets = []; const ST = buildStyles(); const info = [];
  for (const n of wb.order) {
    const ws = {}; let mr = 1, mc = 1; const s = wb.sheets[n]; const cs = {};
    const touch = (r, c) => { mr = Math.max(mr, r); mc = Math.max(mc, c); };
    for (const k in s.cells) { const [r, c] = k.split(',').map(Number); const cell = s.cells[k]; const v = wb.value(n, r, c); const o = {}; if (cell.f) o.f = cell.f.slice(1); if (E.isErr(v)) { o.t = 'e'; o.v = 0; o.w = v.code; } else if (typeof v === 'number') { o.t = 'n'; o.v = v; } else if (typeof v === 'boolean') { o.t = 'b'; o.v = v; } else { o.t = 's'; o.v = v === null ? '' : String(v); } ws[E.addr(r, c)] = o; touch(r, c); }
    for (const k in wb.spill) { if (k.startsWith(n + '|')) { const [r, c] = k.slice(n.length + 1).split(',').map(Number); const v = wb.spill[k]; ws[E.addr(r, c)] = { t: typeof v === 'number' ? 'n' : 's', v: E.isErr(v) ? v.code : v }; touch(r, c); } }
    // styled / formatted empty cells
    for (const k of new Set([...Object.keys(s.style || {}), ...Object.keys(s.fmt || {})])) { const [r, c] = k.split(',').map(Number); if (r > 1000 || c > 100) continue; const a = E.addr(r, c); if (!ws[a]) { ws[a] = { t: 's', v: '' }; touch(r, c); } }
    for (const a in ws) { if (a[0] === '!') continue; const m = /^([A-Z]+)(\d+)$/.exec(a); const r = +m[2], c = E.colToNum(m[1]); const k = r + ',' + c; const st = s.style && s.style[k], fm = s.fmt && s.fmt[k]; if (st || fm) cs[a] = ST.get(st, fm); }
    const lk = (s.ext && s.ext.links) || {}; for (const k in lk) { const [r, c] = k.split(',').map(Number); const a = E.addr(r, c); if (!ws[a]) { ws[a] = { t: 's', v: lk[k].text || lk[k].url }; touch(r, c); } ws[a].l = { Target: lk[k].url }; }
    const nt = (s.ext && s.ext.notes) || {}; for (const k in nt) { const [r, c] = k.split(',').map(Number); const a = E.addr(r, c); if (!ws[a]) { ws[a] = { t: 's', v: '' }; touch(r, c); } const tx = typeof nt[k] === 'string' ? nt[k] : (nt[k].text || nt[k].t || ''); ws[a].c = [{ a: 'Office HUB', t: tx }]; }
    if (s.merges && s.merges.length) ws['!merges'] = s.merges.map(m => ({ s: { r: m.r1 - 1, c: m.c1 - 1 }, e: { r: m.r2 - 1, c: m.c2 - 1 } }));
    ws['!cols'] = []; for (let c = 1; c <= mc; c++) { const wd = s.colW[c]; ws['!cols'].push(wd === 0 ? { hidden: true, wpx: 10 } : { wpx: wd !== undefined ? wd : DW }); }
    ws['!rows'] = []; for (let r = 1; r <= mr; r++) { const h = s.rowH && s.rowH[r]; ws['!rows'].push(h === 0 ? { hidden: true } : h !== undefined ? { hpx: h } : {}); }
    ws['!ref'] = 'A1:' + E.addr(mr, mc); XLSX.utils.book_append_sheet(w, ws, n.slice(0, 31)); info.push({ n, s, cs });
  }
  const Names = []; for (const nm in (book.names || {})) Names.push({ Name: nm, Ref: String(book.names[nm]).replace(/^=/, '') });
  info.forEach((x, i) => { const P = x.s.ext && x.s.ext.page; if (P && P.area) Names.push({ Name: '_xlnm.Print_Area', Sheet: i, Ref: P.area.replace(/^=/, '') }); if (P && P.titles) Names.push({ Name: '_xlnm.Print_Titles', Sheet: i, Ref: P.titles.replace(/^=/, '') }); });
  if (Names.length) w.Workbook = { Names };
  const out = XLSX.write(w, { bookType: 'xlsx', type: 'array', cellStyles: true }); let blob;
  try {
    const z = await JSZip.loadAsync(out); z.file('xl/styles.xml', ST.xml());
    for (let i = 0; i < info.length; i++) {
      const path = 'xl/worksheets/sheet' + (i + 1) + '.xml'; let x = await z.file(path).async('string'); const { s, cs } = info[i];
      x = x.replace(/<c r="([A-Z]+\d+)"([^>]*?)(\/?)>/g, (m, a, attrs, sl) => { attrs = attrs.replace(/\s+s="\d+"/, ''); return cs[a] !== undefined ? `<c r="${a}"${attrs} s="${cs[a]}"${sl}>` : `<c r="${a}"${attrs}${sl}>`; });
      const F = s.frz; const grid = s.grid === false || s.hideGrid ? ' showGridLines="0"' : ''; const zm = ''; let pane = '';
      if (F && (F.r || F.c)) pane = `<pane${F.c ? ` xSplit="${F.c}"` : ''}${F.r ? ` ySplit="${F.r}"` : ''} topLeftCell="${E.addr(F.r + 1, F.c + 1)}" activePane="${F.r && F.c ? 'bottomRight' : F.r ? 'bottomLeft' : 'topRight'}" state="frozen"/>`;
      x = x.replace(/<sheetView([^>]*?)\/>/, (m, at) => `<sheetView${at}${grid}${pane ? '>' + pane + '</sheetView>' : '/>'}`);
      const af = s.ext && s.ext.af; if (af && !/<autoFilter/.test(x)) x = x.replace(/<\/sheetData>/, `</sheetData><autoFilter ref="${E.addr(af.r, af.c1)}:${E.addr(Math.max(af.r + 1, lastRowOf(af.c1, af.c2, af.r + 1)), af.c2)}"/>`);
      const P = pageXml(s); const dv = dvXml(s);
      x = x.replace(/<pageMargins[^>]*\/>/, ''); // rebuild tail in schema order
      const tail = dv + (P.opt || '') + (P.mar || '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>') + (P.setup || '') + (P.hf || '');
      if (/<legacyDrawing|<drawing/.test(x)) x = x.replace(/(<legacyDrawing|<drawing)/, tail + '$1'); else x = x.replace(/<\/worksheet>/, tail + '</worksheet>');
      z.file(path, x);
    }
    blob = await z.generateAsync({ type: 'blob', mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  } catch (e) { console.log('xlsx post-process failed', e); blob = new Blob([out], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }); }
  window.__lastXlsx = blob; RB.saveBlob(book.name + '.xlsx', blob);
  const code = Object.entries(book.modules).map(([k, v]) => `' ---- ${k} ----\n${v}`).join('\n\n'); if (code.trim()) RB.saveBlob(book.name + '_macros.bas', new Blob([code], { type: 'text/plain' }), 'text/plain');
}

/* ---------- Page Setup & Print ---------- */
const PAPER = { Letter: [8.5, 11], Legal: [8.5, 14], A4: [8.27, 11.69], A3: [11.69, 16.54], A5: [5.83, 8.27], Folio: [8.5, 13] };
function pg() { const e = X(); return e.page = e.page || { orient: 'portrait', paper: 'A4', margins: [0.7, 0.7, 0.75, 0.75], fit: false, gridlines: false, center: false, header: '', footer: '&CPage &P of &N', area: '', titles: '', scale: 100 }; }
function parseRng(t) { const m = /^\$?([A-Z]+)\$?(\d+)(?::\$?([A-Z]+)\$?(\d+))?$/i.exec(String(t || '').replace(/^.*!/, '').trim()); if (!m) return null; const c1 = E.colToNum(m[1].toUpperCase()), r1 = +m[2]; return { r1, c1, r2: m[4] ? +m[4] : r1, c2: m[3] ? E.colToNum(m[3].toUpperCase()) : c1 }; }
function pageDlg() {
  const P = pg(); const o = (a, v) => a.map(x => `<option ${x === v ? 'selected' : ''}>${x}</option>`).join('');
  gen('Page Setup', `<label>Orientation</label><select id="psO"><option value="portrait" ${P.orient === 'portrait' ? 'selected' : ''}>Portrait</option><option value="landscape" ${P.orient === 'landscape' ? 'selected' : ''}>Landscape</option></select><label>Paper size</label><select id="psP">${o(Object.keys(PAPER), P.paper)}</select><label>Margins (inches) left, right, top, bottom</label><input type="text" id="psM" value="${P.margins.join(', ')}"><label>Print area (e.g. A1:F40, blank = all data)</label><input type="text" id="psA" value="${esc(P.area)}"><label>Rows to repeat at top (e.g. $1:$2)</label><input type="text" id="psT" value="${esc(P.titles)}"><label>Header text (&amp;L &amp;C &amp;R sections, &amp;D date, &amp;A sheet, &amp;F file)</label><input type="text" id="psH" value="${esc(P.header)}"><label>Footer text (&amp;P page, &amp;N pages)</label><input type="text" id="psF" value="${esc(P.footer)}"><label><input type="checkbox" id="psFit" ${P.fit ? 'checked' : ''}> Fit all columns on one page</label><label>Scale %</label><input type="text" id="psS" value="${P.scale || 100}"><label><input type="checkbox" id="psG" ${P.gridlines ? 'checked' : ''}> Print gridlines</label><label><input type="checkbox" id="psC" ${P.center ? 'checked' : ''}> Center horizontally</label>`, [
    ['Save', () => { const m = $('psM').value.split(/[ ,]+/).map(Number).filter(x => !isNaN(x)); Object.assign(P, { orient: $('psO').value, paper: $('psP').value, margins: m.length === 4 ? m : P.margins, area: $('psA').value.trim(), titles: $('psT').value.trim(), header: $('psH').value, footer: $('psF').value, fit: $('psFit').checked, scale: +$('psS').value || 100, gridlines: $('psG').checked, center: $('psC').checked }); save(); closeGen(); }],
    ['Print…', () => { $('psO') && Object.assign(P, { orient: $('psO').value, paper: $('psP').value, area: $('psA').value.trim(), titles: $('psT').value.trim(), header: $('psH').value, footer: $('psF').value, fit: $('psFit').checked, scale: +$('psS').value || 100, gridlines: $('psG').checked, center: $('psC').checked }); const m = $('psM').value.split(/[ ,]+/).map(Number).filter(x => !isNaN(x)); if (m.length === 4) P.margins = m; save(); closeGen(); printSheet(); }]]);
}
function setPrintArea(clear) { const P = pg(); P.area = clear ? '' : E.addr(sel.r1, sel.c1) + ':' + E.addr(sel.r2, sel.c2); save(); stMsg(clear ? 'Print area cleared' : 'Print area set to ' + P.area); }
function hfText(t, sn) { const secs = { L: '', C: '', R: '' }; let cur0 = 'C'; String(t || '').split(/(&[LCR])/).forEach(p => { if (/^&[LCR]$/.test(p)) cur0 = p[1]; else secs[cur0] += p; }); const rep = s => s.replace(/&D/g, new Date().toLocaleDateString()).replace(/&A/g, sn).replace(/&F/g, book.name).replace(/&&/g, '&'); const cn = s => { const parts = rep(s).split(/(&P|&N)/); return parts.map(x => x === '&P' ? 'counter(page)' : x === '&N' ? 'counter(pages)' : '"' + x.replace(/"/g, '\\"') + '"').join(' ') || '""'; }; return { L: cn(secs.L), C: cn(secs.C), R: cn(secs.R) }; }
function printSheet() {
  const P = pg(), S = sh(); let R = parseRng(P.area); if (!R) { let mr = 1, mc = 1; for (const k in S.cells) { const [r, c] = k.split(',').map(Number); mr = Math.max(mr, r); mc = Math.max(mc, c); } R = { r1: 1, c1: 1, r2: mr, c2: mc }; }
  const cp = (td, extra) => { const cs = getComputedStyle(td); return `${extra || ''}color:${cs.color};background:${cs.backgroundColor};font-weight:${cs.fontWeight};font-style:${cs.fontStyle};text-decoration:${cs.textDecorationLine};text-align:${cs.textAlign};font-family:${cs.fontFamily};font-size:${cs.fontSize};border-top:${cs.borderTopWidth} ${cs.borderTopStyle} ${cs.borderTopColor};border-bottom:${cs.borderBottomWidth} ${cs.borderBottomStyle} ${cs.borderBottomColor};border-left:${cs.borderLeftWidth} ${cs.borderLeftStyle} ${cs.borderLeftColor};border-right:${cs.borderRightWidth} ${cs.borderRightStyle} ${cs.borderRightColor}`; };
  const hidR = r => S.rowH && S.rowH[r] === 0, hidC = c => S.colW[c] === 0; let tw = 0; const cols = []; for (let c = R.c1; c <= R.c2; c++) if (!hidC(c)) { cols.push(c); tw += colW(c); }
  const T = parseRng(P.titles && /^\$?\d/.test(P.titles) ? 'A' + P.titles.replace(/\$/g, '').replace(':', ':A') : ''); const tRows = T ? [T.r1, T.r2] : null;
  const rowHtml = r => { const rh = rowHt(r); let h = `<tr style="height:${rh}px">`; for (const c of cols) { const td = tds[r] && tds[r][c]; const gr = P.gridlines ? 'border:1px solid #999;' : ''; if (MM[key(r, c)] === 'h') continue; const sp = MM[key(r, c)]; const txt = td ? td.innerHTML : esc(cellText(r, c)); h += `<td${sp && sp.rs ? ` rowspan="${sp.rs}" colspan="${sp.cs}"` : ''} style="${td ? cp(td, gr) : gr}white-space:nowrap;overflow:hidden;padding:0 4px;height:${rh}px;max-width:${colW(c)}px">${txt}</td>`; } return h + '</tr>'; };
  let head = '', body = ''; if (tRows) for (let r = tRows[0]; r <= tRows[1]; r++) if (!hidR(r)) head += rowHtml(r); for (let r = R.r1; r <= R.r2; r++) { if (hidR(r)) continue; if (tRows && r >= tRows[0] && r <= tRows[1] && R.r1 <= tRows[1]) { if (r === R.r1 && R.r1 === tRows[0]) { } } body += rowHtml(r); }
  const pp = PAPER[P.paper] || PAPER.A4, w = P.orient === 'landscape' ? pp[1] : pp[0], availPx = (w - P.margins[0] - P.margins[1]) * 96; const zoom = P.fit ? Math.min(1, availPx / (tw + 4)) : (P.scale || 100) / 100;
  const hf = hfText(P.header, cur), ff = hfText(P.footer, cur); let el = document.getElementById('prn'); if (el) el.remove(); el = document.createElement('div'); el.id = 'prn';
  el.innerHTML = `<style id="prnCss">@page{size:${P.paper === 'Letter' || !P.paper ? 'letter' : P.paper} ${P.orient};margin:${P.margins[2]}in ${P.margins[1]}in ${P.margins[3]}in ${P.margins[0]}in;@top-left{content:${hf.L}}@top-center{content:${hf.C}}@top-right{content:${hf.R}}@bottom-left{content:${ff.L}}@bottom-center{content:${ff.C}}@bottom-right{content:${ff.R}}}@media screen{#prn{display:none}}@media print{body>*:not(#prn){display:none!important}#prn{display:block!important;position:static;background:#fff}#prn table{border-collapse:collapse;table-layout:fixed;zoom:${zoom};${P.center ? 'margin:0 auto;' : ''}}#prn thead{display:table-header-group}#prn tr{page-break-inside:avoid}}</style><table><colgroup>${cols.map(c => `<col style="width:${colW(c)}px">`).join('')}</colgroup>${head ? `<thead>${head}</thead>` : ''}<tbody>${body}</tbody></table>`;
  document.body.appendChild(el); setTimeout(() => { try { window.print(); } catch (e) { alert('Printing is not available here.'); } setTimeout(() => el.remove(), 1000); }, 50);
}
(function () { const o = document.createElement('style'); o.textContent = '#prn{display:none}'; document.head.appendChild(o); })();

/* ---------- Outline / Group / Subtotal ---------- */
function OG() { const e = X(); return e.og = e.og || { rows: {}, cols: {}, hid: {} }; }
function groupSel(dir, on) {
  const G = OG(); const tgt = dir === 'c' ? G.cols : G.rows; const a = dir === 'c' ? [sel.c1, sel.c2] : [sel.r1, sel.r2]; if (!(sel.c1 === 1 && sel.c2 >= COLS) && !(sel.r1 === 1 && sel.r2 >= ROWS) && dir === undefined) dir = 'r';
  for (let i = a[0]; i <= a[1]; i++) { const lv = (tgt[i] || 0) + (on ? 1 : -1); if (lv <= 0) delete tgt[i]; else tgt[i] = Math.min(8, lv); } if (!on) showDetail(dir, true); save(); build(); stMsg(on ? 'Grouped' : 'Ungrouped');
}
function groupAuto() { const whole = sel.c1 === 1 && sel.c2 >= COLS; groupSel(whole ? 'r' : (sel.r1 === 1 && sel.r2 >= ROWS ? 'c' : 'r'), true); }
function ungroupAuto() { const whole = sel.c1 === 1 && sel.c2 >= COLS; groupSel(whole ? 'r' : (sel.r1 === 1 && sel.r2 >= ROWS ? 'c' : 'r'), false); }
function showDetail(dir, show) {
  const G = OG(), S = sh(); const tgt = dir === 'c' ? G.cols : G.rows; const arr = dir === 'c' ? S.colW : (S.rowH = S.rowH || {}); const lo = dir === 'c' ? sel.c1 : sel.r1, hi = dir === 'c' ? sel.c2 : sel.r2;
  for (const k in tgt) { const i = +k; if (i < lo || i > hi) continue; const hk = dir + i; if (show) { if (G.hid[hk] !== undefined) { if (G.hid[hk] === null) delete arr[i]; else arr[i] = G.hid[hk]; delete G.hid[hk]; } else if (arr[i] === 0) delete arr[i]; } else { if (G.hid[hk] === undefined) G.hid[hk] = arr[i] === undefined ? null : arr[i]; arr[i] = 0; } }
  save(); build();
}
function groupFromCell() { // Alt+Shift+Right/Left in Excel apply to the selection
  if (sel.r1 === sel.r2 && sel.c1 === sel.c2) { stMsg('Select rows or columns to group'); return; } groupAuto();
}
function outlineLvl(dir, i) { const G = OG(); return (dir === 'c' ? G.cols : G.rows)[i] || 0; }
function subtotalDlg() {
  const R = regionSel(); if (R.r2 - R.r1 < 2) { alert('Select the data (with a header row), sorted by the column to group on.'); return; } const hd = []; for (let c = R.c1; c <= R.c2; c++) hd.push(String(cellVal(R.r1, c) ?? E.numToCol(c)));
  gen('Subtotal', `<label>At each change in</label><select id="stG">${hd.map((h, i) => `<option value="${i}">${esc(h)}</option>`).join('')}</select><label>Use function</label><select id="stF"><option value="9">Sum</option><option value="2">Count</option><option value="1">Average</option><option value="4">Max</option><option value="5">Min</option></select><label>Add subtotal to</label><select id="stC">${hd.map((h, i) => `<option value="${i}" ${i === hd.length - 1 ? 'selected' : ''}>${esc(h)}</option>`).join('')}</select>`, [['OK', () => { const g = +$('stG').value + R.c1, f = +$('stF').value, tc = +$('stC').value + R.c1; closeGen(); doSubtotal(R, g, f, tc); }], ['Remove All', () => { closeGen(); removeSubtotals(R); }]]);
}
function doSubtotal(R, g, f, tc) {
  removeSubtotals(R, true); const rows = []; for (let r = R.r1 + 1; r <= R.r2; r++) { const row = []; for (let c = R.c1; c <= R.c2; c++) row.push(wb.getRaw(cur, r, c)); rows.push({ src: r, row, st: [] }); } const nC = R.c2 - R.c1 + 1;
  const old = []; for (let r = R.r1 + 1; r <= R.r2; r++) old.push(r); const styles = {}; old.forEach((r, i) => { styles[i] = {}; for (let c = R.c1; c <= R.c2; c++) { const k = key(r, c); styles[i][c] = [sh().style[k], sh().fmt[k]]; } });
  const out = []; let start = 0; const flush = (end) => { if (end < start) return; out.push({ sub: true, from: start, to: end, label: String(rows[end].row[g - R.c1] ?? '') }); }; rows.forEach((x, i) => { out.push({ data: x, i }); const nxt = rows[i + 1]; if (!nxt || String(nxt.row[g - R.c1]) !== String(x.row[g - R.c1])) { flush(i); start = i + 1; } });
  const lastR = R.r1 + out.length + 1; for (let r = R.r1 + 1; r <= Math.max(R.r2, lastR); r++) for (let c = R.c1; c <= R.c2; c++) { wb.setRaw(cur, r, c, ''); delete sh().style[key(r, c)]; delete sh().fmt[key(r, c)]; }
  const G = OG(); G.rows = {}; let r = R.r1 + 1; const pos = {}; const dataStart = {}; const subs = [];
  out.forEach(o => { if (o.data) { o.data.row.forEach((v, j) => { if (v !== '' && v !== undefined) wb.setRaw(cur, r, R.c1 + j, v); const s0 = styles[o.i][R.c1 + j]; if (s0[0]) sh().style[key(r, R.c1 + j)] = s0[0]; if (s0[1]) sh().fmt[key(r, R.c1 + j)] = s0[1]; }); pos[o.i] = r; G.rows[r] = 1; r++; } else { const a = pos[o.from], b = pos[o.to]; wb.setRaw(cur, r, g, o.label + ' ' + ({ 9: 'Total', 2: 'Count', 1: 'Average', 4: 'Max', 5: 'Min' }[f])); wb.setRaw(cur, r, tc, `=SUBTOTAL(${f},${E.addr(a, tc)}:${E.addr(b, tc)})`); for (let c = R.c1; c <= R.c2; c++) sh().style[key(r, c)] = { b: 1 }; subs.push(r); r++; } });
  wb.setRaw(cur, r, g, 'Grand ' + ({ 9: 'Total', 2: 'Count', 1: 'Average', 4: 'Max', 5: 'Min' }[f])); wb.setRaw(cur, r, tc, `=SUBTOTAL(${f},${E.addr(R.r1 + 1, tc)}:${E.addr(r - 1, tc)})`); for (let c = R.c1; c <= R.c2; c++) sh().style[key(r, c)] = { b: 1, bt: '1px solid #000000' };
  save(); build(); stMsg('Subtotals inserted');
}
function removeSubtotals(R, silent) {
  const rows = []; let max = R.r2 + 200; for (let r = R.r1 + 1; r <= max; r++) { const raws = []; for (let c = R.c1; c <= R.c2; c++) raws.push(wb.getRaw(cur, r, c)); rows.push(raws); } const keep = rows.filter(row => !row.some(v => typeof v === 'string' && /^=SUBTOTAL\(/i.test(v)));
  if (keep.length === rows.length) { if (!silent) stMsg('No subtotals found'); return; }
  for (let i = 0; i < rows.length; i++) for (let c = R.c1; c <= R.c2; c++) { const k = key(R.r1 + 1 + i, c); wb.setRaw(cur, R.r1 + 1 + i, c, i < keep.length ? keep[i][c - R.c1] : ''); if (i >= keep.length) { delete sh().style[k]; delete sh().fmt[k]; } }
  OG().rows = {}; if (!silent) { save(); build(); stMsg('Subtotals removed'); }
}

/* ---------- Behaviour polish ---------- */
// #### for numbers that do not fit
POST.push((td, r, c) => { const v = wb.value(cur, r, c); if (typeof v !== 'number' || td.classList.contains('typing')) return; const w = td.clientWidth; if (w > 8 && td.scrollWidth > w + 1) td.textContent = '#'.repeat(Math.max(1, Math.floor(w / 9))); });
// Ctrl+wheel zoom
document.addEventListener('wheel', e => { if (e.ctrlKey && e.target.closest && e.target.closest('#wrap')) { e.preventDefault(); setZoom(Math.max(0.3, Math.min(3, +(zoom + (e.deltaY < 0 ? 0.1 : -0.1)).toFixed(2)))); } }, { passive: false, capture: true });
// Enter / Tab move inside a multi-cell selection
window.addEventListener('keydown', e => {
  if ((e.key !== 'Enter' && e.key !== 'Tab') || e.ctrlKey || e.altKey || e.metaKey || editing || editCell || /INPUT|TEXTAREA|SELECT/.test((e.target || {}).tagName || '')) return; if (sel.r1 === sel.r2 && sel.c1 === sel.c2) return; if (document.getElementById('mGen') && document.getElementById('mGen').style.display === 'flex') return;
  e.preventDefault(); e.stopImmediatePropagation(); const back = e.shiftKey, vert = e.key === 'Enter'; let { r, c } = anchor;
  if (vert) { r += back ? -1 : 1; if (r > sel.r2) { r = sel.r1; c++; if (c > sel.c2) c = sel.c1; } if (r < sel.r1) { r = sel.r2; c--; if (c < sel.c1) c = sel.c2; } } else { c += back ? -1 : 1; if (c > sel.c2) { c = sel.c1; r++; if (r > sel.r2) r = sel.r1; } if (c < sel.c1) { c = sel.c2; r--; if (r < sel.r1) r = sel.r2; } }
  setSel(sel.r1, sel.c1, sel.r2, sel.c2, r, c); try { scrollToCell && scrollToCell(r, c); } catch (x) { } updateFx();
}, true);
// Status bar options (right-click)
const SBO = (() => { try { return JSON.parse(localStorage.getItem('sbo') || 'null'); } catch (e) { return null; } })() || { avg: 1, cnt: 1, sum: 1, min: 0, max: 0, nc: 0 };
const _uf = updateFx; updateFx = function () { _uf(); const vals = [], all = []; for (let r = sel.r1; r <= sel.r2 && r <= ROWS; r++)for (let c = sel.c1; c <= sel.c2 && c <= COLS; c++) { const v = wb.value(cur, r, c); if (typeof v === 'number') vals.push(v); if (v !== null && v !== '' && v !== undefined) all.push(v); }
  if (!(sel.r1 === sel.r2 && sel.c1 === sel.c2) && all.length) { const p = []; const sum = vals.reduce((a, b) => a + b, 0); if (SBO.avg && vals.length) p.push('Average: ' + E.numStr(+(sum / vals.length).toPrecision(8))); if (SBO.cnt) p.push('Count: ' + all.length); if (SBO.nc && vals.length) p.push('Numerical Count: ' + vals.length); if (SBO.min && vals.length) p.push('Min: ' + E.numStr(Math.min(...vals))); if (SBO.max && vals.length) p.push('Max: ' + E.numStr(Math.max(...vals))); if (SBO.sum && vals.length) p.push('Sum: ' + E.numStr(+sum.toPrecision(12))); $('stR').textContent = p.join('   '); } else $('stR').textContent = ''; };
(function () { const el = document.getElementById('stR'); if (!el) return; el.addEventListener('contextmenu', e => { e.preventDefault(); const m = menuEl('ctxm'); m.innerHTML = [['avg', 'Average'], ['cnt', 'Count'], ['nc', 'Numerical Count'], ['min', 'Minimum'], ['max', 'Maximum'], ['sum', 'Sum']].map(([k, t]) => `<div data-k="${k}">${SBO[k] ? '✓' : '&nbsp;&nbsp;'} ${t}</div>`).join(''); m.querySelectorAll('div').forEach(x => x.onclick = () => { SBO[x.dataset.k] = SBO[x.dataset.k] ? 0 : 1; localStorage.setItem('sbo', JSON.stringify(SBO)); closeMenus(); updateFx(); }); showAt(m, e.clientX, e.clientY - 150); }); })();
// shortcuts: Alt+Shift+Right/Left group/ungroup, Ctrl+8 toggles detail, Alt+Shift+? none
window.addEventListener('keydown', e => { if (editing || editCell) return; if (e.altKey && e.shiftKey && e.key === 'ArrowRight') { e.preventDefault(); e.stopImmediatePropagation(); groupFromCell(); } else if (e.altKey && e.shiftKey && e.key === 'ArrowLeft') { e.preventDefault(); e.stopImmediatePropagation(); ungroupAuto(); } }, true);
/* ribbon */
(function () {
  const rib = document.getElementById('rib'); if (!rib) return; const grp = (label, btns) => `<div class="grp"><div class="gb">${btns.map(b => `<button class="big" onclick="${b[1]}"><i>${b[2]}</i>${b[0]}</button>`).join('')}</div><div class="gl">${label}</div></div>`; const add = (tab, h) => { const t = rib.querySelector(`[data-t="${tab}"]`); if (t) t.insertAdjacentHTML('beforeend', h); };
  add('Page Layout', grp('Page Setup', [['Page Setup', 'pageDlg()', '&#128209;'], ['Print Area', 'setPrintArea()', '&#9635;'], ['Clear Area', 'setPrintArea(true)', '&#10006;'], ['Print', 'printSheet()', '&#128424;']]));
  add('Data', grp('Outline', [['Group', 'groupAuto()', '&#10133;'], ['Ungroup', 'ungroupAuto()', '&#10134;'], ['Hide Detail', "showDetail(sel.c1===1&&sel.c2>=COLS||!(sel.r1===1&&sel.r2>=ROWS)?'r':'c',false)", '&#9660;'], ['Show Detail', "showDetail(sel.c1===1&&sel.c2>=COLS||!(sel.r1===1&&sel.r2>=ROWS)?'r':'c',true)", '&#9650;'], ['Subtotal', 'subtotalDlg()', '&#931;']]));
})();
