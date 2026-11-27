// PDF extraction + conversion to Word / Excel / PowerPoint
pdfjsLib.GlobalWorkerOptions.workerSrc = 'libs/pdf.worker.min.js';
const COL = {word:'#2b579a', excel:'#217346', ppt:'#c43e1c', img:'#0f766e', tools:'#475569'};
let mode = 'word', pdfDoc = null, pdfName = 'document', outBlob = null, outName = '', handoff = null;
const $ = id => document.getElementById(id);

function setMode(m){
  mode = m;
  $('pdfcard').style.display = (m==='img'||m==='tools') ? 'none' : ''; $('imgcard').style.display = m==='img' ? '' : 'none'; $('toolcard').style.display = m==='tools' ? '' : 'none'; if (m==='tools') TOOLS.setTool($('tl-sel').value);
  ['word','excel','ppt','img','tools'].forEach(k => { const b = $('t-'+k); b.classList.toggle('on', k===m); b.style.background = k===m ? COL[k] : '#fff'; });
  $('top').style.background = COL[m]; $('go').style.background = COL[(m==='img'||m==='tools')?'word':m]; $('dl').style.background = COL[m]; $('openapp').style.background = '#e5e7eb';
  $('o-excel').classList.toggle('show', m==='excel'); $('o-ppt').classList.toggle('show', m==='ppt');
  $('openapp').textContent = 'Open in ' + ({word:'Word editor',excel:'Excel',ppt:'PowerPoint'}[m]||'');
  $('result').style.display = 'none';
}
function pickPdf(){
  RB.pickFile('.pdf,application/pdf', async (name, buf) => {
    try {
      $('status').textContent = 'Loading PDF...';
      pdfDoc = await pdfjsLib.getDocument({data: new Uint8Array(buf)}).promise;
      pdfName = name.replace(/\.pdf$/i, '');
      $('fname').textContent = name + ' — ' + pdfDoc.numPages + ' page(s)';
      $('go').disabled = false; $('status').textContent = '';
    } catch(e){ $('status').textContent = 'Could not open PDF: ' + e.message; }
  });
}
function pageList(n){
  const t = $('range').value.trim(); if (!t) return Array.from({length:n}, (_,i)=>i+1);
  const s = new Set();
  t.split(',').forEach(p => { const m = p.trim().match(/^(\d+)(?:\s*-\s*(\d+))?$/); if (!m) return; const a = +m[1], b = m[2] ? +m[2] : a; for (let i=Math.min(a,b); i<=Math.max(a,b); i++) if (i>=1 && i<=n) s.add(i); });
  return [...s].sort((a,b)=>a-b);
}
function progress(i, n, msg){ $('prog').style.display = 'block'; $('bar').style.width = Math.round(i/n*100) + '%'; $('status').textContent = msg; }

// ---- text extraction: items -> lines -> cells ----
async function extractPage(pageNo){
  const page = await pdfDoc.getPage(pageNo);
  const vp = page.getViewport({scale:1});
  const tc = await page.getTextContent();
  let items = tc.items.filter(it => it.str && it.str.trim() !== '' || (it.str === ' ')).map(it => {
    const size = Math.abs(it.transform[3]) || it.height || 10;
    const st = tc.styles && tc.styles[it.fontName] || {};
    const fn = (st.fontFamily || '') + ' ' + (it.fontName || '');
    return {s: it.str, x: it.transform[4], y: vp.height - it.transform[5], w: it.width, size, bold: /bold|black|heavy/i.test(fn), italic: /italic|oblique/i.test(fn)};
  }).filter(it => it.s.trim() !== '');
  const om = $('ocr-mode') ? $('ocr-mode').value : 'auto'; let ocred = false;
  const chars = items.reduce((a,i)=>a+i.s.trim().length,0);
  if (om === 'always' || (om === 'auto' && chars < 8)) {
    const sc = Math.min(3, 2200 / Math.max(vp.width, vp.height));
    const v2 = page.getViewport({scale: sc}); const cv = document.createElement('canvas'); cv.width = Math.round(v2.width); cv.height = Math.round(v2.height);
    const g = cv.getContext('2d'); g.fillStyle = '#fff'; g.fillRect(0,0,cv.width,cv.height);
    await page.render({canvasContext: g, viewport: v2}).promise;
    progress(0, 1, 'OCR page ' + pageNo + ' (first run loads the OCR engine)...');
    items = await OCR.pageItems(cv, sc, p => progress(p, 1, 'OCR page ' + pageNo + ': ' + Math.round(p*100) + '%'));
    cv.width = cv.height = 1; ocred = true;
  }
  // group into lines by y
  items.sort((a,b) => a.y - b.y || a.x - b.x);
  const lines = [];
  items.forEach(it => {
    let L = lines[lines.length-1];
    if (L && Math.abs(it.y - L.y) <= Math.max(2, Math.min(it.size, L.size) * 0.5)) { L.items.push(it); L.y = (L.y + it.y) / 2; L.size = Math.max(L.size, it.size); }
    else lines.push({y: it.y, size: it.size, items: [it]});
  });
  lines.forEach(L => {
    L.items.sort((a,b) => a.x - b.x);
    // text pieces ("segments") split by wide gaps
    const segs = []; let cur = null;
    L.items.forEach(it => {
      if (cur && it.x - cur.end <= Math.max(L.size * 1.2, 7)) {
        const gap = it.x - cur.end; cur.s += (gap > L.size * 0.15 ? ' ' : '') + it.s; cur.end = it.x + it.w; cur.bold = cur.bold && it.bold; cur.italic = cur.italic && it.italic;
      } else { cur = {s: it.s, x: it.x, end: it.x + it.w, bold: it.bold, italic: it.italic, size: it.size}; segs.push(cur); }
    });
    L.segs = segs; L.x = segs[0].x; L.end = segs[segs.length-1].end;
    L.text = segs.map(s => s.s).join(' ');
  });
  return {pageNo, width: vp.width, height: vp.height, lines, page, ocred};
}

// ---- Word ----
function buildWordModel(pages){
  const sizes = []; pages.forEach(p => p.lines.forEach(L => { for (let i=0;i<L.text.length;i++) sizes.push(L.size); }));
  sizes.sort((a,b)=>a-b); const body = sizes.length ? sizes[Math.floor(sizes.length/2)] : 11;
  const blocks = [], html = [];
  const esc = s => s.replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
  pages.forEach((p, pi) => {
    let para = null, prev = null;
    const flush = () => {
      if (!para) return;
      const txt = para.text.join(' ');
      const big = para.size / body;
      const heading = (big >= 1.5 && txt.length < 120) ? 1 : (big >= 1.2 && txt.length < 120) ? 2 : 0;
      blocks.push({type:'p', heading: heading || undefined, align: para.align, runs: [{t: txt, b: para.bold && !heading ? true : undefined, i: para.italic || undefined, sz: heading ? undefined : Math.max(8, Math.min(20, Math.round(para.size)))}]});
      const tag = heading ? 'h' + heading : 'p';
      html.push('<' + tag + (para.align && para.align !== 'left' ? ' style="text-align:' + para.align + '"' : '') + '>' + (para.bold && !heading ? '<b>' : '') + esc(txt) + (para.bold && !heading ? '</b>' : '') + '</' + tag + '>');
      para = null;
    };
    // group consecutive multi-column lines (2+ segments) into tables
    const seq = []; let run = [];
    const endRun = () => { if (run.length >= 2) seq.push({table: run}); else run.forEach(l => seq.push(l)); run = []; };
    p.lines.forEach(L => { if (L.segs.length >= 2) run.push(L); else { endRun(); seq.push(L); } });
    endRun();
    seq.forEach(L => {
      if (L.table) {
        flush(); prev = null;
        const rows = rowsFromLines(L.table).map(r => r.map(v => String(v)));
        blocks.push({type:'table', rows});
        html.push('<table border="1" style="border-collapse:collapse;width:100%">' + rows.map(r => '<tr>' + r.map(c => '<td>' + esc(c) + '</td>').join('') + '</tr>').join('') + '</table><p><br></p>');
        return;
      }
      const centered = Math.abs((L.x + L.end) / 2 - p.width / 2) < 14 && L.x > 60;
      const align = centered ? 'center' : 'left';
      const gap = prev ? L.y - prev.y : 0;
      const bullet = /^([••●▪\-–*]|\d+[.)])\s/.test(L.text);
      const newPara = !para || gap > prev.size * 1.7 || Math.abs(L.size - para.size) > 1.5 || align !== para.align || bullet || (para.text.length && /[.!?:]$/.test(para.text[para.text.length-1]) && L.x > prev.x + 8);
      if (newPara) { flush(); para = {text: [L.text], size: L.size, align, bold: L.segs.every(s => s.bold), italic: L.segs.every(s => s.italic)}; }
      else { para.text.push(L.text); para.bold = para.bold && L.segs.every(s => s.bold); }
      prev = L;
    });
    flush();
    if (pi < pages.length - 1) { blocks.push({type:'pagebreak'}); html.push('<hr style="page-break-after:always;border:0">'); }
  });
  return {blocks, html: html.join('\n'), page: {w: 12240, h: 15840}};
}

// ---- Excel ----
function buildRows(p){ return rowsFromLines(p.lines); }
function rowsFromLines(lines){
  // cluster segment x-starts into columns
  const xs = []; lines.forEach(L => L.segs.forEach(s => xs.push(s.x))); xs.sort((a,b)=>a-b);
  const cols = []; xs.forEach(x => { const c = cols[cols.length-1]; if (c && x - c.max <= 6) { c.max = x; c.sum += x; c.n++; } else cols.push({min:x, max:x, sum:x, n:1}); });
  const colX = cols.map(c => c.sum / c.n);
  const rows = lines.map(L => {
    const row = Array(colX.length).fill('');
    L.segs.forEach(s => { let best = 0, bd = 1e9; colX.forEach((cx, i) => { const d = Math.abs(cx - s.x); if (d < bd) { bd = d; best = i; } }); row[best] = row[best] ? row[best] + ' ' + s.s : s.s; });
    return row;
  });
  // drop fully empty columns
  const keep = colX.map((_, i) => rows.some(r => r[i] !== ''));
  return rows.map(r => r.filter((_, i) => keep[i])).map(r => r.map(v => { const t = String(v).trim().replace(/,/g,''); return (t !== '' && /^[-+]?\d+(\.\d+)?$/.test(t)) ? Number(t) : v; }));
}

// ---- PowerPoint ----
const SLIDE_W = 960;
async function buildPpt(pages, asImage){
  const p0 = pages[0]; const wIn = p0.width / 72, hIn = p0.height / 72;
  const pptx = new PptxGenJS(); pptx.defineLayout({name:'PDF', width: wIn, height: hIn}); pptx.layout = 'PDF';
  const deck = {w: SLIDE_W, h: Math.round(SLIDE_W * p0.height / p0.width), slides: []};
  for (let i = 0; i < pages.length; i++) {
    const p = pages[i]; const slide = pptx.addSlide(); const k = SLIDE_W / p.width; const s = {bg:'#ffffff', items: []};
    if (asImage) {
      const vp = p.page.getViewport({scale: 2}); const cv = document.createElement('canvas'); cv.width = vp.width; cv.height = vp.height;
      await p.page.render({canvasContext: cv.getContext('2d'), viewport: vp}).promise;
      const data = cv.toDataURL('image/jpeg', 0.85);
      slide.addImage({data, x:0, y:0, w: p.width/72, h: p.height/72});
      s.items.push({type:'image', x:0, y:0, w: deck.w, h: deck.h, src: data});
    } else {
      p.lines.forEach(L => {
        const sz = Math.max(6, Math.round(L.size)); const x = L.x, w = Math.max(20, L.end - L.x) + 10, h = L.size * 1.35;
        slide.addText(L.text, {x: x/72, y: (L.y - L.size)/72, w: w/72, h: h/72, fontSize: sz, bold: L.segs.every(q=>q.bold), italic: L.segs.every(q=>q.italic), margin: 0, valign:'top', fit:'none', wrap:false});
        s.items.push({type:'text', x: Math.round(x*k), y: Math.round((L.y - L.size)*k), w: Math.round(w*k), h: Math.round(h*k), text: L.text, size: Math.round(sz*k*0.75*10)/10, bold: L.segs.every(q=>q.bold), italic: L.segs.every(q=>q.italic), color:'#000000', align:'left'});
      });
    }
    deck.slides.push(s);
    progress(i+1, pages.length, 'Building slide ' + (i+1) + ' of ' + pages.length);
  }
  return {pptx, deck};
}

async function convert(){
  if (!pdfDoc) return;
  $('go').disabled = true; $('result').style.display = 'none';
  try {
    const nums = pageList(pdfDoc.numPages); if (!nums.length) throw new Error('No valid pages in range');
    const pages = [];
    for (let i = 0; i < nums.length; i++) { progress(i, nums.length, 'Reading page ' + nums[i] + '...'); pages.push(await extractPage(nums[i])); }
    const nText = pages.reduce((a,p)=>a+p.lines.length, 0);
    if (!nText && !(mode==='ppt' && $('pp-mode').value==='image')) throw new Error('No readable text found, even with OCR. The scan may be too blurry or not English text.');
    let info = nums.length + ' page(s), ' + nText + ' text lines.' + (pages.some(p=>p.ocred) ? ' Scanned pages were read with OCR - please proof-read numbers.' : '');
    if (mode === 'word') {
      const m = buildWordModel(pages);
      outBlob = await Docx.build({blocks: m.blocks, page: m.page}); outName = pdfName + '.docx';
      handoff = {key:'hub_import_word', data: m.html, page: 'word.html'};
    } else if (mode === 'excel') {
      const wb = XLSX.utils.book_new(); const sheets = [];
      if ($('xl-layout').value === 'per') pages.forEach(p => { const r = buildRows(p); sheets.push({name: 'Page ' + p.pageNo, rows: r}); });
      else { let all = []; pages.forEach(p => { all = all.concat(buildRows(p)); }); sheets.push({name: 'Sheet1', rows: all}); }
      sheets.forEach(sh => XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(sh.rows.length ? sh.rows : [[]]), sh.name));
      const arr = XLSX.write(wb, {type:'array', bookType:'xlsx'});
      outBlob = new Blob([arr], {type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'}); outName = pdfName + '.xlsx';
      handoff = {key:'hub_import_sheet', data: JSON.stringify({sheets}), page: 'sheet.html'};
    } else {
      const img = $('pp-mode').value === 'image';
      const r = await buildPpt(pages, img);
      const arr = await r.pptx.write({outputType:'arraybuffer'});
      outBlob = new Blob([arr], {type:'application/vnd.openxmlformats-officedocument.presentationml.presentation'}); outName = pdfName + '.pptx';
      handoff = {key:'hub_import_slides', data: JSON.stringify(r.deck), page: 'slides.html'};
    }
    $('rtitle').textContent = outName; $('rinfo').textContent = info; $('result').style.display = 'block'; progress(1, 1, 'Done.');
  } catch(e){ $('status').textContent = 'Error: ' + e.message; }
  $('go').disabled = false;
}
function download(){ if (outBlob) RB.saveBlob(outName, outBlob); }
function openInApp(){
  if (!handoff) return;
  try { localStorage.setItem(handoff.key, handoff.data); location.href = handoff.page; }
  catch(e){ $('status').textContent = 'Too large to open directly - please download instead.'; }
}
setMode('word');
