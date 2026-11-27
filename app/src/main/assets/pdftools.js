// PDF tools: merge, split/extract/rotate, images -> PDF, PDF -> text, canvases -> PDF
const TOOLS = (function () {
  const $ = id => document.getElementById(id);
  let list = [], tool = 'merge';
  const T = { merge: ['Merge PDFs', 'Add PDF files in the order you want, then press Run.', '.pdf', true], split: ['Split / extract / rotate pages', 'Choose one PDF. Pages: e.g. 1-3,5 (blank = all).', '.pdf', false], img2pdf: ['Images to PDF', 'Add JPG / PNG pictures in order, then press Run.', 'image/*', true], pdf2txt: ['PDF to Text (.txt)', 'Uses OCR automatically for scanned pages.', '.pdf', false] };
  function setTool(t) { tool = t; list = []; $('tl-info').textContent = T[t][1]; $('tl-split').style.display = t === 'split' ? 'block' : 'none'; $('tl-add').textContent = T[t][3] ? 'Add file' : 'Choose file'; draw(); $('tl-status').textContent = ''; }
  function add() { RB.pickFile(T[tool][2], (name, buf) => { if (!T[tool][3]) list = []; list.push({ name, buf }); draw(); }); }
  function draw() { $('tl-list').innerHTML = list.map((f, i) => `<div style="display:flex;gap:6px;align-items:center;padding:6px 0;border-bottom:1px solid #eee;font-size:13px"><span style="flex:1;overflow:hidden;text-overflow:ellipsis">${i + 1}. ${f.name.replace(/</g, '&lt;')}</span>${T[tool][3] ? `<button onclick="TOOLS.mv(${i},-1)">&uarr;</button><button onclick="TOOLS.mv(${i},1)">&darr;</button>` : ''}<button onclick="TOOLS.rm(${i})">&times;</button></div>`).join(''); }
  function mv(i, d) { const j = i + d; if (j < 0 || j >= list.length) return; [list[i], list[j]] = [list[j], list[i]]; draw(); }
  function rm(i) { list.splice(i, 1); draw(); }
  function pages(txt, n) { if (!txt.trim()) return Array.from({ length: n }, (_, i) => i); const s = new Set(); txt.split(',').forEach(p => { const m = p.trim().match(/^(\d+)(?:\s*-\s*(\d+))?$/); if (!m) return; const a = +m[1], b = m[2] ? +m[2] : a; for (let i = Math.min(a, b); i <= Math.max(a, b); i++) if (i >= 1 && i <= n) s.add(i - 1); }); return [...s].sort((a, b) => a - b); }
  async function run() {
    if (!list.length) { $('tl-status').textContent = 'Add a file first.'; return; }
    const st = m => $('tl-status').textContent = m; const PL = PDFLib;
    try {
      if (tool === 'merge') {
        if (list.length < 2) { st('Add at least two PDFs.'); return; }
        const out = await PL.PDFDocument.create();
        for (let i = 0; i < list.length; i++) { st('Merging ' + (i + 1) + '/' + list.length); const d = await PL.PDFDocument.load(list[i].buf, { ignoreEncryption: true }); (await out.copyPages(d, d.getPageIndices())).forEach(p => out.addPage(p)); }
        RB.saveBlob('merged.pdf', new Blob([await out.save()], { type: 'application/pdf' }), 'application/pdf'); st('Done: merged.pdf (' + out.getPageCount() + ' pages)');
      } else if (tool === 'split') {
        const d = await PL.PDFDocument.load(list[0].buf, { ignoreEncryption: true }); const idx = pages($('tl-range').value, d.getPageCount()); if (!idx.length) { st('No valid pages.'); return; }
        const rot = +$('tl-rot').value, mode = $('tl-smode').value, base = list[0].name.replace(/\.pdf$/i, '');
        if (mode === 'each') { const z = new JSZip(); for (const i of idx) { const o = await PL.PDFDocument.create(); const [p] = await o.copyPages(d, [i]); if (rot) p.setRotation(PL.degrees((p.getRotation().angle + rot) % 360)); o.addPage(p); z.file(base + '_page' + (i + 1) + '.pdf', await o.save()); st('Splitting ' + (i + 1)); } RB.saveBlob(base + '_pages.zip', await z.generateAsync({ type: 'blob' }), 'application/zip'); st('Done: ' + idx.length + ' files in a ZIP.'); }
        else { const o = await PL.PDFDocument.create(); (await o.copyPages(d, idx)).forEach(p => { if (rot) p.setRotation(PL.degrees((p.getRotation().angle + rot) % 360)); o.addPage(p); }); RB.saveBlob(base + '_extract.pdf', new Blob([await o.save()], { type: 'application/pdf' }), 'application/pdf'); st('Done: ' + idx.length + ' page(s).'); }
      } else if (tool === 'img2pdf') {
        const out = await PL.PDFDocument.create();
        for (let i = 0; i < list.length; i++) { st('Adding image ' + (i + 1)); const f = list[i]; let bytes = new Uint8Array(f.buf); let img;
          const isPng = bytes[0] === 0x89 && bytes[1] === 0x50, isJpg = bytes[0] === 0xFF && bytes[1] === 0xD8;
          if (isPng) img = await out.embedPng(bytes); else if (isJpg) img = await out.embedJpg(bytes); else { const bm = await createImageBitmap(new Blob([bytes])); const cv = document.createElement('canvas'); cv.width = bm.width; cv.height = bm.height; cv.getContext('2d').drawImage(bm, 0, 0); img = await out.embedJpg(new Uint8Array(await (await new Promise(r => cv.toBlob(r, 'image/jpeg', 0.92))).arrayBuffer())); }
          const pg = out.addPage([img.width, img.height]); pg.drawImage(img, { x: 0, y: 0, width: img.width, height: img.height }); }
        RB.saveBlob('images.pdf', new Blob([await out.save()], { type: 'application/pdf' }), 'application/pdf'); st('Done: images.pdf');
      } else if (tool === 'pdf2txt') {
        pdfDoc = await pdfjsLib.getDocument({ data: new Uint8Array(list[0].buf) }).promise; let txt = '';
        for (let i = 1; i <= pdfDoc.numPages; i++) { st('Reading page ' + i + '/' + pdfDoc.numPages); const p = await extractPage(i); txt += p.lines.map(L => L.segs.map(s => s.s).join('\t')).join('\n') + '\n\n'; }
        RB.saveBlob(list[0].name.replace(/\.pdf$/i, '') + '.txt', new Blob([txt], { type: 'text/plain' }), 'text/plain'); st('Done: ' + pdfDoc.numPages + ' page(s) as text.');
      }
    } catch (e) { st('Error: ' + (e.message || e)); }
  }
  // canvases (JPEG) -> single PDF blob
  async function canvasesToPdf(canvases) {
    const out = await PDFLib.PDFDocument.create();
    for (const cv of canvases) { const b = await new Promise(r => cv.toBlob(r, 'image/jpeg', 0.9)); const img = await out.embedJpg(new Uint8Array(await b.arrayBuffer())); const pg = out.addPage([cv.width * 0.75, cv.height * 0.75]); pg.drawImage(img, { x: 0, y: 0, width: cv.width * 0.75, height: cv.height * 0.75 }); }
    return new Blob([await out.save()], { type: 'application/pdf' });
  }
  return { setTool, add, run, mv, rm, canvasesToPdf };
})();
