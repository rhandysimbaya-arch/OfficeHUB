// Office -> Image: render .docx / .xlsx / .pptx / .pdf pages to PNG or JPG
const IMG = (function () {
  let pages = [], baseName = 'document', busy = false;
  const $ = id => document.getElementById(id);
  const PW = 816, PH = 1056;
  const esc = s => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');

  function setStatus(t) { $('im-status').textContent = t; }
  function stage() { let s = document.getElementById('im-stage'); if (!s) { s = document.createElement('div'); s.id = 'im-stage'; s.style.cssText = 'position:absolute;left:-99999px;top:0;'; document.body.appendChild(s); } return s; }

  // render an HTML element (inline-styled, self-contained) to a canvas through SVG foreignObject
  function toCanvas(el, w, h, scale) {
    return new Promise((res, rej) => {
      const clone = el.cloneNode(true); clone.setAttribute('xmlns', 'http://www.w3.org/1999/xhtml');
      const xml = new XMLSerializer().serializeToString(clone);
      const svg = '<svg xmlns="http://www.w3.org/2000/svg" width="' + w + '" height="' + h + '"><foreignObject x="0" y="0" width="' + w + '" height="' + h + '">' + xml + '</foreignObject></svg>';
      const img = new Image();
      img.onload = () => { const cv = document.createElement('canvas'); cv.width = Math.round(w * scale); cv.height = Math.round(h * scale); const g = cv.getContext('2d'); g.fillStyle = '#fff'; g.fillRect(0, 0, cv.width, cv.height); g.scale(scale, scale); g.drawImage(img, 0, 0, w, h); res(cv); };
      img.onerror = () => rej(new Error('Could not render page'));
      img.src = 'data:image/svg+xml;charset=utf-8,' + encodeURIComponent(svg);
    });
  }

  // slice a tall block of html into pages at block boundaries
  async function paginate(innerHtml, css, width, scale) {
    const st = stage(); st.innerHTML = '<div id="im-m" style="width:' + width + 'px;box-sizing:border-box;padding:72px;background:#fff;' + css + '">' + innerHtml + '</div>';
    const m = st.firstChild; await new Promise(r => setTimeout(r, 60));
    const total = m.scrollHeight; const blocks = [...m.children].map(c => ({ top: c.offsetTop, bottom: c.offsetTop + c.offsetHeight }));
    const cuts = [0]; let start = 0;
    while (total - start > PH + 4) {
      let end = start + PH; const fit = blocks.filter(b => b.bottom <= end - 20 && b.bottom > start + 200); if (fit.length) end = fit[fit.length - 1].bottom + 12;
      cuts.push(end); start = end;
    }
    cuts.push(total); const out = [];
    for (let i = 0; i < cuts.length - 1; i++) {
      const top = cuts[i], h = Math.max(PH, cuts[i + 1] - top > PH ? PH : cuts[i + 1] - top + (i === 0 ? 0 : 0));
      const hh = i === cuts.length - 2 ? Math.max(Math.min(PH, cuts[i + 1] - top + 60), 200) : cuts[i + 1] - top;
      const w = document.createElement('div'); w.setAttribute('style', 'width:' + width + 'px;height:' + hh + 'px;overflow:hidden;position:relative;background:#fff');
      const inner = m.cloneNode(true); inner.removeAttribute('id'); inner.style.position = 'absolute'; inner.style.left = '0'; inner.style.top = (-top) + 'px'; w.appendChild(inner);
      out.push(await toCanvas(w, width, hh, scale)); setStatus('Rendering page ' + (i + 1) + ' of ' + (cuts.length - 1));
    }
    st.innerHTML = ''; return out;
  }

  async function fromDocx(buf, scale) {
    const r = await mammoth.convertToHtml({ arrayBuffer: buf });
    const css = 'font-family:Calibri,Arial,sans-serif;font-size:15px;line-height:1.35;color:#000;';
    const style = '<style>table{border-collapse:collapse;width:100%}td,th{border:1px solid #999;padding:4px}h1{font-size:26px;margin:14px 0 6px}h2{font-size:21px;margin:12px 0 5px}h3{font-size:17px}img{max-width:100%}p{margin:0 0 8px}ul{list-style:disc;padding-left:30px}ol{list-style:decimal;padding-left:30px}</style>';
    return paginate(style + r.value, css, PW, scale);
  }
  async function fromXlsx(buf, scale) {
    const wb = XLSX.read(buf, { type: 'array' }); const out = [];
    for (const n of wb.SheetNames) {
      const html = XLSX.utils.sheet_to_html(wb.Sheets[n], { header: '', footer: '' }).replace(/<\/?(html|head|body)[^>]*>/g, '').replace(/<meta[^>]*>/g, '').replace(/<title>.*?<\/title>/, '');
      const style = '<style>table{border-collapse:collapse;font-size:13px}td,th{border:1px solid #bbb;padding:3px 8px;white-space:nowrap}</style><h3 style="margin:0 0 8px;color:#217346">' + esc(n) + '</h3>';
      const probe = stage(); probe.innerHTML = '<div style="padding:72px;display:inline-block;font-family:Calibri,Arial">' + style + html + '</div>'; await new Promise(r => setTimeout(r, 40));
      const w = Math.max(PW, probe.firstChild.offsetWidth); probe.innerHTML = '';
      out.push(...await paginate(style + html, 'font-family:Calibri,Arial,sans-serif;', w, scale));
    }
    return out;
  }
  async function pptxToDeck(buf) {
    const z = await JSZip.loadAsync(buf); const pres = await z.file('ppt/presentation.xml').async('string');
    const m = /<p:sldSz[^>]*cx="(\d+)"[^>]*cy="(\d+)"/.exec(pres) || [0, 9144000, 5143500]; const E = 960 / +m[1]; const deck = { w: 960, h: Math.round(+m[2] * E), slides: [] };
    const files = Object.keys(z.files).filter(f => /^ppt\/slides\/slide\d+\.xml$/.test(f)).sort((a, b) => +a.match(/\d+/)[0] - +b.match(/\d+/)[0]);
    for (const f of files) {
      const xml = await z.file(f).async('string'); const doc = new DOMParser().parseFromString(xml, 'application/xml'); const rf = 'ppt/slides/_rels/' + f.split('/').pop() + '.rels'; const rels = {};
      if (z.file(rf)) { const rx = await z.file(rf).async('string'); [...rx.matchAll(/<Relationship [^>]*Id="([^"]+)"[^>]*Target="([^"]+)"/g)].forEach(r => rels[r[1]] = r[2]); }
      const s = { bg: '#ffffff', items: [] }; const bgm = /<p:bg>[\s\S]*?srgbClr val="([0-9A-Fa-f]{6})"/.exec(xml); if (bgm) s.bg = '#' + bgm[1]; const g = (el, t) => el.getElementsByTagName(t);
      for (const el of [...g(doc, 'p:sp'), ...g(doc, 'p:pic')]) {
        const off = g(el, 'a:off')[0], ext = g(el, 'a:ext')[0]; if (!off || !ext) continue;
        const it = { x: Math.round(+off.getAttribute('x') * E), y: Math.round(+off.getAttribute('y') * E), w: Math.round(+ext.getAttribute('cx') * E), h: Math.round(+ext.getAttribute('cy') * E) };
        if (el.nodeName === 'p:pic') { const bl = g(el, 'a:blip')[0]; const t = rels[bl && bl.getAttribute('r:embed')]; if (t) { const path = t.startsWith('/') ? t.slice(1) : 'ppt/' + t.replace('../', ''); const fl = z.file(path); if (fl) { const e2 = path.split('.').pop().toLowerCase(); it.type = 'image'; it.src = 'data:image/' + (e2 === 'jpg' ? 'jpeg' : e2) + ';base64,' + await fl.async('base64'); s.items.push(it); } } continue; }
        const text = [...g(el, 'a:p')].map(p => [...g(p, 'a:t')].map(t => t.textContent).join('')).join('\n'); const rp = g(el, 'a:rPr')[0];
        if (text.trim()) { it.type = 'text'; it.text = text; it.size = rp && rp.getAttribute('sz') ? +rp.getAttribute('sz') / 100 : 20; it.bold = !!(rp && rp.getAttribute('b') === '1'); const cl = rp && rp.getElementsByTagName('a:srgbClr')[0]; if (cl) it.color = '#' + cl.getAttribute('val'); const al = g(el, 'a:pPr')[0]; if (al) { const a = al.getAttribute('algn'); it.align = a === 'ctr' ? 'center' : a === 'r' ? 'right' : 'left'; } s.items.push(it); }
        else { const sf = el.getElementsByTagName('a:srgbClr')[0]; if (sf) { it.type = 'shape'; const pgm = g(el, 'a:prstGeom')[0]; it.shape = pgm && pgm.getAttribute('prst') === 'ellipse' ? 'ellipse' : 'rect'; it.fill = '#' + sf.getAttribute('val'); s.items.push(it); } }
      }
      deck.slides.push(s);
    }
    return deck;
  }
  function slideHtml(s, deck) {
    let h = '<div style="position:relative;width:' + deck.w + 'px;height:' + deck.h + 'px;overflow:hidden;background:' + s.bg + ';font-family:Calibri,Arial,sans-serif">';
    s.items.forEach(it => {
      let st = 'position:absolute;box-sizing:border-box;left:' + it.x + 'px;top:' + it.y + 'px;width:' + it.w + 'px;height:' + it.h + 'px;';
      if (it.type === 'text') h += '<div style="' + st + 'overflow:hidden;line-height:1.2;white-space:pre-wrap;padding:2px 4px;font-size:' + ((it.size || 20) / 0.75) + 'px;font-weight:' + (it.bold ? 700 : 400) + ';font-style:' + (it.italic ? 'italic' : 'normal') + ';color:' + (it.color || '#000') + ';text-align:' + (it.align || 'left') + (it.fill ? ';background:' + it.fill : '') + '">' + esc(it.text) + '</div>';
      else if (it.type === 'image') h += '<img style="' + st + '" src="' + it.src + '"/>';
      else if (it.type === 'shape') h += '<div style="' + st + 'background:' + (it.fill || '#c43e1c') + (it.shape === 'ellipse' ? ';border-radius:50%' : '') + '"></div>';
    });
    return h + '</div>';
  }
  async function fromPptx(buf, scale) {
    const deck = await pptxToDeck(buf); const out = [];
    for (let i = 0; i < deck.slides.length; i++) { const d = document.createElement('div'); d.innerHTML = slideHtml(deck.slides[i], deck); out.push(await toCanvas(d.firstChild, deck.w, deck.h, scale)); setStatus('Rendering slide ' + (i + 1) + ' of ' + deck.slides.length); }
    return out;
  }
  async function fromPdf(buf, scale) {
    const doc = await pdfjsLib.getDocument({ data: new Uint8Array(buf) }).promise; const out = [];
    for (let i = 1; i <= doc.numPages; i++) { const p = await doc.getPage(i); const vp = p.getViewport({ scale: 1.5 * scale }); const cv = document.createElement('canvas'); cv.width = vp.width; cv.height = vp.height; const g = cv.getContext('2d'); g.fillStyle = '#fff'; g.fillRect(0, 0, cv.width, cv.height); await p.render({ canvasContext: g, viewport: vp }).promise; out.push(cv); setStatus('Rendering page ' + i + ' of ' + doc.numPages); }
    return out;
  }

  function pick() {
    RB.pickFile('.docx,.xlsx,.xls,.pptx,.pdf', async (name, buf) => {
      if (busy) return; busy = true; pages = []; $('im-grid').innerHTML = ''; $('im-actions').style.display = 'none';
      baseName = name.replace(/\.[^.]+$/, ''); $('im-fname').textContent = name; const scale = +$('im-scale').value;
      try {
        setStatus('Rendering...'); const ext = name.split('.').pop().toLowerCase();
        pages = ext === 'docx' ? await fromDocx(buf, scale) : (ext === 'xlsx' || ext === 'xls') ? await fromXlsx(buf, scale) : ext === 'pptx' ? await fromPptx(buf, scale) : ext === 'pdf' ? await fromPdf(buf, scale) : (() => { throw new Error('Unsupported file type'); })();
        setStatus(pages.length + ' image(s) ready.'); show();
      } catch (e) { setStatus('Error: ' + (e.message || e)); }
      busy = false;
    });
  }
  function fmt() { const f = $('im-fmt').value; return { mime: f === 'jpg' ? 'image/jpeg' : f === 'pdf' ? 'application/pdf' : 'image/png', ext: f }; }
  function blobOf(cv) { const f = fmt(); if (f.ext === 'pdf') return TOOLS.canvasesToPdf([cv]); return new Promise(r => cv.toBlob(r, f.mime, 0.92)); }
  function show() {
    const g = $('im-grid'); g.innerHTML = '';
    pages.forEach((cv, i) => { const d = document.createElement('div'); d.style.cssText = 'background:#fff;border-radius:8px;padding:6px;box-shadow:0 1px 3px rgba(0,0,0,.2);text-align:center'; const im = document.createElement('img'); im.style.cssText = 'width:100%;border:1px solid #ddd'; im.src = cv.toDataURL('image/jpeg', 0.5); d.appendChild(im); const b = document.createElement('button'); b.textContent = 'Save #' + (i + 1); b.style.cssText = 'margin-top:6px;padding:6px 10px;border:0;border-radius:6px;background:#0f766e;color:#fff;font-size:12px'; b.onclick = async () => RB.saveBlob(baseName + '_' + (i + 1) + '.' + fmt().ext, await blobOf(cv), fmt().mime); d.appendChild(b); g.appendChild(d); });
    $('im-actions').style.display = pages.length ? 'block' : 'none';
  }
  async function saveAll() {
    if (fmt().ext === 'pdf') { RB.saveBlob(baseName + '.pdf', await TOOLS.canvasesToPdf(pages), 'application/pdf'); return; }
    if (pages.length === 1) { RB.saveBlob(baseName + '.' + fmt().ext, await blobOf(pages[0]), fmt().mime); return; }
    const z = new JSZip(); for (let i = 0; i < pages.length; i++) z.file(baseName + '_' + String(i + 1).padStart(2, '0') + '.' + fmt().ext, await blobOf(pages[i]));
    RB.saveBlob(baseName + '_images.zip', await z.generateAsync({ type: 'blob' }), 'application/zip');
  }
  return { pick, saveAll, show, _t: { toCanvas, fromDocx, fromXlsx, fromPptx } };
})();
