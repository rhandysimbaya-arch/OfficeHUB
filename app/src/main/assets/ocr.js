// Offline OCR using the bundled Tesseract worker (no network). Files are read via XHR and handed to the worker as blobs,
// so this also works from file:// inside an Android WebView.
const OCR = (function () {
  let w = null, loading = null, jid = 0; const pending = {};
  function xhr(url, type) { const viaX = () => new Promise((res, rej) => { const x = new XMLHttpRequest(); x.open('GET', url); x.responseType = type; x.onload = () => (x.status === 0 || x.status === 200) && x.response ? res(x.response) : rej(new Error('Cannot load ' + url)); x.onerror = () => rej(new Error('Cannot load ' + url)); x.send(); }); return viaX().catch(() => fetch(url).then(r => { if (!r.ok && r.status !== 0) throw new Error('Cannot load ' + url); return type === 'blob' ? r.blob() : r.arrayBuffer(); })); }
  function call(action, payload, transfer, onprog) {
    return new Promise((res, rej) => { const id = ++jid; pending[id] = { res, rej, onprog }; w.postMessage({ workerId: 'hub', jobId: id, action, payload }, transfer || []); });
  }
  async function init() {
    if (w) return w; if (loading) return loading;
    loading = (async () => {
      const wjs = await xhr('libs/ocr/worker.min.js', 'blob'), cjs = await xhr('libs/ocr/core.js', 'blob'), lang = new Uint8Array(await xhr('libs/ocr/eng.lang.bin', 'arraybuffer'));
      w = new Worker(URL.createObjectURL(wjs));
      w.onmessage = e => { const d = e.data, p = pending[d.jobId]; if (!p) return; if (d.status === 'resolve') { delete pending[d.jobId]; p.res(d.data); } else if (d.status === 'reject') { delete pending[d.jobId]; p.rej(new Error(String(d.data))); } else if (d.status === 'progress' && p.onprog && d.data) p.onprog(d.data); };
      w.onerror = e => { Object.keys(pending).forEach(k => { pending[k].rej(new Error('OCR worker error: ' + (e.message || ''))); delete pending[k]; }); };
      await call('load', { options: { lstmOnly: true, corePath: URL.createObjectURL(cjs) + '#core.js', logging: false } });
      await call('loadLanguage', { langs: [{ code: 'eng', data: lang }], options: { cacheMethod: 'none', gzip: false, lstmOnly: true } }, [lang.buffer]);
      await call('initialize', { langs: 'eng', oem: 1, config: {} });
      return w;
    })();
    try { return await loading; } catch (e) { loading = null; if (w) { w.terminate(); w = null; } throw e; }
  }
  function canvasBytes(cv) { return new Promise(r => cv.toBlob(async b => r(new Uint8Array(await b.arrayBuffer())), 'image/png')); }
  // returns word items in PDF-point space: {s,x,y(top-down baseline),w,size,bold,italic}
  async function pageItems(canvas, scale, progress) {
    await init(); const bytes = await canvasBytes(canvas);
    const out = await call('recognize', { image: bytes, options: {}, output: { blocks: true, text: true } }, [bytes.buffer], d => { if (progress && d.status === 'recognizing text') progress(d.progress || 0); });
    const data = out && out.data && (out.data.lines || out.data.blocks) ? out.data : out;
    let lines = data.lines;
    if (!lines && data.blocks) { lines = []; data.blocks.forEach(b => (b.paragraphs || []).forEach(p => (p.lines || []).forEach(l => lines.push(l)))); }
    const items = [];
    (lines || []).forEach(L => {
      const words = (L.words || []).filter(x => x.text && x.text.trim() && x.confidence > 25); if (!words.length) return;
      const hs = words.map(x => x.bbox.y1 - x.bbox.y0).sort((a, b) => a - b); const mh = hs[hs.length >> 1] / scale; const size = Math.max(6, mh * 0.82);
      words.forEach(x => items.push({ s: x.text, x: x.bbox.x0 / scale, y: (L.bbox.y1 / scale) - size * 0.22, w: (x.bbox.x1 - x.bbox.x0) / scale, size, bold: !!x.is_bold, italic: !!x.is_italic }));
    });
    return items;
  }
  return { init, pageItems };
})();
