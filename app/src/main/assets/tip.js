/* Hover tooltips for mouse users: shows a small label when the pointer rests on an icon or control. */
(function () {
  if (window.__tip) return; window.__tip = 1;
  const css = document.createElement('style');
  css.textContent = '#hubtip{position:fixed;z-index:3000;pointer-events:none;background:#2b2b2b;color:#fff;font:12px/1.3 "Segoe UI",Roboto,Arial,sans-serif;padding:5px 9px;border-radius:5px;box-shadow:0 3px 10px rgba(0,0,0,.35);max-width:260px;opacity:0;transition:opacity .12s;white-space:normal}#hubtip.on{opacity:1}#hubtip small{display:block;color:#bbb;margin-top:1px;font-size:11px}';
  document.head.appendChild(css);
  const tip = document.createElement('div'); tip.id = 'hubtip'; document.addEventListener('DOMContentLoaded', () => document.body.appendChild(tip)); if (document.body) document.body.appendChild(tip);
  const FALL = [[/zoomStep\(-|zoomBy\(-/, 'Zoom out'], [/zoomStep\(|zoomBy\(0/, 'Zoom in'], [/setZoom\(1\)/, 'Reset zoom to 100%'], [/wSide\(null\)/, 'Close pane'], [/fsz\(|szPt\(|fmt\('size'/, 'Font size']];
  const human = s => { const m = /([A-Za-z_$][\w$]*)\s*\(/.exec(s || ''); if (!m) return ''; let n = m[1].replace(/^[wx](?=[A-Z])/, '').replace(/([a-z])([A-Z])/g, '$1 $2').replace(/_/g, ' '); return n.charAt(0).toUpperCase() + n.slice(1); };
  function labelOf(el) {
    let t = el.getAttribute('data-tip'); if (t) return t;
    t = (el.getAttribute('title') || '').trim();
    if (t) { el.setAttribute('data-tip', t); el.removeAttribute('title'); return t; }
    t = (el.getAttribute('aria-label') || '').trim(); if (t) return t;
    const letters = (el.textContent || '').replace(/[^A-Za-z]/g, '');
    if (letters.length >= 3 && !el.matches('select')) return '';        // already has a readable text label
    const oc = (el.getAttribute('onclick') || el.getAttribute('onchange') || '');
    for (const [re, l] of FALL) if (re.test(oc)) return l;
    if (el.matches('select') && el.options && el.options[0]) return el.options[0].text.replace(/[^\w ]/g, '').trim();
    return human(oc);
  }
  let timer = 0, cur = null;
  function hide() { clearTimeout(timer); tip.classList.remove('on'); cur = null; }
  function show(el, text) {
    tip.innerHTML = ''; const parts = text.split(/\s*\((Ctrl\+[^)]+|Alt\+[^)]+|F\d+)\)\s*$/); tip.textContent = parts[0]; if (parts[1]) { const s = document.createElement('small'); s.textContent = parts[1]; tip.appendChild(s); }
    tip.style.left = '0px'; tip.style.top = '0px'; tip.classList.add('on');
    const r = el.getBoundingClientRect(), w = tip.offsetWidth, h = tip.offsetHeight;
    let x = Math.max(4, Math.min(r.left + r.width / 2 - w / 2, innerWidth - w - 4)), y = r.bottom + 8; if (y + h > innerHeight - 4) y = Math.max(4, r.top - h - 8);
    tip.style.left = x + 'px'; tip.style.top = y + 'px';
  }
  document.addEventListener('pointerover', e => {
    if (e.pointerType !== 'mouse') return;
    const el = e.target.closest && e.target.closest('button,select,label.cb,.gi,[data-tip],[title],input[type=color],input[type=checkbox],.gl,.tabsrow div,[onclick]'); if (!el || el === cur) return;
    hide(); const t = labelOf(el); if (!t) return; cur = el;
    timer = setTimeout(() => { if (cur === el) show(el, t); }, 480);
  }, true);
  document.addEventListener('pointerout', e => { if (cur && !(e.relatedTarget && cur.contains(e.relatedTarget))) hide(); }, true);
  ['pointerdown', 'wheel', 'keydown', 'scroll'].forEach(ev => document.addEventListener(ev, hide, true));
})();
