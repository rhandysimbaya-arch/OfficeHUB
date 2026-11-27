// Excel-style formula engine: parser + evaluator + ~200 functions + dynamic arrays (spill)
(function (root) {
  'use strict';
  class XErr { constructor(c) { this.code = c; } toString() { return this.code; } }
  const E = c => new XErr(c);
  const isErr = v => v instanceof XErr;
  class Rng { constructor(sheet, r1, c1, r2, c2) { this.sheet = sheet; this.r1 = r1; this.c1 = c1; this.r2 = r2; this.c2 = c2; } }

  // ---------- addresses ----------
  function colToNum(s) { let n = 0; for (const ch of s.toUpperCase()) n = n * 26 + ch.charCodeAt(0) - 64; return n; }
  function numToCol(n) { let s = ''; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; }
  function parseAddr(a) { const m = /^\$?([A-Za-z]{1,3})\$?(\d+)$/.exec(a); return m ? { r: +m[2], c: colToNum(m[1]) } : null; }
  function addr(r, c) { return numToCol(c) + r; }

  // ---------- tokenizer ----------
  function tokenize(src) {
    const t = []; let i = 0; const n = src.length;
    while (i < n) {
      const ch = src[i];
      if (/\s/.test(ch)) { i++; continue; }
      if (ch === '"') { let s = ''; i++; while (i < n) { if (src[i] === '"') { if (src[i + 1] === '"') { s += '"'; i += 2; continue; } break; } s += src[i++]; } i++; t.push({ k: 'str', v: s }); continue; }
      if (ch === "'") { let s = ''; i++; while (i < n && src[i] !== "'") s += src[i++]; i++; if (src[i] === '!') { i++; t.push({ k: 'sheet', v: s }); continue; } throw E('#NAME?'); }
      if (/[0-9.]/.test(ch)) { const m = /^(\d+\.?\d*|\.\d+)([eE][+-]?\d+)?/.exec(src.slice(i)); if (!m) throw E('#VALUE!'); t.push({ k: 'num', v: parseFloat(m[0]) }); i += m[0].length; continue; }
      if (/[A-Za-z_$]/.test(ch)) {
        const m = /^[A-Za-z_$][A-Za-z0-9_.$]*/.exec(src.slice(i)); let w = m[0]; i += w.length;
        if (src[i] === '!') { i++; t.push({ k: 'sheet', v: w }); continue; }
        if (src[i] === '(') { t.push({ k: 'fn', v: w.toUpperCase() }); i++; t.push({ k: 'op', v: '(' }); continue; }
        const u = w.toUpperCase();
        if (u === 'TRUE' || u === 'FALSE') { t.push({ k: 'bool', v: u === 'TRUE' }); continue; }
        // range like A1:B2 / A:A / 1:1
        let rest = src.slice(i); const rm = /^:\$?([A-Za-z]{1,3}\$?\d*|\d+)/.exec(rest);
        if (rm && (parseAddr(w) || /^\$?[A-Za-z]{1,3}$/.test(w))) { t.push({ k: 'ref', v: w + rm[0] }); i += rm[0].length; continue; }
        if (parseAddr(w) || /^\$?[A-Za-z]{1,3}$/.test(w) && false) { t.push({ k: 'ref', v: w }); continue; }
        t.push({ k: 'name', v: w }); continue;
      }
      if (/\d/.test(ch)) { }
      const two = src.substr(i, 2);
      if (['<=', '>=', '<>'].includes(two)) { t.push({ k: 'op', v: two }); i += 2; continue; }
      if ('+-*/^&=<>(),%{};:'.includes(ch)) { t.push({ k: 'op', v: ch }); i++; continue; }
      if (ch === '#') { const m = /^#(N\/A|REF!|VALUE!|DIV\/0!|NAME\?|NUM!|NULL!|SPILL!)/i.exec(src.slice(i)); if (m) { t.push({ k: 'err', v: m[0].toUpperCase() }); i += m[0].length; continue; } }
      throw E('#NAME?');
    }
    return t;
  }
  // numeric row ranges like 1:1 need handling before number tokenizing: handle in parser via num ':' num
  // ---------- parser ----------
  function parse(src) {
    const toks = tokenize(src); let p = 0; let curSheet = null;
    const peek = () => toks[p], next = () => toks[p++];
    const isOp = v => toks[p] && toks[p].k === 'op' && toks[p].v === v;
    function parseCmp() { let l = parseConcat(); while (toks[p] && toks[p].k === 'op' && ['=', '<>', '<', '>', '<=', '>='].includes(toks[p].v)) { const o = next().v; l = { t: 'bin', o, l, r: parseConcat() }; } return l; }
    function parseConcat() { let l = parseAdd(); while (isOp('&')) { next(); l = { t: 'bin', o: '&', l, r: parseAdd() }; } return l; }
    function parseAdd() { let l = parseMul(); while (isOp('+') || isOp('-')) { const o = next().v; l = { t: 'bin', o, l, r: parseMul() }; } return l; }
    function parseMul() { let l = parsePow(); while (isOp('*') || isOp('/')) { const o = next().v; l = { t: 'bin', o, l, r: parsePow() }; } return l; }
    function parsePow() { let l = parseUnary(); while (isOp('^')) { next(); l = { t: 'bin', o: '^', l, r: parseUnary() }; } return l; }
    function parseUnary() { if (isOp('-')) { next(); return { t: 'un', o: '-', e: parseUnary() }; } if (isOp('+')) { next(); return parseUnary(); } return parsePostfix(); }
    function parsePostfix() { let e = parsePrimary(); while (isOp('%')) { next(); e = { t: 'pct', e }; } return e; }
    function refNode(txt, sheet) {
      let m = /^\$?([A-Za-z]{1,3})\$?(\d+)(?::\$?([A-Za-z]{1,3})\$?(\d+))?$/.exec(txt);
      if (m) { const c1 = colToNum(m[1]), r1 = +m[2]; if (m[3]) { const c2 = colToNum(m[3]), r2 = +m[4]; return { t: 'ref', sheet, r1: Math.min(r1, r2), c1: Math.min(c1, c2), r2: Math.max(r1, r2), c2: Math.max(c1, c2) }; } return { t: 'ref', sheet, r1, c1, r2: r1, c2: c1 }; }
      m = /^\$?([A-Za-z]{1,3}):\$?([A-Za-z]{1,3})$/.exec(txt);
      if (m) { const a = colToNum(m[1]), b = colToNum(m[2]); return { t: 'ref', sheet, r1: 1, c1: Math.min(a, b), r2: Infinity, c2: Math.max(a, b) }; }
      m = /^\$?([A-Za-z]{1,3})\$?(\d*):\$?([A-Za-z]{1,3})\$?(\d*)$/.exec(txt);
      if (m) { throw E('#REF!'); }
      throw E('#REF!');
    }
    function parsePrimary() {
      const tk = next(); if (!tk) throw E('#VALUE!');
      switch (tk.k) {
        case 'num': { if (isOp(':') && toks[p + 1] && toks[p + 1].k === 'num') { next(); const b = next().v; return { t: 'ref', sheet: curSheet, r1: Math.min(tk.v, b), c1: 1, r2: Math.max(tk.v, b), c2: 16384 }; } return { t: 'num', v: tk.v }; }
        case 'str': return { t: 'str', v: tk.v };
        case 'bool': return { t: 'bool', v: tk.v };
        case 'err': return { t: 'err', v: tk.v };
        case 'sheet': { const sh = tk.v; const nx = next(); if (!nx || (nx.k !== 'ref' && nx.k !== 'name')) throw E('#REF!'); curSheet = sh; const node = refNode(nx.v, sh); curSheet = null; return node; }
        case 'ref': return refNode(tk.v, null);
        case 'name': return { t: 'name', v: tk.v };
        case 'fn': { next(); const args = []; if (isOp(')')) { next(); return { t: 'fn', n: tk.v, a: args }; } while (true) { if (isOp(',') || isOp(')')) args.push({ t: 'empty' }); else args.push(parseCmp()); if (isOp(',')) { next(); continue; } if (isOp(')')) { next(); break; } throw E('#VALUE!'); } return { t: 'fn', n: tk.v, a: args }; }
        case 'op':
          if (tk.v === '(') { const e = parseCmp(); if (!isOp(')')) throw E('#VALUE!'); next(); return e; }
          if (tk.v === '{') { const rows = [[]]; while (true) { rows[rows.length - 1].push(parseCmp()); if (isOp(',')) { next(); continue; } if (isOp(';')) { next(); rows.push([]); continue; } if (isOp('}')) { next(); break; } throw E('#VALUE!'); } return { t: 'arr', rows }; }
        // fallthrough
        default: throw E('#VALUE!');
      }
    }
    const e = parseCmp(); if (p < toks.length) throw E('#VALUE!'); return e;
  }

  // ---------- coercion ----------
  function toNum(v) { if (isErr(v)) throw v; if (v === null || v === undefined || v === '') return 0; if (typeof v === 'number') return v; if (typeof v === 'boolean') return v ? 1 : 0; if (typeof v === 'string') { const s = v.trim().replace(/,/g, ''); if (s === '') return 0; if (/%$/.test(s) && !isNaN(+s.slice(0, -1))) return +s.slice(0, -1) / 100; const n = Number(s); if (!isNaN(n)) return n; const d = parseDateStr(v); if (d !== null) return d; throw E('#VALUE!'); } throw E('#VALUE!'); }
  function toStr(v) { if (isErr(v)) throw v; if (v === null || v === undefined) return ''; if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE'; if (typeof v === 'number') return numStr(v); return String(v); }
  function toBool(v) { if (isErr(v)) throw v; if (typeof v === 'boolean') return v; if (typeof v === 'number') return v !== 0; if (v === null || v === '') return false; if (typeof v === 'string') { const u = v.toUpperCase(); if (u === 'TRUE') return true; if (u === 'FALSE') return false; } throw E('#VALUE!'); }
  function numStr(n) { if (!isFinite(n)) return n > 0 ? 'Infinity' : '-Infinity'; if (Number.isInteger(n) && Math.abs(n) < 1e15) return String(n); let s = String(+n.toPrecision(15)); return s; }
  const EPOCH = Date.UTC(1899, 11, 30);
  function dateToSerial(d) { return (Date.UTC(d.getFullYear(), d.getMonth(), d.getDate()) - EPOCH) / 864e5 + (d.getHours() * 3600 + d.getMinutes() * 60 + d.getSeconds()) / 86400; }
  function serialToDate(s) { const ms = Math.round(s * 864e5); const d = new Date(EPOCH + ms); return { y: d.getUTCFullYear(), m: d.getUTCMonth() + 1, d: d.getUTCDate(), h: d.getUTCHours(), mi: d.getUTCMinutes(), s: d.getUTCSeconds(), wd: d.getUTCDay() }; }
  function mkSerial(y, m, d) { return (Date.UTC(y, m - 1, d) - EPOCH) / 864e5; }
  function parseDateStr(s) { s = String(s).trim(); let m = /^(\d{4})-(\d{1,2})-(\d{1,2})$/.exec(s); if (m) return mkSerial(+m[1], +m[2], +m[3]); m = /^(\d{1,2})\/(\d{1,2})\/(\d{2,4})$/.exec(s); if (m) { let y = +m[3]; if (y < 100) y += y < 30 ? 2000 : 1900; return mkSerial(y, +m[1], +m[2]); } m = /^(\d{1,2}):(\d{2})(?::(\d{2}))?$/.exec(s); if (m) return (+m[1] * 3600 + +m[2] * 60 + (+m[3] || 0)) / 86400; return null; }
  const isNum = v => typeof v === 'number';

  // compare for = < > (Excel: numbers < text < booleans; text case-insensitive)
  function cmpVals(a, b) {
    if (isErr(a)) throw a; if (isErr(b)) throw b;
    if (a === null || a === undefined) a = (typeof b === 'string') ? '' : (typeof b === 'boolean' ? false : 0);
    if (b === null || b === undefined) b = (typeof a === 'string') ? '' : (typeof a === 'boolean' ? false : 0);
    const ta = typeof a === 'number' ? 0 : typeof a === 'string' ? 1 : 2, tb = typeof b === 'number' ? 0 : typeof b === 'string' ? 1 : 2;
    if (ta !== tb) return ta < tb ? -1 : 1;
    if (ta === 1) { const x = a.toLowerCase(), y = b.toLowerCase(); return x < y ? -1 : x > y ? 1 : 0; }
    return a < b ? -1 : a > b ? 1 : 0;
  }

  // ---------- criteria ----------
  function makeCrit(c) {
    if (isErr(c)) throw c;
    let op = '=', val = c;
    if (typeof c === 'string') { const m = /^(<=|>=|<>|<|>|=)(.*)$/.exec(c); if (m) { op = m[1]; val = m[2]; } }
    let num = null; if (typeof val === 'number') num = val; else if (typeof val === 'string' && val.trim() !== '' && !isNaN(Number(val))) num = Number(val); else if (typeof val === 'boolean') num = null;
    const isWild = typeof val === 'string' && /[*?]/.test(val);
    let rx = null; if (isWild) rx = new RegExp('^' + val.replace(/[.+^${}()|[\]\\]/g, '\\$&').replace(/~\*/g, '\u0001').replace(/~\?/g, '\u0002').replace(/\*/g, '.*').replace(/\?/g, '.').replace(/\u0001/g, '\\*').replace(/\u0002/g, '\\?') + '$', 'i');
    return v => {
      if (isErr(v)) return false;
      if (num !== null && typeof val !== 'string' || (num !== null && op !== '=' && op !== '<>') || (num !== null)) {
        if (typeof v === 'number') { switch (op) { case '=': return v === num; case '<>': return v !== num; case '<': return v < num; case '>': return v > num; case '<=': return v <= num; case '>=': return v >= num; } }
        if (op === '<>') return true; if (op === '=') return typeof v === 'string' && v.trim() !== '' && Number(v) === num && false; return false;
      }
      if (val === '' && op === '=') return v === null || v === '' || v === undefined;
      if (val === '' && op === '<>') return !(v === null || v === '' || v === undefined);
      const s = (v === null || v === undefined) ? '' : typeof v === 'boolean' ? (v ? 'TRUE' : 'FALSE') : String(v);
      if (typeof v === 'number' && op !== '<>' && op !== '=') return false;
      if (rx) return op === '<>' ? !rx.test(s) : rx.test(s);
      const a = s.toLowerCase(), b = String(val).toLowerCase();
      switch (op) { case '=': return a === b; case '<>': return a !== b; case '<': return a < b; case '>': return a > b; case '<=': return a <= b; case '>=': return a >= b; }
      return false;
    };
  }

  // ---------- Workbook ----------
  class Workbook {
    constructor() { this.sheets = {}; this.order = []; this.addSheet('Sheet1'); this.spill = {}; this.spillPrev = {}; this.now = new Date(); }
    addSheet(name) { if (!this.sheets[name]) { this.sheets[name] = { name, cells: {}, vals: {}, fmt: {}, style: {}, colW: {}, rowH: {}, merges: [] }; this.order.push(name); } return this.sheets[name]; }
    removeSheet(name) { delete this.sheets[name]; this.order = this.order.filter(n => n !== name); }
    key(r, c) { return r + ',' + c; }
    // raw input: number | string | boolean | '=formula'
    setRaw(sheet, r, c, raw) {
      const sh = this.addSheet(sheet); const k = this.key(r, c);
      if (raw === null || raw === undefined || raw === '') { delete sh.cells[k]; return; }
      if (typeof raw === 'string' && raw[0] === '=' && raw.length > 1) sh.cells[k] = { f: raw };
      else if (typeof raw === 'string') { const p = this.parseInput(raw); sh.cells[k] = { v: p.v }; if (p.fmt && !sh.fmt[k]) sh.fmt[k] = p.fmt; }
      else sh.cells[k] = { v: raw };
    }
    parseInput(s) {
      const t = s.trim(); if (t === '') return { v: '' };
      if (/^[-+]?(\d+\.?\d*|\.\d+)(e[-+]?\d+)?$/i.test(t)) return { v: Number(t) };
      if (/^[-+]?\d{1,3}(,\d{3})+(\.\d+)?$/.test(t)) return { v: Number(t.replace(/,/g, '')), fmt: '#,##0.00' };
      if (/^[-+]?\d+(\.\d+)?%$/.test(t)) return { v: Number(t.slice(0, -1)) / 100, fmt: '0%' };
      if (/^(TRUE|FALSE)$/i.test(t)) return { v: t.toUpperCase() === 'TRUE' };
      const d = /^\d{4}-\d{1,2}-\d{1,2}$|^\d{1,2}\/\d{1,2}\/\d{2,4}$/.test(t) ? parseDateStr(t) : null; if (d !== null) return { v: d, fmt: 'yyyy-mm-dd' };
      return { v: s };
    }
    getRaw(sheet, r, c) { const sh = this.sheets[sheet]; if (!sh) return ''; const cell = sh.cells[this.key(r, c)]; if (!cell) return ''; if (cell.f) return cell.f; return cell.v === null || cell.v === undefined ? '' : cell.v; }
    usedRange(sheet) { const sh = this.sheets[sheet]; let mr = 0, mc = 0; if (!sh) return { rows: 0, cols: 0 }; for (const k in sh.cells) { const [r, c] = k.split(',').map(Number); if (r > mr) mr = r; if (c > mc) mc = c; } for (const k in this.spill) { if (k.startsWith(sheet + '|')) { const [r, c] = k.split('|')[1].split(',').map(Number); if (r > mr) mr = r; if (c > mc) mc = c; } } return { rows: mr, cols: mc }; }
    // computed value of a cell
    getVal(sheet, r, c) {
      const sh = this.sheets[sheet]; if (!sh) throw E('#REF!');
      const k = this.key(r, c); const cell = sh.cells[k];
      if (cell) {
        if (!cell.f) return cell.v === undefined ? null : cell.v;
        if (k in sh.vals) { const v = sh.vals[k]; return v; }
        if (this._stack.has(sheet + '|' + k)) return E('#REF!');
        this._stack.add(sheet + '|' + k);
        let v;
        try { v = this.evalFormula(sheet, r, c, cell); } catch (e) { v = isErr(e) ? e : E('#VALUE!'); if (!isErr(e)) { v = E('#VALUE!'); } }
        this._stack.delete(sheet + '|' + k);
        sh.vals[k] = v; return v;
      }
      const sp = this.spill[sheet + '|' + k] !== undefined ? this.spill[sheet + '|' + k] : this.spillPrev[sheet + '|' + k];
      if (sp !== undefined) return sp;
      return null;
    }
    evalFormula(sheet, r, c, cell) {
      if (!cell.ast) { try { cell.ast = parse((this.pre ? this.pre(sheet, r, c, cell.f) : cell.f).slice(1)); } catch (e) { cell.ast = { t: 'err', v: isErr(e) ? e.code : '#NAME?' }; } }
      const ctx = { sheet, r, c, wb: this, locals: {} };
      let v = evalNode(cell.ast, ctx);
      if (v instanceof Rng) { if (v.r1 === v.r2 && v.c1 === v.c2) return this.getVal(v.sheet || sheet, v.r1, v.c1); v = rngGrid(v, ctx); }
      if (Array.isArray(v)) {
        const rows = v.length, cols = v[0].length; if (rows === 1 && cols === 1) return v[0][0];
        // spill
        const sh = this.sheets[sheet]; let blocked = false;
        for (let i = 0; i < rows && !blocked; i++) for (let j = 0; j < cols; j++) { if (i === 0 && j === 0) continue; if (sh.cells[this.key(r + i, c + j)]) { blocked = true; break; } }
        if (blocked) return E('#SPILL!');
        for (let i = 0; i < rows; i++) for (let j = 0; j < cols; j++) { if (i === 0 && j === 0) continue; this.spill[sheet + '|' + this.key(r + i, c + j)] = v[i][j]; }
        return v[0][0];
      }
      if (v === undefined) return null;
      return v;
    }
    recalc() {
      this.now = new Date();
      for (let pass = 0; pass < 3; pass++) {
        const prev = this.spill; this.spillPrev = prev; this.spill = {};
        for (const n of this.order) { this.sheets[n].vals = {}; }
        this._stack = new Set();
        for (const n of this.order) { const sh = this.sheets[n]; for (const k in sh.cells) { if (sh.cells[k].f) { const [r, c] = k.split(',').map(Number); this.getVal(n, r, c); } } }
        const a = Object.keys(this.spill).sort().join('|') + JSON.stringify(Object.values(this.spill).map(x => isErr(x) ? x.code : x));
        const b = Object.keys(prev).sort().join('|') + JSON.stringify(Object.values(prev).map(x => isErr(x) ? x.code : x));
        if (Object.keys(this.spill).length === 0 && Object.keys(prev).length === 0) break;
        if (a === b || pass === 2) break;
      }
      this.spillPrev = {};
    }
    // display value (computed)
    value(sheet, r, c) { try { return this.getVal(sheet, r, c); } catch (e) { return isErr(e) ? e : E('#VALUE!'); } }
    toJSON() { const o = { order: this.order, sheets: {} }; for (const n of this.order) { const sh = this.sheets[n]; const cells = {}; for (const k in sh.cells) { const c = sh.cells[k]; cells[k] = c.f ? c.f : c.v; } o.sheets[n] = { cells, fmt: sh.fmt, style: sh.style, colW: sh.colW, rowH: sh.rowH, merges: sh.merges || [], cf: sh.cf || [], frz: sh.frz || null, hid: sh.hid || {}, flt: sh.flt || {}, defSz: sh.defSz, defFont: sh.defFont, grid: sh.grid, dv: sh.dv || [], ext: sh.ext || {} }; } return o; }
    static fromJSON(o) { const wb = new Workbook(); wb.sheets = {}; wb.order = []; for (const n of o.order) { const sh = wb.addSheet(n); const s = o.sheets[n]; for (const k in s.cells) { const v = s.cells[k]; sh.cells[k] = (typeof v === 'string' && v[0] === '=') ? { f: v } : { v }; } sh.fmt = s.fmt || {}; sh.style = s.style || {}; sh.colW = s.colW || {}; sh.rowH = s.rowH || {}; sh.merges = s.merges || []; sh.cf = s.cf || []; sh.frz = s.frz || null; sh.hid = s.hid || {}; sh.flt = s.flt || {}; sh.defSz = s.defSz; sh.defFont = s.defFont; sh.grid = s.grid; sh.dv = s.dv || []; sh.ext = s.ext || {}; } if (!wb.order.length) wb.addSheet('Sheet1'); return wb; }
  }

  // ---------- evaluation ----------
  function rngBounds(rg, ctx) {
    let r2 = rg.r2, c2 = rg.c2; const sheet = rg.sheet || ctx.sheet;
    if (r2 === Infinity || c2 === 16384) { const u = ctx.wb.usedRange(sheet); if (r2 === Infinity) r2 = Math.max(rg.r1, u.rows); if (c2 === 16384) c2 = Math.max(rg.c1, u.cols); }
    return { sheet, r1: rg.r1, c1: rg.c1, r2, c2 };
  }
  function rngGrid(rg, ctx) {
    const b = rngBounds(rg, ctx); const out = [];
    for (let r = b.r1; r <= b.r2; r++) { const row = []; for (let c = b.c1; c <= b.c2; c++) row.push(ctx.wb.getVal(b.sheet, r, c)); out.push(row); }
    return out;
  }
  function grid(x, ctx) { if (x instanceof Rng) return rngGrid(x, ctx); if (Array.isArray(x)) return x; return [[x]]; }
  function scalar(x, ctx) { if (x instanceof Rng) { const b = rngBounds(x, ctx); const v = ctx.wb.getVal(b.sheet, b.r1, b.c1); return v === undefined ? null : v; } if (Array.isArray(x)) return x[0][0]; return x; }
  function isMulti(x, ctx) { if (x instanceof Rng) { const b = rngBounds(x, ctx); return b.r1 !== b.r2 || b.c1 !== b.c2; } return Array.isArray(x) && (x.length > 1 || x[0].length > 1); }
  function flatNums(args, ctx, strict) { // numbers from args; range cells ignore text/bool
    const out = [];
    for (const a of args) {
      if (a instanceof Rng || Array.isArray(a)) { for (const row of grid(a, ctx)) for (const v of row) { if (isErr(v)) throw v; if (typeof v === 'number') out.push(v); } }
      else { if (isErr(a)) throw a; if (a === null || a === undefined) continue; out.push(toNum(a)); }
    }
    return out;
  }
  function broadcast(a, b, fn, ctx) {
    if (!isMulti(a, ctx) && !isMulti(b, ctx)) return fn(scalar(a, ctx), scalar(b, ctx));
    const A = grid(a, ctx), B = grid(b, ctx); const R = Math.max(A.length, B.length), C = Math.max(A[0].length, B[0].length); const out = [];
    for (let i = 0; i < R; i++) { const row = []; for (let j = 0; j < C; j++) { const x = A[A.length === 1 ? 0 : i] && A[A.length === 1 ? 0 : i][A[0].length === 1 ? 0 : j]; const y = B[B.length === 1 ? 0 : i] && B[B.length === 1 ? 0 : i][B[0].length === 1 ? 0 : j]; if (x === undefined || y === undefined) row.push(E('#N/A')); else { try { row.push(fn(x, y)); } catch (e) { row.push(isErr(e) ? e : E('#VALUE!')); } } } out.push(row); }
    return out;
  }
  function mapArr(a, fn, ctx) { if (!isMulti(a, ctx)) { return fn(scalar(a, ctx)); } return grid(a, ctx).map(row => row.map(v => { try { return fn(v); } catch (e) { return isErr(e) ? e : E('#VALUE!'); } })); }

  function evalNode(n, ctx) {
    switch (n.t) {
      case 'num': case 'str': case 'bool': return n.v;
      case 'err': return E(n.v);
      case 'empty': return null;
      case 'ref': return new Rng(n.sheet, n.r1, n.c1, n.r2, n.c2);
      case 'name': { if (n.v.toUpperCase() in ctx.locals) return ctx.locals[n.v.toUpperCase()]; const nm = ctx.wb.names && ctx.wb.names[n.v.toUpperCase()]; if (nm) return evalNode(parse(nm), ctx); throw E('#NAME?'); }
      case 'arr': return n.rows.map(r => r.map(e => scalar(evalNode(e, ctx), ctx)));
      case 'pct': return mapArr(evalNode(n.e, ctx), v => toNum(v) / 100, ctx);
      case 'un': return mapArr(evalNode(n.e, ctx), v => -toNum(v), ctx);
      case 'bin': {
        const a = evalNode(n.l, ctx), b = evalNode(n.r, ctx); const o = n.o;
        const f = { '+': (x, y) => toNum(x) + toNum(y), '-': (x, y) => toNum(x) - toNum(y), '*': (x, y) => toNum(x) * toNum(y), '/': (x, y) => { const d = toNum(y); if (d === 0) throw E('#DIV/0!'); return toNum(x) / d; }, '^': (x, y) => { const r = Math.pow(toNum(x), toNum(y)); if (isNaN(r)) throw E('#NUM!'); return r; }, '&': (x, y) => toStr(x) + toStr(y), '=': (x, y) => cmpVals(x, y) === 0, '<>': (x, y) => cmpVals(x, y) !== 0, '<': (x, y) => cmpVals(x, y) < 0, '>': (x, y) => cmpVals(x, y) > 0, '<=': (x, y) => cmpVals(x, y) <= 0, '>=': (x, y) => cmpVals(x, y) >= 0 }[o];
        return broadcast(a, b, f, ctx);
      }
      case 'fn': {
        const lazy = LAZY[n.n]; if (lazy) return lazy(n.a, ctx);
        const fn = FN[n.n]; if (!fn) { const nm = ctx.wb.names && ctx.wb.names[n.n]; if (nm) { const l = evalNode(parse(nm), ctx); if (l && l.lam) return callLam(l, n.a.map(x => evalNode(x, ctx))); } throw E('#NAME?'); }
        const args = n.a.map(a => evalNode(a, ctx));
        if (!RAW[n.n]) for (const a of args) { if (isErr(a)) throw a; }
        const r = fn(args, ctx); return r === undefined ? null : r;
      }
    }
    throw E('#VALUE!');
  }

  // ---------- number formatting ----------
  const MONTHS = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
  const DAYS = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];
  function fmtNumber(v, fmt) {
    if (isErr(v)) return v.code; if (v === null || v === undefined) return ''; if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE'; if (typeof v === 'string') return v;
    if (!fmt || fmt === 'General') return numStr(Math.abs(v) >= 1e11 || (Math.abs(v) < 1e-9 && v !== 0) ? +v.toPrecision(10) : +v.toPrecision(12) + 0);
    let cf = null; const CRE = /\[(<=|>=|<>|<|>|=)\s*(-?[\d.]+)\]/;
    if (CRE.test(fmt)) { const secs = fmt.split(';'); const cmpc = (o, b) => o === '<' ? v < b : o === '>' ? v > b : o === '<=' ? v <= b : o === '>=' ? v >= b : o === '=' ? v === b : v !== b; for (let i = 0; i < secs.length; i++) { const m = CRE.exec(secs[i]); if (!m || cmpc(m[1], +m[2])) { cf = secs[i].replace(CRE, ''); break; } } if (cf === null) cf = ''; }
    const parts = cf !== null ? [cf] : fmt.split(';'); let f = parts[0]; let val = v;
    if (parts.length > 1 && v < 0) { f = parts[1]; val = -v; } else if (parts.length > 2 && v === 0) { f = parts[2]; }
    { const sc = /([0#?])(,+)(?![0#?,])/.exec(f.replace(/"[^"]*"/g, m => 'Q'.repeat(m.length))); if (sc) { val = val / Math.pow(1000, sc[2].length); f = f.slice(0, sc.index + 1) + f.slice(sc.index + 1 + sc[2].length); } }
    f = f.replace(/\[\$[^\]]*\]/g, '').replace(/\[(Red|Blue|Green|Black|White|Yellow|Magenta|Cyan|Color\d+)\]/gi, '');
    if (f === '@') return String(v);
    if (/0E[+-]0+/i.test(f)) { const dec = ((/\.([0#]+)E/i.exec(f) || [, ''])[1]).length; if (val === 0) return (0).toFixed(dec) + 'E+00'; let ex = Math.floor(Math.log10(Math.abs(val))); let mt = val / Math.pow(10, ex); if (+Math.abs(mt).toFixed(dec) >= 10) { mt /= 10; ex++; } return (v < 0 && parts.length < 2 ? '-' : '') + Math.abs(mt).toFixed(dec) + 'E' + (ex < 0 ? '-' : '+') + String(Math.abs(ex)).padStart(2, '0'); }
    const fr = /(\?+)\/(\?+|\d+)/.exec(f);
    if (fr) { const ad = Math.abs(val); let wh = Math.floor(ad), rem = ad - wh; const maxD = /^\d+$/.test(fr[2]) ? +fr[2] : Math.pow(10, fr[2].length) - 1; let bn = 0, bd = 1, be = rem; for (let d = 1; d <= maxD; d++) { const n = Math.round(rem * d); const e = Math.abs(rem - n / d); if (e < be - 1e-12) { be = e; bn = n; bd = d; } } if (bn === bd) { wh++; bn = 0; } const sg = v < 0 && parts.length < 2 ? '-' : ''; return sg + (bn === 0 ? String(wh) : (wh ? wh + ' ' : '') + bn + '/' + bd); }
    const clean = f.replace(/"[^"]*"/g, '').replace(/\\./g, '').replace(/\[[^\]]*\]/g, '');
    if (/[ymdhs]/i.test(clean) && !/^[#0,.\s%$-]*$/.test(clean) && !/[#0]/.test(clean.replace(/[ymdhs]/gi, ''))) return fmtDate(val, f);
    // numeric
    let pct = /%/.test(clean); if (pct) val = val * 100;
    const m = /([#0?,]*[#0?])(\.[0#?]+)?/.exec(f); if (!m) return f.replace(/"/g, '');
    const intPart = m[1], decPart = m[2] ? m[2].slice(1) : ''; const thousands = /,/.test(intPart);
    const minDec = (decPart.match(/0/g) || []).length, maxDec = decPart.length;
    let s = Math.abs(val).toFixed(maxDec); if (maxDec > minDec) { s = s.replace(/0+$/, m => m.slice(0, Math.max(0, m.length - (maxDec - minDec)))); let [a, b = ''] = s.split('.'); b = b.replace(/0+$/, ''); while (b.length < minDec) b += '0'; s = b ? a + '.' + b : a; }
    let [ip, dp = ''] = s.split('.'); const minInt = (intPart.replace(/,/g, '').match(/0/g) || []).length; while (ip.length < minInt) ip = '0' + ip; if (minInt === 0 && ip === '0' && dp) ip = '';
    if (thousands) ip = ip.replace(/\B(?=(\d{3})+(?!\d))/g, ',');
    let out = ip + (dp ? '.' + dp : '');
    const pre = f.slice(0, m.index).replace(/"/g, '').replace(/\\(.)/g, '$1').replace(/_.|\*./g, ''), post = f.slice(m.index + m[0].length).replace(/"/g, '').replace(/\\(.)/g, '$1').replace(/_.|\*./g, '');
    const neg = val < 0 || (v < 0 && parts.length < 2 && Math.round(Math.abs(val) * Math.pow(10, maxDec)) !== 0);
    return (v < 0 && parts.length < 2 ? '-' : '') + pre + out + post;
  }
  function fmtDate(serial, f) {
    const d = serialToDate(serial); const ampm = /AM\/PM|A\/P/i.test(f);
    let h12 = d.h % 12 || 12;
    const toks = []; let i = 0; const s = f;
    while (i < s.length) {
      const rest = s.slice(i);
      let m;
      if ((m = /^"([^"]*)"/.exec(rest))) { toks.push({ lit: m[1] }); i += m[0].length; continue; }
      if ((m = /^(AM\/PM|A\/P)/i.exec(rest))) { toks.push({ t: 'ap', v: m[0] }); i += m[0].length; continue; }
      if ((m = /^(yyyy|yy|mmmmm|mmmm|mmm|mm|m|dddd|ddd|dd|d|hh|h|ss|s)/i.exec(rest))) { toks.push({ t: m[0].toLowerCase(), raw: m[0] }); i += m[0].length; continue; }
      toks.push({ lit: s[i] === '\\' ? s[++i] : s[i] }); i++;
    }
    // minute vs month: m after h or before s => minutes
    toks.forEach((t, idx) => { if (t.t === 'm' || t.t === 'mm') { let prev = null, nxt = null; for (let j = idx - 1; j >= 0; j--) if (toks[j].t) { prev = toks[j].t; break; } for (let j = idx + 1; j < toks.length; j++) if (toks[j].t) { nxt = toks[j].t; break; } if (prev === 'h' || prev === 'hh' || nxt === 's' || nxt === 'ss') t.min = true; } });
    const p2 = n => String(n).padStart(2, '0');
    return toks.map(t => {
      if (t.lit !== undefined) return t.lit;
      switch (t.t) {
        case 'yyyy': return String(d.y); case 'yy': return p2(d.y % 100);
        case 'mmmmm': return MONTHS[d.m - 1][0]; case 'mmmm': return MONTHS[d.m - 1]; case 'mmm': return MONTHS[d.m - 1].slice(0, 3);
        case 'mm': return t.min ? p2(d.mi) : p2(d.m); case 'm': return t.min ? String(d.mi) : String(d.m);
        case 'dddd': return DAYS[d.wd]; case 'ddd': return DAYS[d.wd].slice(0, 3); case 'dd': return p2(d.d); case 'd': return String(d.d);
        case 'hh': return p2(ampm ? h12 : d.h); case 'h': return String(ampm ? h12 : d.h);
        case 'ss': return p2(d.s); case 's': return String(d.s);
        case 'ap': return d.h < 12 ? (t.v.length > 3 ? 'AM' : 'A') : (t.v.length > 3 ? 'PM' : 'P');
      }
      return '';
    }).join('');
  }

  // ---------- functions ----------
  const FN = {}, LAZY = {}, RAW = {};
  const def = (names, fn, raw) => { for (const n of names.split(' ')) { FN[n] = fn; if (raw) RAW[n] = true; } };
  const num = (a, ctx) => toNum(scalar(a, ctx));
  const argv = (args, i, d) => (i < args.length && args[i] !== null && args[i] !== undefined) ? args[i] : d;
  const vecMap = (a, ctx, fn) => mapArr(a, fn, ctx);

  // math
  def('HYPERLINK', (a, c) => (a.length > 1 ? scalar(a[1], c) : scalar(a[0], c)), true);
  def('SUM', (a, c) => flatNums(a, c).reduce((x, y) => x + y, 0));
  def('AVERAGE', (a, c) => { const v = flatNums(a, c); if (!v.length) throw E('#DIV/0!'); return v.reduce((x, y) => x + y, 0) / v.length; });
  def('MIN', (a, c) => { const v = flatNums(a, c); return v.length ? Math.min(...v) : 0; });
  def('MAX', (a, c) => { const v = flatNums(a, c); return v.length ? Math.max(...v) : 0; });
  def('COUNT', (a, c) => { let n = 0; for (const x of a) { if (x instanceof Rng || Array.isArray(x)) { for (const r of grid(x, c)) for (const v of r) if (typeof v === 'number') n++; } else if (typeof x === 'number' || (typeof x === 'string' && x !== '' && !isNaN(Number(x)))) n++; } return n; }, true);
  def('COUNTA', (a, c) => { let n = 0; for (const x of a) { if (x instanceof Rng || Array.isArray(x)) { for (const r of grid(x, c)) for (const v of r) if (v !== null && v !== '' && v !== undefined) n++; } else if (x !== null) n++; } return n; }, true);
  def('COUNTBLANK', (a, c) => { let n = 0; for (const r of grid(a[0], c)) for (const v of r) if (v === null || v === '') n++; return n; });
  def('PRODUCT', (a, c) => flatNums(a, c).reduce((x, y) => x * y, 1));
  def('SUMSQ', (a, c) => flatNums(a, c).reduce((x, y) => x + y * y, 0));
  def('ABS', (a, c) => vecMap(a[0], c, v => Math.abs(toNum(v))));
  def('SIGN', (a, c) => vecMap(a[0], c, v => Math.sign(toNum(v))));
  def('SQRT', (a, c) => vecMap(a[0], c, v => { const x = toNum(v); if (x < 0) throw E('#NUM!'); return Math.sqrt(x); }));
  def('EXP', (a, c) => vecMap(a[0], c, v => Math.exp(toNum(v))));
  def('LN', (a, c) => vecMap(a[0], c, v => { const x = toNum(v); if (x <= 0) throw E('#NUM!'); return Math.log(x); }));
  def('LOG10', (a, c) => vecMap(a[0], c, v => { const x = toNum(v); if (x <= 0) throw E('#NUM!'); return Math.log10(x); }));
  def('LOG', (a, c) => { const x = num(a[0], c), b = a.length > 1 ? num(a[1], c) : 10; if (x <= 0 || b <= 0) throw E('#NUM!'); return Math.log(x) / Math.log(b); });
  def('PI', () => Math.PI);
  def('INT', (a, c) => vecMap(a[0], c, v => Math.floor(toNum(v))));
  def('TRUNC', (a, c) => { const d = a.length > 1 ? num(a[1], c) : 0, f = Math.pow(10, d); return vecMap(a[0], c, v => Math.trunc(toNum(v) * f) / f); });
  const rnd = (x, d, mode) => { const f = Math.pow(10, d); const y = x * f; let r; if (mode === 'up') r = Math.sign(y) * Math.ceil(Math.abs(y) - 1e-12); else if (mode === 'down') r = Math.sign(y) * Math.floor(Math.abs(y) + 1e-12); else r = Math.sign(y) * Math.round(Math.abs(y) + 1e-12); return r / f; };
  def('ROUND', (a, c) => { const d = a.length > 1 ? num(a[1], c) : 0; return vecMap(a[0], c, v => rnd(toNum(v), d)); });
  def('ROUNDUP', (a, c) => { const d = a.length > 1 ? num(a[1], c) : 0; return vecMap(a[0], c, v => rnd(toNum(v), d, 'up')); });
  def('ROUNDDOWN', (a, c) => { const d = a.length > 1 ? num(a[1], c) : 0; return vecMap(a[0], c, v => rnd(toNum(v), d, 'down')); });
  def('CEILING', (a, c) => { const s = a.length > 1 ? num(a[1], c) : 1; return vecMap(a[0], c, v => s === 0 ? 0 : Math.ceil(toNum(v) / s) * s); });
  def('FLOOR', (a, c) => { const s = a.length > 1 ? num(a[1], c) : 1; return vecMap(a[0], c, v => s === 0 ? 0 : Math.floor(toNum(v) / s) * s); });
  def('MROUND', (a, c) => { const m = num(a[1], c); return m === 0 ? 0 : Math.round(num(a[0], c) / m) * m; });
  def('EVEN', (a, c) => { const x = num(a[0], c); const r = Math.ceil(Math.abs(x) / 2) * 2; return x < 0 ? -r : r; });
  def('ODD', (a, c) => { const x = num(a[0], c); let r = Math.ceil(Math.abs(x)); if (r % 2 === 0) r++; return x < 0 ? -r : r; });
  def('MOD', (a, c) => { const d = num(a[1], c); if (d === 0) throw E('#DIV/0!'); const n = num(a[0], c); return n - d * Math.floor(n / d); });
  def('QUOTIENT', (a, c) => { const d = num(a[1], c); if (d === 0) throw E('#DIV/0!'); return Math.trunc(num(a[0], c) / d); });
  def('POWER', (a, c) => Math.pow(num(a[0], c), num(a[1], c)));
  def('FACT', (a, c) => { const n = Math.floor(num(a[0], c)); if (n < 0) throw E('#NUM!'); let r = 1; for (let i = 2; i <= n; i++) r *= i; return r; });
  const gcd = (x, y) => { x = Math.abs(x); y = Math.abs(y); while (y) { [x, y] = [y, x % y]; } return x; };
  def('GCD', (a, c) => flatNums(a, c).reduce(gcd, 0));
  def('LCM', (a, c) => flatNums(a, c).reduce((x, y) => x && y ? Math.abs(x * y) / gcd(x, y) : 0, 1));
  def('SIN', (a, c) => vecMap(a[0], c, v => Math.sin(toNum(v)))); def('COS', (a, c) => vecMap(a[0], c, v => Math.cos(toNum(v)))); def('TAN', (a, c) => vecMap(a[0], c, v => Math.tan(toNum(v))));
  def('ASIN', (a, c) => Math.asin(num(a[0], c))); def('ACOS', (a, c) => Math.acos(num(a[0], c))); def('ATAN', (a, c) => Math.atan(num(a[0], c))); def('ATAN2', (a, c) => Math.atan2(num(a[1], c), num(a[0], c)));
  def('DEGREES', (a, c) => num(a[0], c) * 180 / Math.PI); def('RADIANS', (a, c) => num(a[0], c) * Math.PI / 180);
  def('RAND', () => Math.random());
  def('RANDBETWEEN', (a, c) => { const lo = Math.ceil(num(a[0], c)), hi = Math.floor(num(a[1], c)); return lo + Math.floor(Math.random() * (hi - lo + 1)); });
  def('SUMPRODUCT', (a, c) => { const gs = a.map(x => grid(x, c)); const R = gs[0].length, C = gs[0][0].length; let s = 0; for (let i = 0; i < R; i++) for (let j = 0; j < C; j++) { let p = 1; for (const g of gs) { if (!g[i] || g[i][j] === undefined) throw E('#VALUE!'); const v = g[i][j]; if (isErr(v)) throw v; p *= typeof v === 'number' ? v : (typeof v === 'boolean' ? (v ? 1 : 0) : 0); } s += p; } return s; });
  // stats
  const sorted = (a, c) => flatNums(a, c).sort((x, y) => x - y);
  def('MEDIAN', (a, c) => { const v = sorted(a, c); if (!v.length) throw E('#NUM!'); const m = v.length >> 1; return v.length % 2 ? v[m] : (v[m - 1] + v[m]) / 2; });
  def('MODE MODE.SNGL', (a, c) => { const v = flatNums(a, c); const cnt = {}; let best = null, bc = 1; for (const x of v) { cnt[x] = (cnt[x] || 0) + 1; if (cnt[x] > bc) { bc = cnt[x]; best = x; } } if (best === null) throw E('#N/A'); return best; });
  def('LARGE', (a, c) => { const v = sorted([a[0]], c); const k = num(a[1], c); if (k < 1 || k > v.length) throw E('#NUM!'); return v[v.length - k]; });
  def('SMALL', (a, c) => { const v = sorted([a[0]], c); const k = num(a[1], c); if (k < 1 || k > v.length) throw E('#NUM!'); return v[k - 1]; });
  def('RANK RANK.EQ', (a, c) => { const x = num(a[0], c); const v = flatNums([a[1]], c); const asc = a.length > 2 && num(a[2], c) !== 0; if (!v.includes(x)) throw E('#N/A'); return 1 + v.filter(y => asc ? y < x : y > x).length; });
  const mean = v => v.reduce((x, y) => x + y, 0) / v.length;
  const varS = v => { if (v.length < 2) throw E('#DIV/0!'); const m = mean(v); return v.reduce((s, x) => s + (x - m) ** 2, 0) / (v.length - 1); };
  const varP = v => { if (!v.length) throw E('#DIV/0!'); const m = mean(v); return v.reduce((s, x) => s + (x - m) ** 2, 0) / v.length; };
  def('VAR VAR.S', (a, c) => varS(flatNums(a, c))); def('VAR.P VARP', (a, c) => varP(flatNums(a, c)));
  def('STDEV STDEV.S', (a, c) => Math.sqrt(varS(flatNums(a, c)))); def('STDEV.P STDEVP', (a, c) => Math.sqrt(varP(flatNums(a, c))));
  const pctile = (v, k) => { v.sort((x, y) => x - y); if (!v.length || k < 0 || k > 1) throw E('#NUM!'); const pos = k * (v.length - 1), lo = Math.floor(pos), hi = Math.ceil(pos); return v[lo] + (v[hi] - v[lo]) * (pos - lo); };
  def('PERCENTILE PERCENTILE.INC', (a, c) => pctile(flatNums([a[0]], c), num(a[1], c)));
  def('QUARTILE QUARTILE.INC', (a, c) => pctile(flatNums([a[0]], c), num(a[1], c) / 4));
  const pairs = (a, c) => { const x = flatNums([a[0]], c), y = flatNums([a[1]], c); if (x.length !== y.length || !x.length) throw E('#N/A'); return [x, y]; };
  def('CORREL', (a, c) => { const [y, x] = pairs(a, c); const mx = mean(x), my = mean(y); let sxy = 0, sxx = 0, syy = 0; for (let i = 0; i < x.length; i++) { sxy += (x[i] - mx) * (y[i] - my); sxx += (x[i] - mx) ** 2; syy += (y[i] - my) ** 2; } return sxy / Math.sqrt(sxx * syy); });
  def('SLOPE', (a, c) => { const [y, x] = pairs(a, c); const mx = mean(x), my = mean(y); let sxy = 0, sxx = 0; for (let i = 0; i < x.length; i++) { sxy += (x[i] - mx) * (y[i] - my); sxx += (x[i] - mx) ** 2; } return sxy / sxx; });
  def('INTERCEPT', (a, c) => { const [y, x] = pairs(a, c); const mx = mean(x), my = mean(y); let sxy = 0, sxx = 0; for (let i = 0; i < x.length; i++) { sxy += (x[i] - mx) * (y[i] - my); sxx += (x[i] - mx) ** 2; } return my - sxy / sxx * mx; });
  def('FORECAST FORECAST.LINEAR', (a, c) => { const x0 = num(a[0], c); const [y, x] = pairs([a[1], a[2]], c); const mx = mean(x), my = mean(y); let sxy = 0, sxx = 0; for (let i = 0; i < x.length; i++) { sxy += (x[i] - mx) * (y[i] - my); sxx += (x[i] - mx) ** 2; } const s = sxy / sxx; return my - s * mx + s * x0; });
  def('GEOMEAN', (a, c) => { const v = flatNums(a, c); return Math.exp(mean(v.map(Math.log))); });
  def('AVEDEV', (a, c) => { const v = flatNums(a, c); const m = mean(v); return mean(v.map(x => Math.abs(x - m))); });
  // conditional aggregates
  function condIdx(pairsArr, ctx) { // pairsArr: [range, criteria, ...]
    let idx = null; let shape = null;
    for (let i = 0; i < pairsArr.length; i += 2) {
      const g = grid(pairsArr[i], ctx); const crit = makeCrit(scalar(pairsArr[i + 1], ctx));
      if (!shape) { shape = [g.length, g[0].length]; idx = g.map(r => r.map(() => true)); } else if (g.length !== shape[0] || g[0].length !== shape[1]) throw E('#VALUE!');
      for (let r = 0; r < g.length; r++) for (let cc = 0; cc < g[0].length; cc++) if (idx[r][cc] && !crit(g[r][cc])) idx[r][cc] = false;
    }
    return idx;
  }
  const pick = (g, idx) => { const o = []; for (let r = 0; r < idx.length; r++) for (let c = 0; c < idx[0].length; c++) if (idx[r][c] && g[r] && g[r][c] !== undefined) o.push(g[r][c]); return o; };
  def('COUNTIF', (a, c) => { const idx = condIdx([a[0], a[1]], c); return idx.reduce((s, r) => s + r.filter(Boolean).length, 0); });
  def('COUNTIFS', (a, c) => { const idx = condIdx(a, c); return idx.reduce((s, r) => s + r.filter(Boolean).length, 0); });
  def('SUMIF', (a, c) => { const idx = condIdx([a[0], a[1]], c); const sg = a.length > 2 ? grid(a[2], c) : grid(a[0], c); return pick(sg, idx).filter(isNum).reduce((x, y) => x + y, 0); });
  def('SUMIFS', (a, c) => { const idx = condIdx(a.slice(1), c); return pick(grid(a[0], c), idx).filter(isNum).reduce((x, y) => x + y, 0); });
  def('AVERAGEIF', (a, c) => { const idx = condIdx([a[0], a[1]], c); const v = pick(a.length > 2 ? grid(a[2], c) : grid(a[0], c), idx).filter(isNum); if (!v.length) throw E('#DIV/0!'); return mean(v); });
  def('AVERAGEIFS', (a, c) => { const idx = condIdx(a.slice(1), c); const v = pick(grid(a[0], c), idx).filter(isNum); if (!v.length) throw E('#DIV/0!'); return mean(v); });
  def('MAXIFS', (a, c) => { const idx = condIdx(a.slice(1), c); const v = pick(grid(a[0], c), idx).filter(isNum); return v.length ? Math.max(...v) : 0; });
  def('MINIFS', (a, c) => { const idx = condIdx(a.slice(1), c); const v = pick(grid(a[0], c), idx).filter(isNum); return v.length ? Math.min(...v) : 0; });
  // logical
  LAZY.IF = (a, c) => { const cond = evalNode(a[0], c); const f = v => { if (toBool(v)) return a.length > 1 ? evalNode(a[1], c) : true; return a.length > 2 ? evalNode(a[2], c) : false; }; if (isMulti(cond, c)) { const g = grid(cond, c); const T = a.length > 1 ? grid(evalNode(a[1], c), c) : [[true]], F = a.length > 2 ? grid(evalNode(a[2], c), c) : [[false]]; return g.map((row, i) => row.map((v, j) => { try { const s = toBool(v) ? T : F; return s[s.length === 1 ? 0 : i][s[0].length === 1 ? 0 : j]; } catch (e) { return isErr(e) ? e : E('#VALUE!'); } })); } const sc = scalar(cond, c); if (isErr(sc)) throw sc; return f(sc); };
  LAZY.IFS = (a, c) => { for (let i = 0; i + 1 < a.length; i += 2) { if (toBool(scalar(evalNode(a[i], c), c))) return evalNode(a[i + 1], c); } throw E('#N/A'); };
  LAZY.IFERROR = (a, c) => { try { const v = evalNode(a[0], c); const s = v instanceof Rng || Array.isArray(v) ? v : v; if (!isMulti(s, c)) { const x = scalar(s, c); if (isErr(x)) return evalNode(a[1], c); return x; } const alt = evalNode(a[1], c); return grid(s, c).map(r => r.map(x => isErr(x) ? scalar(alt, c) : x)); } catch (e) { if (isErr(e)) return evalNode(a[1], c); throw e; } };
  LAZY.IFNA = (a, c) => { try { const v = scalar(evalNode(a[0], c), c); if (isErr(v) && v.code === '#N/A') return evalNode(a[1], c); return v; } catch (e) { if (isErr(e) && e.code === '#N/A') return evalNode(a[1], c); throw e; } };
  LAZY.SWITCH = (a, c) => { const v = scalar(evalNode(a[0], c), c); let i = 1; for (; i + 1 < a.length; i += 2) { if (cmpVals(v, scalar(evalNode(a[i], c), c)) === 0) return evalNode(a[i + 1], c); } if (i < a.length) return evalNode(a[i], c); throw E('#N/A'); };
  LAZY.CHOOSE = (a, c) => { const i = Math.floor(toNum(scalar(evalNode(a[0], c), c))); if (i < 1 || i >= a.length) throw E('#VALUE!'); return evalNode(a[i], c); };
  LAZY.LET = (a, c) => { const locals = Object.assign({}, c.locals); const c2 = Object.assign({}, c, { locals }); for (let i = 0; i + 1 < a.length; i += 2) { if (a[i].t !== 'name') throw E('#NAME?'); let v = evalNode(a[i + 1], c2); if (v instanceof Rng) v = v; locals[a[i].v.toUpperCase()] = v; } return evalNode(a[a.length - 1], c2); };
  def('AND', (a, c) => { let r = true, any = false; for (const x of a) { for (const row of grid(x, c)) for (const v of row) { if (isErr(v)) throw v; if (typeof v === 'string' && (x instanceof Rng)) continue; if (v === null) continue; any = true; if (!toBool(v)) r = false; } } if (!any) throw E('#VALUE!'); return r; });
  def('OR', (a, c) => { let r = false, any = false; for (const x of a) { for (const row of grid(x, c)) for (const v of row) { if (isErr(v)) throw v; if (typeof v === 'string' && (x instanceof Rng)) continue; if (v === null) continue; any = true; if (toBool(v)) r = true; } } if (!any) throw E('#VALUE!'); return r; });
  def('XOR', (a, c) => { let n = 0; for (const x of a) for (const row of grid(x, c)) for (const v of row) if (v !== null && toBool(v)) n++; return n % 2 === 1; });
  def('NOT', (a, c) => vecMap(a[0], c, v => !toBool(v)));
  def('TRUE', () => true); def('FALSE', () => false); def('NA', () => { throw E('#N/A'); });
  def('ISBLANK', (a, c) => vecMap(a[0], c, v => v === null || v === undefined), true);
  def('ISNUMBER', (a, c) => vecMap(a[0], c, v => typeof v === 'number'), true);
  def('ISTEXT', (a, c) => vecMap(a[0], c, v => typeof v === 'string'), true);
  def('ISNONTEXT', (a, c) => vecMap(a[0], c, v => typeof v !== 'string'), true);
  def('ISLOGICAL', (a, c) => vecMap(a[0], c, v => typeof v === 'boolean'), true);
  def('ISERROR', (a, c) => vecMap(a[0], c, v => isErr(v)), true);
  def('ISERR', (a, c) => vecMap(a[0], c, v => isErr(v) && v.code !== '#N/A'), true);
  def('ISNA', (a, c) => vecMap(a[0], c, v => isErr(v) && v.code === '#N/A'), true);
  def('ISEVEN', (a, c) => vecMap(a[0], c, v => Math.floor(Math.abs(toNum(v))) % 2 === 0));
  def('ISODD', (a, c) => vecMap(a[0], c, v => Math.floor(Math.abs(toNum(v))) % 2 === 1));
  def('N', (a, c) => { const v = scalar(a[0], c); return typeof v === 'number' ? v : typeof v === 'boolean' ? (v ? 1 : 0) : 0; });
  def('T', (a, c) => { const v = scalar(a[0], c); return typeof v === 'string' ? v : ''; });
  def('TYPE', (a, c) => { const v = scalar(a[0], c); return isErr(v) ? 16 : typeof v === 'number' ? 1 : typeof v === 'string' ? 2 : typeof v === 'boolean' ? 4 : 1; }, true);
  // text
  const sm = (a, c, fn) => vecMap(a, c, v => fn(toStr(v)));
  def('LEN', (a, c) => sm(a[0], c, s => s.length));
  def('UPPER', (a, c) => sm(a[0], c, s => s.toUpperCase())); def('LOWER', (a, c) => sm(a[0], c, s => s.toLowerCase()));
  def('PROPER', (a, c) => sm(a[0], c, s => s.toLowerCase().replace(/(^|[^a-z0-9'])([a-z])/g, (m, p, l) => p + l.toUpperCase())));
  def('TRIM', (a, c) => sm(a[0], c, s => s.trim().replace(/ +/g, ' ')));
  def('CLEAN', (a, c) => sm(a[0], c, s => s.replace(/[\x00-\x1f]/g, '')));
  def('LEFT', (a, c) => { const n = a.length > 1 ? num(a[1], c) : 1; if (n < 0) throw E('#VALUE!'); return sm(a[0], c, s => s.slice(0, n)); });
  def('RIGHT', (a, c) => { const n = a.length > 1 ? num(a[1], c) : 1; if (n < 0) throw E('#VALUE!'); return sm(a[0], c, s => n === 0 ? '' : s.slice(-n)); });
  def('MID', (a, c) => { const st = num(a[1], c), n = num(a[2], c); if (st < 1 || n < 0) throw E('#VALUE!'); return sm(a[0], c, s => s.substr(st - 1, n)); });
  def('CONCATENATE', (a, c) => a.map(x => toStr(scalar(x, c))).join(''));
  def('CONCAT', (a, c) => a.map(x => grid(x, c).map(r => r.map(toStr).join('')).join('')).join(''));
  def('TEXTJOIN', (a, c) => { const d = toStr(scalar(a[0], c)), ig = toBool(scalar(a[1], c)); const o = []; for (const x of a.slice(2)) for (const r of grid(x, c)) for (const v of r) { if (isErr(v)) throw v; const s = toStr(v); if (ig && s === '') continue; o.push(s); } return o.join(d); });
  def('REPT', (a, c) => toStr(scalar(a[0], c)).repeat(Math.max(0, num(a[1], c))));
  def('EXACT', (a, c) => toStr(scalar(a[0], c)) === toStr(scalar(a[1], c)));
  def('CHAR', (a, c) => String.fromCharCode(num(a[0], c))); def('CODE', (a, c) => toStr(scalar(a[0], c)).charCodeAt(0));
  def('UNICHAR', (a, c) => String.fromCodePoint(num(a[0], c))); def('UNICODE', (a, c) => toStr(scalar(a[0], c)).codePointAt(0));
  def('SUBSTITUTE', (a, c) => { const old = toStr(scalar(a[1], c)), nw = toStr(scalar(a[2], c)); const inst = a.length > 3 ? num(a[3], c) : 0; return sm(a[0], c, s => { if (old === '') return s; if (!inst) return s.split(old).join(nw); let n = 0, out = '', i = 0; while (i < s.length) { if (s.startsWith(old, i)) { n++; if (n === inst) { out += nw; i += old.length; continue; } out += old; i += old.length; continue; } out += s[i++]; } return out; }); });
  def('REPLACE', (a, c) => { const s = toStr(scalar(a[0], c)), st = num(a[1], c), n = num(a[2], c), nt = toStr(scalar(a[3], c)); return s.slice(0, st - 1) + nt + s.slice(st - 1 + n); });
  def('FIND', (a, c) => { const f = toStr(scalar(a[0], c)), s = toStr(scalar(a[1], c)), st = a.length > 2 ? num(a[2], c) : 1; const i = s.indexOf(f, st - 1); if (i < 0) throw E('#VALUE!'); return i + 1; });
  def('SEARCH', (a, c) => { const f = toStr(scalar(a[0], c)).toLowerCase(), s = toStr(scalar(a[1], c)).toLowerCase(), st = a.length > 2 ? num(a[2], c) : 1; const rx = new RegExp(f.replace(/[.+^${}()|[\]\\]/g, '\\$&').replace(/\*/g, '.*').replace(/\?/g, '.')); const m = rx.exec(s.slice(st - 1)); if (!m) throw E('#VALUE!'); return m.index + st; });
  def('VALUE', (a, c) => vecMap(a[0], c, v => toNum(v)));
  def('NUMBERVALUE', (a, c) => toNum(toStr(scalar(a[0], c)).replace(/\s/g, '')));
  def('TEXT', (a, c) => { const f = toStr(scalar(a[1], c)); return vecMap(a[0], c, v => fmtNumber(typeof v === 'string' && v !== '' && !isNaN(Number(v)) ? Number(v) : v, f)); });
  def('FIXED', (a, c) => { const d = a.length > 1 ? num(a[1], c) : 2; const nc = a.length > 2 && toBool(scalar(a[2], c)); return fmtNumber(num(a[0], c), (nc ? '0' : '#,##0') + (d > 0 ? '.' + '0'.repeat(d) : '')); });
  def('DOLLAR', (a, c) => { const d = a.length > 1 ? num(a[1], c) : 2; return fmtNumber(num(a[0], c), '$#,##0' + (d > 0 ? '.' + '0'.repeat(d) : '')); });
  def('TEXTBEFORE', (a, c) => { const s = toStr(scalar(a[0], c)), d = toStr(scalar(a[1], c)); const i = s.indexOf(d); if (i < 0) throw E('#N/A'); return s.slice(0, i); });
  def('TEXTAFTER', (a, c) => { const s = toStr(scalar(a[0], c)), d = toStr(scalar(a[1], c)); const i = s.indexOf(d); if (i < 0) throw E('#N/A'); return s.slice(i + d.length); });
  def('TEXTSPLIT', (a, c) => { const s = toStr(scalar(a[0], c)); const cd = a[1] ? toStr(scalar(a[1], c)) : ''; const rd = a.length > 2 && a[2] !== null ? toStr(scalar(a[2], c)) : null; const rows = rd ? s.split(rd) : [s]; const out = rows.map(r => cd ? r.split(cd) : [r]); const w = Math.max(...out.map(r => r.length)); return out.map(r => { while (r.length < w) r.push(E('#N/A')); return r; }); });
  // lookup
  function lookupMatch(v, vec, mode) { // mode 0 exact, 1 <=, -1 >=
    if (mode === 0) { const crit = typeof v === 'string' ? makeCrit('=' + v) : null; for (let i = 0; i < vec.length; i++) { const x = vec[i]; if (typeof v === 'string' ? (typeof x === 'string' && (/[*?]/.test(v) ? crit(x) : x.toLowerCase() === v.toLowerCase())) : (x !== null && typeof x === typeof v && x === v)) return i; } return -1; }
    let best = -1; for (let i = 0; i < vec.length; i++) { const x = vec[i]; if (x === null || isErr(x) || typeof x !== typeof v) continue; const cmp = cmpVals(x, v); if (mode === 1 ? cmp <= 0 : cmp >= 0) { if (best < 0 || (mode === 1 ? cmpVals(x, vec[best]) >= 0 : cmpVals(x, vec[best]) <= 0)) best = i; } } return best;
  }
  def('VLOOKUP', (a, c) => { const v = scalar(a[0], c); const g = grid(a[1], c); const ci = Math.floor(num(a[2], c)); const approx = a.length > 3 ? toBool(scalar(a[3], c)) : true; if (ci < 1 || ci > g[0].length) throw E('#REF!'); const col = g.map(r => r[0]); let i; if (approx) { i = -1; for (let k = 0; k < col.length; k++) { if (col[k] === null || typeof col[k] !== typeof v) continue; if (cmpVals(col[k], v) <= 0) i = k; else break; } } else i = lookupMatch(v, col, 0); if (i < 0) throw E('#N/A'); const r = g[i][ci - 1]; return r === null ? 0 : r; });
  def('HLOOKUP', (a, c) => { const v = scalar(a[0], c); const g = grid(a[1], c); const ri = Math.floor(num(a[2], c)); const approx = a.length > 3 ? toBool(scalar(a[3], c)) : true; if (ri < 1 || ri > g.length) throw E('#REF!'); const row = g[0]; let i; if (approx) { i = -1; for (let k = 0; k < row.length; k++) { if (row[k] === null || typeof row[k] !== typeof v) continue; if (cmpVals(row[k], v) <= 0) i = k; else break; } } else i = lookupMatch(v, row, 0); if (i < 0) throw E('#N/A'); const r = g[ri - 1][i]; return r === null ? 0 : r; });
  def('MATCH', (a, c) => { const v = scalar(a[0], c); const g = grid(a[1], c); const vec = g.length === 1 ? g[0] : g.map(r => r[0]); const mt = a.length > 2 ? num(a[2], c) : 1; const i = lookupMatch(v, vec, mt === 0 ? 0 : mt > 0 ? 1 : -1); if (i < 0) throw E('#N/A'); return i + 1; });
  def('XMATCH', (a, c) => { const v = scalar(a[0], c); const g = grid(a[1], c); const vec = g.length === 1 ? g[0] : g.map(r => r[0]); const mt = a.length > 2 ? num(a[2], c) : 0; const rev = a.length > 3 && num(a[3], c) < 0; let i = lookupMatch(v, rev ? vec.slice().reverse() : vec, mt === 0 ? 0 : mt > 0 ? -1 : 1); if (i < 0) throw E('#N/A'); if (rev) i = vec.length - 1 - i; return i + 1; });
  def('XLOOKUP', (a, c) => { const v = scalar(a[0], c); const lg = grid(a[1], c), rg = grid(a[2], c); const vert = lg[0].length === 1 && lg.length > 1 || (lg.length === 1 && lg[0].length === 1 && rg.length > 1); const vec = lg.length === 1 ? lg[0] : lg.map(r => r[0]); const mm = a.length > 4 && a[4] !== null ? num(a[4], c) : 0; const rev = a.length > 5 && a[5] !== null && num(a[5], c) < 0; let i = lookupMatch(v, rev ? vec.slice().reverse() : vec, mm === 0 ? 0 : mm === -1 ? 1 : mm === 1 ? -1 : 0); if (mm === 2 && i < 0) i = lookupMatch(v, vec, 0); if (i < 0) { if (a.length > 3 && a[3] !== null) return a[3]; throw E('#N/A'); } if (rev) i = vec.length - 1 - i; if (lg.length === 1 && lg[0].length > 1) return rg.map(r => [r[i]]).length === 1 ? rg[0][i] : rg.map(r => [r[i]]); return rg[i].length === 1 ? rg[i][0] : [rg[i]]; });
  def('LOOKUP', (a, c) => { const v = scalar(a[0], c); const lg = grid(a[1], c); const vec = lg.length === 1 ? lg[0] : lg.map(r => r[0]); let i = -1; for (let k = 0; k < vec.length; k++) { if (vec[k] !== null && typeof vec[k] === typeof v && cmpVals(vec[k], v) <= 0) i = k; } if (i < 0) throw E('#N/A'); const rg = a.length > 2 ? grid(a[2], c) : lg; const rv = rg.length === 1 ? rg[0] : rg.map(r => r[0]); return rv[i]; });
  def('INDEX', (a, c) => { const g = grid(a[0], c); let r = a.length > 1 ? Math.floor(num(a[1], c)) : 0, cc = a.length > 2 ? Math.floor(num(a[2], c)) : 0; if (g.length === 1 && a.length === 2) { cc = r; r = 1; } if (r < 0 || cc < 0 || r > g.length || cc > g[0].length) throw E('#REF!'); if (r === 0 && cc === 0) return g; if (r === 0) return g.map(row => [row[cc - 1]]); if (cc === 0) return [g[r - 1]]; const v = g[r - 1][cc - 1]; return v === null ? 0 : v; });
  def('ROW', (a, c) => { if (!a.length) return c.r; if (a[0] instanceof Rng) { const rg = a[0]; if (rg.r1 === rg.r2) return rg.r1; const o = []; for (let r = rg.r1; r <= rngBounds(rg, c).r2; r++) o.push([r]); return o; } throw E('#VALUE!'); }, true);
  def('COLUMN', (a, c) => { if (!a.length) return c.c; if (a[0] instanceof Rng) { const rg = a[0]; if (rg.c1 === rg.c2) return rg.c1; const o = []; for (let k = rg.c1; k <= rngBounds(rg, c).c2; k++) o.push(k); return [o]; } throw E('#VALUE!'); }, true);
  def('ROWS', (a, c) => grid(a[0], c).length, true); def('COLUMNS', (a, c) => grid(a[0], c)[0].length, true);
  def('TRANSPOSE', (a, c) => { const g = grid(a[0], c); return g[0].map((_, j) => g.map(r => r[j])); });
  def('ADDRESS', (a, c) => { const r = num(a[0], c), cc = num(a[1], c), abs = a.length > 2 ? num(a[2], c) : 1; const s = (abs === 1 || abs === 3 ? '$' : '') + numToCol(cc) + (abs === 1 || abs === 2 ? '$' : '') + r; return a.length > 4 && a[4] ? toStr(scalar(a[4], c)) + '!' + s : s; });
  def('INDIRECT', (a, c) => { const t = toStr(scalar(a[0], c)); const m = /^(?:'?([^'!]+)'?!)?(.+)$/.exec(t); const n = parse((m[1] ? "'" + m[1] + "'!" : '') + m[2]); return evalNode(n, c); });
  def('OFFSET', (a, c) => { const rg = a[0]; if (!(rg instanceof Rng)) throw E('#VALUE!'); const dr = num(a[1], c), dc = num(a[2], c); const h = a.length > 3 && a[3] !== null ? num(a[3], c) : rg.r2 - rg.r1 + 1, w = a.length > 4 && a[4] !== null ? num(a[4], c) : rg.c2 - rg.c1 + 1; const r1 = rg.r1 + dr, c1 = rg.c1 + dc; if (r1 < 1 || c1 < 1) throw E('#REF!'); return new Rng(rg.sheet, r1, c1, r1 + h - 1, c1 + w - 1); }, true);
  // dynamic arrays
  def('FILTER', (a, c) => { const g = grid(a[0], c), inc = grid(a[1], c); const rowMode = inc[0].length === 1 || inc.length === g.length && inc[0].length === 1; let out; if (inc.length === g.length && inc[0].length === 1) { out = g.filter((r, i) => { const v = inc[i][0]; if (isErr(v)) throw v; return toBool(v === null ? false : v) && v !== 0; }); } else if (inc.length === 1 && inc[0].length === g[0].length) { const keep = inc[0].map(v => toBool(v)); out = g.map(r => r.filter((_, j) => keep[j])); if (!out[0].length) out = []; } else throw E('#VALUE!'); if (!out.length) { if (a.length > 2 && a[2] !== null) return a[2]; throw E('#CALC!'); } return out; });
  def('SORT', (a, c) => { const g = grid(a[0], c).map(r => r.slice()); const ix = a.length > 1 && a[1] !== null ? num(a[1], c) - 1 : 0; const ord = a.length > 2 && a[2] !== null ? num(a[2], c) : 1; const byCol = a.length > 3 && a[3] !== null && toBool(scalar(a[3], c)); if (byCol) { const t = g[0].map((_, j) => g.map(r => r[j])); t.sort((x, y) => ord * cmpVals(x[ix], y[ix])); return t[0].map((_, j) => t.map(r => r[j])); } g.sort((x, y) => ord * cmpVals(x[ix], y[ix])); return g; });
  def('SORTBY', (a, c) => { const g = grid(a[0], c); const keys = []; for (let i = 1; i < a.length; i += 2) keys.push({ v: grid(a[i], c).map(r => r[0]), o: i + 1 < a.length ? num(a[i + 1], c) : 1 }); const idx = g.map((_, i) => i); idx.sort((x, y) => { for (const k of keys) { const r = k.o * cmpVals(k.v[x], k.v[y]); if (r) return r; } return x - y; }); return idx.map(i => g[i]); });
  def('UNIQUE', (a, c) => { const g = grid(a[0], c); const byCol = a.length > 1 && a[1] !== null && toBool(scalar(a[1], c)); const once = a.length > 2 && a[2] !== null && toBool(scalar(a[2], c)); const src = byCol ? g[0].map((_, j) => g.map(r => r[j])) : g; const key = r => JSON.stringify(r.map(v => typeof v === 'string' ? v.toLowerCase() : v)); const cnt = {}; src.forEach(r => { cnt[key(r)] = (cnt[key(r)] || 0) + 1; }); const seen = new Set(); const out = []; for (const r of src) { const k = key(r); if (once ? cnt[k] === 1 : !seen.has(k)) out.push(r); seen.add(k); } if (byCol) return out[0].map((_, j) => out.map(r => r[j])); return out; });
  def('SEQUENCE', (a, c) => { const R = num(a[0], c), C = a.length > 1 && a[1] !== null ? num(a[1], c) : 1, s = a.length > 2 && a[2] !== null ? num(a[2], c) : 1, st = a.length > 3 && a[3] !== null ? num(a[3], c) : 1; const o = []; let v = s; for (let i = 0; i < R; i++) { const row = []; for (let j = 0; j < C; j++) { row.push(v); v += st; } o.push(row); } return o; });
  def('RANDARRAY', (a, c) => { const R = a.length ? num(a[0], c) : 1, C = a.length > 1 ? num(a[1], c) : 1; return Array.from({ length: R }, () => Array.from({ length: C }, () => Math.random())); });
  def('TAKE', (a, c) => { const g = grid(a[0], c); const r = num(a[1], c), cc = a.length > 2 && a[2] !== null ? num(a[2], c) : null; let o = r >= 0 ? g.slice(0, r) : g.slice(r); if (cc !== null) o = o.map(row => cc >= 0 ? row.slice(0, cc) : row.slice(cc)); return o; });
  def('DROP', (a, c) => { const g = grid(a[0], c); const r = num(a[1], c), cc = a.length > 2 && a[2] !== null ? num(a[2], c) : 0; let o = r >= 0 ? g.slice(r) : g.slice(0, r); if (cc) o = o.map(row => cc >= 0 ? row.slice(cc) : row.slice(0, cc)); if (!o.length || !o[0].length) throw E('#CALC!'); return o; });
  def('VSTACK', (a, c) => { let o = []; const w = Math.max(...a.map(x => grid(x, c)[0].length)); for (const x of a) for (const r of grid(x, c)) { const row = r.slice(); while (row.length < w) row.push(E('#N/A')); o.push(row); } return o; });
  def('HSTACK', (a, c) => { const gs = a.map(x => grid(x, c)); const h = Math.max(...gs.map(g => g.length)); const o = []; for (let i = 0; i < h; i++) { let row = []; for (const g of gs) { const r = g[i] || Array(g[0].length).fill(E('#N/A')); row = row.concat(r); } o.push(row); } return o; });
  def('TOCOL', (a, c) => { const o = []; for (const r of grid(a[0], c)) for (const v of r) o.push([v]); return o; });
  def('TOROW', (a, c) => { const o = []; for (const r of grid(a[0], c)) for (const v of r) o.push(v); return [o]; });
  def('CHOOSECOLS', (a, c) => { const g = grid(a[0], c); const ix = a.slice(1).map(x => num(x, c) - 1); return g.map(r => ix.map(i => r[i < 0 ? r.length + i + 1 : i])); });
  def('CHOOSEROWS', (a, c) => { const g = grid(a[0], c); return a.slice(1).map(x => g[num(x, c) - 1]); });
  def('WRAPROWS', (a, c) => { const v = [].concat(...grid(a[0], c)); const w = num(a[1], c); const o = []; for (let i = 0; i < v.length; i += w) { const r = v.slice(i, i + w); while (r.length < w) r.push(E('#N/A')); o.push(r); } return o; });
  // date/time
  def('TODAY', (a, c) => Math.floor(dateToSerial(c.wb.now)));
  def('NOW', (a, c) => dateToSerial(c.wb.now));
  def('DATE', (a, c) => { let y = num(a[0], c), m = num(a[1], c), d = num(a[2], c); if (y < 1900) y += 1900; return Math.round((Date.UTC(y, m - 1, d) - EPOCH) / 864e5); });
  def('TIME', (a, c) => (num(a[0], c) * 3600 + num(a[1], c) * 60 + num(a[2], c)) / 86400);
  def('YEAR', (a, c) => vecMap(a[0], c, v => serialToDate(toNum(v)).y)); def('MONTH', (a, c) => vecMap(a[0], c, v => serialToDate(toNum(v)).m)); def('DAY', (a, c) => vecMap(a[0], c, v => serialToDate(toNum(v)).d));
  def('HOUR', (a, c) => serialToDate(num(a[0], c)).h); def('MINUTE', (a, c) => serialToDate(num(a[0], c)).mi); def('SECOND', (a, c) => serialToDate(num(a[0], c)).s);
  def('WEEKDAY', (a, c) => { const wd = serialToDate(num(a[0], c)).wd; const t = a.length > 1 ? num(a[1], c) : 1; return t === 2 ? (wd + 6) % 7 + 1 : t === 3 ? (wd + 6) % 7 : wd + 1; });
  def('WEEKNUM', (a, c) => { const d = serialToDate(num(a[0], c)); const jan1 = mkSerial(d.y, 1, 1); const j = serialToDate(jan1).wd; return Math.floor((num(a[0], c) - jan1 + j) / 7) + 1; });
  def('DATEVALUE', (a, c) => { const d = parseDateStr(toStr(scalar(a[0], c))); if (d === null) throw E('#VALUE!'); return Math.floor(d); });
  def('TIMEVALUE', (a, c) => { const d = parseDateStr(toStr(scalar(a[0], c))); if (d === null) throw E('#VALUE!'); return d % 1; });
  def('DAYS', (a, c) => num(a[0], c) - num(a[1], c));
  def('EDATE', (a, c) => { const d = serialToDate(num(a[0], c)); const t = d.m - 1 + num(a[1], c); const y = d.y + Math.floor(t / 12), m = ((t % 12) + 12) % 12; const dim = new Date(Date.UTC(y, m + 1, 0)).getUTCDate(); return mkSerial(y, m + 1, Math.min(d.d, dim)); });
  def('EOMONTH', (a, c) => { const d = serialToDate(num(a[0], c)); const t = d.m - 1 + num(a[1], c); const y = d.y + Math.floor(t / 12), m = ((t % 12) + 12) % 12; return mkSerial(y, m + 1, new Date(Date.UTC(y, m + 1, 0)).getUTCDate()); });
  def('DATEDIF', (a, c) => { const s = num(a[0], c), e = num(a[1], c); const u = toStr(scalar(a[2], c)).toUpperCase(); if (e < s) throw E('#NUM!'); const A = serialToDate(s), B = serialToDate(e); let months = (B.y - A.y) * 12 + B.m - A.m - (B.d < A.d ? 1 : 0); switch (u) { case 'D': return Math.floor(e - s); case 'M': return months; case 'Y': return Math.floor(months / 12); case 'MD': { let d = B.d - A.d; if (d < 0) d += new Date(Date.UTC(B.y, B.m - 1, 0)).getUTCDate(); return d; } case 'YM': return months % 12; case 'YD': { let y = B.y; let st = mkSerial(y, A.m, A.d); if (st > e) st = mkSerial(y - 1, A.m, A.d); return Math.floor(e - st); } } throw E('#NUM!'); });
  def('NETWORKDAYS', (a, c) => { let s = Math.floor(num(a[0], c)), e = Math.floor(num(a[1], c)); const hol = a.length > 2 ? new Set(flatNums([a[2]], c).map(Math.floor)) : new Set(); let n = 0; const sg = s <= e ? 1 : -1; for (let d = s; sg > 0 ? d <= e : d >= e; d += sg) { const wd = serialToDate(d).wd; if (wd !== 0 && wd !== 6 && !hol.has(d)) n++; } return n * sg; });
  def('WORKDAY', (a, c) => { let d = Math.floor(num(a[0], c)); let n = num(a[1], c); const hol = a.length > 2 ? new Set(flatNums([a[2]], c).map(Math.floor)) : new Set(); const sg = n >= 0 ? 1 : -1; n = Math.abs(n); while (n > 0) { d += sg; const wd = serialToDate(d).wd; if (wd !== 0 && wd !== 6 && !hol.has(d)) n--; } return d; });
  // financial
  def('PMT', (a, c) => { const r = num(a[0], c), n = num(a[1], c), pv = num(a[2], c), fv = a.length > 3 ? num(a[3], c) : 0, t = a.length > 4 ? num(a[4], c) : 0; if (r === 0) return -(pv + fv) / n; return -(r * (pv * Math.pow(1 + r, n) + fv)) / ((1 + r * t) * (Math.pow(1 + r, n) - 1)); });
  def('FV', (a, c) => { const r = num(a[0], c), n = num(a[1], c), p = num(a[2], c), pv = a.length > 3 ? num(a[3], c) : 0, t = a.length > 4 ? num(a[4], c) : 0; if (r === 0) return -(pv + p * n); return -(pv * Math.pow(1 + r, n) + p * (1 + r * t) * (Math.pow(1 + r, n) - 1) / r); });
  def('PV', (a, c) => { const r = num(a[0], c), n = num(a[1], c), p = num(a[2], c), fv = a.length > 3 ? num(a[3], c) : 0, t = a.length > 4 ? num(a[4], c) : 0; if (r === 0) return -(fv + p * n); return -(fv + p * (1 + r * t) * (Math.pow(1 + r, n) - 1) / r) / Math.pow(1 + r, n); });
  def('NPER', (a, c) => { const r = num(a[0], c), p = num(a[1], c), pv = num(a[2], c), fv = a.length > 3 ? num(a[3], c) : 0, t = a.length > 4 ? num(a[4], c) : 0; if (r === 0) return -(pv + fv) / p; return Math.log((p * (1 + r * t) - fv * r) / (pv * r + p * (1 + r * t))) / Math.log(1 + r); });
  def('NPV', (a, c) => { const r = num(a[0], c); const v = flatNums(a.slice(1), c); return v.reduce((s, x, i) => s + x / Math.pow(1 + r, i + 1), 0); });
  def('IRR', (a, c) => { const v = flatNums([a[0]], c); let r = a.length > 1 ? num(a[1], c) : 0.1; for (let i = 0; i < 100; i++) { let f = 0, d = 0; v.forEach((x, k) => { f += x / Math.pow(1 + r, k); d -= k * x / Math.pow(1 + r, k + 1); }); const nr = r - f / d; if (!isFinite(nr)) throw E('#NUM!'); if (Math.abs(nr - r) < 1e-10) return nr; r = nr; } throw E('#NUM!'); });
  def('RATE', (a, c) => { const n = num(a[0], c), p = num(a[1], c), pv = num(a[2], c), fv = a.length > 3 ? num(a[3], c) : 0, t = a.length > 4 ? num(a[4], c) : 0; let r = a.length > 5 ? num(a[5], c) : 0.1; for (let i = 0; i < 100; i++) { const f = pv * Math.pow(1 + r, n) + p * (1 + r * t) * (Math.pow(1 + r, n) - 1) / r + fv; const h = 1e-6; const r2 = r + h; const f2 = pv * Math.pow(1 + r2, n) + p * (1 + r2 * t) * (Math.pow(1 + r2, n) - 1) / r2 + fv; const nr = r - f * h / (f2 - f); if (Math.abs(nr - r) < 1e-10) return nr; r = nr; } throw E('#NUM!'); });
  def('SLN', (a, c) => (num(a[0], c) - num(a[1], c)) / num(a[2], c));
  def('IPMT', (a, c) => { const r = num(a[0], c), per = num(a[1], c), n = num(a[2], c), pv = num(a[3], c); const pmt = r === 0 ? -pv / n : -(r * pv * Math.pow(1 + r, n)) / (Math.pow(1 + r, n) - 1); const bal = pv * Math.pow(1 + r, per - 1) + pmt * (Math.pow(1 + r, per - 1) - 1) / r; return -bal * r; });
  def('CONVERT', (a, c) => num(a[0], c)); def('ISFORMULA', (a, c) => false);
  def('SUBTOTAL', (a, c) => { const fc = num(a[0], c), f = fc % 100, skipH = fc > 100; const v = []; let cnt = 0; for (const x of a.slice(1)) { if (isErr(x)) throw x; if (x instanceof Rng) { const sn = x.sheet || c.sheet, S = c.wb.sheets[sn]; for (let r = x.r1; r <= x.r2; r++) { if (skipH && S.rowH && S.rowH[r] === 0) continue; for (let cc = x.c1; cc <= x.c2; cc++) { const cell = S.cells[c.wb.key(r, cc)]; if (cell && cell.f && /SUBTOTAL\(/i.test(cell.f)) continue; const val = c.wb.getVal(sn, r, cc); if (val !== null && val !== '' && val !== undefined) cnt++; if (typeof val === 'number') v.push(val); } } } else if (Array.isArray(x)) { for (const row of x) for (const val of row) { if (typeof val === 'number') v.push(val); if (val !== null && val !== '') cnt++; } } else if (typeof x === 'number') { v.push(x); cnt++; } } switch (f) { case 1: return mean(v); case 2: return v.length; case 3: return cnt; case 4: return v.length ? Math.max(...v) : 0; case 5: return v.length ? Math.min(...v) : 0; case 6: return v.reduce((x, y) => x * y, 1); case 9: return v.reduce((x, y) => x + y, 0); case 7: case 8: case 10: case 11: { const m = v.length ? mean(v) : 0; const ss = v.reduce((t, y) => t + (y - m) * (y - m), 0); const d = (f === 7 || f === 8) ? (f === 7 ? v.length - 1 : v.length) : (f === 10 ? v.length - 1 : v.length); if (d <= 0) throw E('#DIV/0!'); return (f === 7 || f === 8) ? Math.sqrt(ss / d) : ss / d; } } throw E('#VALUE!'); }, true);
  def('AGGREGATE', (a, c) => FN.SUBTOTAL([a[0], ...a.slice(2)], c));
  def('ERROR.TYPE', (a, c) => { const v = scalar(a[0], c); const L = ['#NULL!', '#DIV/0!', '#VALUE!', '#REF!', '#NAME?', '#NUM!', '#N/A']; const i = L.indexOf(v.code); if (i < 0) throw E('#N/A'); return i + 1; }, true);
  def('COUNTUNIQUE', (a, c) => new Set([].concat(...grid(a[0], c)).map(v => typeof v === 'string' ? v.toLowerCase() : v)).size);

  // ---------- public API ----------
  function evaluate(wb, sheet, formula) { const ctx = { sheet, r: 1, c: 1, wb, locals: {} }; wb._stack = wb._stack || new Set(); try { let v = evalNode(parse(formula[0] === '=' ? formula.slice(1) : formula), ctx); if (v instanceof Rng || Array.isArray(v)) v = grid(v, ctx)[0][0]; return v; } catch (e) { return isErr(e) ? e : E('#VALUE!'); } }
  // ---------- LAMBDA family, MMULT, misc ----------
  LAZY.LAMBDA = (a, c) => { const params = a.slice(0, -1).map(x => { if (x.t !== 'name') throw E('#VALUE!'); return x.v.toUpperCase(); }); return { lam: true, params, body: a[a.length - 1], ctx: c }; };
  const callLam = (l, args) => { if (!l || !l.lam) throw E('#VALUE!'); const locals = Object.assign({}, l.ctx.locals); l.params.forEach((p, i) => { locals[p] = args[i] === undefined ? null : args[i]; }); const c2 = Object.assign({}, l.ctx, { locals }); const v = evalNode(l.body, c2); if (v instanceof Rng) return scalar(v, c2); return v; };
  const arr2 = (v, c) => { const g = grid(v, c); return g; };
  LAZY.MAP = (a, c) => { const lam = evalNode(a[a.length - 1], c); const gs = a.slice(0, -1).map(x => arr2(evalNode(x, c), c)); const R = gs[0].length, C = gs[0][0].length; const o = []; for (let i = 0; i < R; i++) { const row = []; for (let j = 0; j < C; j++) { let v; try { v = callLam(lam, gs.map(g => g[i] && g[i][j])); } catch (e) { v = isErr(e) ? e : E('#VALUE!'); } row.push(Array.isArray(v) ? v[0][0] : v); } o.push(row); } return o; };
  LAZY.REDUCE = (a, c) => { const lam = evalNode(a[2], c); let acc = a[0].t === 'empty' ? 0 : evalNode(a[0], c); const g = arr2(evalNode(a[1], c), c); for (const row of g) for (const v of row) acc = callLam(lam, [acc, v]); return acc; };
  LAZY.SCAN = (a, c) => { const lam = evalNode(a[2], c); let acc = a[0].t === 'empty' ? 0 : evalNode(a[0], c); const g = arr2(evalNode(a[1], c), c); return g.map(row => row.map(v => { acc = callLam(lam, [acc, v]); return acc; })); };
  LAZY.BYROW = (a, c) => { const lam = evalNode(a[1], c); return arr2(evalNode(a[0], c), c).map(row => [callLam(lam, [[row]])]); };
  LAZY.BYCOL = (a, c) => { const lam = evalNode(a[1], c); const g = arr2(evalNode(a[0], c), c); return [g[0].map((_, j) => callLam(lam, [g.map(r => [r[j]])]))]; };
  LAZY.MAKEARRAY = (a, c) => { const R = toNum(scalar(evalNode(a[0], c), c)), C = toNum(scalar(evalNode(a[1], c), c)), lam = evalNode(a[2], c); const o = []; for (let i = 1; i <= R; i++) { const row = []; for (let j = 1; j <= C; j++) row.push(callLam(lam, [i, j])); o.push(row); } return o; };
  def('MMULT', (a, c) => { const A = grid(a[0], c), B = grid(a[1], c); if (A[0].length !== B.length) throw E('#VALUE!'); return A.map(r => B[0].map((_, j) => r.reduce((s, v, k) => s + toNum(v) * toNum(B[k][j]), 0))); });
  def('FORMULATEXT', (a, c) => { const x = a[0]; if (!(x instanceof Rng)) throw E('#N/A'); const b = rngBounds(x, c); const raw = c.wb.getRaw(b.sheet, b.r1, b.c1); if (typeof raw === 'string' && raw[0] === '=') return raw; throw E('#N/A'); }, true);
  def('ISOWEEKNUM', (a, c) => { const d = serialToDate(Math.floor(num(a[0], c))); const t = new Date(Date.UTC(d.y, d.m - 1, d.d)); const day = t.getUTCDay() || 7; t.setUTCDate(t.getUTCDate() + 4 - day); const y0 = new Date(Date.UTC(t.getUTCFullYear(), 0, 1)); return Math.ceil(((t - y0) / 86400000 + 1) / 7); });
  def('YEARFRAC', (a, c) => { let s = num(a[0], c), e = num(a[1], c); const basis = a.length > 2 && a[2] !== null ? num(a[2], c) : 0; if (s > e) [s, e] = [e, s]; const d1 = serialToDate(Math.floor(s)), d2 = serialToDate(Math.floor(e)); if (basis === 0 || basis === 4) { let D1 = Math.min(d1.d, 30), D2 = d2.d; if (D2 === 31 && D1 === 30) D2 = 30; return ((d2.y - d1.y) * 360 + (d2.m - d1.m) * 30 + (D2 - D1)) / 360; } if (basis === 2) return (e - s) / 360; if (basis === 3) return (e - s) / 365; return (e - s) / 365.25; });
  def('EXPAND', (a, c) => { const g = grid(a[0], c), R = num(a[1], c), C = a.length > 2 && a[2] !== null ? num(a[2], c) : g[0].length, pad = a.length > 3 ? scalar(a[3], c) : E('#N/A'); const o = []; for (let i = 0; i < Math.max(R, g.length); i++) { const row = []; for (let j = 0; j < Math.max(C, g[0].length); j++) row.push(g[i] && g[i][j] !== undefined ? g[i][j] : pad); o.push(row); } return o; });
  def('WRAPCOLS', (a, c) => { const v = grid(a[0], c).flat(), n = num(a[1], c), pad = a.length > 2 ? scalar(a[2], c) : E('#N/A'); const cols = Math.ceil(v.length / n); const o = []; for (let i = 0; i < n; i++) { const row = []; for (let j = 0; j < cols; j++) { const x = v[j * n + i]; row.push(x === undefined ? pad : x); } o.push(row); } return o; });
  const api = { Workbook, XErr, Rng, isErr, parse, evaluate, fmtNumber, colToNum, numToCol, parseAddr, addr, toNum, toStr, toBool, numStr, dateToSerial, serialToDate, FN, E };
  if (typeof module !== 'undefined' && module.exports) module.exports = api; else root.Engine = api;
})(typeof window !== 'undefined' ? window : globalThis);
