// VBA interpreter (practical subset): lexer, parser, tree-walking runtime, builtin library, Collection/Dictionary
(function (root) {
  'use strict';
  const Engine = (typeof require !== 'undefined' && typeof module !== 'undefined') ? require('./engine.js') : root.Engine;

  class VBAError extends Error { constructor(num, desc, src) { super(desc); this.number = num; this.description = desc; this.source = src || 'VBA'; } }
  class ExitSignal { constructor(kind) { this.kind = kind; } }
  class GotoSignal { constructor(label) { this.label = label; } }
  class VDate { constructor(s) { this.s = s; } valueOf() { return this.s; } }
  const NOTHING = { vbNothing: true }, NULLV = { vbNull: true };
  class VArr {
    constructor(dims, fill) { this.dims = dims; let n = 1; for (const d of dims) n *= (d.hi - d.lo + 1); this.data = new Array(Math.max(0, n)).fill(fill === undefined ? null : fill); }
    idx(ix) { let o = 0; for (let k = 0; k < this.dims.length; k++) { const d = this.dims[k], i = ix[k]; if (i < d.lo || i > d.hi) throw new VBAError(9, 'Subscript out of range'); o = o * (d.hi - d.lo + 1) + (i - d.lo); } return o; }
    get(ix) { if (ix.length !== this.dims.length) throw new VBAError(9, 'Subscript out of range'); return this.data[this.idx(ix)]; }
    set(ix, v) { this.data[this.idx(ix)] = v; }
    static from1D(list, lo) { const a = new VArr([{ lo: lo || 0, hi: (lo || 0) + list.length - 1 }]); a.data = list.slice(); return a; }
  }
  const isObj = v => v !== null && typeof v === 'object' && typeof v.vbGet === 'function';

  // ---------------- value helpers ----------------
  const isDateV = v => v instanceof VDate;
  function toNumber(v) {
    if (v === null || v === undefined) return 0; if (typeof v === 'number') return v; if (typeof v === 'boolean') return v ? -1 : 0;
    if (v instanceof VDate) return v.s;
    if (typeof v === 'string') { const t = v.trim(); if (t === '') throw new VBAError(13, 'Type mismatch'); const n = Number(t.replace(/,/g, '')); if (!isNaN(n)) return n; if (/^&H[0-9a-f]+$/i.test(t)) return parseInt(t.slice(2), 16); throw new VBAError(13, 'Type mismatch'); }
    if (v === NULLV) throw new VBAError(94, 'Invalid use of Null');
    if (isObj(v)) return toNumber(v.vbGet('Value', []));
    throw new VBAError(13, 'Type mismatch');
  }
  function toStr(v) {
    if (v === null || v === undefined) return ''; if (typeof v === 'string') return v; if (typeof v === 'boolean') return v ? 'True' : 'False';
    if (typeof v === 'number') return Engine.numStr(v);
    if (v instanceof VDate) return fmtVBA(v, v.s % 1 === 0 ? 'Short Date' : 'General Date');
    if (v === NOTHING) throw new VBAError(91, 'Object variable or With block variable not set');
    if (v instanceof VArr) throw new VBAError(13, 'Type mismatch');
    if (isObj(v)) return toStr(v.vbGet('Value', []));
    return String(v);
  }
  function toBoolean(v) { if (typeof v === 'boolean') return v; if (v === null || v === undefined) return false; if (typeof v === 'string') { const l = v.toLowerCase(); if (l === 'true') return true; if (l === 'false') return false; return toNumber(v) !== 0; } return toNumber(v) !== 0; }
  const isNumericStr = s => typeof s === 'string' && s.trim() !== '' && !isNaN(Number(s.trim().replace(/,/g, '')));
  function bankers(x) { const f = Math.floor(x), d = x - f; if (Math.abs(d - 0.5) < 1e-12) return f % 2 === 0 ? f : f + 1; return Math.round(x); }
  function vround(x, n) { const m = Math.pow(10, n || 0); const y = x * m; const f = Math.floor(y), d = y - f; let r; if (Math.abs(d - 0.5) < 1e-9) r = f % 2 === 0 ? f : f + 1; else r = Math.round(y); return r / m; }
  function fmtVBA(v, f) {
    if (v === null || v === undefined) return ''; const isD = v instanceof VDate; const n = isD ? v.s : v;
    const named = { 'general number': 'General', 'currency': '$#,##0.00', 'fixed': '0.00', 'standard': '#,##0.00', 'percent': '0.00%', 'scientific': '0.00E+00', 'short date': 'm/d/yyyy', 'long date': 'dddd, mmmm d, yyyy', 'medium date': 'dd-mmm-yy', 'short time': 'hh:mm', 'long time': 'h:mm:ss AM/PM', 'general date': 'm/d/yyyy h:mm:ss AM/PM' };
    const lf = String(f).toLowerCase();
    if (lf === 'yes/no') return toNumber(v) !== 0 ? 'Yes' : 'No'; if (lf === 'true/false') return toNumber(v) !== 0 ? 'True' : 'False'; if (lf === 'on/off') return toNumber(v) !== 0 ? 'On' : 'Off';
    if (lf === 'general date') return isD && n % 1 === 0 ? Engine.fmtNumber(n, 'm/d/yyyy') : Engine.fmtNumber(n, 'm/d/yyyy h:mm:ss AM/PM');
    f = named[lf] || f;
    if (typeof v === 'string') { if (isNumericStr(v)) return Engine.fmtNumber(Number(v), f); return /^[<>@!&]/.test(f) ? (f[0] === '>' ? v.toUpperCase() : v.toLowerCase()) : v; }
    if (typeof v === 'boolean') return v ? 'True' : 'False';
    return Engine.fmtNumber(n, f);
  }
  function cmp(a, b, textCmp) {
    if (a === null || a === undefined) a = (typeof b === 'string') ? '' : 0; if (b === null || b === undefined) b = (typeof a === 'string') ? '' : 0;
    if (a instanceof VDate) a = a.s; if (b instanceof VDate) b = b.s; if (isObj(a)) a = a.vbGet('Value', []); if (isObj(b)) b = b.vbGet('Value', []);
    if (typeof a === 'boolean') a = a ? -1 : 0; if (typeof b === 'boolean') b = b ? -1 : 0;
    if (typeof a === 'string' && typeof b === 'string') { if (textCmp) { a = a.toLowerCase(); b = b.toLowerCase(); } return a < b ? -1 : a > b ? 1 : 0; }
    if (typeof a === 'number' && typeof b === 'number') return a < b ? -1 : a > b ? 1 : 0;
    if (typeof a === 'string') { if (isNumericStr(a)) a = Number(a.replace(/,/g, '')); else return 1; } if (typeof b === 'string') { if (isNumericStr(b)) b = Number(b.replace(/,/g, '')); else return -1; }
    return a < b ? -1 : a > b ? 1 : 0;
  }
  function likeMatch(s, pat, textCmp) {
    let rx = ''; for (let i = 0; i < pat.length; i++) { const ch = pat[i]; if (ch === '*') rx += '.*'; else if (ch === '?') rx += '.'; else if (ch === '#') rx += '\\d'; else if (ch === '[') { const j = pat.indexOf(']', i + 1); if (j < 0) { rx += '\\['; continue; } let set = pat.slice(i + 1, j); if (set[0] === '!') set = '^' + set.slice(1); rx += '[' + set.replace(/\\/g, '\\\\') + ']'; i = j; } else rx += ch.replace(/[.+^${}()|\\\]]/g, '\\$&'); }
    return new RegExp('^' + rx + '$', textCmp ? 'i' : '').test(s);
  }
  function defaultFor(type) {
    switch ((type || '').toLowerCase()) { case 'integer': case 'long': case 'longlong': case 'double': case 'single': case 'byte': case 'currency': case 'decimal': return 0; case 'string': return ''; case 'boolean': return false; case 'date': return new VDate(0); case 'object': return NOTHING; default: return null; }
  }
  function coerceType(type, v) {
    if (!type) return v; const t = type.toLowerCase();
    switch (t) { case 'integer': case 'long': case 'longlong': case 'byte': { if (v === null) return 0; const n = bankers(toNumber(v)); return n; } case 'double': case 'single': case 'currency': case 'decimal': return v === null ? 0 : toNumber(v); case 'string': return v === null ? '' : toStr(v); case 'boolean': return toBoolean(v); case 'date': return v instanceof VDate ? v : v === null ? new VDate(0) : (typeof v === 'string' ? new VDate(Engine.FN.DATEVALUE([v], { wb: null })) : new VDate(toNumber(v))); }
    return v;
  }

  // ---------------- tokenizer ----------------
  function lex(src) {
    const t = []; let i = 0; const n = src.length; let line = 1;
    const push = (k, v, extra) => t.push(Object.assign({ k, v, line }, extra || {}));
    while (i < n) {
      const ch = src[i];
      if (ch === '\r') { i++; continue; }
      if (ch === '\n') { push('eol', '\n'); line++; i++; continue; }
      if (ch === ' ' || ch === '\t') { // line continuation " _"
        let j = i; while (j < n && (src[j] === ' ' || src[j] === '\t')) j++;
        if (src[j] === '_' && (src[j + 1] === '\n' || src[j + 1] === '\r' || j + 1 >= n)) { i = j + 1; if (src[i] === '\r') i++; if (src[i] === '\n') { i++; line++; } continue; }
        i++; continue;
      }
      if (ch === "'") { while (i < n && src[i] !== '\n') i++; continue; }
      if (ch === '"') { let s = ''; i++; while (i < n) { if (src[i] === '"') { if (src[i + 1] === '"') { s += '"'; i += 2; continue; } break; } s += src[i++]; } i++; push('str', s); continue; }
      if (ch === '#') { const m = /^#([^#\n]+)#/.exec(src.slice(i)); if (m) { push('date', m[1]); i += m[0].length; continue; } }
      if (/[0-9]/.test(ch) || (ch === '.' && /[0-9]/.test(src[i + 1] || ''))) { const m = /^(\d+\.?\d*|\.\d+)([eE][+-]?\d+)?[%&#!@]?/.exec(src.slice(i)); push('num', parseFloat(m[0])); i += m[0].length; continue; }
      if (ch === '&' && /^&[hH][0-9a-fA-F]+&?/.test(src.slice(i))) { const m = /^&[hH]([0-9a-fA-F]+)&?/.exec(src.slice(i)); push('num', parseInt(m[1], 16)); i += m[0].length; continue; }
      if (/[A-Za-z_À-￿]/.test(ch) || (ch === '[')) {
        if (ch === '[') { const j = src.indexOf(']', i); push('id', src.slice(i + 1, j), { bracket: true }); i = j + 1; continue; }
        const m = /^[A-Za-z_À-￿][A-Za-z0-9_À-￿]*[$%&#!@]?/.exec(src.slice(i)); let w = m[0]; i += w.length;
        if (/[$%&#!@]$/.test(w) && !(w.endsWith('&') && src[i] === '=')) w = w.slice(0, -1); else if (/[$%&#!@]$/.test(w)) { w = w.slice(0, -1); i--; }
        if (w.toLowerCase() === 'rem' && (t.length === 0 || t[t.length - 1].k === 'eol')) { while (i < n && src[i] !== '\n') i++; continue; }
        push('id', w); continue;
      }
      const two = src.substr(i, 2);
      if ([':=', '<=', '>=', '<>'].includes(two)) { push('op', two); i += 2; continue; }
      if ('+-*/\\^&=<>(),.:;!'.includes(ch)) { push('op', ch); i++; continue; }
      throw new VBAError(1001, 'Unexpected character "' + ch + '" on line ' + line);
    }
    push('eol', '\n'); push('eof', '');
    return t;
  }

  // ---------------- parser ----------------
  const BLOCK_END = { if: 1, select: 1, with: 1, sub: 1, function: 1, type: 1, property: 1 };
  function parseProgram(src) {
    const toks = lex(src); let p = 0;
    const procs = {}, globals = [], consts = [], enums = [];
    const peek = (o) => toks[p + (o || 0)], next = () => toks[p++];
    const lw = t => (t && t.k === 'id') ? t.v.toLowerCase() : null;
    const isKw = (w, o) => lw(peek(o)) === w;
    const isOp = (v, o) => { const t = peek(o); return t && t.k === 'op' && t.v === v; };
    const err = (msg) => { const t = peek(); throw new VBAError(1002, 'Syntax error on line ' + (t ? t.line : '?') + ': ' + msg); };
    const skipEol = () => { while (peek().k === 'eol' || isOp(':')) p++; };
    const expectKw = (w) => { if (!isKw(w)) err('expected ' + w); p++; };
    const atEnd = () => peek().k === 'eol' || peek().k === 'eof' || isOp(':');

    function parseTypeName() { // As [New] Type[.Sub]
      let isNew = false; if (isKw('new')) { p++; isNew = true; }
      let name = next().v; while (isOp('.')) { p++; name += '.' + next().v; }
      if (isOp('*')) { p++; next(); } // String * N
      return { name, isNew };
    }
    function parseDimDecls(scope) {
      // after Dim/Private/Public/Static: name[(dims)] [As type] , ...
      const out = [];
      while (true) {
        if (isKw('withevents')) p++;
        const name = next().v; let dims = null;
        if (isOp('(')) { p++; dims = []; if (!isOp(')')) { while (true) { let a = parseExpr(); let lo = null, hi = a; if (isKw('to')) { p++; lo = a; hi = parseExpr(); } dims.push({ lo, hi }); if (isOp(',')) { p++; continue; } break; } } if (!isOp(')')) err('expected )'); p++; }
        let type = null, isNew = false; if (isKw('as')) { p++; const tn = parseTypeName(); type = tn.name; isNew = tn.isNew; }
        out.push({ t: 'dim', name, dims, type, isNew, scope });
        if (isOp(',')) { p++; continue; } break;
      }
      return out;
    }
    function parseParams() {
      const params = []; if (!isOp('(')) return params; p++;
      while (!isOp(')')) {
        const prm = { name: null, byRef: true, optional: false, def: undefined, paramArray: false, type: null };
        while (true) { if (isKw('optional')) { prm.optional = true; p++; } else if (isKw('byval')) { prm.byRef = false; p++; } else if (isKw('byref')) { prm.byRef = true; p++; } else if (isKw('paramarray')) { prm.paramArray = true; p++; } else break; }
        prm.name = next().v; if (isOp('(')) { p++; if (!isOp(')')) err('expected )'); p++; prm.isArray = true; }
        if (isKw('as')) { p++; prm.type = parseTypeName().name; }
        if (isOp('=')) { p++; prm.def = parseExpr(); }
        params.push(prm); if (isOp(',')) { p++; continue; } break;
      }
      if (!isOp(')')) err('expected )'); p++; return params;
    }
    // ---- expressions ----
    function parseExpr() { return parseImp(); }
    function parseImp() { let l = parseEqv(); while (isKw('imp')) { p++; l = { t: 'bin', o: 'imp', l, r: parseEqv() }; } return l; }
    function parseEqv() { let l = parseXor(); while (isKw('eqv')) { p++; l = { t: 'bin', o: 'eqv', l, r: parseXor() }; } return l; }
    function parseXor() { let l = parseOr(); while (isKw('xor')) { p++; l = { t: 'bin', o: 'xor', l, r: parseOr() }; } return l; }
    function parseOr() { let l = parseAnd(); while (isKw('or')) { p++; l = { t: 'bin', o: 'or', l, r: parseAnd() }; } return l; }
    function parseAnd() { let l = parseNot(); while (isKw('and')) { p++; l = { t: 'bin', o: 'and', l, r: parseNot() }; } return l; }
    function parseNot() { if (isKw('not')) { p++; return { t: 'not', e: parseNot() }; } return parseCmp(); }
    function parseCmp() {
      let l = parseConcat();
      while (true) {
        const t = peek(); let o = null;
        if (t.k === 'op' && ['=', '<>', '<', '>', '<=', '>='].includes(t.v)) { o = t.v; p++; }
        else if (isKw('like')) { o = 'like'; p++; } else if (isKw('is')) { o = 'is'; p++; } else break;
        l = { t: 'bin', o, l, r: parseConcat() };
      }
      return l;
    }
    function parseConcat() { let l = parseAdd(); while (isOp('&')) { p++; l = { t: 'bin', o: '&', l, r: parseAdd() }; } return l; }
    function parseAdd() { let l = parseMod(); while (isOp('+') || isOp('-')) { const o = next().v; l = { t: 'bin', o, l, r: parseMod() }; } return l; }
    function parseMod() { let l = parseIntDiv(); while (isKw('mod')) { p++; l = { t: 'bin', o: 'mod', l, r: parseIntDiv() }; } return l; }
    function parseIntDiv() { let l = parseMul(); while (isOp('\\')) { p++; l = { t: 'bin', o: '\\', l, r: parseMul() }; } return l; }
    function parseMul() { let l = parseUnary(); while (isOp('*') || isOp('/')) { const o = next().v; l = { t: 'bin', o, l, r: parseUnary() }; } return l; }
    function parseUnary() { if (isOp('-')) { p++; return { t: 'neg', e: parseUnary() }; } if (isOp('+')) { p++; return parseUnary(); } return parsePow(); }
    function parsePow() { let l = parsePrimary(); while (isOp('^')) { p++; let r = parseUnaryForPow(); l = { t: 'bin', o: '^', l, r }; } return l; }
    function parseUnaryForPow() { if (isOp('-')) { p++; return { t: 'neg', e: parseUnaryForPow() }; } return parsePrimary(); }
    function parseArgs() { // after '(' consumed; returns list of {e, name} until ')'
      const args = [];
      while (!isOp(')')) {
        if (isOp(',')) { p++; args.push({ e: null }); continue; }
        let name = null; if (peek().k === 'id' && isOp(':=', 1)) { name = next().v; p++; }
        const e = parseExpr(); args.push({ e, name });
        if (isOp(',')) { p++; if (isOp(')')) args.push({ e: null }); continue; }
        if (!isOp(')')) err('expected ) or ,');
      }
      p++; return args;
    }
    function parsePostfix(node) {
      while (true) {
        if (isOp('.')) { p++; const t = next(); if (t.k !== 'id') err('expected member name'); node = { t: 'mem', o: node, n: t.v, args: null }; if (isOp('(') ) { p++; node.args = parseArgs(); } continue; }
        if (isOp('!') && node.t !== 'num') { p++; const t = next(); node = { t: 'mem', o: node, n: 'Item', args: [{ e: { t: 'str', v: t.v } }] }; continue; }
        if (isOp('(') && (node.t === 'mem' || node.t === 'id') && node.args !== null && node.called) { p++; node = { t: 'call', f: node, args: parseArgs() }; continue; }
        break;
      }
      return node;
    }
    function parsePrimary() {
      const t = next();
      switch (t.k) {
        case 'num': return { t: 'num', v: t.v };
        case 'str': return { t: 'str', v: t.v };
        case 'date': return { t: 'dlit', v: t.v };
        case 'op':
          if (t.v === '(') { const e = parseExpr(); if (!isOp(')')) err('expected )'); p++; return parsePostfix({ t: 'paren', e }); }
          if (t.v === '.') { const m = next(); let node = { t: 'wmem', n: m.v, args: null }; if (isOp('(')) { p++; node.args = parseArgs(); } return parsePostfix(node); }
          err('unexpected ' + t.v); break;
        case 'id': {
          const w = t.v.toLowerCase();
          if (!t.bracket) {
            if (w === 'true') return { t: 'bool', v: true }; if (w === 'false') return { t: 'bool', v: false };
            if (w === 'nothing') return { t: 'nothing' }; if (w === 'empty') return { t: 'empty' }; if (w === 'null') return { t: 'null' };
            if (w === 'new') { const tn = parseTypeName(); return { t: 'new', type: tn.name }; }
            if (w === 'typeof') { const e = parseConcat(); expectKw('is'); const tn = parseTypeName(); return { t: 'typeof', e, type: tn.name }; }
          }
          let node = { t: 'id', n: t.v, args: null };
          if (isOp('(')) { p++; node.args = parseArgs(); }
          return parsePostfix(node);
        }
      }
      err('unexpected token ' + (t.v === undefined ? t.k : t.v)); return null;
    }
    // ---- statements ----
    function parseBlock(terminators) {
      const stmts = [];
      while (true) {
        skipEol(); const t = peek(); if (t.k === 'eof') break;
        const w = lw(t);
        if (w && terminators.some(tm => matchTerm(tm))) break;
        const s = parseStatement(); if (s) { if (Array.isArray(s)) stmts.push(...s); else stmts.push(s); }
      }
      return stmts;
    }
    function matchTerm(tm) { // tm is array of words e.g. ['end','if'] or ['else']
      for (let k = 0; k < tm.length; k++) if (lw(peek(k)) !== tm[k]) return false; return true;
    }
    function parseCondThen() { const c = parseExpr(); if (isKw('then')) p++; return c; }
    function parseSimpleStmtsUntilEol() { const out = []; while (!(peek().k === 'eol' || peek().k === 'eof') && !isKw('else')) { const s = parseStatement(true); if (s) out.push(...(Array.isArray(s) ? s : [s])); if (isOp(':')) p++; } return out; }
    function parseIf() {
      const cond = parseCondThen();
      if (peek().k !== 'eol' && peek().k !== 'eof') { // single-line if
        const thenS = parseSimpleStmtsUntilEol(); let elseS = null;
        if (isKw('else')) { p++; elseS = parseSimpleStmtsUntilEol(); }
        return { t: 'if', branches: [{ cond, body: thenS }], else: elseS };
      }
      const branches = []; let elseB = null; let body = parseBlock([['elseif'], ['else'], ['end', 'if']]); branches.push({ cond, body });
      while (true) {
        if (isKw('elseif')) { p++; const c2 = parseCondThen(); const b2 = parseBlock([['elseif'], ['else'], ['end', 'if']]); branches.push({ cond: c2, body: b2 }); continue; }
        if (isKw('else')) { p++; elseB = parseBlock([['end', 'if']]); }
        break;
      }
      expectKw('end'); expectKw('if'); return { t: 'if', branches, else: elseB };
    }
    function parseFor() {
      if (isKw('each')) { p++; const v = parsePrimary(); expectKw('in'); const coll = parseExpr(); const body = parseBlock([['next']]); expectKw('next'); if (!atEnd()) parseExpr(); return { t: 'foreach', v, coll, body }; }
      const v = parsePrimary(); if (!isOp('=')) err('expected ='); p++; const from = parseExpr(); expectKw('to'); const to = parseExpr(); let step = null; if (isKw('step')) { p++; step = parseExpr(); }
      const body = parseBlock([['next']]); expectKw('next'); while (!atEnd()) { parseExpr(); if (isOp(',')) p++; }
      return { t: 'for', v, from, to, step, body };
    }
    function parseDo() {
      let pre = null; if (isKw('while') || isKw('until')) { const kind = next().v.toLowerCase(); pre = { kind, cond: parseExpr() }; }
      const body = parseBlock([['loop']]); expectKw('loop'); let post = null; if (isKw('while') || isKw('until')) { const kind = next().v.toLowerCase(); post = { kind, cond: parseExpr() }; }
      return { t: 'do', pre, post, body };
    }
    function parseSelect() {
      expectKw('case'); const subj = parseExpr(); const cases = []; skipEol();
      while (isKw('case')) {
        p++; const conds = []; let isElse = false;
        if (isKw('else')) { p++; isElse = true; }
        else while (true) {
          if (isKw('is')) { p++; const op = next().v; conds.push({ k: 'is', op, e: parseExpr() }); }
          else { const a = parseExpr(); if (isKw('to')) { p++; conds.push({ k: 'range', a, b: parseExpr() }); } else conds.push({ k: 'eq', e: a }); }
          if (isOp(',')) { p++; continue; } break;
        }
        const body = parseBlock([['case'], ['end', 'select']]); cases.push({ conds, isElse, body });
      }
      expectKw('end'); expectKw('select'); return { t: 'select', subj, cases };
    }
    function parseWith() { const e = parseExpr(); const body = parseBlock([['end', 'with']]); expectKw('end'); expectKw('with'); return { t: 'with', e, body }; }
    function parseStatement(inline) {
      const t = peek(); const line = t.line;
      if (t.k !== 'id') { if (isOp('.')) return parseCallOrAssign(line); err('unexpected ' + (t.v === undefined ? t.k : t.v)); }
      const w = t.v.toLowerCase();
      if (!t.bracket) switch (w) {
        case 'dim': case 'static': { p++; return parseDimDecls('local'); }
        case 'private': case 'public': case 'global': { // could be Sub/Function/Const/Dim at module level
          p++; if (isKw('sub') || isKw('function') || isKw('property') || isKw('declare')) return parseProc(); if (isKw('const')) { p++; return parseConst(); } if (isKw('enum')) { p++; return parseEnum(); } if (isKw('type')) { p++; skipType(); return null; } return parseDimDecls('module'); }
        case 'const': p++; return parseConst();
        case 'redim': { p++; let preserve = false; if (isKw('preserve')) { p++; preserve = true; } const out = []; while (true) { const name = next().v; p++; const dims = []; while (true) { const a = parseExpr(); let lo = null, hi = a; if (isKw('to')) { p++; lo = a; hi = parseExpr(); } dims.push({ lo, hi }); if (isOp(',')) { p++; continue; } break; } p++; if (isKw('as')) { p++; parseTypeName(); } out.push({ t: 'redim', name, dims, preserve }); if (isOp(',')) { p++; continue; } break; } return out; }
        case 'erase': { p++; const n = next().v; return { t: 'erase', name: n }; }
        case 'if': p++; return { ...parseIf(), line };
        case 'for': p++; return { ...parseFor(), line };
        case 'do': p++; return { ...parseDo(), line };
        case 'while': { p++; const cond = parseExpr(); const body = parseBlock([['wend']]); expectKw('wend'); return { t: 'do', pre: { kind: 'while', cond }, post: null, body }; }
        case 'select': p++; return { ...parseSelect(), line };
        case 'with': p++; return { ...parseWith(), line };
        case 'set': { p++; const target = parsePrimaryChain(); if (!isOp('=')) err('expected ='); p++; const e = parseExpr(); return { t: 'assign', target, e, set: true, line }; }
        case 'let': { p++; return parseCallOrAssign(line); }
        case 'call': { p++; const node = parsePrimaryChain(); return { t: 'exprstmt', e: node, line }; }
        case 'exit': { p++; const k = next().v.toLowerCase(); return { t: 'exit', kind: k }; }
        case 'goto': { p++; return { t: 'goto', label: next().v }; }
        case 'gosub': case 'return': err('GoSub/Return is not supported'); break;
        case 'resume': { p++; if (isKw('next')) { p++; return { t: 'resume', next: true }; } if (atEnd()) return { t: 'resume' }; return { t: 'goto', label: next().v }; }
        case 'on': {
          p++; if (isKw('error')) { p++; if (isKw('resume')) { p++; expectKw('next'); return { t: 'onerror', mode: 'next' }; } expectKw('goto'); const tk = next(); if (tk.k === 'num' && tk.v === 0) return { t: 'onerror', mode: 'off' }; if (tk.k === 'num' && tk.v === -1) return { t: 'onerror', mode: 'off' }; return { t: 'onerror', mode: 'label', label: tk.v }; }
          err('unsupported On statement'); break; }
        case 'option': { while (!atEnd()) p++; return null; }
        case 'attribute': { while (!atEnd()) p++; return null; }
        case 'randomize': { p++; if (!atEnd()) parseExpr(); return null; }
        case 'stop': case 'end': { if (w === 'end' && (lw(peek(1)) === 'sub' || lw(peek(1)) === 'function' || lw(peek(1)) === 'if')) break; p++; return { t: 'end' }; }
        case 'unload': { p++; const e = parseExpr(); return { t: 'exprstmt', e: { t: 'id', n: '__Unload', args: [{ e }] }, line }; }
        case 'load': { p++; parseExpr(); return null; }
        case 'debug': break;
        case 'sub': case 'function': case 'property': return parseProc();
        case 'enum': p++; return parseEnum();
        case 'type': p++; skipType(); return null;
        case 'declare': { while (!atEnd()) p++; return null; }
      }
      // label?
      if (isOp(':', 1) && (toks[p + 2].k === 'eol' || toks[p + 2].k === 'eof') && !inline) { const name = next().v; p++; return { t: 'label', name }; }
      return parseCallOrAssign(line);
    }
    function skipType() { while (!(isKw('end') && lw(peek(1)) === 'type')) { if (peek().k === 'eof') err('Type without End Type'); p++; } p += 2; }
    function parseConst() { const out = []; while (true) { const name = next().v; if (isKw('as')) { p++; parseTypeName(); } if (!isOp('=')) err('expected ='); p++; const e = parseExpr(); out.push({ t: 'const', name, e }); if (isOp(',')) { p++; continue; } break; } return out; }
    function parseEnum() { const name = next().v; const members = []; skipEol(); let cur = 0; while (!(isKw('end') && lw(peek(1)) === 'enum')) { const n = next().v; let e = null; if (isOp('=')) { p++; e = parseExpr(); } members.push({ name: n, e }); skipEol(); } p += 2; enums.push({ name, members }); return null; }
    function parsePrimaryChain() { return parsePrimary(); }
    function parseCallOrAssign(line) {
      const target = parsePrimaryChain();
      if (isOp('=') ) { p++; const e = parseExpr(); return { t: 'assign', target, e, line }; }
      // command call without parens:  Foo a, b   /  obj.Method a, b
      if (!atEnd() && !isKw('else')) {
        const args = []; // parse comma separated exprs (with named args)
        while (true) {
          if (isOp(',')) { p++; args.push({ e: null }); continue; }
          let name = null; if (peek().k === 'id' && isOp(':=', 1)) { name = next().v; p++; }
          args.push({ e: parseExpr(), name }); if (isOp(',')) { p++; if (atEnd()) args.push({ e: null }); continue; } break;
        }
        const call = { ...target, args: (target.args || []).concat(args), called: true };
        return { t: 'exprstmt', e: call, line };
      }
      return { t: 'exprstmt', e: target, line };
    }
    function parseProc() {
      const kind = next().v.toLowerCase(); // sub | function | property
      let propKind = null; if (kind === 'property') { propKind = next().v.toLowerCase(); }
      const nameTok = next(); const params = parseParams(); let retType = null; if (isKw('as')) { p++; retType = parseTypeName().name; }
      const endW = kind === 'property' ? 'property' : kind;
      const body = parseBlock([['end', endW]]); expectKw('end'); p++;
      const proc = { name: nameTok.v, kind: kind === 'property' ? 'property' : kind, params, body, retType, propKind };
      const key = (propKind ? propKind + ':' : '') + nameTok.v.toLowerCase(); procs[key] = proc; return null;
    }
    // top-level program
    const topStmts = [];
    while (true) {
      skipEol(); if (peek().k === 'eof') break;
      const s = parseStatement(); if (s) { const arr = Array.isArray(s) ? s : [s]; for (const x of arr) { if (x.t === 'dim' || x.t === 'const') globals.push(x); else topStmts.push(x); } }
    }
    return { procs, globals, enums };
  }

  // ---------------- runtime ----------------
  class Scope { constructor(parent) { this.vars = Object.create(null); this.parent = parent; } }

  class Runtime {
    constructor(host) {
      this.host = host || {}; this.procs = {}; this.globals = Object.create(null); this.consts = Object.create(null); this.out = [];
      this.errObj = this.makeErr(); this.callDepth = 0; this.withStack = []; this.steps = 0; this.maxSteps = 5e7; this.state = null; this.stopFlag = false;
      this.installConstants();
    }
    makeErr() { const self = this; return { Number: 0, Description: '', Source: '', vbGet(name, args) { const n = String(name).toLowerCase(); if (n === 'number') return this.Number; if (n === 'description' || n === '') return this.Description; if (n === 'source') return this.Source; if (n === 'clear') { this.Number = 0; this.Description = ''; return null; } if (n === 'raise') { throw new VBAError(toNumber(args[0]), args.length > 2 ? toStr(args[2]) : 'Application-defined or object-defined error', args.length > 1 ? toStr(args[1]) : ''); } throw new VBAError(438, 'Object doesn\'t support this property or method'); }, vbSet(name, args, v) { const n = String(name).toLowerCase(); if (n === 'number') this.Number = toNumber(v); else if (n === 'description') this.Description = toStr(v); else throw new VBAError(438, 'Object doesn\'t support this property or method'); } }; }
    installConstants() {
      const c = this.consts; const set = (o) => { for (const k in o) c[k.toLowerCase()] = o[k]; };
      set({ vbCrLf: '\r\n', vbCr: '\r', vbLf: '\n', vbNewLine: '\n', vbTab: '\t', vbNullString: '', vbNullChar: '\0', vbOKOnly: 0, vbOKCancel: 1, vbAbortRetryIgnore: 2, vbYesNoCancel: 3, vbYesNo: 4, vbRetryCancel: 5, vbCritical: 16, vbQuestion: 32, vbExclamation: 48, vbInformation: 64, vbDefaultButton1: 0, vbDefaultButton2: 256, vbOK: 1, vbCancel: 2, vbAbort: 3, vbRetry: 4, vbIgnore: 5, vbYes: 6, vbNo: 7, vbBinaryCompare: 0, vbTextCompare: 1, vbEmpty: 0, vbNull: 1, vbInteger: 2, vbLong: 3, vbSingle: 4, vbDouble: 5, vbCurrency: 6, vbDate: 7, vbString: 8, vbObject: 9, vbBoolean: 11, vbVariant: 12, vbArray: 8192, vbSunday: 1, vbMonday: 2, vbUseSystemDayOfWeek: 0, vbGeneralDate: 0, vbLongDate: 1, vbShortDate: 2, vbUpperCase: 1, vbLowerCase: 2, vbProperCase: 3, pi: Math.PI,
        xlUp: -4162, xlDown: -4121, xlToLeft: -4159, xlToRight: -4161, xlLeft: -4131, xlRight: -4152, xlCenter: -4108, xlTop: -4160, xlBottom: -4107, xlAscending: 1, xlDescending: 2, xlYes: 1, xlNo: 2, xlGuess: 0, xlValues: -4163, xlFormulas: -4123, xlWhole: 1, xlPart: 2, xlByRows: 1, xlByColumns: 2, xlCalculationAutomatic: -4105, xlCalculationManual: -4135, xlCellTypeLastCell: 11, xlCellTypeBlanks: 4, xlCellTypeConstants: 2, xlCellTypeFormulas: -4123, xlContinuous: 1, xlThin: 2, xlMedium: -4138, xlThick: 4, xlNone: -4142, xlAutomatic: -4105, xlShiftDown: -4121, xlShiftToRight: -4161, xlShiftUp: -4162, xlShiftToLeft: -4159, xlWorksheet: -4167, xlSheetVisible: -1, xlSheetHidden: 0, xlPasteAll: -4104, xlPasteValues: -4163, xlPasteFormats: -4122, xlPasteFormulas: -4123, xlEdgeBottom: 9, xlEdgeTop: 8, xlEdgeLeft: 7, xlEdgeRight: 10, xlInsideHorizontal: 12, xlInsideVertical: 11,
        vbBlack: 0, vbRed: 255, vbGreen: 65280, vbBlue: 16711680, vbYellow: 65535, vbMagenta: 16711935, vbCyan: 16776960, vbWhite: 16777215 });
    }
    load(src, opts) { // parse & register
      const prog = parseProgram(src);
      for (const k in prog.procs) { const p = prog.procs[k]; p.module = opts && opts.module; this.procs[k] = p; }
      for (const en of prog.enums) { let cur = 0; for (const m of en.members) { if (m.e) cur = toNumber(this.evalConstExpr(m.e)); this.consts[m.name.toLowerCase()] = cur; this.consts[(en.name + '.' + m.name).toLowerCase()] = cur; cur++; } }
      this.pendingGlobals = (this.pendingGlobals || []).concat(prog.globals);
      return Object.keys(prog.procs);
    }
    evalConstExpr(node) { const scope = new Scope(null); return this.eval(node, { scope, proc: null }); }
    initGlobals() { const gl = this.pendingGlobals || []; this.pendingGlobals = []; const ctx = { scope: new Scope(null), proc: null }; for (const g of gl) { if (g.t === 'const') { this.consts[g.name.toLowerCase()] = this.eval(g.e, ctx); } else this.declare(g, ctx, this.globals); } }
    hasProc(name) { return !!this.procs[name.toLowerCase()]; }
    listProcs() { return Object.values(this.procs).filter(p => p.kind === 'sub' && !p.propKind).map(p => ({ name: p.name, params: p.params.length })); }
    print(s) { this.out.push(s); if (this.host.print) this.host.print(s); }
    // Run a named procedure with JS arg values; returns result
    run(name, args) {
      this.initGlobals(); const key = name.toLowerCase(); const proc = this.procs[key]; if (!proc) throw new VBAError(35, 'Sub or Function not defined: ' + name);
      this.steps = 0; this.stopFlag = false; this.errObj.Number = 0; this.errObj.Description = '';
      try { return this.callProc(proc, (args || []).map(a => ({ v: a })), null); }
      catch (e) { if (e instanceof ExitSignal && e.kind === 'end') return null; throw e; }
    }
    callProc(proc, argRefs, argNodesCtx) {
      if (this.callDepth > 200) throw new VBAError(28, 'Out of stack space');
      const scope = new Scope(null); const params = proc.params;
      for (let i = 0; i < params.length; i++) {
        const prm = params[i]; const nm = prm.name.toLowerCase();
        if (prm.paramArray) { const rest = argRefs.slice(i).map(r => r.v); scope.vars[nm] = { v: VArr.from1D(rest, 0), type: null }; break; }
        let ref = argRefs[i];
        if (ref === undefined || ref.missing) {
          if (prm.def !== undefined) scope.vars[nm] = { v: this.eval(prm.def, { scope, proc }), type: prm.type };
          else if (prm.optional) scope.vars[nm] = { v: defaultFor(prm.type) === null ? null : defaultFor(prm.type), type: prm.type, missing: true };
          else throw new VBAError(449, 'Argument not optional');
        } else if (prm.byRef && ref.cell) scope.vars[nm] = ref.cell;
        else scope.vars[nm] = { v: prm.type && !isObj(ref.v) && !(ref.v instanceof VArr) ? coerceType(prm.type, ref.v) : ref.v, type: prm.type };
      }
      const retName = proc.name.toLowerCase();
      if (proc.kind === 'function' || proc.propKind === 'get') scope.vars[retName] = { v: proc.retType ? defaultFor(proc.retType) : null, type: proc.retType, isRet: true };
      const ctx = { scope, proc, withs: [] }; const saveState = this.state; this.state = { mode: null, label: null }; this.callDepth++;
      try { this.execBlock(proc.body, ctx); }
      catch (e) { if (e instanceof ExitSignal && (e.kind === 'sub' || e.kind === 'function' || e.kind === 'property')) { /* ok */ } else throw e; }
      finally { this.callDepth--; this.state = saveState; }
      return (proc.kind === 'function' || proc.propKind === 'get') ? scope.vars[retName].v : null;
    }
    // ---- variables ----
    lookup(name, ctx) { const k = name.toLowerCase(); let s = ctx.scope; const v = s.vars[k]; if (v !== undefined) return v; const g = this.globals[k]; if (g !== undefined) return g; return undefined; }
    declare(d, ctx, store) {
      const k = d.name.toLowerCase(); const target = store || ctx.scope.vars;
      if (d.dims) {
        const dims = d.dims.map(x => ({ lo: x.lo ? toNumber(this.eval(x.lo, ctx)) : 0, hi: toNumber(this.eval(x.hi, ctx)) }));
        target[k] = { v: new VArr(dims, defaultFor(d.type)), type: d.type, isArr: true };
      } else if (d.isNew) { target[k] = { v: this.createObject(d.type), type: d.type }; }
      else if (d.dims === null && d.type && /\(\)$/.test(d.type)) target[k] = { v: null, type: null };
      else target[k] = { v: defaultFor(d.type), type: d.type };
    }
    setVar(name, value, ctx, isSet) {
      let ref = this.lookup(name, ctx);
      if (!ref) { ref = { v: null, type: null }; ctx.scope.vars[name.toLowerCase()] = ref; }
      if (ref.type && !isSet && !isObj(value) && !(value instanceof VArr)) value = coerceType(ref.type, value);
      if (ref.isArr && value instanceof VArr === false && !ref.type) { }
      ref.v = value;
    }
    createObject(type) {
      const t = String(type).toLowerCase().replace(/^vba\./, '');
      if (t === 'collection') return new VColl(); if (t === 'scripting.dictionary' || t === 'dictionary') return new VDict();
      if (this.host.createObject) { const o = this.host.createObject(type); if (o) return o; }
      throw new VBAError(429, 'ActiveX component can\'t create object: ' + type);
    }
    // ---- execution ----
    tick() { if (++this.steps > this.maxSteps) throw new VBAError(1000, 'Macro stopped: too many steps (possible infinite loop)'); if (this.stopFlag) throw new VBAError(18, 'User interrupt occurred'); }
    execBlock(stmts, ctx) {
      let i = 0; const n = stmts.length;
      while (i < n) {
        const s = stmts[i];
        try { this.execStmt(s, ctx); i++; }
        catch (e) {
          if (e instanceof GotoSignal) { const li = stmts.findIndex(x => x.t === 'label' && x.name.toLowerCase() === e.label.toLowerCase()); if (li >= 0) { i = li + 1; continue; } throw e; }
          if (e instanceof ExitSignal) throw e;
          if (e instanceof VBAError || e instanceof RangeError || e instanceof TypeError) {
            const ve = e instanceof VBAError ? e : new VBAError(e instanceof RangeError ? 28 : 5, e.message);
            if (ve.number === 1000) throw ve;
            this.errObj.Number = ve.number; this.errObj.Description = ve.description; this.errObj.Source = ve.source;
            const st = this.state || {};
            if (st.mode === 'next') { i++; continue; }
            if (st.mode === 'label') { throw new GotoSignal(st.label); }
            throw ve;
          }
          throw e;
        }
      }
    }
    execStmt(s, ctx) {
      this.tick();
      switch (s.t) {
        case 'dim': this.declare(s, ctx); return;
        case 'const': ctx.scope.vars[s.name.toLowerCase()] = { v: this.eval(s.e, ctx), type: null, isConst: true }; return;
        case 'redim': {
          const ref = this.lookup(s.name, ctx); const dims = s.dims.map(x => ({ lo: x.lo ? toNumber(this.eval(x.lo, ctx)) : 0, hi: toNumber(this.eval(x.hi, ctx)) }));
          const na = new VArr(dims, defaultFor(ref && ref.type)); if (s.preserve && ref && ref.v instanceof VArr) { const old = ref.v; const copy = (ix, k) => { if (k === old.dims.length) { try { na.set(ix, old.get(ix)); } catch (e) { } return; } for (let q = old.dims[k].lo; q <= old.dims[k].hi; q++) copy(ix.concat([q]), k + 1); }; copy([], 0); }
          if (!ref) ctx.scope.vars[s.name.toLowerCase()] = { v: na, type: null, isArr: true }; else ref.v = na; return;
        }
        case 'erase': { const ref = this.lookup(s.name, ctx); if (ref && ref.v instanceof VArr) ref.v = new VArr(ref.v.dims, defaultFor(ref.type)); return; }
        case 'label': return;
        case 'assign': this.assign(s, ctx); return;
        case 'exprstmt': { this.evalCall(s.e, ctx, true); return; }
        case 'if': { for (const b of s.branches) { if (toBoolean(this.eval(b.cond, ctx))) { this.execBlock(b.body, ctx); return; } } if (s.else) this.execBlock(s.else, ctx); return; }
        case 'for': this.execFor(s, ctx); return;
        case 'foreach': this.execForEach(s, ctx); return;
        case 'do': this.execDo(s, ctx); return;
        case 'select': this.execSelect(s, ctx); return;
        case 'with': { const o = this.eval(s.e, ctx); ctx.withs.push(o); try { this.execBlock(s.body, ctx); } finally { ctx.withs.pop(); } return; }
        case 'exit': throw new ExitSignal(s.kind);
        case 'goto': throw new GotoSignal(s.label);
        case 'onerror': { if (s.mode === 'off') this.state = { mode: null }; else if (s.mode === 'next') this.state = { mode: 'next' }; else this.state = { mode: 'label', label: s.label }; return; }
        case 'resume': return;
        case 'end': throw new ExitSignal('end');
      }
      throw new VBAError(1003, 'Unsupported statement: ' + s.t);
    }
    execFor(s, ctx) {
      const from = toNumber(this.eval(s.from, ctx)), to = toNumber(this.eval(s.to, ctx)), step = s.step ? toNumber(this.eval(s.step, ctx)) : 1;
      if (step === 0) throw new VBAError(5, 'Invalid procedure call or argument');
      try { for (let i = from; step > 0 ? i <= to + 1e-12 : i >= to - 1e-12; i += step) { this.assignTo(s.v, i, ctx); try { this.execBlock(s.body, ctx); } catch (e) { if (e instanceof ExitSignal && e.kind === 'for') throw e; throw e; } } }
      catch (e) { if (e instanceof ExitSignal && e.kind === 'for') return; throw e; }
    }
    iterate(coll) {
      if (coll instanceof VArr) return coll.data.slice(); if (coll && typeof coll.vbEnum === 'function') return coll.vbEnum();
      throw new VBAError(438, 'Object doesn\'t support this property or method');
    }
    execForEach(s, ctx) { const items = this.iterate(this.eval(s.coll, ctx)); try { for (const it of items) { this.assignTo(s.v, it, ctx, true); this.execBlock(s.body, ctx); } } catch (e) { if (e instanceof ExitSignal && e.kind === 'for') return; throw e; } }
    execDo(s, ctx) {
      try {
        while (true) {
          if (s.pre) { const c = toBoolean(this.eval(s.pre.cond, ctx)); if (s.pre.kind === 'while' ? !c : c) break; }
          this.execBlock(s.body, ctx); this.tick();
          if (s.post) { const c = toBoolean(this.eval(s.post.cond, ctx)); if (s.post.kind === 'while' ? !c : c) break; }
        }
      } catch (e) { if (e instanceof ExitSignal && e.kind === 'do') return; throw e; }
    }
    execSelect(s, ctx) {
      const v = this.eval(s.subj, ctx);
      for (const c of s.cases) {
        let hit = c.isElse;
        if (!hit) for (const cd of c.conds) {
          if (cd.k === 'eq') { if (cmp(v, this.eval(cd.e, ctx), false) === 0) { hit = true; break; } }
          else if (cd.k === 'range') { if (cmp(v, this.eval(cd.a, ctx)) >= 0 && cmp(v, this.eval(cd.b, ctx)) <= 0) { hit = true; break; } }
          else { const r = cmp(v, this.eval(cd.e, ctx)); const ok = { '=': r === 0, '<>': r !== 0, '<': r < 0, '>': r > 0, '<=': r <= 0, '>=': r >= 0 }[cd.op]; if (ok) { hit = true; break; } }
        }
        if (hit) { this.execBlock(c.body, ctx); return; }
      }
    }
    assign(s, ctx) { const v = this.eval(s.e, ctx); this.assignTo(s.target, v, ctx, s.set); }
    assignTo(target, value, ctx, isSet) {
      if (target.t === 'id') {
        if (target.args === null || target.args === undefined) { this.setVar(target.n, value, ctx, isSet); return; }
        const ref = this.lookup(target.n, ctx);
        if (ref && ref.v instanceof VArr) { ref.v.set(target.args.map(a => Math.round(toNumber(this.eval(a.e, ctx)))), value); return; }
        if (ref && isObj(ref.v)) { ref.v.vbSet('', target.args.map(a => this.eval(a.e, ctx)), value); return; }
        // procedure-call target e.g. Cells(1,1) = 5, or function name assignment with parens
        const obj = this.evalCall(target, ctx); if (isObj(obj)) { obj.vbSet('Value', [], value); return; }
        throw new VBAError(451, 'Property let procedure not defined and property get procedure did not return an object');
      }
      if (target.t === 'mem' || target.t === 'wmem') {
        const obj = target.t === 'wmem' ? this.withObj(ctx) : this.eval(target.o, ctx);
        const args = target.args ? this.evalArgs(target.args, ctx) : [];
        if (obj === NOTHING) throw new VBAError(91, 'Object variable or With block variable not set');
        if (isObj(obj)) { obj.vbSet(target.n, args, value, isSet); return; }
        throw new VBAError(424, 'Object required');
      }
      if (target.t === 'call') { const base = this.eval(target.f, ctx); if (base instanceof VArr) { base.set(target.args.map(a => Math.round(toNumber(this.eval(a.e, ctx)))), value); return; } if (isObj(base)) { base.vbSet('', target.args.map(a => this.eval(a.e, ctx)), value); return; } }
      throw new VBAError(1004, 'Invalid assignment target');
    }
    withObj(ctx) { const w = ctx.withs; if (!w || !w.length) throw new VBAError(1005, 'Invalid use of . outside With'); return w[w.length - 1]; }
    evalArgs(args, ctx) { const out = []; let named = null; for (const a of args) { if (a.e === null) { out.push(undefined); continue; } const v = this.eval(a.e, ctx); if (a.name) { named = named || {}; named[a.name.toLowerCase()] = v; out.push(v); } else out.push(v); } if (named) out.named = named; return out; }
    // evaluate a call-ish node; asStmt => result discarded
    evalCall(node, ctx, asStmt) {
      if (node.t === 'id') {
        const ref = this.lookup(node.n, ctx);
        if (ref !== undefined && !(ref.isRet && node.args !== null && node.args !== undefined && this.procs[node.n.toLowerCase()])) {
          if (node.args === null || node.args === undefined) return ref.v;
          const v = ref.v; const a = this.evalArgs(node.args, ctx);
          if (v instanceof VArr) { if (a.length === 0) return v; return v.get(a.map(x => Math.round(toNumber(x)))); }
          if (isObj(v)) return v.vbGet('', a);
          throw new VBAError(13, 'Type mismatch');
        }
        const key = node.n.toLowerCase(); const proc = this.procs[key];
        if (proc) return this.callUser(proc, node.args || [], ctx);
        const bf = BUILTINS[key]; if (bf) return bf.call(this, this.evalArgs(node.args || [], ctx), ctx, node);
        if (key in this.consts) return this.consts[key];
        if (this.host.global) { const a = node.args ? this.evalArgs(node.args, ctx) : []; const r = this.host.global(key, a, node.args !== null && node.args !== undefined, ctx); if (r !== undefined) return r; }
        throw new VBAError(35, 'Sub or Function not defined: ' + node.n);
      }
      if (node.t === 'mem' || node.t === 'wmem') {
        let obj;
        if (node.t === 'wmem') obj = this.withObj(ctx);
        else {
          // enum member e.g. Color.Red  / vba constant "vbCrLf" etc
          if (node.o.t === 'id' && !node.o.args) { const ek = (node.o.n + '.' + node.n).toLowerCase(); if (ek in this.consts && !this.lookup(node.o.n, ctx)) return this.consts[ek]; }
          obj = this.eval(node.o, ctx);
        }
        const args = node.args ? this.evalArgs(node.args, ctx) : [];
        if (obj === NOTHING || obj === null) throw new VBAError(91, 'Object variable or With block variable not set');
        if (isObj(obj)) return obj.vbGet(node.n, args, node.args !== null);
        if (obj instanceof VArr || typeof obj !== 'object') throw new VBAError(424, 'Object required');
        throw new VBAError(438, 'Object doesn\'t support this property or method');
      }
      if (node.t === 'call') { const f = this.evalCall(node.f, ctx); const a = this.evalArgs(node.args, ctx); if (f instanceof VArr) return f.get(a.map(x => Math.round(toNumber(x)))); if (isObj(f)) return f.vbGet('', a); throw new VBAError(13, 'Type mismatch'); }
      return this.eval(node, ctx);
    }
    callUser(proc, argNodes, ctx) {
      const refs = argNodes.map((a, i) => {
        if (a.e === null) return { missing: true };
        const prm = proc.params[i]; if (a.name) { /* named arg: place by name */ }
        if (a.e.t === 'id' && !a.e.args) { const cell = this.lookup(a.e.n, ctx); if (cell) return { v: cell.v, cell }; }
        return { v: this.eval(a.e, ctx) };
      });
      // named args
      const named = argNodes.filter(a => a.name);
      if (named.length) { const ordered = []; argNodes.forEach((a, i) => { if (!a.name) ordered[i] = refs[i]; }); for (const a of named) { const pi = proc.params.findIndex(p => p.name.toLowerCase() === a.name.toLowerCase()); if (pi < 0) throw new VBAError(448, 'Named argument not found'); ordered[pi] = { v: this.eval(a.e, ctx) }; } for (let i = 0; i < ordered.length; i++) if (!ordered[i]) ordered[i] = { missing: true }; return this.callProc(proc, ordered, ctx); }
      return this.callProc(proc, refs, ctx);
    }
    eval(n, ctx) {
      switch (n.t) {
        case 'num': return n.v; case 'str': return n.v; case 'bool': return n.v;
        case 'nothing': return NOTHING; case 'empty': return null; case 'null': return NULLV;
        case 'dlit': { const d = Engine.FN.DATEVALUE([n.v], { wb: null, sheet: '' }); let tv = 0; const tm = /(\d{1,2}):(\d{2})(?::(\d{2}))?\s*(AM|PM)?/i.exec(n.v); if (tm) { let h = +tm[1]; if (tm[4] && /pm/i.test(tm[4]) && h < 12) h += 12; tv = (h * 3600 + +tm[2] * 60 + (+tm[3] || 0)) / 86400; } let ds = 0; try { ds = /\d+[\/-]\d+[\/-]\d+/.test(n.v) ? Engine.FN.DATEVALUE([n.v.replace(/\s.*$/, '')], {}) : 0; } catch (e) { ds = 0; } return new VDate(ds + tv); }
        case 'paren': return this.eval(n.e, ctx);
        case 'id': case 'mem': case 'wmem': case 'call': { const v = this.evalCall(n, ctx); return this.unwrapDefault(v); }
        case 'new': return this.createObject(n.type);
        case 'neg': { const v = this.eval(n.e, ctx); return v instanceof VDate ? new VDate(-v.s) : -toNumber(v); }
        case 'not': { const v = this.eval(n.e, ctx); if (typeof v === 'boolean') return !v; return ~Math.round(toNumber(v)); }
        case 'typeof': { const v = this.eval(n.e, ctx); return isObj(v) && String(v.vbType || '').toLowerCase() === n.type.toLowerCase(); }
        case 'bin': return this.binary(n, ctx);
      }
      throw new VBAError(1003, 'Unsupported expression ' + n.t);
    }
    unwrapDefault(v) { return v; }
    binary(n, ctx) {
      const o = n.o;
      if (o === 'and' || o === 'or') { // short-circuit not in VBA; evaluate both
        const a = this.eval(n.l, ctx), b = this.eval(n.r, ctx);
        if (typeof a === 'boolean' && typeof b === 'boolean') return o === 'and' ? a && b : a || b;
        const x = Math.round(toNumber(a)), y = Math.round(toNumber(b)); return o === 'and' ? x & y : x | y;
      }
      const a = this.eval(n.l, ctx), b = this.eval(n.r, ctx);
      switch (o) {
        case 'xor': if (typeof a === 'boolean' && typeof b === 'boolean') return a !== b; return Math.round(toNumber(a)) ^ Math.round(toNumber(b));
        case 'eqv': return typeof a === 'boolean' && typeof b === 'boolean' ? a === b : ~(Math.round(toNumber(a)) ^ Math.round(toNumber(b)));
        case 'imp': return typeof a === 'boolean' && typeof b === 'boolean' ? (!a || b) : (~Math.round(toNumber(a)) | Math.round(toNumber(b)));
        case '&': { if (a === NULLV && b === NULLV) return NULLV; return toStr(a === NULLV ? '' : a) + toStr(b === NULLV ? '' : b); }
        case '+': { if (typeof a === 'string' && typeof b === 'string') return a + b; if (a instanceof VDate && typeof b !== 'string') return new VDate(a.s + toNumber(b)); if (b instanceof VDate && typeof a !== 'string') return new VDate(toNumber(a) + b.s); if ((typeof a === 'string' && !isNumericStr(a)) || (typeof b === 'string' && !isNumericStr(b))) { if (typeof a === 'string' && typeof b === 'string') return a + b; throw new VBAError(13, 'Type mismatch'); } return toNumber(a) + toNumber(b); }
        case '-': { if (a instanceof VDate && b instanceof VDate) return a.s - b.s; if (a instanceof VDate) return new VDate(a.s - toNumber(b)); return toNumber(a) - toNumber(b); }
        case '*': return toNumber(a) * toNumber(b);
        case '/': { const d = toNumber(b); if (d === 0) throw new VBAError(11, 'Division by zero'); return toNumber(a) / d; }
        case '\\': { const d = Math.round(toNumber(b)); if (d === 0) throw new VBAError(11, 'Division by zero'); return Math.trunc(Math.round(toNumber(a)) / d); }
        case 'mod': { const d = Math.round(toNumber(b)); if (d === 0) throw new VBAError(11, 'Division by zero'); const x = Math.round(toNumber(a)); return x % d; }
        case '^': return Math.pow(toNumber(a), toNumber(b));
        case '=': return cmp(a, b, this.textCompare) === 0; case '<>': return cmp(a, b, this.textCompare) !== 0;
        case '<': return cmp(a, b, this.textCompare) < 0; case '>': return cmp(a, b, this.textCompare) > 0; case '<=': return cmp(a, b, this.textCompare) <= 0; case '>=': return cmp(a, b, this.textCompare) >= 0;
        case 'like': return likeMatch(toStr(a), toStr(b), this.textCompare);
        case 'is': return a === b || (isObj(a) && isObj(b) && (a === b || (a.vbIs && a.vbIs(b))));
      }
      throw new VBAError(1003, 'Unsupported operator ' + o);
    }
  }

  // ---------------- Collection / Dictionary ----------------
  class VColl {
    constructor() { this.items = []; this.keys = []; this.vbType = 'Collection'; }
    vbEnum() { return this.items.slice(); }
    vbGet(name, args) { const n = String(name).toLowerCase(); switch (n) {
      case 'add': { const v = args[0]; const key = args[1] !== undefined && args[1] !== null ? toStr(args[1]) : null; if (key !== null && this.keys.includes(key)) throw new VBAError(457, 'This key is already associated with an element of this collection'); let pos = this.items.length; if (args[2] !== undefined) pos = Math.round(toNumber(args[2])) - 1; else if (args[3] !== undefined) pos = Math.round(toNumber(args[3])); this.items.splice(pos, 0, v); this.keys.splice(pos, 0, key); return null; }
      case 'count': return this.items.length;
      case '': case 'item': { const k = args[0]; if (typeof k === 'string') { const i = this.keys.findIndex(x => x !== null && x.toLowerCase() === k.toLowerCase()); if (i < 0) throw new VBAError(5, 'Invalid procedure call or argument'); return this.items[i]; } const i = Math.round(toNumber(k)) - 1; if (i < 0 || i >= this.items.length) throw new VBAError(9, 'Subscript out of range'); return this.items[i]; }
      case 'remove': { const k = args[0]; let i; if (typeof k === 'string') i = this.keys.findIndex(x => x !== null && x.toLowerCase() === k.toLowerCase()); else i = Math.round(toNumber(k)) - 1; if (i < 0 || i >= this.items.length) throw new VBAError(9, 'Subscript out of range'); this.items.splice(i, 1); this.keys.splice(i, 1); return null; }
    } throw new VBAError(438, 'Object doesn\'t support this property or method'); }
    vbSet() { throw new VBAError(438, 'Object doesn\'t support this property or method'); }
  }
  class VDict {
    constructor() { this.map = new Map(); this.vbType = 'Dictionary'; this.mode = 0; }
    k(v) { const s = typeof v === 'number' ? 'n' + v : 's' + toStr(v); return this.mode === 1 && typeof v === 'string' ? 's' + v.toLowerCase() : s; }
    vbEnum() { return Array.from(this.map.values()).map(e => e.key); }
    vbGet(name, args) { const n = String(name).toLowerCase(); switch (n) {
      case 'add': { const k = this.k(args[0]); if (this.map.has(k)) throw new VBAError(457, 'This key is already associated with an element of this collection'); this.map.set(k, { key: args[0], v: args[1] }); return null; }
      case 'exists': return this.map.has(this.k(args[0]));
      case '': case 'item': { const e = this.map.get(this.k(args[0])); return e ? e.v : null; }
      case 'count': return this.map.size;
      case 'keys': return VArr.from1D(Array.from(this.map.values()).map(e => e.key), 0);
      case 'items': return VArr.from1D(Array.from(this.map.values()).map(e => e.v), 0);
      case 'remove': { if (!this.map.delete(this.k(args[0]))) throw new VBAError(32811, 'Element not found'); return null; }
      case 'removeall': this.map.clear(); return null;
      case 'comparemode': return this.mode;
    } throw new VBAError(438, 'Object doesn\'t support this property or method'); }
    vbSet(name, args, v) { const n = String(name).toLowerCase(); if (n === '' || n === 'item') { const k = this.k(args[0]); const e = this.map.get(k); if (e) e.v = v; else this.map.set(k, { key: args[0], v }); return; } if (n === 'comparemode') { this.mode = toNumber(v); return; } throw new VBAError(438, 'Object doesn\'t support this property or method'); }
  }

  // ---------------- builtins ----------------
  const B = {}; const BUILTINS = B;
  const defb = (names, fn) => { for (const n of names.split(' ')) B[n.toLowerCase()] = fn; };
  const S1 = f => function (a) { return f(toStr(a[0])); };
  defb('Len', a => { if (a[0] === null) return 0; return toStr(a[0]).length; });
  defb('Left', a => toStr(a[0]).slice(0, Math.max(0, toNumber(a[1])))); defb('Right', a => { const n = Math.max(0, toNumber(a[1])); return n === 0 ? '' : toStr(a[0]).slice(-n); });
  defb('Mid', a => { const s = toStr(a[0]), st = toNumber(a[1]); if (st < 1) throw new VBAError(5, 'Invalid procedure call or argument'); return a[2] === undefined ? s.slice(st - 1) : s.substr(st - 1, toNumber(a[2])); });
  defb('UCase', S1(s => s.toUpperCase())); defb('LCase', S1(s => s.toLowerCase()));
  defb('Trim', S1(s => s.replace(/^ +| +$/g, ''))); defb('LTrim', S1(s => s.replace(/^ +/, ''))); defb('RTrim', S1(s => s.replace(/ +$/, '')));
  defb('Space', a => ' '.repeat(toNumber(a[0]))); defb('String', a => { const c = typeof a[1] === 'string' ? a[1][0] : String.fromCharCode(toNumber(a[1])); return c.repeat(toNumber(a[0])); });
  defb('StrReverse', S1(s => s.split('').reverse().join('')));
  defb('InStr', function (a) { let st = 1, i = 0; if (typeof a[0] === 'number' && a.length >= 3) { st = a[0]; i = 1; } const s = toStr(a[i]), f = toStr(a[i + 1]); const tc = a[i + 2] === 1 || this.textCompare; const hay = tc ? s.toLowerCase() : s, nd = tc ? f.toLowerCase() : f; return hay.indexOf(nd, st - 1) + 1; });
  defb('InStrRev', a => { const s = toStr(a[0]), f = toStr(a[1]); const st = a[2] !== undefined ? toNumber(a[2]) : -1; const idx = st > 0 ? s.lastIndexOf(f, st - f.length) : s.lastIndexOf(f); return idx + 1; });
  defb('Replace', a => { const s = toStr(a[0]), f = toStr(a[1]), r = toStr(a[2]); if (f === '') return s; const start = a[3] !== undefined ? toNumber(a[3]) - 1 : 0; const cnt = a[4] !== undefined ? toNumber(a[4]) : -1; const tc = a[5] === 1; const head = s.slice(0, start); let rest = s.slice(start); if (cnt < 0 && !tc) return head + rest.split(f).join(r); let out = '', n = 0, i = 0; const hay = tc ? rest.toLowerCase() : rest, nd = tc ? f.toLowerCase() : f; while (i < rest.length) { if ((cnt < 0 || n < cnt) && hay.startsWith(nd, i)) { out += r; i += f.length; n++; } else out += rest[i++]; } return head + out; });
  defb('StrComp', a => { const x = toStr(a[0]), y = toStr(a[1]); const tc = a[2] === 1; const p = tc ? x.toLowerCase() : x, q = tc ? y.toLowerCase() : y; return p < q ? -1 : p > q ? 1 : 0; });
  defb('Chr ChrW', a => String.fromCharCode(toNumber(a[0]))); defb('Asc AscW', a => toStr(a[0]).charCodeAt(0));
  defb('Str', a => { const n = toNumber(a[0]); return (n >= 0 ? ' ' : '') + Engine.numStr(n); });
  defb('Val', a => { const m = /^\s*[-+]?(\d+\.?\d*|\.\d+)([eE][-+]?\d+)?/.exec(toStr(a[0]).replace(/\s/g, '')); return m ? parseFloat(m[0]) : 0; });
  defb('CStr', a => toStr(a[0])); defb('CInt CLng CByte CLngLng', a => bankers(toNumber(a[0]))); defb('CDbl CSng CCur CDec', a => toNumber(a[0])); defb('CBool', a => toBoolean(a[0])); defb('CVar', a => a[0]);
  defb('CDate', a => { const v = a[0]; if (v instanceof VDate) return v; if (typeof v === 'string') { const d = Engine.FN.DATEVALUE([v.replace(/\s+\d+:\d+.*$/, '')], {}); const tm = /(\d{1,2}):(\d{2})(?::(\d{2}))?/.exec(v); return new VDate(d + (tm ? (+tm[1] * 3600 + +tm[2] * 60 + (+tm[3] || 0)) / 86400 : 0)); } return new VDate(toNumber(v)); });
  defb('Abs', a => Math.abs(toNumber(a[0]))); defb('Sqr', a => { const x = toNumber(a[0]); if (x < 0) throw new VBAError(5, 'Invalid procedure call or argument'); return Math.sqrt(x); });
  defb('Int', a => Math.floor(toNumber(a[0]))); defb('Fix', a => Math.trunc(toNumber(a[0]))); defb('Sgn', a => Math.sign(toNumber(a[0])));
  defb('Round', a => vround(toNumber(a[0]), a[1] !== undefined ? toNumber(a[1]) : 0));
  defb('Exp', a => Math.exp(toNumber(a[0]))); defb('Log', a => { const x = toNumber(a[0]); if (x <= 0) throw new VBAError(5, 'Invalid procedure call or argument'); return Math.log(x); });
  defb('Sin', a => Math.sin(toNumber(a[0]))); defb('Cos', a => Math.cos(toNumber(a[0]))); defb('Tan', a => Math.tan(toNumber(a[0]))); defb('Atn', a => Math.atan(toNumber(a[0])));
  defb('Rnd', a => Math.random()); defb('Hex', a => Math.round(toNumber(a[0])).toString(16).toUpperCase()); defb('Oct', a => Math.round(toNumber(a[0])).toString(8));
  defb('Format', a => fmtVBA(a[0], a[1] === undefined ? 'General' : toStr(a[1])));
  defb('FormatNumber', a => { const d = a[1] !== undefined ? toNumber(a[1]) : 2; return Engine.fmtNumber(toNumber(a[0]), '#,##0' + (d > 0 ? '.' + '0'.repeat(d) : '')); });
  defb('FormatCurrency', a => { const d = a[1] !== undefined ? toNumber(a[1]) : 2; return Engine.fmtNumber(toNumber(a[0]), '$#,##0' + (d > 0 ? '.' + '0'.repeat(d) : '')); });
  defb('FormatPercent', a => { const d = a[1] !== undefined ? toNumber(a[1]) : 2; return Engine.fmtNumber(toNumber(a[0]), '0' + (d > 0 ? '.' + '0'.repeat(d) : '') + '%'); });
  defb('FormatDateTime', a => { const f = a[1] === undefined ? 0 : toNumber(a[1]); const fm = ['m/d/yyyy h:mm:ss AM/PM', 'dddd, mmmm d, yyyy', 'm/d/yyyy', 'h:mm:ss AM/PM', 'hh:mm'][f] || 'm/d/yyyy'; return Engine.fmtNumber(toNumber(a[0]), fm); });
  // dates
  const nowSerial = () => Engine.dateToSerial(new Date());
  defb('Now', () => new VDate(nowSerial())); defb('Date', () => new VDate(Math.floor(nowSerial()))); defb('Time', () => new VDate(nowSerial() % 1)); defb('Timer', () => { const d = new Date(); return d.getHours() * 3600 + d.getMinutes() * 60 + d.getSeconds() + d.getMilliseconds() / 1000; });
  const sd = v => Engine.serialToDate(toNumber(v));
  defb('Year', a => sd(a[0]).y); defb('Month', a => sd(a[0]).m); defb('Day', a => sd(a[0]).d); defb('Hour', a => sd(a[0]).h); defb('Minute', a => sd(a[0]).mi); defb('Second', a => sd(a[0]).s);
  defb('Weekday', a => { const wd = sd(a[0]).wd; const f = a[1] !== undefined ? toNumber(a[1]) : 1; return f === 2 ? (wd + 6) % 7 + 1 : wd + 1; });
  defb('WeekdayName', a => ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'][(toNumber(a[0]) - 1 + (a[2] === 2 ? 1 : 0)) % 7].slice(0, a[1] ? 3 : 99));
  defb('MonthName', a => ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'][toNumber(a[0]) - 1].slice(0, a[1] ? 3 : 99));
  defb('DateSerial', a => new VDate(Engine.FN.DATE([toNumber(a[0]), toNumber(a[1]), toNumber(a[2])], {}))); defb('TimeSerial', a => new VDate((toNumber(a[0]) * 3600 + toNumber(a[1]) * 60 + toNumber(a[2])) / 86400));
  defb('DateValue', a => new VDate(Math.floor(Engine.FN.DATEVALUE([toStr(a[0])], {})))); defb('IsDate', a => { if (a[0] instanceof VDate) return true; if (typeof a[0] !== 'string') return false; try { Engine.FN.DATEVALUE([a[0]], {}); return true; } catch (e) { return false; } });
  defb('DateAdd', a => { const iv = toStr(a[0]).toLowerCase(), n = toNumber(a[1]); const s = toNumber(a[2]); switch (iv) { case 'd': case 'y': case 'w': return new VDate(s + n); case 'ww': return new VDate(s + 7 * n); case 'h': return new VDate(s + n / 24); case 'n': return new VDate(s + n / 1440); case 's': return new VDate(s + n / 86400); case 'm': return new VDate(Engine.FN.EDATE([s, n], { wb: null }) + (s % 1)); case 'q': return new VDate(Engine.FN.EDATE([s, 3 * n], {}) + (s % 1)); case 'yyyy': return new VDate(Engine.FN.EDATE([s, 12 * n], {}) + (s % 1)); } throw new VBAError(5, 'Invalid procedure call or argument'); });
  defb('DateDiff', a => { const iv = toStr(a[0]).toLowerCase(); const x = toNumber(a[1]), y = toNumber(a[2]); const A = Engine.serialToDate(x), Bd = Engine.serialToDate(y); switch (iv) { case 'd': return Math.floor(y) - Math.floor(x); case 'ww': return Math.floor((Math.floor(y) - Math.floor(x)) / 7); case 'h': return Math.floor((y - x) * 24); case 'n': return Math.floor((y - x) * 1440); case 's': return Math.floor((y - x) * 86400); case 'm': return (Bd.y - A.y) * 12 + Bd.m - A.m; case 'q': return Math.floor(((Bd.y - A.y) * 12 + Bd.m - A.m) / 3); case 'yyyy': return Bd.y - A.y; } throw new VBAError(5, 'Invalid procedure call or argument'); });
  defb('DatePart', a => { const iv = toStr(a[0]).toLowerCase(); const d = Engine.serialToDate(toNumber(a[1])); switch (iv) { case 'yyyy': return d.y; case 'm': return d.m; case 'd': return d.d; case 'h': return d.h; case 'n': return d.mi; case 's': return d.s; case 'q': return Math.ceil(d.m / 3); case 'w': return d.wd + 1; } return 0; });
  // type tests
  defb('IsNumeric', a => typeof a[0] === 'number' || (typeof a[0] === 'string' && isNumericStr(a[0])) || typeof a[0] === 'boolean' && false); defb('IsEmpty', a => a[0] === null); defb('IsNull', a => a[0] === NULLV); defb('IsArray', a => a[0] instanceof VArr); defb('IsObject', a => isObj(a[0]) || a[0] === NOTHING); defb('IsError', a => false); defb('IsMissing', a => a[0] === undefined);
  defb('TypeName', a => { const v = a[0]; if (v === null) return 'Empty'; if (v === NULLV) return 'Null'; if (v === NOTHING) return 'Nothing'; if (typeof v === 'string') return 'String'; if (typeof v === 'boolean') return 'Boolean'; if (v instanceof VDate) return 'Date'; if (typeof v === 'number') return Number.isInteger(v) ? (Math.abs(v) < 32768 ? 'Integer' : 'Long') : 'Double'; if (v instanceof VArr) return 'Variant()'; if (isObj(v)) return v.vbType || 'Object'; return 'Variant'; });
  defb('VarType', a => { const v = a[0]; if (v === null) return 0; if (v === NULLV) return 1; if (typeof v === 'string') return 8; if (typeof v === 'boolean') return 11; if (v instanceof VDate) return 7; if (typeof v === 'number') return Number.isInteger(v) ? 3 : 5; if (v instanceof VArr) return 8204; return 9; });
  // arrays
  defb('Array', a => VArr.from1D(a, 0)); defb('UBound', a => { const d = a[0].dims[(a[1] !== undefined ? toNumber(a[1]) : 1) - 1]; if (!d) throw new VBAError(9, 'Subscript out of range'); return d.hi; }); defb('LBound', a => { const d = a[0].dims[(a[1] !== undefined ? toNumber(a[1]) : 1) - 1]; if (!d) throw new VBAError(9, 'Subscript out of range'); return d.lo; });
  defb('Split', a => { const s = toStr(a[0]); const d = a[1] === undefined ? ' ' : toStr(a[1]); const lim = a[2] !== undefined ? toNumber(a[2]) : -1; if (s === '') return new VArr([{ lo: 0, hi: -1 }]); let parts = d === '' ? [s] : s.split(d); if (lim > 0 && parts.length > lim) { parts = parts.slice(0, lim - 1).concat([parts.slice(lim - 1).join(d)]); } return VArr.from1D(parts, 0); });
  defb('Join', a => { const arr = a[0]; if (!(arr instanceof VArr)) throw new VBAError(13, 'Type mismatch'); return arr.data.map(toStr).join(a[1] === undefined ? ' ' : toStr(a[1])); });
  defb('Filter', a => VArr.from1D(a[0].data.filter(x => toStr(x).toLowerCase().includes(toStr(a[1]).toLowerCase()) === (a[2] === undefined ? true : toBoolean(a[2]))), 0));
  // misc
  defb('IIf', a => toBoolean(a[0]) ? a[1] : a[2]); defb('Choose', a => { const i = Math.round(toNumber(a[0])); return i >= 1 && i < a.length ? a[i] : null; });
  defb('Switch', a => { for (let i = 0; i + 1 < a.length; i += 2) if (toBoolean(a[i])) return a[i + 1]; return null; });
  defb('Nz', a => (a[0] === null || a[0] === NULLV) ? (a[1] === undefined ? 0 : a[1]) : a[0]);
  defb('DoEvents Beep', () => 0); defb('RGB', a => toNumber(a[0]) + 256 * toNumber(a[1]) + 65536 * toNumber(a[2]));
  defb('MsgBox', function (a) { const msg = toStr(a[0]); const btn = a[1] !== undefined ? toNumber(a[1]) : 0; const kind = btn & 7; const title = a[2] !== undefined ? toStr(a[2]) : ''; if (this.host.msgBox) return this.host.msgBox(msg, btn, title); if (kind === 0) return 1; return 6; });
  defb('InputBox', function (a) { if (this.host.inputBox) return this.host.inputBox(toStr(a[0]), a[1] !== undefined ? toStr(a[1]) : '', a[2] !== undefined ? toStr(a[2]) : ''); return ''; });
  defb('CreateObject', function (a) { return this.createObject(toStr(a[0])); }); defb('GetObject', () => { throw new VBAError(429, 'ActiveX component can\'t create object'); });
  defb('Environ', () => ''); defb('Shell', () => { throw new VBAError(70, 'Permission denied: Shell is not available'); });
  defb('__Unload', function (a) { if (this.host.unload) this.host.unload(a[0]); return null; });
  // Debug.Print support via 'Debug' object
  const DebugObj = { vbType: 'Debug', vbGet(name, args) { if (String(name).toLowerCase() === 'print') { return null; } throw new VBAError(438, 'Object doesn\'t support this property or method'); }, vbSet() { } };
  Runtime.prototype.debugObj = function () { const self = this; return { vbType: 'Debug', vbGet(name, args) { if (String(name).toLowerCase() === 'print') { self.print(args.map(x => x === undefined ? '' : toStr(x)).join('\t')); return null; } throw new VBAError(438, 'Object doesn\'t support this property or method'); }, vbSet() { } }; };
  const origEvalCall = Runtime.prototype.evalCall;
  Runtime.prototype.evalCall = function (node, ctx, asStmt) {
    if (node.t === 'id' && node.args === null && node.n.toLowerCase() === 'debug' && !this.lookup(node.n, ctx)) return this.debugObj();
    if (node.t === 'id' && node.n.toLowerCase() === 'err' && !this.lookup('err', ctx)) { return this.errObj; }
    if (node.t === 'id' && node.n.toLowerCase() === 'me' && !this.lookup('me', ctx) && this.host.me) return this.host.me(ctx);
    return origEvalCall.call(this, node, ctx, asStmt);
  };

  const api = { Runtime, VBAError, VArr, VDate, VColl, VDict, NOTHING, NULLV, toNumber, toStr, toBoolean, isObj, parseProgram, fmtVBA, defaultFor };
  if (typeof module !== 'undefined' && module.exports) module.exports = api; else root.VBA = api;
})(typeof window !== 'undefined' ? window : globalThis);
