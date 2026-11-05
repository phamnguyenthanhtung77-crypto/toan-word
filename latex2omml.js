/*
 * latex2omml.js — chuyển LaTeX (tập con dùng cho Toán THCS/THPT) sang OMML,
 * tức định dạng công thức gốc của Microsoft Word.
 * Dùng được trong trình duyệt (window.latexToOmml) và Node (module.exports).
 */
(function (root) {
  'use strict';

  // ---------- Bảng ký hiệu ----------
  var SYM = {
    // phép toán, quan hệ
    cdot: '⋅', times: '×', div: '÷', pm: '±', mp: '∓',
    le: '≤', leq: '≤', ge: '≥', geq: '≥', ne: '≠', neq: '≠',
    leqslant: '⩽', geqslant: '⩾', approx: '≈', equiv: '≡', sim: '∼',
    simeq: '≃', cong: '≅', propto: '∝', ll: '≪', gg: '≫',
    Rightarrow: '⇒', Leftarrow: '⇐', Leftrightarrow: '⇔', iff: '⇔',
    implies: '⇒', rightarrow: '→', to: '→', leftarrow: '←',
    leftrightarrow: '↔', mapsto: '↦', longrightarrow: '⟶',
    Longrightarrow: '⟹', Longleftrightarrow: '⟺',
    in: '∈', notin: '∉', ni: '∋', subset: '⊂', supset: '⊃',
    subseteq: '⊆', supseteq: '⊇', cup: '∪', cap: '∩',
    emptyset: '∅', varnothing: '∅', setminus: '∖', forall: '∀',
    exists: '∃', neg: '¬', lnot: '¬', land: '∧', lor: '∨',
    wedge: '∧', vee: '∨',
    parallel: '∥', nparallel: '∦', perp: '⊥', angle: '∠',
    measuredangle: '∡', triangle: '△', bigtriangleup: '△', square: '□',
    Box: '□', circ: '∘', bullet: '∙', degree: '°', infty: '∞',
    partial: '∂', nabla: '∇', prime: '′', ldots: '…', dots: '…',
    cdots: '⋯', vdots: '⋮', ddots: '⋱', therefore: '∴', because: '∵',
    mid: '∣', nmid: '∤', colon: ':', vert: '|', Vert: '‖', lvert: '|', rvert: '|',
    lbrace: '{', rbrace: '}', lbrack: '[', rbrack: ']', langle: '⟨', rangle: '⟩',
    backslash: '\\', '%': '%', '$': '$', '#': '#', '&': '&', '_': '_', '{': '{', '}': '}',
    '|': '‖', ',': ' ', ';': ' ', ':': ' ', '!': '', ' ': ' ',
    quad: ' ', qquad: '  ', enspace: ' ', thinspace: ' ',
    // chữ Hy Lạp
    alpha: 'α', beta: 'β', gamma: 'γ', delta: 'δ', epsilon: 'ϵ',
    varepsilon: 'ε', zeta: 'ζ', eta: 'η', theta: 'θ', vartheta: 'ϑ',
    iota: 'ι', kappa: 'κ', lambda: 'λ', mu: 'μ', nu: 'ν',
    xi: 'ξ', pi: 'π', varpi: 'ϖ', rho: 'ρ', sigma: 'σ',
    tau: 'τ', upsilon: 'υ', phi: 'ϕ', varphi: 'φ', chi: 'χ',
    psi: 'ψ', omega: 'ω', Gamma: 'Γ', Delta: 'Δ', Theta: 'Θ',
    Lambda: 'Λ', Xi: 'Ξ', Pi: 'Π', Sigma: 'Σ', Phi: 'Φ',
    Psi: 'Ψ', Omega: 'Ω'
  };
  var BB = { R: 'ℝ', N: 'ℕ', Z: 'ℤ', Q: 'ℚ', C: 'ℂ', P: 'ℙ' };
  var FUNCS = ['sin', 'cos', 'tan', 'cot', 'tg', 'cotg', 'sec', 'csc', 'log', 'ln', 'lg',
    'exp', 'arcsin', 'arccos', 'arctan', 'sinh', 'cosh', 'tanh', 'min', 'max', 'gcd',
    'deg', 'det', 'dim', 'ker', 'arg', 'mod', 'sgn'];
  var LIMFUNCS = ['lim', 'limsup', 'liminf', 'sup', 'inf'];
  var NARY = { sum: '∑', prod: '∏', int: '∫', iint: '∬', oint: '∮',
    bigcup: '⋃', bigcap: '⋂' };
  var ACC = { vec: '⃗', overrightarrow: '⃗', hat: '̂', widehat: '̂',
    tilde: '̃', widetilde: '̃', dot: '̇', ddot: '̈', bar: '̅',
    check: '̌', breve: '̆', acute: '́', grave: '̀',
    overleftarrow: '⃖', overleftrightarrow: '⃡' };
  var DELIM = { '(': '(', ')': ')', '[': '[', ']': ']', '\\{': '{', '\\}': '}',
    '\\lbrace': '{', '\\rbrace': '}', '\\lbrack': '[', '\\rbrack': ']', '|': '|',
    '\\vert': '|', '\\lvert': '|', '\\rvert': '|', '\\|': '‖', '\\Vert': '‖',
    '\\langle': '⟨', '\\rangle': '⟩', '\\lfloor': '⌊', '\\rfloor': '⌋',
    '\\lceil': '⌈', '\\rceil': '⌉', '.': '', '<': '⟨', '>': '⟩' };
  var IGNORE = ['displaystyle', 'textstyle', 'scriptstyle', 'limits', 'nolimits', 'left.',
    'big', 'Big', 'bigg', 'Bigg', 'bigl', 'bigr', 'Bigl', 'Bigr', 'nonumber', 'notag',
    'mathstrut', 'strut', 'relax'];

  // ---------- Tiện ích XML ----------
  function esc(s) {
    return s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
  }
  var OPTS = { rpr: '' };
  function run(text, sty) {
    if (text === '') return '';
    var mr = '';
    if (sty === 'nor') mr = '<m:rPr><m:nor/></m:rPr>';
    else if (sty) mr = '<m:rPr><m:sty m:val="' + sty + '"/></m:rPr>';
    return '<m:r>' + mr + OPTS.rpr + '<m:t xml:space="preserve">' + esc(text) + '</m:t></m:r>';
  }
  function e(x) { return '<m:e>' + x + '</m:e>'; }

  // ---------- Tách token ----------
  function tokenize(s) {
    var t = [], i = 0;
    while (i < s.length) {
      var c = s[i];
      if (c === '\\') {
        var m = /^\\([A-Za-z]+)\*?/.exec(s.slice(i));
        if (m) { t.push({ k: 'cmd', v: m[1] }); i += m[0].length; continue; }
        var nc = s[i + 1];
        if (nc === '\\') { t.push({ k: 'nl' }); i += 2; continue; }
        t.push({ k: 'cmd', v: nc === undefined ? '' : nc }); i += 2; continue;
      }
      if (c === '%') { while (i < s.length && s[i] !== '\n') i++; continue; }
      if (/\s/.test(c)) { t.push({ k: 'sp' }); while (i < s.length && /\s/.test(s[i])) i++; continue; }
      if (c === '{') { t.push({ k: '{' }); i++; continue; }
      if (c === '}') { t.push({ k: '}' }); i++; continue; }
      if (c === '^' || c === '_' || c === '&') { t.push({ k: c }); i++; continue; }
      t.push({ k: 'ch', v: c }); i++;
    }
    return t;
  }

  // ---------- Bộ phân tích ----------
  function Parser(src) { this.t = tokenize(src); this.p = 0; }
  var P = Parser.prototype;
  P.peek = function () { return this.t[this.p]; };
  P.next = function () { return this.t[this.p++]; };
  P.skipSp = function () { while (this.peek() && this.peek().k === 'sp') this.p++; };
  P.eof = function () { return this.p >= this.t.length; };

  // Đọc một đối số bắt buộc: {…} hoặc một token.
  P.argTokensRaw = function () {
    this.skipSp();
    var tk = this.peek();
    if (!tk) return '';
    if (tk.k === '{') {
      this.p++;
      var depth = 1, out = '';
      while (!this.eof()) {
        var x = this.next();
        if (x.k === '{') { depth++; out += '{'; continue; }
        if (x.k === '}') { depth--; if (depth === 0) break; out += '}'; continue; }
        out += tokText(x);
      }
      return out;
    }
    this.p++;
    return tokText(tk);
  };
  function tokText(x) {
    switch (x.k) {
      case 'cmd': return /^[A-Za-z]+$/.test(x.v) ? '\\' + x.v + ' ' : '\\' + x.v;
      case 'ch': return x.v;
      case 'sp': return ' ';
      case 'nl': return '\\\\';
      default: return x.k;
    }
  }
  P.arg = function () { return convertInner(this.argTokensRaw()); };
  P.optArg = function () {
    this.skipSp();
    var tk = this.peek();
    if (!tk || tk.k !== 'ch' || tk.v !== '[') return null;
    this.p++;
    var depth = 0, out = '';
    while (!this.eof()) {
      var x = this.next();
      if (x.k === 'ch' && x.v === '[') depth++;
      if (x.k === 'ch' && x.v === ']') { if (depth === 0) break; depth--; }
      out += tokText(x);
    }
    return convertInner(out);
  };
  P.readDelim = function () {
    this.skipSp();
    var x = this.next();
    if (!x) return '';
    var key = x.k === 'cmd' ? '\\' + x.v : (x.v || x.k);
    if (key in DELIM) return DELIM[key];
    return key.replace(/^\\/, '');
  };

  // Phân tích một dãy nguyên tố đến khi gặp điều kiện dừng; trả về mảng atom.
  // atom: {x: chuỗi OMML, plain: true nếu là ký tự thường (gộp được)}
  P.seq = function (stop) {
    var atoms = [];
    while (!this.eof()) {
      var tk = this.peek();
      if (stop && stop(tk, this)) break;
      if (tk.k === '^' || tk.k === '_') { this.p++; this.script(atoms, tk.k); continue; }
      var a = this.atom();
      if (a === null) continue;
      if (Array.isArray(a)) { for (var i = 0; i < a.length; i++) atoms.push(a[i]); }
      else atoms.push(a);
    }
    return atoms;
  };

  P.script = function (atoms, kind) {
    var base = atoms.length ? atoms.pop() : { x: '' };
    var sub = null, sup = null;
    var first = this.scriptArg();
    if (kind === '^') sup = first; else sub = first;
    this.skipSp();
    var tk = this.peek();
    if (tk && (tk.k === '^' || tk.k === '_') && tk.k !== kind) {
      this.p++;
      var second = this.scriptArg();
      if (tk.k === '^') sup = second; else sub = second;
    }
    // ^\circ hoặc ^{\circ} → dấu độ
    if (sup && sup.raw && /^\s*\\circ\s*$/.test(sup.raw) && !sub) {
      atoms.push(base); atoms.push({ x: run('°'), plain: true, t: '°' }); return;
    }
    if (base.nary) { // ∑, ∫ với cận
      atoms.push(base.withLimits(sub ? sub.x : null, sup ? sup.x : null)); return;
    }
    if (base.limfn) { atoms.push(base.withLim(sub ? sub.x : '')); return; }
    var bx = base.x;
    var out;
    if (sub && sup) out = '<m:sSubSup>' + e(bx) + '<m:sub>' + sub.x + '</m:sub><m:sup>' + sup.x + '</m:sup></m:sSubSup>';
    else if (sup) out = '<m:sSup>' + e(bx) + '<m:sup>' + sup.x + '</m:sup></m:sSup>';
    else out = '<m:sSub>' + e(bx) + '<m:sub>' + sub.x + '</m:sub></m:sSub>';
    var res = { x: out };
    if (base.fn) { res.fn = true; }
    atoms.push(res);
  };
  P.scriptArg = function () {
    this.skipSp();
    var tk = this.peek();
    if (tk && tk.k === 'cmd' && tk.v === 'prime') { this.p++; return { x: run('′'), raw: '\\prime' }; }
    var raw = this.argTokensRaw();
    return { x: convertInner(raw), raw: raw };
  };

  P.atom = function () {
    var tk = this.next();
    switch (tk.k) {
      case 'sp': return null;
      case '{': {
        var inner = this.seq(function (t) { return t.k === '}'; });
        this.next();
        return { x: join(inner) };
      }
      case '}': return null;
      case 'nl': return null;
      case '&': return { x: run(' '), plain: true, t: ' ' };
      case 'ch': return this.charAtom(tk.v);
      case 'cmd': return this.cmdAtom(tk.v);
    }
    return null;
  };

  P.charAtom = function (c) {
    if (c === '(' || c === '[') {
      // tự ghép ngoặc thường thành ngoặc co giãn nếu tìm thấy ngoặc đóng tương ứng
      var close = c === '(' ? ')' : ']';
      var save = this.p, depth = 0, found = false;
      for (var j = this.p; j < this.t.length; j++) {
        var t = this.t[j];
        if (t.k === 'ch' && t.v === c) depth++;
        if (t.k === 'ch' && t.v === close) { if (depth === 0) { found = true; break; } depth--; }
        if (t.k === 'cmd' && (t.v === 'right' || t.v === 'end')) break;
        if (t.k === '}' || t.k === 'nl' || t.k === '&') { if (depth === 0) break; }
      }
      if (found && !(c === '[' && this.inInterval(save))) {
        var inner = this.seq(function (t) { return t.k === 'ch' && t.v === close; });
        this.next();
        return { x: delimXml(c, close, [join(inner)]) };
      }
      return { x: run(c), plain: true, t: c };
    }
    var map = { '-': '−', '*': '⋅', "'": '′', '~': ' ' };
    var ch = map[c] || c;
    return { x: run(ch), plain: true, t: ch };
  };
  // [a; b) — khoảng nửa mở thì không tự ghép
  P.inInterval = function (from) {
    for (var j = from; j < this.t.length; j++) {
      var t = this.t[j];
      if (t.k === 'ch' && t.v === ']') return false;
      if (t.k === 'ch' && t.v === ')') return true;
    }
    return false;
  };

  P.cmdAtom = function (name) {
    var self = this;
    if (name === '' ) return null;
    if (IGNORE.indexOf(name) >= 0) return null;
    if (name === 'frac' || name === 'dfrac' || name === 'tfrac' || name === 'cfrac') {
      var num = this.arg(), den = this.arg();
      return { x: '<m:f><m:num>' + num + '</m:num><m:den>' + den + '</m:den></m:f>' };
    }
    if (name === 'binom' || name === 'dbinom' || name === 'tbinom') {
      var n1 = this.arg(), n2 = this.arg();
      return { x: delimXml('(', ')', ['<m:f><m:fPr><m:type m:val="noBar"/></m:fPr><m:num>' + n1 + '</m:num><m:den>' + n2 + '</m:den></m:f>']) };
    }
    if (name === 'sqrt') {
      var deg = this.optArg();
      var body = this.arg();
      if (deg === null) return { x: '<m:rad><m:radPr><m:degHide m:val="1"/></m:radPr><m:deg/>' + e(body) + '</m:rad>' };
      return { x: '<m:rad><m:deg>' + deg + '</m:deg>' + e(body) + '</m:rad>' };
    }
    if (name === 'root') { // \root n \of {x}
      var dg = this.arg(); this.skipSp(); if (this.peek() && this.peek().v === 'of') this.p++;
      return { x: '<m:rad><m:deg>' + dg + '</m:deg>' + e(this.arg()) + '</m:rad>' };
    }
    if (name === 'left') {
      var open = this.readDelim();
      var parts = [], cur = [];
      while (!this.eof()) {
        var tk = this.peek();
        if (tk.k === 'cmd' && tk.v === 'right') { this.p++; break; }
        if (tk.k === 'cmd' && tk.v === 'middle') { this.p++; this.readDelim(); parts.push(join(cur)); cur = []; continue; }
        var got = this.seq(function (t) { return t.k === 'cmd' && (t.v === 'right' || t.v === 'middle'); });
        cur = cur.concat(got);
      }
      parts.push(join(cur));
      var closeD = this.readDelim();
      return { x: delimXml(open, closeD, parts, true) };
    }
    if (name === 'right' || name === 'middle') { this.readDelim(); return null; }
    if (name === 'begin') return this.env(this.argTokensRaw().trim());
    if (name === 'end') { this.argTokensRaw(); return null; }
    if (name === 'text' || name === 'textrm' || name === 'mbox' || name === 'textnormal' || name === 'textit' || name === 'textbf') {
      var raw = this.argTokensRaw();
      return { x: run(raw.replace(/\\ /g, ' ').replace(/\\([{}%$&#_])/g, '$1'), 'nor') };
    }
    if (name === 'mathrm' || name === 'operatorname' || name === 'mathup' || name === 'rm') {
      var r2 = this.argTokensRaw();
      if (/^[A-Za-z0-9 .,]+$/.test(r2)) return { x: run(r2.trim(), 'p'), fn: name === 'operatorname' };
      return { x: convertInner(r2) };
    }
    if (name === 'mathbf' || name === 'boldsymbol' || name === 'bm') {
      var r3 = this.argTokensRaw();
      if (/^[A-Za-z0-9 ]+$/.test(r3)) return { x: run(r3.trim(), 'b') };
      return { x: convertInner(r3) };
    }
    if (name === 'mathit') return { x: this.arg() };
    if (name === 'mathbb' || name === 'Bbb') {
      var r4 = this.argTokensRaw().trim();
      return { x: run(r4.split('').map(function (c) { return BB[c] || c; }).join('')), plain: true };
    }
    if (name === 'overline' || name === 'underline') {
      var b1 = this.arg();
      return { x: '<m:bar><m:barPr><m:pos m:val="' + (name === 'overline' ? 'top' : 'bot') + '"/></m:barPr>' + e(b1) + '</m:bar>' };
    }
    if (name in ACC) {
      var b2 = this.arg();
      return { x: '<m:acc><m:accPr><m:chr m:val="' + ACC[name] + '"/></m:accPr>' + e(b2) + '</m:acc>' };
    }
    if (name === 'overset' || name === 'stackrel') {
      var top = this.arg(), bas = this.arg();
      return { x: '<m:limUpp>' + e(bas) + '<m:lim>' + top + '</m:lim></m:limUpp>' };
    }
    if (name === 'underset') {
      var bot = this.arg(), bas2 = this.arg();
      return { x: '<m:limLow>' + e(bas2) + '<m:lim>' + bot + '</m:lim></m:limLow>' };
    }
    if (name === 'placeholder') { this.optArg(); this.argTokensRaw(); return { x: '' }; }
    if (name === 'hspace' || name === 'hskip' || name === 'mkern' || name === 'kern') { this.argTokensRaw(); return { x: run(' '), plain: true }; }
    if (name === 'color' || name === 'textcolor') { this.argTokensRaw(); return null; }
    if (name === 'boxed' || name === 'fbox') {
      return { x: '<m:borderBox>' + e(this.arg()) + '</m:borderBox>' };
    }
    if (name === 'not') {
      this.skipSp(); var nt = this.next();
      var base = nt.k === 'cmd' ? (SYM[nt.v] || nt.v) : nt.v;
      var negs = { '=': '≠', '∈': '∉', '∥': '∦', '<': '≮', '>': '≯', '≤': '≰', '≥': '≱', '⊂': '⊄' };
      return { x: run(negs[base] || base + '̸'), plain: true };
    }
    if (FUNCS.indexOf(name) >= 0) return this.func(run(name, 'p'));
    if (LIMFUNCS.indexOf(name) >= 0) {
      var nm = name;
      return {
        limfn: true, x: run(nm, 'p'),
        withLim: function (lim) {
          var fname = lim ? '<m:limLow>' + e(run(nm, 'p')) + '<m:lim>' + lim + '</m:lim></m:limLow>' : run(nm, 'p');
          var arg = self.funcArg();
          return { x: '<m:func><m:fName>' + fname + '</m:fName>' + e(arg) + '</m:func>' };
        }
      };
    }
    if (name in NARY) {
      var chr = NARY[name];
      var isInt = /int$/.test(name);
      return {
        nary: true, x: run(chr),
        withLimits: function (lo, hi) {
          var body = self.funcArg(true);
          var pr = '<m:naryPr><m:chr m:val="' + chr + '"/>' + (isInt ? '<m:limLoc m:val="subSup"/>' : '<m:limLoc m:val="undOvr"/>') +
            (lo === null ? '<m:subHide m:val="1"/>' : '') + (hi === null ? '<m:supHide m:val="1"/>' : '') + '</m:naryPr>';
          return { x: '<m:nary>' + pr + '<m:sub>' + (lo || '') + '</m:sub><m:sup>' + (hi || '') + '</m:sup>' + e(body) + '</m:nary>' };
        }
      };
    }
    if (name in SYM) {
      var ch = SYM[name];
      if (ch === '') return null;
      if (/^[A-Z]/.test(name) && /[\u0391-\u03A9]/.test(ch)) return { x: run(ch, 'p') }; // chữ Hy Lạp hoa: đứng
      return { x: run(ch), plain: true, t: ch };
    }
    // lệnh lạ: hiện nguyên tên để người dùng thấy và sửa
    return { x: run('\\' + name, 'p'), plain: false };
  };

  // Đối số của hàm (sin x, sin(x), lim …): lấy một atom kèm chỉ số trên/dưới nếu có.
  P.funcArg = function (greedy) {
    this.skipSp();
    if (this.eof()) return '';
    var tk = this.peek();
    if (tk.k === '}' || tk.k === 'nl' || tk.k === '&') return '';
    if (greedy) {
      var rest = this.seq(function (t) { return t.k === '}' || t.k === 'nl' || t.k === '&' || (t.k === 'cmd' && (t.v === 'right' || t.v === 'end')); });
      return join(rest);
    }
    var atoms = [];
    var a = this.atom();
    if (a) atoms.push(a);
    // gom tiếp các chữ/số liền nhau (sin 2x, cos alpha)
    while (!this.eof()) {
      var t2 = this.peek();
      if (t2.k === '^' || t2.k === '_') { this.p++; this.script(atoms, t2.k); continue; }
      if (t2.k === 'ch' && /[A-Za-z0-9.,]/.test(t2.v) && atoms.length && atoms[atoms.length - 1].plain) { atoms.push(this.atom()); continue; }
      if (t2.k === 'cmd' && /^(alpha|beta|gamma|theta|varphi|phi|pi|omega|x)$/.test(t2.v) && atoms.length && atoms[atoms.length - 1].plain) { atoms.push(this.atom()); continue; }
      break;
    }
    return join(atoms);
  };
  P.func = function (nameXml) {
    var fname = nameXml;
    this.skipSp();
    var tk = this.peek();
    if (tk && (tk.k === '^' || tk.k === '_')) {
      this.p++;
      var arr = [{ x: nameXml }];
      this.script(arr, tk.k);
      fname = arr[0].x;
    }
    var arg = this.funcArg();
    return { x: '<m:func><m:fName>' + fname + '</m:fName>' + e(arg) + '</m:func>' };
  };

  P.env = function (name) {
    var rows = [[]], self = this;
    if (name === 'array' || name === 'tabular') this.argTokensRaw();
    // Đọc nội dung đến \end{name}
    var depth = 0, buf = '', cells = [[]];
    while (!this.eof()) {
      var tk = this.next();
      if (tk.k === 'cmd' && tk.v === 'begin') { depth++; buf += '\\begin'; continue; }
      if (tk.k === 'cmd' && tk.v === 'end') {
        if (depth === 0) { this.argTokensRaw(); break; }
        depth--; buf += '\\end'; continue;
      }
      if (depth === 0 && tk.k === '&') { cells[cells.length - 1].push(buf); buf = ''; continue; }
      if (depth === 0 && tk.k === 'nl') { cells[cells.length - 1].push(buf); buf = ''; cells.push([]); continue; }
      buf += tokText(tk);
    }
    cells[cells.length - 1].push(buf);
    // bỏ dòng rỗng cuối
    if (cells.length > 1 && cells[cells.length - 1].join('').trim() === '') cells.pop();
    var conv = cells.map(function (r) { return r.map(function (c) { return convertInner(c); }); });

    var matrix = { matrix: ['', ''], pmatrix: ['(', ')'], bmatrix: ['[', ']'], Bmatrix: ['{', '}'], vmatrix: ['|', '|'], Vmatrix: ['‖', '‖'], smallmatrix: ['', ''] };
    if (name in matrix || name === 'array') {
      var ncol = Math.max.apply(null, conv.map(function (r) { return r.length; }));
      var mx = '<m:m><m:mPr><m:mcs><m:mc><m:mcPr><m:count m:val="' + ncol + '"/><m:mcJc m:val="center"/></m:mcPr></m:mc></m:mcs></m:mPr>' +
        conv.map(function (r) {
          var cs = r.slice(); while (cs.length < ncol) cs.push('');
          return '<m:mr>' + cs.map(e).join('') + '</m:mr>';
        }).join('') + '</m:m>';
      var d = matrix[name];
      if (d && (d[0] || d[1])) return { x: delimXml(d[0], d[1], [mx], true) };
      return { x: mx };
    }
    // cases, aligned, gathered, align, split, eqnarray… → mảng phương trình
    if (name === 'cases' || name === 'dcases' || name === 'rcases') {
      var nc = Math.max.apply(null, conv.map(function (r) { return r.length; }));
      var mm = '<m:m><m:mPr><m:mcs><m:mc><m:mcPr><m:count m:val="' + nc + '"/><m:mcJc m:val="left"/></m:mcPr></m:mc></m:mcs><m:cGpRule m:val="3"/><m:cGp m:val="12"/></m:mPr>' +
        conv.map(function (r) {
          var cs = r.slice(); while (cs.length < nc) cs.push('');
          return '<m:mr>' + cs.map(e).join('') + '</m:mr>';
        }).join('') + '</m:m>';
      return { x: name === 'rcases' ? delimXml('', '}', [mm]) : delimXml('{', '', [mm]) };
    }
    var eq = '<m:eqArr>' + conv.map(function (r) {
      return e(r.join(run(' ')));
    }).join('') + '</m:eqArr>';
    if (name === 'cases' || name === 'dcases') return { x: delimXml('{', '', [eq], true) };
    if (name === 'rcases') return { x: delimXml('', '}', [eq], true) };
    return { x: eq };
  };

  function delimXml(open, close, parts) {
    var pr = '<m:dPr><m:begChr m:val="' + esc(open) + '"/><m:endChr m:val="' + esc(close) + '"/>' +
      (parts.length > 1 ? '<m:sepChr m:val="|"/>' : '') + '</m:dPr>';
    return '<m:d>' + pr + parts.map(e).join('') + '</m:d>';
  }

  function join(atoms) {
    var out = '', buf = '';
    for (var i = 0; i < atoms.length; i++) {
      var a = atoms[i];
      if (a.plain && a.t !== undefined && !a.fn) { buf += a.t; continue; }
      if (buf) { out += run(buf); buf = ''; }
      out += a.x;
    }
    if (buf) out += run(buf);
    return out;
  }

  function convertInner(src) {
    var p = new Parser(src);
    return join(p.seq(null));
  }

  // ---------- API ----------
  // latexToOmml(latex, {display:boolean, fontSizePt:number, font:string}) → chuỗi <m:oMath> hoặc <m:oMathPara>
  function latexToOmml(latex, opts) {
    opts = opts || {};
    var rpr = '';
    if (opts.fontSizePt || opts.font) {
      rpr = '<w:rPr>' + (opts.font ? '<w:rFonts w:ascii="' + esc(opts.font) + '" w:hAnsi="' + esc(opts.font) + '"/>' : '') +
        (opts.fontSizePt ? '<w:sz w:val="' + Math.round(opts.fontSizePt * 2) + '"/><w:szCs w:val="' + Math.round(opts.fontSizePt * 2) + '"/>' : '') + '</w:rPr>';
    }
    OPTS.rpr = rpr;
    var src = String(latex).trim()
      .replace(/^\$\$([\s\S]*)\$\$$/, '$1').replace(/^\$([\s\S]*)\$$/, '$1')
      .replace(/^\\\[([\s\S]*)\\\]$/, '$1').replace(/^\\\(([\s\S]*)\\\)$/, '$1')
      .replace(/\{,\}/g, ',').replace(/\{\.\}/g, '.');
    var body;
    try { body = convertInner(src); } finally { OPTS.rpr = ''; }
    var math = '<m:oMath>' + body + '</m:oMath>';
    if (opts.display) {
      return '<m:oMathPara><m:oMathParaPr><m:jc m:val="' + (opts.align || 'center') + '"/></m:oMathParaPr>' + math + '</m:oMathPara>';
    }
    return math;
  }

  var NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"';
  // Gói OOXML để chèn bằng Office.js: range.insertOoxml(...)
  function ommlToOoxmlPackage(ommlParagraphContent) {
    return '<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">' +
      '<pkg:part pkg:name="/_rels/.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml"><pkg:xmlData>' +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>' +
      '</Relationships></pkg:xmlData></pkg:part>' +
      '<pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"><pkg:xmlData>' +
      '<w:document ' + NS + '><w:body><w:p>' + ommlParagraphContent + '</w:p></w:body></w:document>' +
      '</pkg:xmlData></pkg:part></pkg:package>';
  }

  var api = { latexToOmml: latexToOmml, ommlToOoxmlPackage: ommlToOoxmlPackage, NS: NS };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  else root.L2O = api;
})(this);
