/* Khung soạn công thức cho Word — chèn công thức gốc của Word (OMML). */
(function () {
  'use strict';

  // ---------- Bảng nút nhanh (LaTeX; #0 = phần đang chọn, #? = ô trống) ----------
  var PALETTE = [
    { title: 'Căn & phân số', items: [
      ['\\frac{#0}{#?}', '\\frac{a}{b}', 'a/b'], ['\\sqrt{#0}', '\\sqrt{x}', '√x'], ['\\sqrt[3]{#0}', '\\sqrt[3]{x}', '∛x'],
      ['\\sqrt[#?]{#0}', '\\sqrt[n]{x}', 'ⁿ√x'], ['#0^{2}', 'x^2', 'x²'], ['#0^{#?}', 'x^n', 'xⁿ'], ['#0_{#?}', 'x_n', 'xₙ'],
      ['\\left(#0\\right)', '(\\square)', '(□)'], ['\\left|#0\\right|', '|\\square|', '|□|'], ['\\left[#0\\right]', '[\\square]', '[□]'],
      ['\\left(#0\\right)^{2}', '(\\square)^2', '(□)²'], ['\\sqrt{\\left(#0\\right)^{2}}', '\\sqrt{(\\square)^2}', '√(□)²']
    ]},
    { title: 'Quan hệ & phép toán', items: [
      ['=', '=', '='], ['\\ne', '\\ne', '≠'], ['<', '<', '<'], ['>', '>', '>'], ['\\le', '\\le', '≤'], ['\\ge', '\\ge', '≥'],
      ['\\approx', '\\approx', '≈'], ['\\pm', '\\pm', '±'], ['\\cdot', '\\cdot', '·'], ['\\times', '\\times', '×'], [':', ':', ':'], ['\\infty', '\\infty', '∞']
    ]},
    { title: 'Logic & tập hợp', items: [
      ['\\Rightarrow', '\\Rightarrow', '⇒'], ['\\Leftrightarrow', '\\Leftrightarrow', '⇔'], ['\\in', '\\in', '∈'], ['\\notin', '\\notin', '∉'],
      ['\\subset', '\\subset', '⊂'], ['\\cup', '\\cup', '∪'], ['\\cap', '\\cap', '∩'], ['\\varnothing', '\\varnothing', '∅'],
      ['\\mathbb{R}', '\\mathbb{R}', 'ℝ'], ['\\mathbb{N}', '\\mathbb{N}', 'ℕ'], ['\\mathbb{Z}', '\\mathbb{Z}', 'ℤ'], ['\\mathbb{Q}', '\\mathbb{Q}', 'ℚ']
    ]},
    { title: 'Hệ & cấu trúc', items: [
      ['\\begin{cases}#0\\\\#?\\end{cases}', '\\begin{cases}\\square\\\\\\square\\end{cases}', '{ hệ 2'],
      ['\\begin{cases}#0\\\\#?\\\\#?\\end{cases}', '\\begin{cases}\\square\\\\\\square\\\\\\square\\end{cases}', '{ hệ 3'],
      ['\\left[\\begin{array}{l}#0\\\\#?\\end{array}\\right.', '\\left[\\begin{array}{l}\\square\\\\\\square\\end{array}\\right.', '[ hoặc'],
      ['\\overline{#0}', '\\overline{AB}', 'AB̄'], ['\\text{#0}', '\\text{chữ}', 'chữ']
    ]},
    { title: 'Hình học & lượng giác', items: [
      ['\\widehat{#0}', '\\widehat{ABC}', 'ABC^'], ['\\overrightarrow{#0}', '\\overrightarrow{AB}', 'AB→'], ['^{\\circ}', 'x^\\circ', 'x°'],
      ['\\angle', '\\angle', '∠'], ['\\triangle', '\\triangle', '△'], ['\\perp', '\\perp', '⊥'], ['\\parallel', '\\parallel', '∥'],
      ['\\backsim', '\\backsim', '∽'], ['\\pi', '\\pi', 'π'], ['\\alpha', '\\alpha', 'α'], ['\\Delta', '\\Delta', 'Δ'],
      ['\\sin #0', '\\sin', 'sin'], ['\\cos #0', '\\cos', 'cos'], ['\\tan #0', '\\tan', 'tan'], ['\\cot #0', '\\cot', 'cot']
    ]}
  ];

  var $ = function (id) { return document.getElementById(id); };
  var mf = $('mf'), latexBox = $('latex');
  var officeReady = false;

  function setStatus(msg, kind) {
    var s = $('status');
    s.textContent = msg || '';
    s.className = 'status ' + (kind || '');
    if (msg && kind !== 'err') setTimeout(function () { if (s.textContent === msg) s.textContent = ''; }, 4000);
  }

  // ---------- Ô soạn MathLive ----------
  var mathliveOk = false;
  function initMathLive() {
    if (!window.MathfieldElement) return false;
    mathliveOk = true;
    try {
      mf.mathModeSpace = '\\:';
      mf.smartFence = true;
      mf.smartSuperscript = true;
      mf.inlineShortcuts = Object.assign({}, mf.inlineShortcuts, {
        'can': '\\sqrt{#?}', 'can3': '\\sqrt[3]{#?}', 'tg': '\\tan', 'cotg': '\\cot', 'do': '^{\\circ}',
        'goc': '\\widehat{#?}', 'vecto': '\\overrightarrow{#?}', 'he': '\\begin{cases}#?\\\\#?\\end{cases}',
        '<=': '\\le', '>=': '\\ge', '!=': '\\ne', '=>': '\\Rightarrow', '<=>': '\\Leftrightarrow'
      });
    } catch (e) { /* phiên bản cũ có thể thiếu thuộc tính */ }
    mf.addEventListener('input', function () { latexBox.value = mf.value; });
    latexBox.addEventListener('input', function () { mf.value = latexBox.value; });
    $('kbd').onclick = function () {
      var k = window.mathVirtualKeyboard;
      if (!k) return;
      if (k.visible) k.hide(); else { mf.focus(); k.show(); }
    };
    return true;
  }
  function currentLatex() { return mathliveOk ? mf.value : latexBox.value; }

  // ---------- Bảng nút ----------
  function buildPalette() {
    var host = $('palette');
    PALETTE.forEach(function (grp) {
      var h = document.createElement('h2'); h.textContent = grp.title; host.appendChild(h);
      var box = document.createElement('div'); box.className = 'grid';
      grp.items.forEach(function (it) {
        var b = document.createElement('button');
        b.className = 'key'; b.title = it[0].replace(/#0|#\?/g, '□');
        var done = false;
        if (window.MathLive && MathLive.convertLatexToMarkup) {
          try { b.innerHTML = MathLive.convertLatexToMarkup(it[1]); done = true; } catch (e) { /* dùng nhãn chữ */ }
        }
        if (!done) b.textContent = it[2];
        b.onclick = function () {
          if (mathliveOk) { mf.focus(); mf.insert(it[0], { selectionMode: 'placeholder', format: 'latex' }); latexBox.value = mf.value; }
          else { latexBox.value += it[0].replace(/#0|#\?/g, ''); }
        };
        box.appendChild(b);
      });
      host.appendChild(box);
    });
  }

  // ---------- Chèn vào Word ----------
  function needOffice() {
    if (!officeReady) { setStatus('Hãy mở trang này trong Word (tab TOÁN).', 'err'); return true; }
    return false;
  }

  async function insertFormula() {
    var tex = currentLatex().trim();
    if (!tex) { setStatus('Chưa có công thức.', 'err'); return; }
    if (needOffice()) return;
    var display = $('display').checked;
    try {
      await Word.run(async function (ctx) {
        var sel = ctx.document.getSelection();
        sel.load('font/size');
        await ctx.sync();
        var size = sel.font.size || null;
        var omml = L2O.latexToOmml(tex, { display: display, fontSizePt: size });
        var where = display ? 'After' : 'Replace';
        if (display) {
          var p = sel.paragraphs.getLast();
          p.insertParagraph('', 'After').insertOoxml(L2O.ommlToOoxmlPackage(omml), 'Replace');
        } else {
          sel.insertOoxml(L2O.ommlToOoxmlPackage(omml), where);
        }
        await ctx.sync();
      });
      setStatus('Đã chèn ✓', 'ok');
      mf.value = ''; latexBox.value = '';
      if (mathliveOk) mf.focus();
    } catch (err) {
      setStatus('Lỗi: ' + (err && err.message ? err.message : err), 'err');
    }
  }

  // ---------- Chuyển hàng loạt $…$ ----------
  async function convertDollar(onlySelection) {
    if (needOffice()) return;
    setStatus('Đang chuyển…');
    var total = 0, failed = 0;
    try {
      await Word.run(async function (ctx) {
        var scope = onlySelection ? ctx.document.getSelection() : ctx.document.body;
        // Lượt 1: $$…$$ (công thức riêng dòng) ; lượt 2: $…$
        var passes = [{ pat: '$$*$$', display: true, strip: 2 }, { pat: '$*$', display: false, strip: 1 }];
        for (var k = 0; k < passes.length; k++) {
          var ps = passes[k];
          var res = scope.search(ps.pat, { matchWildcards: true });
          res.load('items/text,items/font/size');
          await ctx.sync();
          var items = res.items;
          for (var i = 0; i < items.length; i++) {
            var r = items[i], txt = r.text;
            if (txt.length <= ps.strip * 2) continue;
            var tex = txt.slice(ps.strip, txt.length - ps.strip);
            if (!tex.trim() || tex.indexOf('$') >= 0) continue;
            try {
              var omml = L2O.latexToOmml(tex, { display: false, fontSizePt: r.font.size || null });
              r.insertOoxml(L2O.ommlToOoxmlPackage(omml), 'Replace');
              total++;
            } catch (e) { failed++; }
            if (total % 40 === 0) await ctx.sync();
          }
          await ctx.sync();
        }
      });
      setStatus('Đã chuyển ' + total + ' công thức' + (failed ? (', lỗi ' + failed) : '') + ' ✓', failed ? 'err' : 'ok');
    } catch (err) {
      setStatus('Lỗi: ' + (err && err.message ? err.message : err), 'err');
    }
  }

  // ---------- Khởi động ----------
  function start() {
    buildPalette();
    $('insert').onclick = insertFormula;
    $('clear').onclick = function () { mf.value = ''; latexBox.value = ''; if (mathliveOk) mf.focus(); };
    $('convertAll').onclick = function () { convertDollar(false); };
    $('convertSel').onclick = function () { convertDollar(true); };
    document.addEventListener('keydown', function (ev) {
      if (ev.key === 'Enter' && (ev.ctrlKey || ev.metaKey)) { ev.preventDefault(); insertFormula(); }
    });
  }

  function whenMathLive(cb) {
    var tries = 0;
    (function wait() {
      if (window.MathfieldElement || tries > 50) return cb();
      tries++; setTimeout(wait, 100);
    })();
  }

  document.addEventListener('DOMContentLoaded', function () {
    whenMathLive(function () {
      if (!initMathLive()) {
        mf.style.display = 'none';
        document.querySelector('details').open = true;
        setStatus('Không tải được ô soạn trực quan — dùng ô LaTeX.', 'err');
      }
      start();
    });
  });

  if (window.Office) {
    Office.onReady(function (info) {
      officeReady = !!(info && info.host === Office.HostType.Word);
    });
  }
})();
