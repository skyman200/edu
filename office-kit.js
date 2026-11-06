/*!
 * OfficeKit 1.0.0 - 브라우저 안에서 DOCX / XLSX 파일을 만드는 작은 도구.
 * 서버도 네트워크도 쓰지 않습니다. 의존: 전역 JSZip (vendor/jszip/jszip.min.js).
 * 모델이 만든 글과 업로드한 파일 글은 모두 적대적 입력으로 취급합니다:
 * XML 이스케이프, XML 1.0 금지 제어문자 제거, 길이 상한, eval 없음.
 */
(function (root, factory) {
  var api = factory(root);
  if (typeof module === 'object' && module && module.exports) module.exports = api;
  else root.OfficeKit = api;
})(typeof globalThis !== 'undefined' ? globalThis : typeof self !== 'undefined' ? self : this, function (root) {
  'use strict';

  var VERSION = '1.0.0';
  var LIM = {
    html: 1500000, para: 2000, head: 120, cell: 2000, slides: 40, depth: 80,
    tableRows: 500, tableCols: 30, sheets: 30, rows: 10000, cols: 200, cellChars: 32000
  };
  var MIME_DOCX = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';
  var MIME_XLSX = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
  var XML_HEAD = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
  var PAGE_W = 11906, PAGE_H = 16838, MARGIN = 1134; // A4, 20 mm
  var BODY_W = PAGE_W - 2 * MARGIN; // 9638 twips

  /* ------------------------------------------------------------------ */
  /* 공통 도우미                                                          */
  /* ------------------------------------------------------------------ */

  function getJSZip() {
    if (root.JSZip) return root.JSZip;
    if (typeof require === 'function') { try { return require('jszip'); } catch (e) { /* fall through */ } }
    throw new Error('JSZip 라이브러리를 불러오지 못했습니다.');
  }

  var CTRL_RE = /[\u0000-\u0008\u000B\u000C\u000E-\u001F\u007F-\u009F\uFFFE\uFFFF]/g;
  function fixSurrogates(s) {
    var o = '';
    for (var i = 0; i < s.length; i++) {
      var c = s.charCodeAt(i);
      if (c >= 0xD800 && c <= 0xDBFF) {
        var d = s.charCodeAt(i + 1);
        if (d >= 0xDC00 && d <= 0xDFFF) { o += s.charAt(i) + s.charAt(i + 1); i++; }
      } else if (!(c >= 0xDC00 && c <= 0xDFFF)) o += s.charAt(i);
    }
    return o;
  }
  /** XML 1.0 에서 허용되지 않는 문자를 제거한다. */
  function clean(s) {
    s = s == null ? '' : String(s);
    s = s.replace(CTRL_RE, '');
    if (/[\uD800-\uDFFF]/.test(s)) s = fixSurrogates(s);
    return s;
  }
  function esc(s) {
    return clean(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;').replace(/'/g, '&apos;');
  }
  function clone(o, extra) {
    var r = {}, k;
    if (o) for (k in o) r[k] = o[k];
    if (extra) for (k in extra) r[k] = extra[k];
    return r;
  }
  function repeat(s, n) { var o = ''; while (n-- > 0) o += s; return o; }
  function cutStr(s, n) {
    if (s.length <= n) return s;
    var c = s.charCodeAt(n - 1);
    if (c >= 0xD800 && c <= 0xDBFF) n--;
    return s.slice(0, n);
  }
  /** 표시 폭: 한글/한자/전각 = 2, 그 밖 = 1. */
  function dispWidth(s) {
    var w = 0;
    for (var i = 0; i < s.length; i++) {
      var c = s.charCodeAt(i);
      w += (c >= 0x1100 && (c <= 0x11FF || (c >= 0x2E80 && c <= 0xD7FF) || (c >= 0xF900 && c <= 0xFAFF) ||
        (c >= 0xFF00 && c <= 0xFF60) || (c >= 0xFFE0 && c <= 0xFFE6))) ? 2 : 1;
    }
    return w;
  }
  function zipBlob(files, mime) {
    var JSZip = getJSZip();
    var zip = new JSZip();
    files.forEach(function (f) { zip.file(f[0], f[1]); });
    return zip.generateAsync({
      type: 'uint8array', compression: 'DEFLATE', compressionOptions: { level: 6 }, mimeType: mime
    }).then(function (u8) {
      return typeof Blob === 'function' ? new Blob([u8], { type: mime }) : u8;
    });
  }
  function isoNow() { return new Date().toISOString().replace(/\.\d+Z$/, 'Z'); }

  /* ------------------------------------------------------------------ */
  /* 너그러운 HTML 파서 (DOM 없이)                                         */
  /* ------------------------------------------------------------------ */

  var ENT = {
    amp: '&', lt: '<', gt: '>', quot: '"', apos: "'", nbsp: '\u00a0', middot: '\u00b7', hellip: '\u2026',
    ndash: '\u2013', mdash: '\u2014', lsquo: '\u2018', rsquo: '\u2019', ldquo: '\u201c', rdquo: '\u201d',
    times: '\u00d7', rarr: '\u2192', larr: '\u2190', bull: '\u2022', copy: '\u00a9', laquo: '\u00ab',
    raquo: '\u00bb', ensp: '\u2002', emsp: '\u2003', thinsp: '\u2009', shy: '', zwj: '', zwnj: '',
    check: '\u2713', deg: '\u00b0', plusmn: '\u00b1', divide: '\u00f7', sup2: '\u00b2', frac12: '\u00bd'
  };
  function decodeEnt(s) {
    if (s.indexOf('&') < 0) return s;
    return s.replace(/&(#x[0-9a-f]+|#[0-9]+|[a-z][a-z0-9]*);/gi, function (m, g) {
      if (g.charAt(0) === '#') {
        var cp = g.charAt(1).toLowerCase() === 'x' ? parseInt(g.slice(2), 16) : parseInt(g.slice(1), 10);
        if (!(cp > 0 && cp <= 0x10FFFF) || (cp >= 0xD800 && cp <= 0xDFFF)) return '';
        return String.fromCodePoint(cp);
      }
      var v = ENT[g.toLowerCase()];
      return v == null ? m : v;
    });
  }
  function mkEl(tag, attrs) {
    var c = attrs && attrs['class'] ? ' ' + String(attrs['class']).toLowerCase().replace(/\s+/g, ' ').trim() + ' ' : ' ';
    return { tag: tag, attrs: attrs || {}, cls: c, kids: [] };
  }
  function hc(n, c) { return !!n.cls && n.cls.indexOf(' ' + c + ' ') >= 0; }

  var VOID = { br: 1, hr: 1, img: 1, input: 1, meta: 1, link: 1, wbr: 1, col: 1, area: 1, base: 1, source: 1, embed: 1 };
  var RAWTEXT = { script: 1, style: 1 };
  var BLOCK_START = {
    div: 1, section: 1, article: 1, header: 1, footer: 1, aside: 1, h1: 1, h2: 1, h3: 1, h4: 1, h5: 1, h6: 1,
    ul: 1, ol: 1, table: 1, blockquote: 1, p: 1, dl: 1, hr: 1, figure: 1, pre: 1
  };
  var INLINE_T = { span: 1, b: 1, strong: 1, em: 1, i: 1, small: 1, mark: 1, a: 1, u: 1, s: 1, sub: 1, sup: 1, code: 1, abbr: 1, cite: 1, font: 1 };
  var AUTOCLOSE = {
    li: [['li'], ['ul', 'ol']],
    td: [['td', 'th'], ['tr', 'table']], th: [['td', 'th'], ['tr', 'table']],
    tr: [['tr'], ['table']],
    thead: [['thead', 'tbody', 'tfoot'], ['table']], tbody: [['thead', 'tbody', 'tfoot'], ['table']], tfoot: [['thead', 'tbody', 'tfoot'], ['table']],
    dt: [['dt', 'dd'], ['dl']], dd: [['dt', 'dd'], ['dl']]
  };
  function autoClose(stack, tag) {
    var r = AUTOCLOSE[tag], i, t;
    if (r) {
      for (i = stack.length - 1; i > 0; i--) {
        t = stack[i].tag;
        if (r[0].indexOf(t) >= 0) { stack.length = i; return; }
        if (r[1].indexOf(t) >= 0) return;
      }
      return;
    }
    if (BLOCK_START[tag]) {
      for (i = stack.length - 1; i > 0; i--) {
        t = stack[i].tag;
        if (t === 'p') { stack.length = i; return; }
        if (!INLINE_T[t]) return;
      }
    }
  }

  function parseAttrs(str) {
    var attrs = {}, re = /([^\s=\/"'<>]+)(?:\s*=\s*(?:"([^"]*)"|'([^']*)'|([^\s"'>]+)))?/g, m;
    while ((m = re.exec(str))) {
      var v = m[2] != null ? m[2] : m[3] != null ? m[3] : m[4] != null ? m[4] : '';
      attrs[m[1].toLowerCase()] = decodeEnt(v);
    }
    return attrs;
  }

  /** HTML 문자열 -> {tag:'#root', kids:[...]} (텍스트 노드는 {text}) */
  function parseHtml(html) {
    html = clean(html);
    if (html.length > LIM.html) html = html.slice(0, LIM.html);
    var rootNode = mkEl('#root', {});
    var stack = [rootNode], skipped = {}, i = 0, n = html.length;
    var tagRe = /<(\/?)([A-Za-z][A-Za-z0-9:-]*)((?:"[^"]*"|'[^']*'|[^>"'])*)>/y;
    function addText(t) {
      if (!t) return;
      stack[stack.length - 1].kids.push({ text: decodeEnt(t) });
    }
    while (i < n) {
      var lt = html.indexOf('<', i);
      if (lt < 0) { addText(html.slice(i)); break; }
      if (lt > i) addText(html.slice(i, lt));
      if (html.substr(lt, 4) === '<!--') {
        var ce = html.indexOf('-->', lt + 4);
        i = ce < 0 ? n : ce + 3; continue;
      }
      var c1 = html.charAt(lt + 1);
      if (c1 === '!' || c1 === '?') {
        var ge = html.indexOf('>', lt);
        i = ge < 0 ? n : ge + 1; continue;
      }
      tagRe.lastIndex = lt;
      var m = tagRe.exec(html);
      if (!m) { addText('<'); i = lt + 1; continue; }
      i = lt + m[0].length;
      var closing = m[1] === '/', tag = m[2].toLowerCase();
      if (closing) {
        if (skipped[tag] > 0) { skipped[tag]--; continue; }
        for (var k = stack.length - 1; k > 0; k--) {
          if (stack[k].tag === tag) { stack.length = k; break; }
        }
        continue;
      }
      var selfClose = /\/\s*$/.test(m[3]);
      var attrs = parseAttrs(m[3].replace(/\/\s*$/, ''));
      if (RAWTEXT[tag]) {
        var re2 = new RegExp('</' + tag + '\\b[^>]*>', 'i');
        var rest = html.slice(i), mm = re2.exec(rest);
        i = mm ? i + mm.index + mm[0].length : n;
        continue;
      }
      autoClose(stack, tag);
      var el = mkEl(tag, attrs);
      if (VOID[tag] || selfClose) { stack[stack.length - 1].kids.push(el); continue; }
      if (stack.length >= LIM.depth) { skipped[tag] = (skipped[tag] || 0) + 1; continue; }
      stack[stack.length - 1].kids.push(el);
      stack.push(el);
    }
    return rootNode;
  }

  /** 브라우저 DOM(Element / ShadowRoot / DocumentFragment) -> 같은 트리 */
  function fromDom(node, depth) {
    depth = depth || 0;
    var t = node.nodeType;
    if (t === 3 || t === 4) return { text: clean(node.nodeValue || node.data || '') };
    if (t !== 1 && t !== 11 && t !== 9) return null;
    var tag = t === 1 ? String(node.tagName || '').toLowerCase() : '#root';
    var attrs = {};
    if (t === 1) {
      var cl = node.getAttribute ? node.getAttribute('class') : null;
      if (cl) attrs['class'] = cl;
      ['data-w', 'colspan', 'rowspan', 'start'].forEach(function (a) {
        var v = node.getAttribute ? node.getAttribute(a) : null;
        if (v != null) attrs[a] = v;
      });
    }
    var el = mkEl(tag, attrs);
    if (RAWTEXT[tag] || depth > LIM.depth) return el;
    var cn = node.childNodes || [];
    for (var i = 0; i < cn.length; i++) {
      var k = fromDom(cn[i], depth + 1);
      if (k) el.kids.push(k);
    }
    return el;
  }

  function toTree(input) {
    if (input && typeof input === 'object' && typeof input.nodeType === 'number') return fromDom(input, 0);
    if (input && typeof input === 'object' && Array.isArray(input.kids)) return input;
    return parseHtml(input);
  }

  /* ------------------------------------------------------------------ */
  /* 표 -> 2차원 배열 (xlsx 용)                                            */
  /* ------------------------------------------------------------------ */

  var SKIPT = { script: 1, style: 1, head: 1, svg: 1, iframe: 1, object: 1, embed: 1, noscript: 1, template: 1, img: 1, input: 1, button: 1, canvas: 1, video: 1, audio: 1, meta: 1, link: 1, '#comment': 1 };
  var BLOCK_TAGS = {
    div: 1, section: 1, article: 1, header: 1, footer: 1, aside: 1, h1: 1, h2: 1, h3: 1, h4: 1, h5: 1, h6: 1,
    p: 1, ul: 1, ol: 1, li: 1, table: 1, thead: 1, tbody: 1, tfoot: 1, tr: 1, td: 1, th: 1, blockquote: 1,
    dl: 1, dt: 1, dd: 1, hr: 1, figure: 1, figcaption: 1, pre: 1, address: 1, nav: 1, main: 1, form: 1,
    fieldset: 1, caption: 1, details: 1, summary: 1, center: 1, body: 1, html: 1
  };
  var BLOCK_CLS = ['kicker', 'sub', 'meta', 'n', 'qt', 'big', 'src', 'pg', 'ans', 'pt'];

  /** 셀 안의 글만 뽑는다 (줄바꿈 보존). */
  function cellText(node) {
    var out = '';
    function nl() { if (out && out.charAt(out.length - 1) !== '\n') out += '\n'; }
    function walk(n, listBullet) {
      if (n.text != null) { out += n.text.replace(/[ \t\r\n\f]+/g, ' '); return; }
      if (SKIPT[n.tag] || hc(n, 'bar')) return;
      if (n.tag === 'br') { out += '\n'; return; }
      var block = BLOCK_TAGS[n.tag] || BLOCK_CLS.some(function (c) { return hc(n, c); });
      if (block) nl();
      if (n.tag === 'li') out += listBullet === 'ol' ? '' : '\u2022 ';
      var idx = 0;
      for (var i = 0; i < n.kids.length; i++) {
        var k = n.kids[i];
        if (n.tag === 'ol' && k.tag === 'li') { idx++; nl(); out += idx + '. '; walkKids(k); nl(); continue; }
        walk(k, n.tag);
      }
      if (block) nl();
    }
    function walkKids(n) { n.kids.forEach(function (k) { walk(k, 'li'); }); }
    walk(node, '');
    return clean(out).split('\n').map(function (l) { return l.replace(/^[ \u00a0]+|[ \u00a0]+$/g, ''); })
      .filter(function (l, i, a) { return l || (i > 0 && i < a.length - 1 && a[i - 1] !== ''); })
      .join('\n').replace(/^\n+|\n+$/g, '');
  }
  function plainText(node) {
    if (node.text != null) return node.text;
    if (SKIPT[node.tag] || hc(node, 'bar')) return '';
    var s = '';
    for (var i = 0; i < node.kids.length; i++) s += plainText(node.kids[i]) + (BLOCK_TAGS[node.kids[i].tag] ? ' ' : '');
    return s;
  }
  function tableRows(table) {
    var trs = [];
    (function collect(n) {
      n.kids.forEach(function (k) {
        if (k.tag === 'tr') trs.push(k);
        else if (k.tag === 'thead' || k.tag === 'tbody' || k.tag === 'tfoot') collect(k);
      });
    })(table);
    var rows = [], pending = [];
    trs.slice(0, LIM.tableRows).forEach(function (tr) {
      var row = [], col = 0;
      function fillPending() {
        while (pending[col] && pending[col].left > 0) { row.push(pending[col].text); pending[col].left--; col++; }
      }
      tr.kids.forEach(function (c) {
        if (c.tag !== 'td' && c.tag !== 'th') return;
        fillPending();
        var txt = cellText(c);
        var cs = Math.min(LIM.tableCols, Math.max(1, parseInt(c.attrs.colspan, 10) || 1));
        var rs = Math.min(50, Math.max(1, parseInt(c.attrs.rowspan, 10) || 1));
        for (var j = 0; j < cs; j++) {
          row.push(j === 0 ? txt : '');
          if (rs > 1) pending[col] = { text: j === 0 ? txt : '', left: rs - 1 };
          col++;
        }
      });
      fillPending();
      rows.push(row.slice(0, LIM.tableCols * 3));
    });
    return rows;
  }

  var HEADING_TAGS = { h1: 1, h2: 1, h3: 1, h4: 1 };
  function tablesFromElement(el) {
    var roots = [];
    if (el && typeof el === 'object' && el.shadowRoot) roots.push(toTree(el.shadowRoot));
    roots.push(toTree(el));
    var out = [], seen = 0;
    roots.forEach(function (tree) {
      var last = '', found = [];
      (function walk(n) {
        if (n.text != null || SKIPT[n.tag]) return;
        if (HEADING_TAGS[n.tag]) { var t = plainText(n).replace(/\s+/g, ' ').trim(); if (t) last = t; return; }
        if (n.tag === 'table') {
          var cap = n.kids.filter(function (k) { return k.tag === 'caption'; })[0];
          var capT = cap ? plainText(cap).replace(/\s+/g, ' ').trim() : '';
          var rows = tableRows(n);
          if (rows.length) found.push({ name: capT || last || '', rows: rows });
          return;
        }
        n.kids.forEach(walk);
      })(tree);
      if (found.length) { out = out.concat(found); }
    });
    return out.slice(0, LIM.sheets).map(function (t) {
      seen++;
      return { name: t.name || ('표 ' + seen), rows: t.rows };
    });
  }

  function normHead(s) { return String(s == null ? '' : s).replace(/[\s\u00a0]+/g, '').toLowerCase(); }
  /** 첫 행이 강사용 퀴즈 표(문항 | 보기1~4 | 정답 | 해설)인지. */
  function quizRows(rows) {
    if (rows && !Array.isArray(rows) && Array.isArray(rows.rows)) rows = rows.rows;
    if (!Array.isArray(rows) || !Array.isArray(rows[0])) return false;
    var h = rows[0].map(normHead);
    if (h.length < 7) return false;
    if (!/^(문항|문제)/.test(h[0])) return false;
    for (var i = 1; i <= 4; i++) if (!new RegExp('^보기' + i + '$').test(h[i])) return false;
    return /^정답/.test(h[5]) && /^해설/.test(h[6]);
  }

  /* ------------------------------------------------------------------ */
  /* DOCX                                                                */
  /* ------------------------------------------------------------------ */

  var RED = 'E54536', INK = '1A1A1A', GRAY = '6F6F6F', LGRAY = 'BDBDBD', RULE = 'D2D2D2', PAPER = 'F2F2F2';
  var CIRCLED = ['\u2460', '\u2461', '\u2462', '\u2463', '\u2464', '\u2465', '\u2466', '\u2467', '\u2468', '\u2469'];
  var LIST_STEP = 420;

  function rprXml(rp) {
    if (!rp) return '';
    var s = '';
    if (rp.font) s += '<w:rFonts w:ascii="' + rp.font + '" w:hAnsi="' + rp.font + '" w:cs="' + rp.font + '"/>';
    if (rp.b) s += '<w:b/><w:bCs/>';
    if (rp.i) s += '<w:i/><w:iCs/>';
    if (rp.caps) s += '<w:caps/>';
    if (rp.strike) s += '<w:strike/>';
    if (rp.color) s += '<w:color w:val="' + rp.color + '"/>';
    if (rp.spacing) s += '<w:spacing w:val="' + rp.spacing + '"/>';
    if (rp.size) s += '<w:sz w:val="' + rp.size + '"/><w:szCs w:val="' + rp.size + '"/>';
    if (rp.hl) s += '<w:highlight w:val="' + rp.hl + '"/>';
    if (rp.u) s += '<w:u w:val="single"/>';
    if (rp.bdr) s += '<w:bdr w:val="single" w:sz="4" w:space="0" w:color="' + rp.bdr + '"/>';
    if (rp.shd) s += '<w:shd w:val="clear" w:color="auto" w:fill="' + rp.shd + '"/>';
    if (rp.vert) s += '<w:vertAlign w:val="' + rp.vert + '"/>';
    return s ? '<w:rPr>' + s + '</w:rPr>' : '';
  }
  function runXml(text, rp) {
    return '<w:r>' + rprXml(rp) + '<w:t xml:space="preserve">' + esc(text) + '</w:t></w:r>';
  }
  function brXml(rp) { return '<w:r>' + rprXml(rp) + '<w:br/></w:r>'; }

  function pprXml(p) {
    var s = '';
    if (p.style) s += '<w:pStyle w:val="' + p.style + '"/>';
    if (p.keepNext) s += '<w:keepNext/>';
    if (p.keepLines) s += '<w:keepLines/>';
    if (p.pageBreakBefore) s += '<w:pageBreakBefore/>';
    if (p.numId != null) s += '<w:numPr><w:ilvl w:val="' + p.ilvl + '"/><w:numId w:val="' + p.numId + '"/></w:numPr>';
    if (p.bdr) {
      var b = '';
      ['top', 'left', 'bottom', 'right', 'between'].forEach(function (k) {
        var d = p.bdr[k];
        if (d) b += '<w:' + k + ' w:val="' + (d.val || 'single') + '" w:sz="' + d.sz + '" w:space="' + (d.space || 0) + '" w:color="' + d.color + '"/>';
      });
      if (b) s += '<w:pBdr>' + b + '</w:pBdr>';
    }
    if (p.shd) s += '<w:shd w:val="clear" w:color="auto" w:fill="' + p.shd + '"/>';
    if (p.tabs) {
      s += '<w:tabs>' + p.tabs.map(function (t) { return '<w:tab w:val="left" w:leader="underscore" w:pos="' + t + '"/>'; }).join('') + '</w:tabs>';
    }
    if (p.before != null || p.after != null || p.line != null) {
      s += '<w:spacing' + (p.before != null ? ' w:before="' + p.before + '"' : '') +
        (p.after != null ? ' w:after="' + p.after + '"' : '') +
        (p.line != null ? ' w:line="' + p.line + '" w:lineRule="' + (p.rule || 'auto') + '"' : '') + '/>';
    }
    if (p.indL != null || p.indR != null || p.hang != null) {
      s += '<w:ind' + (p.indL != null ? ' w:left="' + p.indL + '"' : '') + (p.indR != null ? ' w:right="' + p.indR + '"' : '') +
        (p.hang != null ? ' w:hanging="' + p.hang + '"' : '') + '/>';
    }
    if (p.jc) s += '<w:jc w:val="' + p.jc + '"/>';
    return s ? '<w:pPr>' + s + '</w:pPr>' : '';
  }

  /* 블록 표현: {p:1, runs:[xml], props} | {x:xml} | {pb:1} */
  function paraXml(b) { return '<w:p>' + pprXml(b.props) + b.runs.join('') + '</w:p>'; }
  function spacer(line) { return { p: 1, runs: [], props: { before: 0, after: 0, line: line || 120, rule: 'exact' } }; }

  function blocksXml(blocks, inCell) {
    var out = '', pendingPb = false;
    blocks.forEach(function (b) {
      if (b.pb) { if (!inCell) pendingPb = true; return; }
      if (pendingPb) {
        out += paraXml({ runs: [], props: { pageBreakBefore: true, before: 0, after: 0, line: 20, rule: 'exact' } });
        pendingPb = false;
      }
      out += b.x != null ? b.x : paraXml(b);
    });
    if (inCell) {
      var last = blocks.filter(function (b) { return !b.pb; }).pop();
      if (!last || last.x != null) out += '<w:p><w:pPr><w:spacing w:before="0" w:after="0"/></w:pPr></w:p>';
    }
    return out;
  }

  function bdrLine(sz, color, space, val) { return { val: val || 'single', sz: sz, space: space == null ? 4 : space, color: color }; }

  /* 인라인 항목 -> 런 xml. 공백 정리, 길이 상한. */
  function buildRuns(items, cap) {
    var segs = [], total = 0, lastSpace = true, visible = false, truncated = false, i;
    for (i = 0; i < items.length && !truncated; i++) {
      var it = items[i];
      if (it.br) { segs.push({ br: true, rp: it.rp }); lastSpace = true; continue; }
      if (it.tab) { segs.push({ tab: true, rp: it.rp }); visible = true; lastSpace = true; continue; }
      var t = clean(it.t).replace(/[ \t\r\n\f]+/g, ' ');
      if (lastSpace && t.charAt(0) === ' ') t = t.slice(1);
      if (!t) continue;
      if (total + t.length > cap) { t = cutStr(t, Math.max(0, cap - total - 1)) + '\u2026'; truncated = true; }
      total += t.length;
      lastSpace = t.charAt(t.length - 1) === ' ';
      if (/[^ ]/.test(t)) visible = true;
      segs.push({ t: t, rp: it.rp });
    }
    if (!visible) return null;
    // 끝의 줄바꿈/공백 제거
    while (segs.length && (segs[segs.length - 1].br || segs[segs.length - 1].t === ' ')) segs.pop();
    var last = segs[segs.length - 1];
    if (last && last.tab) last = null;
    if (last && last.t != null && last.t.charAt(last.t.length - 1) === ' ') last.t = last.t.slice(0, -1);
    while (segs.length && segs[0].br) segs.shift();
    return segs.map(function (s) { return s.br ? brXml(s.rp) : s.tab ? '<w:r>' + rprXml(s.rp) + '<w:tab/></w:r>' : runXml(s.t, s.rp); });
  }

  function makePara(items, ctx, extra) {
    extra = extra || {};
    var runs = buildRuns(items, extra.cap || LIM.para);
    if (!runs) return null;
    var p = clone(ctx.ppr || {});
    if (ctx.ppr && ctx.ppr.bdr) p.bdr = clone(ctx.ppr.bdr);
    for (var k in extra) {
      if (k === 'cap') continue;
      if (k === 'bdr') p.bdr = clone(p.bdr, extra.bdr); else p[k] = extra[k];
    }
    var ps = ctx.pspace;
    if (ps) {
      if (p.before == null) p.before = ps.before;
      if (p.after == null) p.after = ps.after;
      if (p.line == null && ps.line) p.line = ps.line;
    }
    if (ctx.cell && p.style && /^(Heading|Title)/.test(p.style) && extra.before == null) { p.before = 40; p.after = 60; }
    return { p: 1, runs: runs, props: p };
  }

  /* ---- 인라인 ---- */
  function inlineClass(k, ctx, rp) {
    // 반환: {rp, pre:[items], post:[items]} 또는 null
    if (hc(k, 'tag')) {
      var r;
      if (hc(k, 'red')) r = clone(rp, { b: true, size: 18, color: 'FFFFFF', shd: RED, bdr: RED });
      else if (hc(k, 'gray')) r = clone(rp, { b: true, size: 18, color: GRAY, bdr: LGRAY });
      else r = clone(rp, { b: true, size: 18, color: ctx.dark ? 'FFFFFF' : INK, bdr: ctx.dark ? 'FFFFFF' : INK });
      return { rp: r, pre: [{ t: '\u00a0', rp: r }], post: [{ t: '\u00a0', rp: r }, { t: ' ', rp: rp }] };
    }
    if (hc(k, 'no')) return { rp: clone(rp, { b: true, size: 18, color: RED, spacing: 30, font: 'Arial' }), pre: [], post: [{ t: ' ', rp: rp }] };
    if (hc(k, 'fill')) {
      var u = clone(rp, { u: true });
      return { rp: u, pre: [{ t: repeat('\u00a0', 6), rp: u }], post: [{ t: repeat('\u00a0', 6), rp: u }] };
    }
    if (hc(k, 'mute')) return { rp: clone(rp, { color: ctx.dark ? LGRAY : GRAY }), pre: [], post: [] };
    if (hc(k, 'sm')) return { rp: clone(rp, { size: 18, color: ctx.dark ? LGRAY : GRAY }), pre: [], post: [] };
    return null;
  }
  function inline(kids, ctx, rp, out) {
    for (var i = 0; i < kids.length; i++) {
      var k = kids[i];
      if (k.text != null) { out.push({ t: k.text, rp: rp }); continue; }
      var tag = k.tag;
      if (SKIPT[tag] || hc(k, 'bar')) continue;
      if (tag === 'br') { out.push({ br: true, rp: rp }); continue; }
      var r = rp, pre = [], post = [], ic;
      switch (tag) {
        case 'b': case 'strong':
          r = clone(rp, { b: true });
          if (ctx.bcolor) r.color = ctx.bcolor;
          if (ctx.bpost) post = [{ t: ' ', rp: rp }];
          break;
        case 'i': case 'em': case 'cite': case 'dfn': case 'var': r = clone(rp, { i: true }); break;
        case 'small': r = clone(rp, { size: 18, color: ctx.dark ? LGRAY : GRAY }); break;
        case 'mark': r = clone(rp, { hl: 'yellow' }); break;
        case 'u': case 'ins': r = clone(rp, { u: true }); break;
        case 's': case 'del': case 'strike': r = clone(rp, { strike: true }); break;
        case 'sub': r = clone(rp, { vert: 'subscript' }); break;
        case 'sup': r = clone(rp, { vert: 'superscript' }); break;
        default:
      }
      if ((ic = inlineClass(k, ctx, r))) { r = ic.rp; pre = pre.concat(ic.pre); post = ic.post.concat(post); }
      else if (hc(k, 'kicker')) r = clone(r, { b: true, size: 18, color: RED, spacing: 40, font: 'Arial' });
      if (BLOCK_TAGS[tag] && out.length) out.push({ br: true, rp: rp });
      for (var j = 0; j < pre.length; j++) out.push(pre[j]);
      inline(k.kids, ctx, r, out);
      for (j = 0; j < post.length; j++) out.push(post[j]);
      if (BLOCK_TAGS[tag] && out.length) out.push({ br: true, rp: rp });
    }
  }
  function inlineItems(kids, ctx, rp) { var o = []; inline(kids, ctx, rp || ctx.rpr || {}, o); return o; }
  function nodeItems(n, ctx, rp) { var o = []; inline([n], ctx, rp || ctx.rpr || {}, o); return o; }

  function isBlockNode(n, ctx) {
    if (BLOCK_TAGS[n.tag]) return true;
    if (ctx.blockTags && ctx.blockTags[n.tag]) return true;
    for (var i = 0; i < BLOCK_CLS.length; i++) if (hc(n, BLOCK_CLS[i])) return true;
    return false;
  }

  /* ---- 블록 ---- */
  function renderKids(kids, ctx) {
    var out = [], run = [];
    function flush() {
      if (!run.length) return;
      var p = makePara(run, ctx);
      if (p) out.push(p);
      run = [];
    }
    for (var i = 0; i < kids.length; i++) {
      var k = kids[i];
      if (k.text != null) { run.push({ t: k.text, rp: ctx.rpr }); continue; }
      if (SKIPT[k.tag] || hc(k, 'bar')) continue;
      if (isBlockNode(k, ctx)) {
        flush();
        var r = renderBlock(k, ctx);
        for (var j = 0; j < r.length; j++) out.push(r[j]);
      } else {
        inline([k], ctx, ctx.rpr || {}, run);
      }
    }
    flush();
    return out;
  }

  function cellCtx(ctx, width, extra) {
    return clone(ctx, clone({ cell: true, ppr: {}, pspace: { before: 40, after: 40, line: 300 }, width: width, blockTags: null }, extra));
  }

  function tcBordersXml(b) {
    if (!b) return '';
    var s = '';
    ['top', 'left', 'bottom', 'right'].forEach(function (k) {
      var d = b[k];
      s += d ? '<w:' + k + ' w:val="single" w:sz="' + d.sz + '" w:space="0" w:color="' + d.color + '"/>' : '<w:' + k + ' w:val="nil"/>';
    });
    return '<w:tcBorders>' + s + '</w:tcBorders>';
  }
  /* rows: [{header, cells:[{w, span, blocks, fill, vm, borders, mar, valign}]}] */
  function tblXml(widths, rows, o) {
    o = o || {};
    var total = widths.reduce(function (a, b) { return a + b; }, 0);
    var s = '<w:tbl><w:tblPr><w:tblW w:w="' + total + '" w:type="dxa"/><w:tblInd w:w="0" w:type="dxa"/>';
    if (o.grid) {
      var bd = '<w:{k} w:val="single" w:sz="4" w:space="0" w:color="BFBFBF"/>';
      s += '<w:tblBorders>' + ['top', 'left', 'bottom', 'right', 'insideH', 'insideV'].map(function (k) { return bd.replace('{k}', k); }).join('') + '</w:tblBorders>';
    }
    s += '<w:tblLayout w:type="fixed"/><w:tblCellMar><w:top w:w="' + (o.padV != null ? o.padV : 40) + '" w:type="dxa"/><w:left w:w="' + (o.padH != null ? o.padH : 100) +
      '" w:type="dxa"/><w:bottom w:w="' + (o.padV != null ? o.padV : 40) + '" w:type="dxa"/><w:right w:w="' + (o.padH != null ? o.padH : 100) + '" w:type="dxa"/></w:tblCellMar></w:tblPr>';
    s += '<w:tblGrid>' + widths.map(function (w) { return '<w:gridCol w:w="' + w + '"/>'; }).join('') + '</w:tblGrid>';
    rows.forEach(function (r) {
      s += '<w:tr><w:trPr><w:cantSplit/>' + (r.header ? '<w:tblHeader/>' : '') + '</w:trPr>';
      r.cells.forEach(function (c) {
        s += '<w:tc><w:tcPr><w:tcW w:w="' + c.w + '" w:type="dxa"/>' + (c.span > 1 ? '<w:gridSpan w:val="' + c.span + '"/>' : '') +
          (c.vm === 'restart' ? '<w:vMerge w:val="restart"/>' : c.vm === 'cont' ? '<w:vMerge/>' : '') +
          tcBordersXml(c.borders) +
          (c.fill ? '<w:shd w:val="clear" w:color="auto" w:fill="' + c.fill + '"/>' : '') +
          (c.mar ? '<w:tcMar><w:top w:w="' + c.mar.t + '" w:type="dxa"/><w:left w:w="' + c.mar.l + '" w:type="dxa"/><w:bottom w:w="' + c.mar.b + '" w:type="dxa"/><w:right w:w="' + c.mar.r + '" w:type="dxa"/></w:tcMar>' : '') +
          (c.valign ? '<w:vAlign w:val="' + c.valign + '"/>' : '') +
          '</w:tcPr>' + blocksXml(c.blocks || [], true) + '</w:tc>';
      });
      s += '</w:tr>';
    });
    return s + '</w:tbl>';
  }

  /* 한 칸짜리 표로 만든 상자 (card, callout, sheet, box) */
  function boxBlock(n, ctx, o) {
    var mar = o.mar || { t: 100, l: 160, b: 100, r: 160 };
    var inner = ctx.width - mar.l - mar.r;
    var ic = clone(ctx, {
      cell: true, ppr: {}, pspace: { before: 0, after: 80, line: 320 }, width: inner, blockTags: null,
      dark: !!o.dark, rpr: o.dark ? clone(ctx.rpr, { color: 'FFFFFF' }) : ctx.rpr, hcolor: o.dark ? 'FFFFFF' : ctx.hcolor
    });
    var blocks = o.build ? o.build(ic) : renderKids(n.kids, ic);
    if (o.minParas) {
      while (blocks.filter(function (b) { return b.p; }).length < o.minParas) {
        blocks.push({ p: 1, runs: [], props: { before: 0, after: 0, line: 480, rule: 'exact' } });
      }
    }
    var x = tblXml([ctx.width], [{ cells: [{ w: ctx.width, span: 1, blocks: blocks, fill: o.fill, borders: o.borders, mar: mar }] }], { padH: 0, padV: 0 });
    return [{ x: x }, spacer(o.gap || 140)];
  }

  function colWeights(table, ncols, firstCol) {
    var stats = [];
    for (var i = 0; i < ncols; i++) stats.push({ max: 0, sum: 0, n: 0 });
    table.rows.forEach(function (r) {
      r.entries.forEach(function (e) {
        if (!e.cell || e.span !== 1) return;
        var w = dispWidth(plainText(e.cell).replace(/\s+/g, ' ').trim());
        var s = stats[e.col];
        s.max = Math.max(s.max, w); s.sum += w; s.n++;
      });
    });
    return stats.map(function (s) {
      var w = s.n ? (0.6 * s.max + 0.4 * (s.sum / s.n)) : 6;
      return Math.max(6, Math.min(45, w));
    });
  }
  function fitWidths(weights, total, fixedFirst) {
    var n = weights.length, minW = Math.min(650, Math.floor(total / n));
    var w = new Array(n), i, sum;
    var start = 0, avail = total;
    if (fixedFirst && n > 1) { w[0] = Math.max(minW, Math.round(total * fixedFirst)); avail = total - w[0]; start = 1; }
    var wsum = 0;
    for (i = start; i < n; i++) wsum += weights[i];
    for (i = start; i < n; i++) w[i] = Math.max(minW, Math.round(avail * weights[i] / wsum));
    sum = w.reduce(function (a, b) { return a + b; }, 0);
    var big = 0;
    for (i = 1; i < n; i++) if (w[i] > w[big]) big = i;
    w[big] += total - sum;
    return w;
  }

  function dataTable(n, ctx) {
    // 1) 행 수집
    var trs = [];
    (function collect(p, inHead) {
      p.kids.forEach(function (k) {
        if (k.tag === 'tr') trs.push({ tr: k, head: inHead });
        else if (k.tag === 'thead') collect(k, true);
        else if (k.tag === 'tbody' || k.tag === 'tfoot') collect(k, inHead);
      });
    })(n, false);
    trs = trs.slice(0, LIM.tableRows);
    var blocks = [];
    var cap = n.kids.filter(function (k) { return k.tag === 'caption'; })[0];
    if (cap) {
      var cp = makePara(inlineItems(cap.kids, ctx, clone(ctx.rpr, { b: true, size: 20 })), ctx, { keepNext: true, after: 60 });
      if (cp) blocks.push(cp);
    }
    if (!trs.length) return blocks;
    // 2) 격자 배치
    var occ = [], conts = [], ncols = 0, rows = [];
    trs.forEach(function (t, ri) {
      occ[ri] = occ[ri] || []; conts[ri] = conts[ri] || [];
      var entries = [], col = 0;
      conts[ri].forEach(function (c) { entries.push(c); });
      t.tr.kids.forEach(function (c) {
        if (c.tag !== 'td' && c.tag !== 'th') return;
        while (occ[ri][col]) col++;
        var cs = Math.min(LIM.tableCols, Math.max(1, parseInt(c.attrs.colspan, 10) || 1));
        var rs = Math.min(50, Math.max(1, parseInt(c.attrs.rowspan, 10) || 1));
        if (ri + rs > trs.length) rs = trs.length - ri;
        for (var j = 0; j < cs; j++) occ[ri][col + j] = true;
        for (var r2 = 1; r2 < rs; r2++) {
          occ[ri + r2] = occ[ri + r2] || []; conts[ri + r2] = conts[ri + r2] || [];
          for (j = 0; j < cs; j++) occ[ri + r2][col + j] = true;
          conts[ri + r2].push({ col: col, span: cs, cont: true });
        }
        entries.push({ col: col, span: cs, cell: c, rs: rs, th: c.tag === 'th' });
        col += cs;
      });
      entries.sort(function (a, b) { return a.col - b.col; });
      entries.forEach(function (e) { ncols = Math.max(ncols, e.col + e.span); });
      var allTh = entries.length && entries.every(function (e) { return e.cell && e.th; });
      rows.push({ entries: entries, header: t.head || (ri === 0 && allTh) });
    });
    ncols = Math.min(ncols, LIM.tableCols);
    if (!ncols) return blocks;
    var rubric = hc(n, 'rubric');
    var widths = fitWidths(colWeights({ rows: rows }, ncols), ctx.width, rubric && ncols > 2 ? 0.17 : 0);
    // 3) 셀 렌더
    var outRows = [], bodyIdx = 0;
    rows.forEach(function (r) {
      var cells = [], used = 0;
      if (!r.header) bodyIdx++;
      r.entries.forEach(function (e) {
        if (e.col >= ncols) return;
        var span = Math.min(e.span, ncols - e.col);
        var w = 0;
        for (var j = 0; j < span; j++) w += widths[e.col + j];
        used = e.col + span;
        if (e.cont) { cells.push({ w: w, span: span, vm: 'cont', blocks: [] }); return; }
        var isHead = r.header || e.th && r.header;
        var rp = clone(ctx.rpr, { size: 20 });
        var fill;
        if (r.header) { rp.b = true; rp.color = 'FFFFFF'; fill = INK; }
        else {
          if (bodyIdx % 2 === 0) fill = 'FAFAFA';
          if (rubric && e.col === 0) { rp.b = true; fill = PAPER; }
          else if (e.th) rp.b = true;
        }
        var cc = cellCtx(ctx, w - 200, { rpr: rp, dark: !!r.header, hcolor: r.header ? 'FFFFFF' : null });
        cells.push({ w: w, span: span, vm: e.rs > 1 ? 'restart' : null, fill: fill, blocks: renderKids(e.cell.kids, cc), valign: isHead ? 'bottom' : null });
      });
      if (used < ncols) {
        var rest = 0;
        for (var q = used; q < ncols; q++) rest += widths[q];
        cells.push({ w: rest, span: ncols - used, blocks: [] });
      }
      outRows.push({ header: r.header, cells: cells });
    });
    blocks.push({ x: tblXml(widths, outRows, { grid: true }) });
    blocks.push(spacer(140));
    return blocks;
  }

  function colsTable(n, ctx, ncols) {
    var kids = n.kids.filter(function (k) { return k.tag && !SKIPT[k.tag] && !hc(k, 'bar'); }).slice(0, 60);
    if (!kids.length) return [];
    var colW = Math.floor(ctx.width / ncols);
    var widths = [];
    for (var i = 0; i < ncols; i++) widths.push(colW);
    var rows = [];
    for (var s = 0; s < kids.length; s += ncols) {
      var cells = [];
      for (var j = 0; j < ncols; j++) {
        var k = kids[s + j];
        var cc = cellCtx(ctx, colW - 200, { pspace: { before: 0, after: 80, line: 320 } });
        cells.push({ w: colW, span: 1, blocks: k ? renderKids([k], cc) : [] });
      }
      rows.push({ cells: cells });
    }
    return [{ x: tblXml(widths, rows, { padH: 100, padV: 40 }) }, spacer(120)];
  }

  function listBlocks(n, ctx) {
    var st = ctx.st, ordered = n.tag === 'ol', lvl = Math.min(ctx.lvl || 0, 8), numId;
    if (ordered) {
      var start = parseInt(n.attrs.start, 10);
      st.nums.push({ start: start > 0 && start < 10000 ? start : 1 });
      numId = st.nums.length + 1;
    } else numId = 1;
    var out = [];
    n.kids.slice(0, 400).forEach(function (li) {
      if (li.tag !== 'li') return;
      var base = (ctx.ppr && ctx.ppr.indL) || 0;
      var ic = clone(ctx, { lvl: lvl + 1 });
      var blocks = renderKids(li.kids, ic);
      var first = true, left = base + LIST_STEP * (lvl + 1);
      blocks.forEach(function (b) {
        if (b.p && b.props.numId == null && b.props.indL == null || b.p && b.props.indL === base && b.props.numId == null) {
          if (first && b.runs.length) {
            b.props.style = b.props.style || 'ListParagraph';
            b.props.numId = numId; b.props.ilvl = lvl;
            b.props.indL = left; b.props.hang = LIST_STEP;
            if (b.props.after == null) b.props.after = 40;
            first = false;
          } else { b.props.indL = left; }
        }
        out.push(b);
      });
    });
    out.push(spacer(60));
    return out;
  }

  function stepsBlocks(n, ctx) {
    var out = [], idx = 0;
    var bd = { top: bdrLine(4, RULE, 4), bottom: bdrLine(4, RULE, 4), between: bdrLine(4, RULE, 4) };
    n.kids.slice(0, 100).forEach(function (li) {
      if (li.tag !== 'li') return;
      idx++;
      var ic = clone(ctx, { ppr: clone(ctx.ppr, { bdr: bd, indL: 720, hang: 720 }), lvl: 0 });
      var blocks = renderKids(li.kids, ic);
      var num = (idx < 10 ? '0' : '') + idx;
      var pre = [runXml(num, { b: true, size: 28, color: RED, font: 'Arial' }), '<w:r><w:tab/></w:r>'];
      if (blocks.length && blocks[0].p) blocks[0].runs = pre.concat(blocks[0].runs);
      else blocks.unshift({ p: 1, runs: pre, props: clone(ic.ppr) });
      blocks.forEach(function (b) { if (b.p && b.props.indL == null) b.props.indL = 720; out.push(b); });
    });
    out.push(spacer(100));
    return out;
  }

  function optsBlocks(n, ctx) {
    var out = [], idx = 0;
    n.kids.forEach(function (li) {
      if (li.tag !== 'li') return;
      var items = [{ t: (CIRCLED[idx] || String(idx + 1) + '.') , rp: clone(ctx.rpr, { b: true }) }, { t: '\u00a0\u00a0', rp: ctx.rpr }].concat(inlineItems(li.kids, ctx));
      idx++;
      var p = makePara(items, ctx, { indL: 480, hang: 360, after: 40 });
      if (p) out.push(p);
    });
    return out;
  }

  function headingBlock(n, ctx) {
    var map = { h1: 'Title', h2: 'Heading1', h3: 'Heading2', h4: 'Heading3', h5: 'Heading3', h6: 'Heading3' };
    var rp = clone(ctx.rpr);
    if (ctx.hcolor) rp.color = ctx.hcolor;
    var p = makePara(inlineItems(n.kids, ctx, rp), ctx, { style: map[n.tag], cap: LIM.head, keepNext: true });
    if (!p) return [];
    if (n.tag === 'h1' && !ctx.st.title) ctx.st.title = plainText(n).replace(/\s+/g, ' ').trim();
    return [p];
  }

  function wkBlocks(n, ctx) {
    var col = hc(n, 'mid') || hc(n, 'fin') ? RED : null;
    var ic = clone(ctx, { hcolor: col, ppr: clone(ctx.ppr, { bdr: { left: bdrLine(16, INK, 10) }, indL: 360 }) });
    return renderKids(n.kids, ic);
  }

  function rowSpans(n, ctx, mode) {
    var items = [], rp0 = ctx.rpr || {}, first = true;
    var chips = n.kids.some(function (k) { return k.tag && k.tag !== 'b' && k.tag !== 'strong' && k.tag !== 'br' && !SKIPT[k.tag]; });
    var sep = mode === 'meta' ? { t: '   |   ', rp: clone(rp0, { color: LGRAY }) } : { t: '\u00a0\u00a0', rp: rp0 };
    n.kids.forEach(function (k) {
      if (k.text != null) { if (/\S/.test(k.text)) items.push({ t: k.text, rp: rp0 }); return; }
      if (SKIPT[k.tag] || hc(k, 'bar')) return;
      var isLabel = k.tag === 'b' || k.tag === 'strong';
      if ((chips && !isLabel) || (!chips && isLabel)) {
        if (!first) items.push(sep);
        first = false;
      }
      var sub = [];
      inline([k], ctx, rp0, sub);
      items = items.concat(sub);
      if (mode === 'name' && chips) items.push({ t: ' ', rp: rp0 }, { tab: true, rp: rp0 });
    });
    return items;
  }

  function renderBlock(n, ctx) {
    var out = renderBlock0(n, ctx);
    if (hc(n, 'pb')) out.push({ pb: 1 });
    return out;
  }

  function renderBlock0(n, ctx) {
    var t = n.tag, r, ic, p, out;
    if (SKIPT[t] || hc(n, 'bar')) return [];
    if (t === 'hr') return [{ p: 1, runs: [], props: { before: 120, after: 120, line: 20, rule: 'exact', bdr: { bottom: bdrLine(4, RULE, 1) } } }];
    if (t === 'table') return dataTable(n, ctx);
    if (t === 'ul' || t === 'ol') {
      if (hc(n, 'steps')) return stepsBlocks(n, ctx);
      if (hc(n, 'opts')) return optsBlocks(n, ctx);
      return listBlocks(n, ctx);
    }
    if (HEADING_BLOCK[t]) return headingBlock(n, ctx);

    // --- 클래스로 정해지는 구성 요소 ---
    if (hc(n, 'cols2')) return colsTable(n, ctx, 2);
    if (hc(n, 'cols3')) return colsTable(n, ctx, 3);
    if (hc(n, 'cols4')) return colsTable(n, ctx, 4);
    if (hc(n, 'card')) {
      var dark = hc(n, 'dark'), hl = hc(n, 'hl');
      var bc = dark ? INK : RULE;
      var bo = { top: { sz: 4, color: bc }, left: hl ? { sz: 32, color: RED } : { sz: 4, color: bc }, bottom: { sz: 4, color: bc }, right: { sz: 4, color: bc } };
      return boxBlock(n, ctx, { fill: dark ? INK : null, dark: dark, borders: bo });
    }
    if (hc(n, 'callout')) {
      var warn = hc(n, 'warn');
      return boxBlock(n, ctx, { fill: PAPER, borders: { left: { sz: 32, color: warn ? RED : INK } }, mar: { t: 120, l: 220, b: 120, r: 200 } });
    }
    if (hc(n, 'sheet')) {
      return boxBlock(n, ctx, { borders: { top: { sz: 12, color: INK }, left: { sz: 12, color: INK }, bottom: { sz: 12, color: INK }, right: { sz: 12, color: INK } }, mar: { t: 280, l: 360, b: 240, r: 360 }, gap: 160 });
    }
    if (hc(n, 'box')) {
      var bx = { sz: 6, color: INK };
      var hasText = /\S/.test(plainText(n));
      return boxBlock(n, ctx, { borders: { top: bx, left: bx, bottom: bx, right: bx }, mar: { t: 100, l: 160, b: 100, r: 160 }, minParas: hasText ? 2 : 3 });
    }
    if (hc(n, 'line')) {
      return [{ p: 1, runs: [], props: { before: 0, after: 0, line: 560, rule: 'exact', bdr: { bottom: bdrLine(6, INK, 1) } } }];
    }
    if (hc(n, 'cover')) {
      out = [{ p: 1, runs: [], props: { before: 0, after: 200, line: 20, rule: 'exact', bdr: { top: bdrLine(48, INK, 1) } } }];
      renderKids(n.kids, clone(ctx, { cover: true })).forEach(function (b) { out.push(b); });
      out.push({ p: 1, runs: [], props: { before: 120, after: 360, line: 20, rule: 'exact', bdr: { bottom: bdrLine(4, RULE, 1) } } });
      return out;
    }
    if (hc(n, 'lead')) {
      ic = clone(ctx, { rpr: clone(ctx.rpr, { size: 24, color: '333333' }), ppr: clone(ctx.ppr, { bdr: { left: bdrLine(32, RED, 8) }, indL: 240, after: 80 }) });
      return renderKids(n.kids, ic).concat([spacer(120)]);
    }
    if (hc(n, 'say')) {
      return sayBlocks(n, ctx);
    }
    if (hc(n, 'stat')) return statBlocks(n, ctx);
    if (hc(n, 'timeline')) {
      ic = clone(ctx, { ppr: clone(ctx.ppr, { bdr: { left: bdrLine(16, INK, 10) }, indL: 360 }) });
      return renderKids(n.kids, ic).concat([spacer(120)]);
    }
    if (hc(n, 'wk')) return wkBlocks(n, ctx);
    if (hc(n, 'slide')) return slideBlocks(n, ctx);
    if (hc(n, 'kicker')) {
      p = makePara(inlineItems(n.kids, ctx, clone(ctx.rpr, { b: true, size: 18, color: RED, spacing: 40, font: 'Arial' })), ctx, { after: 100, keepNext: true });
      return p ? [p] : [];
    }
    if (hc(n, 'sub')) {
      p = makePara(inlineItems(n.kids, ctx, clone(ctx.rpr, { size: 28, color: '444444' })), ctx, { after: 200 });
      return p ? [p] : [];
    }
    if (hc(n, 'meta')) {
      ic = clone(ctx, { rpr: clone(ctx.rpr, { color: GRAY, size: 21 }), bcolor: INK, bpost: true });
      p = makePara(rowSpans(n, ic, 'meta'), ic, { after: 60 });
      return p ? [p] : [];
    }
    if (hc(n, 'name')) {
      var nch = n.kids.filter(function (k) { return k.tag && !SKIPT[k.tag] && !hc(k, 'bar'); }).length || 1;
      var tabs = [];
      for (var ti = 1; ti <= nch; ti++) tabs.push(Math.floor(ctx.width * ti / nch) - (ti < nch ? 240 : 0));
      p = makePara(rowSpans(n, ctx, 'name'), ctx, { before: 60, after: 200, tabs: tabs });
      return p ? [p] : [];
    }
    if (hc(n, 'n')) {
      p = makePara(inlineItems(n.kids, ctx, clone(ctx.rpr, { b: true, size: 17, color: GRAY, spacing: 30, font: 'Arial' })), ctx, { after: 20, keepNext: true });
      return p ? [p] : [];
    }
    if (hc(n, 'pt')) {
      p = makePara(inlineItems(n.kids, ctx, clone(ctx.rpr, { b: true, size: 18, color: RED, font: 'Arial' })), ctx, { after: 20 });
      return p ? [p] : [];
    }
    if (hc(n, 'qt')) {
      p = makePara(inlineItems(n.kids, ctx, clone(ctx.rpr, { b: true })), ctx, { keepNext: true, before: 120, after: 80 });
      return p ? [p] : [];
    }
    if (hc(n, 'big')) {
      p = makePara(inlineItems(n.kids, ctx, clone(ctx.rpr, { b: true, size: 88, color: RED, font: 'Arial' })), ctx, { after: 60 });
      return p ? [p] : [];
    }
    if (hc(n, 'src') || hc(n, 'pg')) {
      if (hc(n, 'pg')) return [];
      p = makePara(inlineItems(n.kids, ctx, clone(ctx.rpr, { size: 18, color: GRAY })), ctx, {});
      return p ? [p] : [];
    }
    if (hc(n, 'ans')) {
      p = makePara(inlineItems(n.kids, ctx, clone(ctx.rpr, { size: 18, color: ctx.dark ? LGRAY : GRAY })), ctx, { before: 80, after: 60, bdr: { top: bdrLine(4, LGRAY, 4, 'dashed') } });
      return p ? [p] : [];
    }
    if (hc(n, 'foot')) {
      ic = clone(ctx, { rpr: clone(ctx.rpr, { size: 18, color: GRAY }), ppr: clone(ctx.ppr, { bdr: { top: bdrLine(4, RULE, 6) } }) });
      return [spacer(120)].concat(renderKids(n.kids, ic));
    }

    // --- 태그로 정해지는 구성 요소 ---
    if (t === 'blockquote') {
      ic = clone(ctx, { rpr: clone(ctx.rpr, { i: true }), ppr: clone(ctx.ppr, { bdr: { left: bdrLine(18, LGRAY, 8) }, indL: 360 }) });
      return renderKids(n.kids, ic).concat([spacer(80)]);
    }
    if (t === 'dt') {
      p = makePara(inlineItems(n.kids, ctx, clone(ctx.rpr, { b: true })), ctx, { keepNext: true, after: 20 });
      return p ? [p] : [];
    }
    if (t === 'dd') {
      ic = clone(ctx, { ppr: clone(ctx.ppr, { indL: ((ctx.ppr && ctx.ppr.indL) || 0) + 360 }) });
      return renderKids(n.kids, ic);
    }
    if (t === 'figcaption' || t === 'caption') {
      p = makePara(inlineItems(n.kids, ctx, clone(ctx.rpr, { size: 18, color: GRAY })), ctx, { after: 100 });
      return p ? [p] : [];
    }
    if (t === 'li') {
      p = makePara([{ t: '\u2022\u00a0', rp: ctx.rpr }].concat(inlineItems(n.kids, ctx)), ctx, {});
      return p ? [p] : [];
    }
    if (t === 'tr' || t === 'td' || t === 'th' || t === 'thead' || t === 'tbody' || t === 'tfoot') return renderKids(n.kids, ctx);
    r = renderKids(n.kids, ctx);
    return r;
  }
  var HEADING_BLOCK = { h1: 1, h2: 1, h3: 1, h4: 1, h5: 1, h6: 1 };

  function sayBlocks(n, ctx) {
    var ic = clone(ctx, {
      rpr: clone(ctx.rpr, { size: 22 }),
      ppr: clone(ctx.ppr, { bdr: { left: bdrLine(24, RED, 8) }, indL: 240, after: 60 })
    });
    var out = [], rest = [];
    function flushRest() { if (rest.length) { out = out.concat(renderKids(rest, ic)); rest = []; } }
    n.kids.forEach(function (k) {
      if (k.tag === 'b' || k.tag === 'strong') {
        flushRest();
        var p = makePara(inlineItems(k.kids, ic, clone(ic.rpr, { b: true, size: 17, color: GRAY, spacing: 20 })), ic, { after: 20, keepNext: true });
        if (p) out.push(p);
      } else rest.push(k);
    });
    flushRest();
    out.push(spacer(100));
    return out;
  }

  function statBlocks(n, ctx) {
    var out = [], rest = [];
    var labelCtx = clone(ctx, { rpr: clone(ctx.rpr, { size: 19, color: ctx.dark ? LGRAY : GRAY }) });
    function flushRest() { if (rest.length) { out = out.concat(renderKids(rest, labelCtx)); rest = []; } }
    n.kids.forEach(function (k) {
      if (k.tag === 'b' || k.tag === 'strong') {
        flushRest();
        var rp = clone(ctx.rpr, { b: true, size: 56, font: 'Arial', color: ctx.dark ? 'FFFFFF' : INK });
        var p = makePara(inlineItems(k.kids, ctx, rp), ctx, { before: 60, after: 20, keepNext: true, bdr: { top: bdrLine(24, INK, 6) } });
        if (p) out.push(p);
      } else rest.push(k);
    });
    flushRest();
    return out;
  }

  function slideBlocks(n, ctx) {
    var st = ctx.st;
    if (++st.slide > LIM.slides) return [];
    var hasHead = n.kids.some(function (k) { return k.tag === 'h1' || k.tag === 'h2' || k.tag === 'h3'; });
    var out = [];
    if (hasHead) {
      out.push(makePara([{ t: 'SLIDE ' + (st.slide < 10 ? '0' : '') + st.slide, rp: { b: true, size: 16, color: RED, spacing: 40, font: 'Arial' } }], ctx, { before: 240, after: 0, keepNext: true, bdr: { top: bdrLine(6, INK, 6) } }));
    } else {
      out.push(makePara([{ t: '\uc2ac\ub77c\uc774\ub4dc ' + st.slide, rp: ctx.rpr }], ctx, { style: 'Heading2', keepNext: true }));
    }
    var ic = clone(ctx, { slideDark: hc(n, 'dark') });
    return out.concat(renderKids(n.kids, ic)).concat([spacer(120)]);
  }

  /* ---- 문서 조립 ---- */
  function numberingXml(st) {
    var ab0 = '', ab1 = '', i;
    var fmts = ['decimal', 'lowerLetter', 'lowerRoman'];
    var bchars = ['\u2022', '\u2013', '\u2022'];
    for (i = 0; i < 9; i++) {
      var left = LIST_STEP * (i + 1);
      ab0 += '<w:lvl w:ilvl="' + i + '"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="' + bchars[i % 3] +
        '"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="' + left + '" w:hanging="' + LIST_STEP + '"/></w:pPr><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial" w:hint="default"/></w:rPr></w:lvl>';
      ab1 += '<w:lvl w:ilvl="' + i + '"><w:start w:val="1"/><w:numFmt w:val="' + fmts[i % 3] + '"/><w:lvlText w:val="%' + (i + 1) +
        '."/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="' + left + '" w:hanging="' + LIST_STEP + '"/></w:pPr></w:lvl>';
    }
    var s = XML_HEAD + '<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
      '<w:abstractNum w:abstractNumId="0"><w:multiLevelType w:val="hybridMultilevel"/>' + ab0 + '</w:abstractNum>' +
      '<w:abstractNum w:abstractNumId="1"><w:multiLevelType w:val="hybridMultilevel"/>' + ab1 + '</w:abstractNum>' +
      '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>';
    st.nums.forEach(function (nm, idx) {
      s += '<w:num w:numId="' + (idx + 2) + '"><w:abstractNumId w:val="1"/>';
      for (var l = 0; l < 9; l++) s += '<w:lvlOverride w:ilvl="' + l + '"><w:startOverride w:val="' + (l === 0 ? nm.start : 1) + '"/></w:lvlOverride>';
      s += '</w:num>';
    });
    return s + '</w:numbering>';
  }

  function stylesXml() {
    var F = '<w:rFonts w:ascii="Malgun Gothic" w:eastAsia="Malgun Gothic" w:hAnsi="Malgun Gothic" w:cs="Malgun Gothic"/>';
    function ps(id, name, ppr, rpr, extra) {
      return '<w:style w:type="paragraph" w:styleId="' + id + '"' + (id === 'Normal' ? ' w:default="1"' : '') + '><w:name w:val="' + name + '"/>' +
        (id === 'Normal' ? '' : '<w:basedOn w:val="Normal"/><w:next w:val="Normal"/>') + (extra || '') + '<w:qFormat/>' +
        (ppr ? '<w:pPr>' + ppr + '</w:pPr>' : '') + (rpr ? '<w:rPr>' + rpr + '</w:rPr>' : '') + '</w:style>';
    }
    return XML_HEAD + '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
      '<w:docDefaults><w:rPrDefault><w:rPr>' + F + '<w:sz w:val="21"/><w:szCs w:val="21"/><w:lang w:val="ko-KR" w:eastAsia="ko-KR" w:bidi="ar-SA"/></w:rPr></w:rPrDefault>' +
      '<w:pPrDefault><w:pPr><w:spacing w:after="120" w:line="336" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>' +
      ps('Normal', 'Normal', '', '') +
      '<w:style w:type="character" w:default="1" w:styleId="DefaultParagraphFont"><w:name w:val="Default Paragraph Font"/><w:uiPriority w:val="1"/><w:semiHidden/></w:style>' +
      '<w:style w:type="table" w:default="1" w:styleId="TableNormal"><w:name w:val="Normal Table"/><w:uiPriority w:val="99"/><w:semiHidden/><w:tblPr><w:tblInd w:w="0" w:type="dxa"/><w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="108" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr></w:style>' +
      '<w:style w:type="numbering" w:default="1" w:styleId="NoList"><w:name w:val="No List"/><w:uiPriority w:val="99"/><w:semiHidden/></w:style>' +
      ps('Title', 'Title', '<w:keepNext/><w:spacing w:before="120" w:after="120" w:line="288" w:lineRule="auto"/>', '<w:b/><w:bCs/><w:color w:val="1A1A1A"/><w:sz w:val="48"/><w:szCs w:val="48"/>') +
      ps('Heading1', 'heading 1', '<w:keepNext/><w:keepLines/><w:spacing w:before="360" w:after="120" w:line="288" w:lineRule="auto"/><w:outlineLvl w:val="0"/>', '<w:b/><w:bCs/><w:color w:val="1A1A1A"/><w:sz w:val="34"/><w:szCs w:val="34"/>') +
      ps('Heading2', 'heading 2', '<w:keepNext/><w:keepLines/><w:spacing w:before="240" w:after="80" w:line="288" w:lineRule="auto"/><w:outlineLvl w:val="1"/>', '<w:b/><w:bCs/><w:color w:val="1A1A1A"/><w:sz w:val="26"/><w:szCs w:val="26"/>') +
      ps('Heading3', 'heading 3', '<w:keepNext/><w:keepLines/><w:spacing w:before="160" w:after="60" w:line="288" w:lineRule="auto"/><w:outlineLvl w:val="2"/>', '<w:b/><w:bCs/><w:color w:val="1A1A1A"/><w:sz w:val="23"/><w:szCs w:val="23"/>') +
      ps('ListParagraph', 'List Paragraph', '<w:spacing w:after="40"/>', '') +
      '</w:styles>';
  }

  function docx(input, opts) {
    return Promise.resolve().then(function () {
      opts = opts || {};
      var tree = toTree(input);
      var st = { nums: [], slide: 0, title: '' };
      var ctx = { width: BODY_W, ppr: {}, rpr: {}, lvl: 0, st: st };
      var blocks = renderKids(tree.kids, ctx);
      var body = blocksXml(blocks, false);
      // 마지막은 문단이어야 안전 (표 뒤에 문단이 없으면 Word 가 거부)
      var lastB = blocks.filter(function (b) { return !b.pb; }).pop();
      if (!lastB || lastB.x != null) body += '<w:p><w:pPr><w:spacing w:before="0" w:after="0"/></w:pPr></w:p>';
      var title = cutStr(clean(opts.title || st.title || '수업 문서'), LIM.head);
      var author = cutStr(clean(opts.author || '교수자 AI 활용 연수 실습실'), 80);
      var footerText = opts.footer != null ? cutStr(clean(opts.footer), 120) : '';
      var NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
      var sect = '<w:sectPr><w:footerReference w:type="default" r:id="rId5"/><w:pgSz w:w="' + PAGE_W + '" w:h="' + PAGE_H + '"/>' +
        '<w:pgMar w:top="' + MARGIN + '" w:right="' + MARGIN + '" w:bottom="' + MARGIN + '" w:left="' + MARGIN + '" w:header="567" w:footer="567" w:gutter="0"/></w:sectPr>';
      var documentXml = XML_HEAD + '<w:document ' + NS + '><w:body>' + body + sect + '</w:body></w:document>';
      var fr = { size: 16, color: GRAY };
      var footerXml = XML_HEAD + '<w:ftr ' + NS + '><w:p><w:pPr><w:jc w:val="center"/><w:spacing w:before="0" w:after="0"/></w:pPr>' +
        (footerText ? runXml(footerText + '   ', fr) : '') +
        '<w:r>' + rprXml(fr) + '<w:fldChar w:fldCharType="begin"/></w:r><w:r>' + rprXml(fr) + '<w:instrText xml:space="preserve"> PAGE </w:instrText></w:r>' +
        '<w:r>' + rprXml(fr) + '<w:fldChar w:fldCharType="separate"/></w:r>' + runXml('1', fr) + '<w:r>' + rprXml(fr) + '<w:fldChar w:fldCharType="end"/></w:r></w:p></w:ftr>';
      var settingsXml = XML_HEAD + '<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:zoom w:percent="100"/><w:defaultTabStop w:val="720"/>' +
        '<w:characterSpacingControl w:val="doNotCompress"/><w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat>' +
        '<w:themeFontLang w:val="ko-KR" w:eastAsia="ko-KR"/></w:settings>';
      var R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
      var now = isoNow();
      var files = [
        ['[Content_Types].xml', XML_HEAD + '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
          '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>' +
          '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
          '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>' +
          '<Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>' +
          '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>' +
          '<Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>' +
          '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>' +
          '<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/></Types>'],
        ['_rels/.rels', XML_HEAD + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
          '<Relationship Id="rId1" Type="' + R + '/officeDocument" Target="word/document.xml"/>' +
          '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>' +
          '<Relationship Id="rId3" Type="' + R + '/extended-properties" Target="docProps/app.xml"/></Relationships>'],
        ['word/document.xml', documentXml],
        ['word/_rels/document.xml.rels', XML_HEAD + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
          '<Relationship Id="rId1" Type="' + R + '/styles" Target="styles.xml"/>' +
          '<Relationship Id="rId2" Type="' + R + '/numbering" Target="numbering.xml"/>' +
          '<Relationship Id="rId3" Type="' + R + '/settings" Target="settings.xml"/>' +
          '<Relationship Id="rId5" Type="' + R + '/footer" Target="footer1.xml"/></Relationships>'],
        ['word/styles.xml', stylesXml()],
        ['word/numbering.xml', numberingXml(st)],
        ['word/settings.xml', settingsXml],
        ['word/footer1.xml', footerXml],
        ['docProps/core.xml', XML_HEAD + '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">' +
          '<dc:title>' + esc(title) + '</dc:title><dc:creator>' + esc(author) + '</dc:creator><cp:lastModifiedBy>' + esc(author) + '</cp:lastModifiedBy>' +
          '<dcterms:created xsi:type="dcterms:W3CDTF">' + now + '</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">' + now + '</dcterms:modified></cp:coreProperties>'],
        ['docProps/app.xml', XML_HEAD + '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Application>OfficeKit</Application><AppVersion>1.0</AppVersion></Properties>']
      ];
      return zipBlob(files, MIME_DOCX);
    });
  }

  /* ------------------------------------------------------------------ */
  /* XLSX                                                                */
  /* ------------------------------------------------------------------ */

  function colLetter(i) {
    var s = '';
    i++;
    while (i > 0) { var m = (i - 1) % 26; s = String.fromCharCode(65 + m) + s; i = Math.floor((i - 1) / 26); }
    return s;
  }
  function sanitizeSheetNames(sheets) {
    var used = {}, out = [];
    sheets.forEach(function (sh, idx) {
      var nm = clean(sh && sh.name != null ? sh.name : '').replace(/[\[\]:*?\/\\]/g, '_').replace(/\s+/g, ' ').trim();
      nm = nm.replace(/^'+|'+$/g, '').trim();
      if (!nm) nm = '시트' + (idx + 1);
      if (nm.toLowerCase() === 'history') nm = 'History_';
      nm = cutStr(nm, 31);
      var base = nm, n = 1;
      while (used[nm.toLowerCase()]) {
        n++;
        var suf = '_' + n;
        nm = cutStr(base, 31 - suf.length) + suf;
      }
      used[nm.toLowerCase()] = true;
      out.push(nm);
    });
    return out;
  }
  function cellKind(v, header, numbers) {
    if (v == null) return { k: 'e' };
    if (typeof v === 'number') return isFinite(v) ? { k: 'n', v: v } : { k: 'e' };
    if (typeof v === 'boolean') return { k: 's', v: v ? 'TRUE' : 'FALSE' };
    if (v instanceof Date && !isNaN(v)) return { k: 's', v: v.toISOString().slice(0, 10) };
    var s = clean(String(v));
    if (s === '') return { k: 'e' };
    if (numbers && !header && /^-?(0|[1-9]\d{0,14})(\.\d{1,14})?$/.test(s) && s.replace(/[-.]/g, '').length <= 15) return { k: 'n', v: parseFloat(s) };
    return { k: 's', v: cutStr(s, LIM.cellChars) };
  }

  function xlsx(sheets, opts) {
    return Promise.resolve().then(function () {
      opts = opts || {};
      if (!Array.isArray(sheets) || !sheets.length) throw new Error('내보낼 시트가 없습니다.');
      sheets = sheets.slice(0, LIM.sheets);
      var names = sanitizeSheetNames(sheets);
      var sst = [], sstIdx = {}, sstCount = 0;
      function si(s) {
        sstCount++;
        if (sstIdx[s] == null) { sstIdx[s] = sst.length; sst.push(s); }
        return sstIdx[s];
      }
      var sheetXmls = sheets.map(function (sh, sIdx) {
        var rows = (sh && Array.isArray(sh.rows) ? sh.rows : []).slice(0, LIM.rows);
        var header = !(sh && sh.header === false) && rows.length > 0;
        var ncols = 0;
        rows = rows.map(function (r) { r = Array.isArray(r) ? r.slice(0, LIM.cols) : [r]; ncols = Math.max(ncols, r.length); return r; });
        var widths = [], c;
        for (c = 0; c < ncols; c++) widths.push(8);
        var rowXml = '';
        rows.forEach(function (r, ri) {
          var cells = '';
          for (var ci = 0; ci < ncols; ci++) {
            var isHead = header && ri === 0;
            var k = cellKind(r[ci], isHead, !!opts.numbers);
            var ref = colLetter(ci) + (ri + 1), sty = isHead ? 1 : 2;
            if (k.k === 'e') { cells += '<c r="' + ref + '" s="' + sty + '"/>'; continue; }
            if (k.k === 'n') {
              cells += '<c r="' + ref + '" s="' + sty + '"><v>' + String(k.v) + '</v></c>';
              widths[ci] = Math.max(widths[ci], String(k.v).length + 2);
            } else {
              cells += '<c r="' + ref + '" s="' + sty + '" t="s"><v>' + si(k.v) + '</v></c>';
              var lw = 0;
              k.v.split('\n').forEach(function (l) { lw = Math.max(lw, dispWidth(l)); });
              widths[ci] = Math.max(widths[ci], lw + 2 + (isHead ? 1 : 0));
            }
          }
          rowXml += '<row r="' + (ri + 1) + '">' + cells + '</row>';
        });
        var custom = sh && Array.isArray(sh.widths) ? sh.widths : null;
        var colsXml = '';
        for (c = 0; c < ncols; c++) {
          var w = custom && typeof custom[c] === 'number' && custom[c] > 0 ? Math.min(255, custom[c]) : Math.min(60, Math.max(8, widths[c]));
          colsXml += '<col min="' + (c + 1) + '" max="' + (c + 1) + '" width="' + w.toFixed(2) + '" customWidth="1"/>';
        }
        var dim = ncols && rows.length ? 'A1:' + colLetter(ncols - 1) + rows.length : 'A1';
        var view = '<sheetViews><sheetView workbookViewId="0"' + (sIdx === 0 ? ' tabSelected="1"' : '') + '>' +
          (header ? '<pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/><selection pane="bottomLeft" activeCell="A2" sqref="A2"/>' : '') +
          '</sheetView></sheetViews>';
        return XML_HEAD + '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
          '<dimension ref="' + dim + '"/>' + view + '<sheetFormatPr defaultRowHeight="15"/>' + (colsXml ? '<cols>' + colsXml + '</cols>' : '') +
          '<sheetData>' + rowXml + '</sheetData><pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>' +
          '<pageSetup paperSize="9" orientation="landscape" fitToHeight="0"/></worksheet>';
      });
      var R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
      var n = sheets.length, now = isoNow();
      var title = cutStr(clean(opts.title || names[0]), LIM.head), author = cutStr(clean(opts.author || '교수자 AI 활용 연수 실습실'), 80);
      var files = [];
      files.push(['[Content_Types].xml', XML_HEAD + '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>' +
        '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' +
        sheetXmls.map(function (_, i) { return '<Override PartName="/xl/worksheets/sheet' + (i + 1) + '.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'; }).join('') +
        '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>' +
        '<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>' +
        '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>' +
        '<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/></Types>']);
      files.push(['_rels/.rels', XML_HEAD + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        '<Relationship Id="rId1" Type="' + R + '/officeDocument" Target="xl/workbook.xml"/>' +
        '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>' +
        '<Relationship Id="rId3" Type="' + R + '/extended-properties" Target="docProps/app.xml"/></Relationships>']);
      files.push(['xl/workbook.xml', XML_HEAD + '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="' + R + '"><bookViews><workbookView xWindow="0" yWindow="0" windowWidth="24000" windowHeight="12000" activeTab="0"/></bookViews><sheets>' +
        names.map(function (nm, i) { return '<sheet name="' + esc(nm) + '" sheetId="' + (i + 1) + '" r:id="rId' + (i + 1) + '"/>'; }).join('') + '</sheets></workbook>']);
      files.push(['xl/_rels/workbook.xml.rels', XML_HEAD + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        names.map(function (_, i) { return '<Relationship Id="rId' + (i + 1) + '" Type="' + R + '/worksheet" Target="worksheets/sheet' + (i + 1) + '.xml"/>'; }).join('') +
        '<Relationship Id="rId' + (n + 1) + '" Type="' + R + '/styles" Target="styles.xml"/>' +
        '<Relationship Id="rId' + (n + 2) + '" Type="' + R + '/sharedStrings" Target="sharedStrings.xml"/></Relationships>']);
      sheetXmls.forEach(function (x, i) { files.push(['xl/worksheets/sheet' + (i + 1) + '.xml', x]); });
      var F = '<name val="Malgun Gothic"/><family val="2"/><charset val="129"/>';
      files.push(['xl/styles.xml', XML_HEAD + '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
        '<fonts count="2"><font><sz val="10"/>' + F + '</font><font><b/><sz val="10"/><color rgb="FFFFFFFF"/>' + F + '</font></fonts>' +
        '<fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FF1A1A1A"/><bgColor indexed="64"/></patternFill></fill></fills>' +
        '<borders count="2"><border><left/><right/><top/><bottom/><diagonal/></border><border>' +
        '<left style="thin"><color rgb="FFBFBFBF"/></left><right style="thin"><color rgb="FFBFBFBF"/></right><top style="thin"><color rgb="FFBFBFBF"/></top><bottom style="thin"><color rgb="FFBFBFBF"/></bottom><diagonal/></border></borders>' +
        '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
        '<cellXfs count="3"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyFont="1"/>' +
        '<xf numFmtId="0" fontId="1" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment vertical="center" wrapText="1"/></xf>' +
        '<xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyFont="1" applyBorder="1" applyAlignment="1"><alignment vertical="top" wrapText="1"/></xf></cellXfs>' +
        '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>']);
      files.push(['xl/sharedStrings.xml', XML_HEAD + '<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="' + sstCount + '" uniqueCount="' + sst.length + '">' +
        sst.map(function (s) { return '<si><t xml:space="preserve">' + esc(s) + '</t></si>'; }).join('') + '</sst>']);
      files.push(['docProps/core.xml', XML_HEAD + '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">' +
        '<dc:title>' + esc(title) + '</dc:title><dc:creator>' + esc(author) + '</dc:creator><cp:lastModifiedBy>' + esc(author) + '</cp:lastModifiedBy>' +
        '<dcterms:created xsi:type="dcterms:W3CDTF">' + now + '</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">' + now + '</dcterms:modified></cp:coreProperties>']);
      files.push(['docProps/app.xml', XML_HEAD + '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Application>OfficeKit</Application><AppVersion>1.0</AppVersion></Properties>']);
      return zipBlob(files, MIME_XLSX);
    });
  }

  /* ------------------------------------------------------------------ */
  /* 내려받기                                                             */
  /* ------------------------------------------------------------------ */

  function save(blob, filename) {
    if (typeof document === 'undefined' || typeof URL === 'undefined' || !URL.createObjectURL || !document.body) return blob;
    var name = String(filename == null ? '' : filename).replace(/[\u0000-\u001F\u007F\/\\:*?"<>|]/g, '_').replace(/^[\s.]+|[\s.]+$/g, '');
    if (!name) name = 'document';
    var url = URL.createObjectURL(blob);
    var a = document.createElement('a');
    a.href = url; a.download = name; a.style.display = 'none';
    document.body.appendChild(a);
    a.click();
    setTimeout(function () { if (a.parentNode) a.parentNode.removeChild(a); }, 0);
    setTimeout(function () { URL.revokeObjectURL(url); }, 60000);
    return blob;
  }

  return {
    version: VERSION,
    mime: { docx: MIME_DOCX, xlsx: MIME_XLSX },
    docx: docx,
    xlsx: xlsx,
    tablesFromElement: tablesFromElement,
    quizRows: quizRows,
    save: save,
    _internal: { parseHtml: parseHtml, fromDom: fromDom, clean: clean, esc: esc, sanitizeSheetNames: sanitizeSheetNames, colLetter: colLetter, tableRows: tableRows }
  };
});
