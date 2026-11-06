/*!
 * PptKit 1.0.0 - 브라우저 안에서 진짜 .pptx 를 만들고(buildDeck), 올린 .pptx 의 글을 고쳐(applyEdits) 돌려주는 작은 도구.
 * 서버도 네트워크도 쓰지 않습니다. 의존: 전역 JSZip (vendor/jszip/jszip.min.js),
 * buildDeck 에만 전역 PptxGenJS (vendor/pptxgenjs/pptxgen.bundle.js).
 * 모델이 쓴 글과 올린 파일의 글은 모두 적대적 입력으로 취급합니다:
 * XML/HTML 이스케이프, XML 1.0 금지 제어문자 제거, 길이 상한, eval 없음.
 *
 * API (async 는 Promise 를 돌려줍니다)
 *   PptKit.THEMES                       { brutal, clean, navy, warm } -> { label, desc, colors }
 *   PptKit.parseModelJson(text)         -> object | null   (잘린 답은 닫아서 __truncated: true)
 *   PptKit.validateDeck(spec)           -> { spec, warnings }
 *   PptKit.previewHtml(spec)            -> HTML 문자열 (<div class="ppt-prev">, 인라인 스타일만)
 *   PptKit.buildDeck(spec, opts)        async -> Blob(.pptx)   opts: { theme?, morph? }
 *   PptKit.inspect(arrayBuffer)         async -> { slideCount, size:{ w, h (인치), cx, cy (EMU) },
 *                                         slides:[{ n, title, notes, shapes:[{ id, name, kind, text, paragraphs }], objects:[{ id, name, kind }] }] }
 *       n 은 presentation.xml 의 실제 순서. shapes 는 글 있는 도형(자리표시자, 묶음 안, 표 칸 포함), kind: title|body|text|table|other.
 *       id 는 p:cNvPr id(문자열). 표 칸은 "표 id.r행c열"(1부터, 예 "12.r1c2"). objects 는 글 없는 그림·도형·선·묶음(이름 바꾸기용).
 *   PptKit.outline(info, { maxChars })  -> 모델에게 줄 글: [슬라이드 N] / #id kind: "문단1 | 문단2" / #objects: ... / #notes: "..."
 *   PptKit.applyEdits(buf, edits, opts) async -> { blob, applied, skipped:[{ edit, reason }],
 *                                         changed:[{ slide, shape, field: text|name|notes, before, after }], morphed:[슬라이드 번호] }
 *
 * 덱 JSON (validateDeck 가 받는 모양)
 *   { "title", "subtitle", "author", "date", "theme": "brutal|clean|navy|warm", "morph": false, "themeNote",
 *     "slides": [ { "layout", "kicker", "title", "body", "accent", "items":[{ "num", "head", "text" }],
 *                   "stat":{ "value", "label" }, "table":{ "head":[..], "rows":[[..]] },
 *                   "left":{ "head", "items":[..] }, "right":{ "head", "items":[..] },
 *                   "quote":{ "text", "by" }, "notes", "names":{ "title": "!!title" } } ] }
 *   layout: cover section statement text cards stat steps timeline table compare quote closing (모르는 값은 text)
 *
 * 고치기 JSON (applyEdits 가 받는 edits 의 원소)
 *   { "slide": 3, "shape": "12", "text": "한 문단" | ["문단1", "문단2"] }  문단 안 줄바꿈은 \n (<a:br/>)
 *   { "slide": 3, "shape": "12.r2c1", "text": [...] }   표 칸: 표 id + ".r행c열" (1부터)
 *   { "slide": 3, "notes": "발표자 노트", "keepOriginal": true }  \n 마다 새 문단. keepOriginal 이면 원래 노트를 "── 원본 메모 ──" 아래에 남김
 *   { "slide": 3, "shape": "12", "name": "!!title" }    도형 이름 바꾸기 (모핑 짝 맞추기). 그림·도형(#objects)도 됨
 *   applyEdits(buf, edits, { morph: true })             2..N 장의 전환을 모핑으로 바꿈
 *   슬라이드를 지우거나 순서를 바꾸지 않습니다. 손대지 않은 zip 항목은 바이트 그대로입니다.
 *   한계: mc:AlternateContent 안 도형은 mc:Choice 쪽만 고칩니다(옛 버전용 Fallback 은 그대로).
 */
(function (root, factory) {
  var api = factory(root);
  if (typeof module === 'object' && module && module.exports) module.exports = api;
  else root.PptKit = api;
})(typeof globalThis !== 'undefined' ? globalThis : typeof self !== 'undefined' ? self : this, function (root) {
  'use strict';

  var VERSION = '1.0.0';
  var PPTX_MIME = 'application/vnd.openxmlformats-officedocument.presentationml.presentation';
  var SW = 13.333, SH = 7.5;
  var LIM = {
    input: 600000, title: 120, para: 2000, paras: 40, body: 4000, notes: 6000, slides: 40, shapes: 60,
    kicker: 60, head: 80, item: 600, items: 8, cell: 300, cols: 6, rows: 8, accent: 40, name: 80, edits: 2000, meta: 120
  };
  var FONT_KR = 'Malgun Gothic', FONT_LATIN = 'Arial';
  var CSS_KR = "'Malgun Gothic','맑은 고딕','Apple SD Gothic Neo','Noto Sans KR',sans-serif";
  var CSS_LATIN = "Arial,'Helvetica Neue',Helvetica,'Malgun Gothic','Apple SD Gothic Neo',sans-serif";

  /* ------------------------------------------------------------------ */
  /* 공통 도우미                                                          */
  /* ------------------------------------------------------------------ */

  function getJSZip() {
    if (root.JSZip) return root.JSZip;
    if (typeof require === 'function') { try { return require('jszip'); } catch (e) { /* fall through */ } }
    throw new Error('JSZip 라이브러리를 불러오지 못했습니다.');
  }
  function getPptxGen() {
    if (root.PptxGenJS) return root.PptxGenJS;
    if (typeof require === 'function') { try { var m = require('pptxgenjs'); return m && m.default ? m.default : m; } catch (e) { /* fall through */ } }
    throw new Error('PptxGenJS 라이브러리를 불러오지 못했습니다.');
  }

  var CTRL_RE = /[\u0000-\u0008\u000B\u000C\u000E-\u001F\uFFFE\uFFFF]/g;
  function fixSurrogates(s) {
    var out = '';
    for (var i = 0; i < s.length; i++) {
      var c = s.charCodeAt(i);
      if (c >= 0xD800 && c <= 0xDBFF) {
        var d = s.charCodeAt(i + 1);
        if (d >= 0xDC00 && d <= 0xDFFF) { out += s[i] + s[i + 1]; i++; }
      } else if (!(c >= 0xDC00 && c <= 0xDFFF)) out += s[i];
    }
    return out;
  }
  function str(v) {
    if (v == null) return '';
    if (typeof v === 'string') return v;
    if (typeof v === 'number') return isFinite(v) ? String(v) : '';
    if (typeof v === 'boolean') return String(v);
    return '';
  }
  // 제어문자 제거, 줄바꿈 통일, 길이 상한(넘으면 … 로 끝냄)
  function clean(v, max) {
    var s = str(v);
    if (!s) return '';
    s = s.replace(/\r\n?/g, '\n').replace(/\t/g, ' ').replace(CTRL_RE, '');
    if (/[\uD800-\uDFFF]/.test(s)) s = fixSurrogates(s);
    if (max && s.length > max) {
      s = s.slice(0, max - 1);
      if (/[\uD800-\uDBFF]$/.test(s)) s = s.slice(0, -1);
      s = s.replace(/\s+$/, '') + '…';
    }
    return s;
  }
  function text1(v, max) { return clean(str(v).replace(/\s*[\r\n]+\s*/g, ' '), max).trim(); }
  function textN(v, max) { // 여러 줄 글: 문단마다 다듬고 빈 줄은 하나로
    var s = Array.isArray(v) ? v.map(str).join('\n') : str(v);
    s = clean(s, max).split('\n').map(function (l) { return l.replace(/\s+$/, ''); }).join('\n');
    return s.replace(/\n{3,}/g, '\n\n').trim();
  }
  function xmlEsc(v) {
    return clean(v).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/'/g, '&apos;');
  }
  function htmlEsc(v) {
    return clean(v).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/'/g, '&#39;');
  }
  function hasHangul(s) { return /[\u1100-\u11FF\u3130-\u318F\uAC00-\uD7AF]/.test(s); }
  function hasCJK(s) { return /[\u1100-\u11FF\u2E80-\u9FFF\uAC00-\uD7AF\uF900-\uFAFF\uFF00-\uFFEF]/.test(s); }
  var EMOJI_RE = /[\u2600-\u27BF\u2B50\u2B55\u{1F000}-\u{1FAFF}\uFE0F]/u;
  function isObj(v) { return !!v && typeof v === 'object' && !Array.isArray(v); }
  function own(o, k) { return Object.prototype.hasOwnProperty.call(o, k); }
  function round(n, d) { var k = Math.pow(10, d || 3); return Math.round(n * k) / k; }

  /* ------------------------------------------------------------------ */
  /* 테마                                                                 */
  /* ------------------------------------------------------------------ */

  // bg 바탕, ink 제목·선, body 본문, mute 캡션·라벨, hair 옅은 선, ghost 배경 글씨, accent 강조 하나
  var PALETTES = {
    brutal: { bg: 'F2F2F2', ink: '1A1A1A', body: '6F6F6F', mute: '9A9A9A', hair: 'D2D2D2', ghost: 'E8E8E8', accent: 'E54536' },
    clean: { bg: 'FFFFFF', ink: '111827', body: '4B5563', mute: '8B95A5', hair: 'E3E7EE', ghost: 'F3F5F8', accent: '2563EB' },
    navy: { bg: '13214A', ink: 'FFFFFF', body: 'C5CDE3', mute: '8D99BD', hair: '3A4A78', ghost: '1A2A5A', accent: 'EE5A36' },
    warm: { bg: 'FBF5EC', ink: '3A2A20', body: '6E5B4D', mute: 'A28E7E', hair: 'E5D6C4', ghost: 'F4EADD', accent: 'C8553D' }
  };
  var THEMES = {
    brutal: { label: '브루탈 미니멀 (강사 디자인 명세)', desc: '#F2F2F2 종이에 검정 글, 강조는 빨강 #E54536 하나. 아주 큰 굵은 제목, 작은 회색 본문 한 단, 1pt 선.' },
    clean: { label: '클린 화이트', desc: '흰 바탕, 검정 글, 파랑 강조 하나. 강의실 프로젝터에서 가장 또렷합니다.' },
    navy: { label: '네이비 · 오렌지', desc: '짙은 남색 #13214A 바탕에 흰 글, 주황 #EE5A36 강조.' },
    warm: { label: '웜 크림', desc: '크림 바탕, 짙은 갈색 글, 테라코타 강조. 부드러운 인상.' }
  };
  Object.keys(THEMES).forEach(function (k) {
    var c = {}; Object.keys(PALETTES[k]).forEach(function (n) { c[n] = '#' + PALETTES[k][n]; });
    THEMES[k].colors = Object.freeze(c);
    Object.freeze(THEMES[k]);
  });
  Object.freeze(THEMES);

  /* ------------------------------------------------------------------ */
  /* parseModelJson: 모델 답에서 첫 JSON 객체 꺼내기                         */
  /* ------------------------------------------------------------------ */

  // src[at] === '{' 에서 시작해 객체 하나를 훑습니다. 문법이 틀리면 null.
  // 끝까지 닫히지 않았으면(잘린 답) 마지막으로 완결된 값까지 자르고 괄호를 닫아 truncated: true.
  // 최상위 배열(slides, edits) 안의 원소가 반쯤 쓰였으면 그 원소는 버립니다.
  function scanJson(src, at) {
    var out = '', stack = [], i = at, n = src.length, safe = null, arrSafe = null;
    function closers() { var s = ''; for (var k = stack.length - 1; k >= 0; k--) s += stack[k].t === '{' ? '}' : ']'; return s; }
    function mark() {
      var snap = { len: out.length, close: closers(), deep: stack.length > 2 && stack[1].t === '[' };
      safe = snap;
      if (stack.length === 2 && stack[1].t === '[') arrSafe = snap;
    }
    function done() { stack[stack.length - 1].exp = 'comma'; mark(); }
    if (src[at] !== '{') return null;
    stack.push({ t: '{', exp: 'key', key: '' }); out = '{'; i = at + 1; mark();
    while (i < n) {
      var c = src[i], top = stack[stack.length - 1];
      if (c === ' ' || c === '\n' || c === '\r' || c === '\t') { out += c; i++; continue; }
      if (c === '"') {
        var j = i + 1, buf = '"', closed = false;
        while (j < n) {
          var d = src[j];
          if (d === '\\') {
            if (j + 1 >= n) { j = n; break; }
            var e = src[j + 1];
            if (e === 'u') {
              if (j + 6 > n) { j = n; break; }
              if (/^[0-9a-fA-F]{4}$/.test(src.substr(j + 2, 4))) { buf += src.substr(j, 6); j += 6; } else { buf += '\\\\'; j++; }
            } else if ('"\\/bfnrt'.indexOf(e) >= 0) { buf += d + e; j += 2; }
            else if (e === "'") { buf += "'"; j += 2; }
            else { buf += '\\\\'; j++; }
          } else if (d === '"') { buf += '"'; closed = true; j++; break; }
          else if (d === '\n') { buf += '\\n'; j++; }
          else if (d === '\r') { j++; }
          else if (d === '\t') { buf += '\\t'; j++; }
          else if (d < ' ') { j++; }
          else { buf += d; j++; }
        }
        if (!closed) break;
        if (top.t === '{' && top.exp === 'key') {
          try { top.key = JSON.parse(buf); } catch (er) { return null; }
          top.exp = 'colon'; out += buf; i = j; continue;
        }
        if (top.exp !== 'value') return null;
        out += buf; i = j; done(); continue;
      }
      if (c === '{' || c === '[') {
        if (top.exp !== 'value') return null;
        out += c; i++;
        stack.push(c === '{' ? { t: '{', exp: 'key', key: '' } : { t: '[', exp: 'value', key: '' });
        mark(); continue;
      }
      if (c === '}' || c === ']') {
        if ((c === '}') !== (top.t === '{')) return null;
        if (top.t === '{' && (top.exp === 'colon' || top.exp === 'value')) return null;
        out += c; i++; stack.pop();
        if (!stack.length) return { text: out, truncated: false };
        done(); continue;
      }
      if (c === ':') {
        if (top.t !== '{' || top.exp !== 'colon') return null;
        top.exp = 'value'; out += c; i++; continue;
      }
      if (c === ',') {
        if (top.exp !== 'comma') return null;
        var k2 = i + 1; while (k2 < n && /\s/.test(src[k2])) k2++;
        if (k2 < n && (src[k2] === '}' || src[k2] === ']')) { i++; continue; } // 끝에 붙은 쉼표는 버림
        top.exp = top.t === '{' ? 'key' : 'value'; out += c; i++; continue;
      }
      if (top.exp !== 'value') return null;
      var m = /^(?:-?\d+(?:\.\d+)?(?:[eE][+-]?\d+)?|true|false|null)/.exec(src.slice(i, i + 64));
      if (!m) return null;
      if (i + m[0].length >= n) break; // 숫자가 끝에서 잘렸을 수 있음
      out += m[0]; i += m[0].length; done();
    }
    if (!safe) return null;
    var pick = safe.deep && arrSafe ? arrSafe : safe;
    return { text: out.slice(0, pick.len) + pick.close, truncated: true };
  }

  function parseModelJson(text) {
    if (typeof text !== 'string' || !text) return null;
    var src = (text.length > LIM.input ? text.slice(0, LIM.input) : text).replace(/^\uFEFF/, '');
    var starts = [];
    var fence = /```[A-Za-z0-9_-]*[ \t]*\r?\n/.exec(src);
    if (fence) starts.push(fence.index + fence[0].length);
    starts.push(0);
    for (var r = 0; r < starts.length; r++) {
      var from = starts[r];
      for (var tries = 0; tries < 60; tries++) {
        var at = src.indexOf('{', from);
        if (at < 0) break;
        var res = scanJson(src, at);
        if (res) {
          var obj = null;
          try { obj = JSON.parse(res.text); } catch (e) { obj = null; }
          if (isObj(obj) && Object.keys(obj).length) {
            if (res.truncated) obj.__truncated = true;
            return obj;
          }
        }
        from = at + 1;
      }
    }
    return null;
  }

  /* ------------------------------------------------------------------ */
  /* validateDeck: 덱 JSON 다듬기 (모르는 값은 버리고, 빠진 값은 채우고, 상한)   */
  /* ------------------------------------------------------------------ */

  var LAYOUTS = ['cover', 'section', 'statement', 'text', 'cards', 'stat', 'steps', 'timeline', 'table', 'compare', 'quote', 'closing'];
  var LAYOUT_ALIAS = {
    title: 'cover', intro: 'cover', chapter: 'section', divider: 'section', message: 'statement', big: 'statement',
    bullets: 'text', list: 'text', content: 'text', grid: 'cards', columns: 'cards', number: 'stat', kpi: 'stat', metric: 'stat',
    process: 'steps', flow: 'steps', history: 'timeline', schedule: 'timeline', vs: 'compare', versus: 'compare',
    comparison: 'compare', citation: 'quote', end: 'closing', outro: 'closing', thanks: 'closing'
  };
  var ITEM_MAX = { cards: 4, steps: 5, timeline: 8, text: 6, closing: 4 };
  var NAME_ROLES = /^[a-z][a-z0-9]{0,15}$/;

  function normItem(it) {
    if (typeof it === 'string' || typeof it === 'number') return { num: '', head: '', text: textN(it, LIM.item) };
    if (!isObj(it)) return null;
    var o = {
      num: text1(it.num != null ? it.num : it.year != null ? it.year : it.label, 24),
      head: text1(it.head != null ? it.head : it.title, LIM.head),
      text: textN(it.text != null ? it.text : it.body != null ? it.body : it.desc, LIM.item)
    };
    return o.num || o.head || o.text ? o : null;
  }
  function normList(v, max, each) {
    var a = Array.isArray(v) ? v : typeof v === 'string' && v ? v.split('\n') : [];
    var out = [];
    for (var i = 0; i < a.length && out.length < max; i++) { var t = textN(a[i], each); if (t) out.push(t); }
    return out;
  }
  function normSide(v) {
    if (!isObj(v)) return null;
    var items = normList(v.items != null ? v.items : v.body != null ? v.body : v.text, 6, 300);
    var head = text1(v.head != null ? v.head : v.title, LIM.head);
    return head || items.length ? { head: head, items: items } : null;
  }
  function normTable(v, w, label) {
    var head = [], rows = [];
    if (Array.isArray(v)) { var a = v.filter(Array.isArray); head = a[0] || []; rows = a.slice(1); }
    else if (isObj(v)) { head = v.head || v.header || v.columns || []; rows = v.rows || v.body || []; }
    if (!Array.isArray(head)) head = [];
    if (!Array.isArray(rows)) rows = [];
    rows = rows.filter(Array.isArray);
    var cols = Math.max(head.length, rows.reduce(function (m, r) { return Math.max(m, r.length); }, 0));
    if (cols > LIM.cols) w.push(label + '표의 열이 많아 ' + LIM.cols + '열까지만 씁니다.');
    if (rows.length > LIM.rows) w.push(label + '표의 행이 많아 ' + LIM.rows + '행까지만 씁니다.');
    cols = Math.min(cols, LIM.cols);
    if (!cols) return null;
    var cell = function (c) { return textN(c, LIM.cell); };
    var fix = function (r) { var o = []; for (var i = 0; i < cols; i++) o.push(cell(r[i])); return o; };
    var h = head.length ? fix(head) : [];
    var rs = rows.slice(0, LIM.rows).map(fix).filter(function (r) { return r.some(Boolean); });
    return h.some(Boolean) || rs.length ? { head: h.some(Boolean) ? h : [], rows: rs } : null;
  }
  function normNames(v) {
    var o = {};
    if (!isObj(v)) return o;
    Object.keys(v).slice(0, 16).forEach(function (k) {
      var role = String(k).toLowerCase();
      if (!NAME_ROLES.test(role)) return;
      var nm = text1(v[k], 40).replace(/["<>&]/g, '').replace(/^!*/, '');
      if (nm) o[role] = '!!' + nm;
    });
    return o;
  }

  function normSlide(s, i, spec, w) {
    var label = '슬라이드 ' + (i + 1) + ': ';
    if (!isObj(s)) { w.push(label + '내용이 비어 있어 글 장으로 만들었습니다.'); s = {}; }
    var raw = text1(s.layout, 24).toLowerCase();
    var layout = LAYOUTS.indexOf(raw) >= 0 ? raw : own(LAYOUT_ALIAS, raw) ? LAYOUT_ALIAS[raw] : '';
    if (!layout) { if (raw) w.push(label + '모르는 구성 "' + raw + '" 대신 text 를 씁니다.'); layout = 'text'; }
    var o = {
      layout: layout,
      kicker: text1(s.kicker, LIM.kicker),
      title: textN(s.title, LIM.title),
      body: textN(s.body != null ? s.body : s.text, LIM.body),
      accent: text1(s.accent, LIM.accent),
      items: [],
      stat: null, table: null, left: null, right: null, quote: null,
      notes: textN(s.notes, LIM.notes),
      names: normNames(s.names)
    };
    var items = Array.isArray(s.items) ? s.items : Array.isArray(s.steps) ? s.steps : Array.isArray(s.cards) ? s.cards : [];
    var max = own(ITEM_MAX, layout) ? ITEM_MAX[layout] : LIM.items;
    var norm = items.slice(0, LIM.items * 2).map(normItem).filter(Boolean);
    if (norm.length > max) w.push(label + layout + ' 은 항목을 ' + max + '개까지 씁니다. 나머지는 뺐습니다.');
    o.items = norm.slice(0, max);
    var st = isObj(s.stat) ? s.stat : s.value != null || s.number != null ? { value: s.value != null ? s.value : s.number, label: s.label } : null;
    if (st) {
      var val = text1(st.value != null ? st.value : st.number, 24);
      if (val) o.stat = { value: val, label: text1(st.label, LIM.head * 2) };
    }
    if (s.table != null) o.table = normTable(s.table, w, label);
    o.left = normSide(s.left);
    o.right = normSide(s.right);
    if (!o.left && !o.right && Array.isArray(s.columns) && s.columns.length >= 2) { o.left = normSide(s.columns[0]); o.right = normSide(s.columns[1]); }
    var q = isObj(s.quote) ? s.quote : typeof s.quote === 'string' ? { text: s.quote, by: s.by || s.source } : null;
    if (q) { var qt = textN(q.text, 300); if (qt) o.quote = { text: qt, by: text1(q.by != null ? q.by : q.source, LIM.meta) }; }

    // 구성에 꼭 필요한 값이 없으면 글 장으로
    var need = {
      cards: o.items.length > 0, steps: o.items.length > 0, timeline: o.items.length > 0,
      stat: !!o.stat, table: !!o.table, compare: !!(o.left && o.right), quote: !!(o.quote || o.title)
    };
    if (need[layout] === false) {
      w.push(label + layout + ' 에 필요한 값이 없어 text 로 바꿨습니다.');
      if (layout === 'compare') { [o.left, o.right].forEach(function (c) { if (c) o.body = textN([o.body, c.head].concat(c.items).filter(Boolean), LIM.body); }); o.left = o.right = null; }
      o.layout = layout = 'text';
    }
    if (layout === 'quote' && !o.quote) { o.quote = { text: o.title, by: '' }; o.title = ''; }
    if (layout === 'cover' && !o.title) o.title = spec.title;
    if (!o.title && layout !== 'quote') w.push(label + '제목이 없습니다.');
    if (o.accent && layout !== 'quote' && o.title.indexOf(o.accent) < 0) w.push(label + '강조 단어 "' + o.accent + '" 가 제목에 없어 쓰지 않습니다.');
    return o;
  }

  function normalizeDeck(input, w) {
    var src = isObj(input) ? input : {};
    var theme = typeof src.theme === 'string' ? src.theme.toLowerCase().trim() : '';
    if (theme && !own(THEMES, theme)) w.push('모르는 테마 "' + text1(theme, 20) + '" 대신 brutal 을 씁니다.');
    var spec = {
      title: text1(src.title, LIM.title),
      subtitle: textN(src.subtitle, 300),
      author: text1(src.author, LIM.meta),
      date: text1(src.date, 40),
      theme: own(THEMES, theme) ? theme : 'brutal',
      morph: src.morph === true,
      themeNote: textN(src.themeNote, 600),
      slides: []
    };
    var slides = Array.isArray(src.slides) ? src.slides : [];
    if (slides.length > LIM.slides) w.push('슬라이드가 ' + slides.length + '장이라 앞 ' + LIM.slides + '장만 씁니다.');
    if (!spec.title) {
      var first = slides.filter(isObj)[0];
      spec.title = first ? text1(first.title, LIM.title) : '';
    }
    spec.slides = slides.slice(0, LIM.slides).map(function (s, i) { return normSlide(s, i, spec, w); });
    if (!spec.slides.length) {
      w.push('슬라이드가 없어 표지 한 장만 만듭니다.');
      spec.slides.push(normSlide({ layout: 'cover', title: spec.title || '발표 자료', body: spec.subtitle }, 0, spec, w));
    }
    if (!spec.title) spec.title = spec.slides[0].title.replace(/\n/g, ' ') || '발표 자료';
    return spec;
  }

  function slideTexts(s) {
    var a = [s.kicker, s.title, s.body, s.accent, s.notes];
    s.items.forEach(function (it) { a.push(it.num, it.head, it.text); });
    if (s.stat) a.push(s.stat.value, s.stat.label);
    if (s.table) { a = a.concat(s.table.head); s.table.rows.forEach(function (r) { a = a.concat(r); }); }
    [s.left, s.right].forEach(function (c) { if (c) a = a.concat([c.head], c.items); });
    if (s.quote) a.push(s.quote.text, s.quote.by);
    return a.join('\n');
  }

  function validateDeck(input) {
    var w = [];
    var spec = normalizeDeck(input, w);
    var brutal = spec.theme === 'brutal';
    spec.slides.forEach(function (s, i) {
      var label = '슬라이드 ' + (i + 1) + ': ';
      var L = layoutSlide(spec, i);
      L.warn.forEach(function (m) { w.push(label + m); });
      if (!brutal) return;
      // 강사 디자인 명세(brutal) 규칙은 경고만 합니다
      if (i > 0 && spec.slides[i - 1].layout === s.layout) w.push(label + '바로 앞 장과 같은 구성(' + s.layout + ')입니다. 이웃한 장은 구성을 바꾸세요.');
      if (EMOJI_RE.test(slideTexts(s))) w.push(label + '이모지가 있습니다. 디자인 명세는 이모지를 쓰지 않습니다.');
      var reds = countAccent(L, PALETTES.brutal.accent);
      if (reds > 2) w.push(label + '빨간 요소가 ' + reds + '개입니다. 한 장에 2개까지만 씁니다.');
      if (s.body.length > 420) w.push(label + '본문이 깁니다. 한 장에는 메시지 하나만 둡니다.');
      if (s.title.length > 40) w.push(label + '제목이 깁니다. 두 줄 안쪽의 짧은 문장으로 줄이세요.');
    });
    return { spec: spec, warnings: w };
  }

  /* ------------------------------------------------------------------ */
  /* 글자 폭 어림 (보수적): 한글이 상자를 넘지 않게 줄 수를 넉넉히 셉니다         */
  /* ------------------------------------------------------------------ */

  // 글자 폭(em). 한글·한자는 1em, 라틴은 Arial 기준보다 조금 넓게.
  function chW(ch, bold) {
    var c = ch.charCodeAt(0);
    if (c >= 0x1100 && (c <= 0x11FF || (c >= 0x2E80 && c <= 0xA4CF) || (c >= 0xAC00 && c <= 0xD7AF) || (c >= 0xF900 && c <= 0xFAFF) || (c >= 0xFF00 && c <= 0xFF60) || (c >= 0xD800 && c <= 0xDFFF))) return 1.0;
    var w;
    if (ch === ' ') w = 0.3;
    else if ('il.,:;\'!|`ıj'.indexOf(ch) >= 0) w = 0.3;
    else if ('frtI()[]{}/\\-"'.indexOf(ch) >= 0) w = 0.42;
    else if ('mwMW@%'.indexOf(ch) >= 0) w = 0.95;
    else if (c >= 65 && c <= 90) w = 0.72;
    else if (c >= 48 && c <= 57) w = 0.6;
    else if (c >= 97 && c <= 122) w = 0.58;
    else if (c >= 0x2010 && c <= 0x2027) w = 0.6;
    else if (c < 0x2000) w = 0.62;
    else w = 1.0;
    return bold ? w * 1.07 : w;
  }
  function wordW(word, bold, sp) {
    var t = 0;
    for (var i = 0; i < word.length; i++) t += chW(word[i], bold) + sp;
    return t;
  }
  // 한 줄(강제 줄바꿈 없는 글)이 maxEm 폭에서 몇 줄이 되는지: 낱말 단위로 감고, 너무 긴 낱말은 글자 단위로 자릅니다.
  function wrapCount(seg, maxEm, bold, sp) {
    if (!seg) return 1;
    var words = seg.split(' '), lines = 1, cur = 0, spW = chW(' ', bold) + sp;
    for (var i = 0; i < words.length; i++) {
      var wd = words[i], ww = wordW(wd, bold, sp);
      if (cur > 0 && cur + spW + ww <= maxEm) { cur += spW + ww; continue; }
      if (cur > 0) { lines++; cur = 0; }
      if (ww <= maxEm) { cur = ww; continue; }
      for (var k = 0; k < wd.length; k++) {
        var cw = chW(wd[k], bold) + sp;
        if (cur > 0 && cur + cw > maxEm) { lines++; cur = 0; }
        cur += cw;
      }
    }
    return lines;
  }
  function paraText(para) { return para.map(function (r) { return r.text; }).join(''); }
  // 상자 안 글의 높이(인치) 어림. 줄간격은 정확한 pt 값(size * lh)으로 넣으므로 그대로 셉니다.
  function measure(paras, size, w, o) {
    var maxEm = (w * 72 / size) * 0.93, sp = (o.spacing || 0) / size, lines = 0;
    paras.forEach(function (para) {
      paraText(para).split('\n').forEach(function (seg) { lines += wrapCount(seg, maxEm, !!o.bold, sp); });
    });
    return { lines: lines, h: lines * size * o.lh / 72 + Math.max(0, paras.length - 1) * (o.after || 0) / 72 };
  }
  function cutParas(paras, k) { // 앞에서 k 글자만 남기고 … 를 붙임
    var out = [], left = k;
    for (var i = 0; i < paras.length && left > 0; i++) {
      var para = [];
      for (var j = 0; j < paras[i].length && left > 0; j++) {
        var r = paras[i][j], t = r.text.slice(0, left);
        left -= t.length;
        para.push(Object.assign({}, r, { text: t }));
      }
      out.push(para);
    }
    if (!out.length) out.push([{ text: '' }]);
    var last = out[out.length - 1], lr = last[last.length - 1];
    lr.text = lr.text.replace(/[\s.,·]+$/, '') + '…';
    return out;
  }
  // 글자 크기를 max 에서 min 까지 줄여 보고, 그래도 넘치면 글을 잘라 맞춥니다(min 은 11pt 아래로 내려가지 않음).
  function fitText(p) {
    var size = p.max, min = Math.max(11, Math.min(p.min || 11, p.max));
    while (true) {
      var m = measure(p.paras, size, p.w, p);
      if (m.h <= p.h + 1e-6) { p.size = size; p.usedH = m.h; p.lines = m.lines; return p; }
      if (size <= min) break;
      size = Math.max(min, size - (size > 40 ? 4 : size > 20 ? 2 : 1));
    }
    p.size = min;
    var total = p.paras.reduce(function (t, para) { return t + paraText(para).length; }, 0), lo = 0, hi = total, best = cutParas(p.paras, 0);
    while (lo <= hi) {
      var mid = (lo + hi) >> 1, cut = cutParas(p.paras, mid);
      if (measure(cut, min, p.w, p).h <= p.h + 1e-6) { best = cut; lo = mid + 1; } else hi = mid - 1;
    }
    p.paras = best; p.trunc = true;
    var mm = measure(best, min, p.w, p); p.usedH = mm.h; p.lines = mm.lines;
    return p;
  }

  /* ------------------------------------------------------------------ */
  /* 그리기 단위(primitive): rect / ellipse / line / text (인치, pt)        */
  /* ------------------------------------------------------------------ */

  // content: '글\n글' | ['문단', ...] | [[{text, color, bold}], ...]
  function toParas(content) {
    if (content == null) return [];
    var list = Array.isArray(content) ? content : str(content).split('\n');
    var out = [];
    list.forEach(function (para) {
      if (Array.isArray(para)) { var runs = para.filter(function (r) { return r && r.text; }); if (runs.length) out.push(runs); }
      else str(para).split('\n').forEach(function (t) { if (t.trim()) out.push([{ text: t }]); });
    });
    return out;
  }
  function Ctx(spec, idx) {
    this.spec = spec; this.s = spec.slides[idx]; this.idx = idx;
    this.theme = own(PALETTES, spec.theme) ? spec.theme : 'brutal'; this.P = PALETTES[this.theme];
    this.prims = []; this.warn = [];
  }
  Ctx.prototype.rect = function (x, y, w, h, o) {
    o = o || {};
    var p = { t: 'rect', x: x, y: y, w: w, h: h, fill: o.fill || null, line: o.line || null, lw: o.lw || 1, role: o.role || 'shape' };
    this.prims.push(p); return p;
  };
  Ctx.prototype.ellipse = function (x, y, w, h, o) {
    var p = this.rect(x, y, w, h, o); p.t = 'ellipse'; return p;
  };
  Ctx.prototype.line = function (x, y, w, h, o) { // 가로(h=0) 또는 세로(w=0) 선
    o = o || {};
    var p = { t: 'line', x: x, y: y, w: w, h: h, color: o.color || this.P.ink, lw: o.lw || 1, role: o.role || 'line' };
    this.prims.push(p); return p;
  };
  Ctx.prototype.bar = function (x, y, w, h, role) { // 강조색 막대(빨강)
    return this.rect(x, y, w, h || 0.05, { fill: this.P.accent, role: role || 'bar' });
  };
  // 글 상자. 돌려준 p.usedH 로 다음 요소 자리를 잡습니다. 글이 비면 null.
  Ctx.prototype.text = function (x, y, w, h, content, o) {
    o = o || {};
    var paras = toParas(content);
    if (!paras.length) return null;
    var plain = paras.map(paraText).join('\n');
    var p = {
      t: 'text', x: x, y: y, w: w, h: h, paras: paras,
      max: o.size || 12, min: o.min || Math.min(o.size || 12, 11), bold: !!o.bold, color: o.color || this.P.body,
      font: o.font === 'disp' && !hasCJK(plain) ? 'disp' : 'kr', align: o.align || 'left', valign: o.valign || 'top',
      lh: o.lh || 1.4, after: o.after || 0, spacing: o.spacing || 0, rotate: o.rotate || 0, role: o.role || 'text'
    };
    fitText(p);
    if (p.trunc && !o.quiet) this.warn.push((o.what || '글') + '이 길어 상자에 맞게 줄였습니다.');
    this.prims.push(p); return p;
  };
  // 강조 단어 하나만 강조색으로 (첫 번째로 나온 곳만)
  Ctx.prototype.accentParas = function (txt, accent) {
    var paras = str(txt).split('\n').filter(function (t) { return t.trim(); }).map(function (t) { return [{ text: t }]; });
    if (!accent) return paras;
    for (var i = 0; i < paras.length; i++) {
      var t = paras[i][0].text, k = t.indexOf(accent);
      if (k < 0) continue;
      paras[i] = [{ text: t.slice(0, k) }, { text: accent, color: this.P.accent }, { text: t.slice(k + accent.length) }].filter(function (r) { return r.text; });
      break;
    }
    return paras;
  };
  function countAccent(L, accent) {
    var n = 0;
    L.prims.forEach(function (p) {
      if (p.t === 'text') { if (p.color === accent) n++; p.paras.forEach(function (para) { para.forEach(function (r) { if (r.color === accent) n++; }); }); }
      else if (p.fill === accent || p.line === accent || p.color === accent) n++;
    });
    return n;
  }

  /* ------------------------------------------------------------------ */
  /* 구성(layout) 12가지: 각 구성은 한 번만, primitive 목록으로 씁니다           */
  /* 13.333 x 7.5 in, 바깥 여백 왼쪽·위 0.75 in. 쪽번호·꼬리말 없음.           */
  /* ------------------------------------------------------------------ */

  var M = 0.75, RX = SW - 0.75, CW = RX - M, BOT = 7.0;
  function pad2(n) { n = String(n); return n.length < 2 ? '0' + n : n; }
  function latinWords(s) { return s && !hasCJK(s) ? s.toUpperCase().split(/\s+/).filter(function (x) { return /^[A-Z0-9&.+-]{1,12}$/.test(x); }) : []; }
  function bodyW(c, room) { return Math.min(room, c.theme === 'brutal' ? 3.6 : 4.6); }

  function kicker(c, x, y, w, t) {
    t = t == null ? c.s.kicker : t;
    if (!t) return null;
    var latin = !hasCJK(t);
    return c.text(x, y, w, 0.3, latin ? t.toUpperCase() : t, { size: 11, bold: true, color: c.P.mute, font: 'disp', spacing: latin ? 2.4 : 0.6, lh: 1.25, role: 'kicker', what: '라벨' });
  }
  function title(c, x, y, w, h, size, min, o) {
    o = o || {};
    return c.text(x, y, w, h, c.accentParas(o.text != null ? o.text : c.s.title, c.s.accent), {
      size: size, min: min, bold: true, color: c.P.ink, lh: o.lh || 1.12, valign: o.valign || 'top', role: o.role || 'title', what: o.what || '제목'
    });
  }
  function body(c, x, y, w, h, content, o) {
    o = o || {};
    return c.text(x, y, w, h, content, { size: o.size || 12, min: 11, color: o.color || c.P.body, lh: 1.5, after: o.after != null ? o.after : 6, role: o.role || 'body', what: o.what || '본문' });
  }
  function hasAccent(c) { return !!(c.s.accent && c.s.title.indexOf(c.s.accent) >= 0); }
  function itemParas(c, items) {
    return items.map(function (it) {
      var runs = [];
      var head = [it.num, it.head].filter(Boolean).join('  ');
      if (head) runs.push({ text: head, bold: true, color: c.P.ink });
      if (it.text) runs.push({ text: (head ? '  ' : '') + it.text.replace(/\n/g, ' ') });
      return runs;
    });
  }
  // 세로로 쌓는 짧은 글들(사이에 옅은 선). 넘치면 11pt, 그래도 넘치면 칸마다 잘라 맞춤.
  function stackItems(c, x, y, w, sepW, bottom, items, o) {
    o = o || {};
    var n = items.length, gap = 0.3, lh = 1.5;
    if (!n) return y;
    var hs = function (sz) { return items.map(function (t) { return measure(toParas(t), sz, w, { lh: lh }).h; }); };
    var size = 12, h = hs(size), room = bottom - y - gap * (n - 1);
    var sum = function (a) { return a.reduce(function (t, v) { return t + v; }, 0); };
    if (sum(h) > room) { size = 11; h = hs(size); }
    var cap = sum(h) > room ? Math.max(0.25, room / n) : 0;
    items.forEach(function (t, i) {
      if (i > 0) c.line(x, y - gap / 2, sepW, 0, { color: c.P.hair, lw: 0.75, role: 'sep' });
      var p = c.text(x, y, w, cap || h[i] + 0.01, t, { size: size, min: size, color: o.color || c.P.body, lh: lh, role: o.role || 'item', what: o.what || '항목' });
      y += (p ? p.usedH : 0) + gap;
    });
    return y;
  }
  // start 이후에 놓인 요소들의 글 아래 끝을 재고, 남는 세로 공간이 있으면 maxShift 까지 아래로 내립니다(글이 짧은 장의 균형).
  function bottomOf(c, start) {
    var b = 0;
    for (var i = start; i < c.prims.length; i++) { var p = c.prims[i]; b = Math.max(b, p.y + (p.t === 'text' ? p.usedH : p.h)); }
    return b;
  }
  function settle(c, start, maxShift) {
    var dy = Math.min(maxShift, BOT - 0.1 - bottomOf(c, start));
    if (dy <= 0.05) return 0;
    for (var i = start; i < c.prims.length; i++) {
      var p = c.prims[i];
      p.y += dy;
      if (p.t === 'text' && p.y + p.h > BOT) p.h = Math.max(p.usedH + 0.02, BOT - p.y);
    }
    return dy;
  }
  function frame(c, withMark) { // 얇은 세로 프레임 + 안쪽 낱글자 + 작은 빨간 세로 막대 + 돌린 작은 글
    var P = c.P, sp = c.spec, fx = 10.15, fw = RX - fx;
    c.rect(fx, M, fw, 6.0, { line: P.ink, lw: 1.5, role: 'frame' });
    var words = latinWords(c.s.kicker).slice(0, 4);
    if (!words.length && sp.date) words = [sp.date.replace(/[^0-9]/g, '').slice(0, 4)].filter(Boolean);
    if (words.length) c.text(fx + 0.3, M + 0.3, fw - 0.5, 2.3, words.join('\n'), { size: 16, min: 11, bold: true, color: P.ink, font: 'disp', spacing: 5, lh: 1.4, role: 'letters', quiet: true });
    if (withMark) c.bar(fx + 0.3, 5.9, 0.05, 0.45, 'mark');
    var meta = [sp.author, sp.date].filter(Boolean).join('  ·  ');
    if (meta) c.text(RX - 0.3 - 1.4, 5.0 - 0.15, 2.8, 0.3, meta, { size: 11, color: P.mute, spacing: 1.2, lh: 1.25, rotate: 270, role: 'meta', quiet: true });
  }

  var LAYOUT_FN = {
    cover: function (c) {
      var P = c.P, s = c.s, ghost = latinWords(s.kicker).sort(function (a, b) { return b.length - a.length; })[0];
      if (ghost) c.text(0.45, 2.3, 12.4, 2.3, ghost, { size: 150, min: 60, bold: true, color: P.ghost, font: 'disp', spacing: 14, lh: 1.0, valign: 'middle', role: 'ghost', quiet: true });
      kicker(c, M, M, 8.4);
      title(c, M, 1.45, 8.7, 3.4, 60, 28, { valign: 'bottom', lh: 1.08 });
      c.bar(M, 5.1, 1.4);
      body(c, M, 5.38, bodyW(c, 8.6), 1.2, s.body || c.spec.subtitle, { what: '부제' });
      frame(c, !hasAccent(c));
    },
    section: function (c) {
      var P = c.P, s = c.s, n = 0;
      for (var i = 0; i <= c.idx; i++) if (c.spec.slides[i].layout === 'section') n++;
      var num = /^\d{1,2}$/.test(s.kicker) ? pad2(s.kicker) : pad2(n);
      c.text(M - 0.06, 0.4, 6.6, 3.05, num, { size: 170, min: 90, bold: true, color: P.ink, font: 'disp', lh: 1.0, valign: 'bottom', role: 'num', quiet: true });
      c.line(M, 3.75, CW, 0, { role: 'line' });
      c.bar(M, 3.725, 1.6);
      title(c, M, 4.05, 7.5, 2.3, 40, 22);
      if (!/^\d{1,2}$/.test(s.kicker)) kicker(c, 8.85, 4.12, 3.7);
      body(c, 8.85, 4.55, bodyW(c, 3.7), 2.3, s.body);
    },
    statement: function (c) {
      var P = c.P, s = c.s;
      c.ellipse(8.2, 0.6, 6.3, 6.3, { line: P.hair, lw: 0.75, role: 'circle' });
      kicker(c, M, 1.2, 7.4);
      var t = title(c, M, 1.7, 10.5, 3.15, 52, 28, { lh: 1.14, what: '문장' });
      var y = 1.7 + (t ? t.usedH : 0) + 0.34;
      c.bar(M, y, 1.0);
      body(c, M, y + 0.3, bodyW(c, 6), BOT - y - 0.3, s.body);
    },
    text: function (c) {
      var P = c.P, s = c.s;
      kicker(c, M, M, 6.4);
      var t = title(c, M, 1.2, 6.4, 3.0, 40, 22);
      c.bar(M, 1.2 + (t ? t.usedH : 0) + 0.32, 0.9);
      var content = toParas(s.body).concat(itemParas(c, s.items));
      var vl = content.length ? c.line(7.7, M, 0, 6.0, { lw: 0.75, role: 'line' }) : { h: 0 };
      var bt = c.text(8.2, 1.25, bodyW(c, RX - 8.2), BOT - 1.25, content, { size: 12, min: 11, color: P.body, lh: 1.5, after: 9, role: 'body', what: '본문' });
      vl.h = Math.min(6.0, Math.max(3.75, Math.max(1.2 + (t ? t.usedH : 0), 1.25 + (bt ? bt.usedH : 0)) + 0.6 - M));
    },
    cards: function (c) {
      var P = c.P, s = c.s, n = s.items.length, gap = n > 3 ? 0.35 : 0.45, colW = (CW - gap * (n - 1)) / n;
      kicker(c, M, M, 8);
      var t = title(c, M, 1.15, 9.4, 1.75, 36, 22);
      var y0 = Math.min(3.4, Math.max(3.0, 1.15 + (t ? t.usedH : 0) + 0.55)), start = c.prims.length;
      s.items.forEach(function (it, i) {
        var x = M + i * (colW + gap), tw = Math.min(colW, 3.6);
        if (i === 0) c.bar(x, y0 - 0.025, colW); else c.line(x, y0, colW, 0, { role: 'line' });
        c.text(x, y0 + 0.22, tw, 0.3, it.num || pad2(i + 1), { size: 12, bold: true, color: P.ink, font: 'disp', spacing: 1.5, lh: 1.2, role: 'num', quiet: true });
        var hd = c.text(x, y0 + 0.62, tw, 1.0, it.head, { size: 19, min: 13, bold: true, color: P.ink, lh: 1.22, role: 'head', what: '카드 제목' });
        var ty = y0 + 0.62 + (hd ? hd.usedH + 0.16 : 0);
        c.text(x, ty, tw, BOT - ty, it.text, { size: 12, min: 11, color: P.body, lh: 1.5, role: 'item', what: '카드 글' });
      });
      settle(c, start, 1.5);
    },
    stat: function (c) {
      var P = c.P, s = c.s;
      kicker(c, M, M, 7.5);
      title(c, M, 1.15, 8.6, 1.15, 30, 20, { lh: 1.15 });
      c.text(M - 0.05, 2.3, CW, 2.8, s.stat.value, { size: 170, min: 60, bold: true, color: P.ink, font: 'disp', lh: 1.0, valign: 'bottom', role: 'stat', what: '숫자' });
      c.line(M, 5.3, CW, 0, { role: 'line' });
      c.bar(M, 5.275, 2.4);
      c.text(M, 5.55, 6.8, 1.2, s.stat.label, { size: 20, min: 13, bold: true, color: P.ink, lh: 1.25, role: 'label', what: '숫자 설명' });
      body(c, 8.85, 5.55, bodyW(c, 3.7), 1.35, s.body);
    },
    steps: function (c) {
      var P = c.P, s = c.s, n = s.items.length, gap = 0.32, colW = (CW - gap * (n - 1)) / n, dy = n >= 5 ? 0.26 : 0.34, y0 = 2.7;
      kicker(c, M, M, 8);
      title(c, M, 1.15, 8.8, 1.35, 36, 22);
      var start = c.prims.length;
      s.items.forEach(function (it, i) {
        var x = M + i * (colW + gap), y = y0 + i * dy, tw = Math.min(colW - 0.05, 3.6), ly = y + 0.905;
        c.text(x, y, colW, 0.8, it.num || pad2(i + 1), { size: 44, min: 24, bold: true, color: P.ink, font: 'disp', lh: 1.0, valign: 'bottom', role: 'num', quiet: true });
        var lx = x - (i > 0 ? gap / 2 : 0), lw = colW + (i > 0 ? gap / 2 : 0) + (i < n - 1 ? gap / 2 : 0);
        c.line(lx, ly, lw, 0, { role: 'line' });
        if (i < n - 1) c.line(x + colW + gap / 2, ly, 0, dy, { role: 'line' });
        if (i === 0) c.bar(x, ly - 0.025, colW * 0.5);
        var hd = c.text(x, y + 1.08, tw, 0.85, it.head, { size: 17, min: 12, bold: true, color: P.ink, lh: 1.22, role: 'head', what: '단계 제목' });
        var ty = y + 1.08 + (hd ? hd.usedH + 0.12 : 0);
        c.text(x, ty, tw, BOT - ty, it.text, { size: 12, min: 11, color: P.body, lh: 1.45, role: 'item', what: '단계 글' });
      });
      settle(c, start, 0.7);
    },
    timeline: function (c) {
      var P = c.P, s = c.s, n = s.items.length, colW = CW / n, ax = 3.95, big = n <= 5;
      kicker(c, M, M, 8);
      title(c, M, 1.15, 9.6, 1.35, 36, 22);
      c.line(M, ax, CW, 0, { role: 'line' });
      s.items.forEach(function (it, i) {
        var x = M + i * colW, tw = colW - 0.22, last = i === n - 1;
        if (it.num) c.text(x, ax - 0.95, colW - 0.08, 0.75, it.num, { size: big ? 28 : 20, min: 12, bold: true, color: P.ink, font: 'disp', lh: 1.05, valign: 'bottom', role: 'num', what: '연도' });
        c.ellipse(x, ax - 0.075, 0.15, 0.15, { fill: last ? P.accent : P.ink, role: last ? 'dot' : 'node' });
        var hd = c.text(x, ax + 0.32, tw, 0.8, it.head, { size: big ? 15 : 13, min: 11, bold: true, color: P.ink, lh: 1.25, role: 'head', what: '항목 제목' });
        var ty = ax + 0.32 + (hd ? hd.usedH + 0.1 : 0);
        c.text(x, ty, tw, BOT - ty, it.text, { size: big ? 12 : 11, min: 11, color: P.body, lh: 1.45, role: 'item', what: '항목 글' });
      });
    },
    table: function (c) {
      var P = c.P, s = c.s, tb = s.table, cols = Math.max(tb.head.length, tb.rows.length ? tb.rows[0].length : 0);
      kicker(c, M, M, 8);
      var t = title(c, M, 1.15, 10.2, 1.25, 32, 20);
      var top = Math.max(2.55, 1.15 + (t ? t.usedH : 0) + 0.45);
      var wts = [];
      for (var k = 0; k < cols; k++) {
        var len = 4;
        [tb.head].concat(tb.rows).forEach(function (r) { if (r[k]) len = Math.max(len, Math.min(28, r[k].length)); });
        wts.push(len);
      }
      var tot = wts.reduce(function (a, b) { return a + b; }, 0), xs = [], ws = [], acc = M;
      wts.forEach(function (wt) { var cw = CW * wt / tot; xs.push(acc); ws.push(cw); acc += cw; });
      var pad = 0.2, gapY = 0.13;  // gapY: 행 위아래 틈
      c.bar(M, top, 0.9);
      var y = top + 0.2;
      var cellH = function (row, size, bold) {
        return row.reduce(function (m, v, k) { return Math.max(m, v ? measure(toParas(v), size, ws[k] - pad, { lh: 1.4, bold: bold }).h : 0); }, size * 1.4 / 72);
      };
      if (tb.head.length) {
        var hh = Math.min(0.9, cellH(tb.head, 12, true));
        tb.head.forEach(function (v, k) { c.text(xs[k], y, ws[k] - pad, hh, v, { size: 12, min: 11, bold: true, color: P.ink, lh: 1.4, role: 'th', what: '표 머리' }); });
        y += hh + gapY;
        c.line(M, y, CW, 0, { role: 'line' });
        y += gapY;
      }
      var rows = tb.rows, size = 12, sum = function (a) { return a.reduce(function (x, v) { return x + v; }, 0); };
      if (rows.length > 5) gapY = 0.08;
      var room = function () { return BOT - y - rows.length * 2 * gapY; };
      var rh = rows.map(function (r) { return cellH(r, size, false); });
      if (sum(rh) > room()) { size = 11; rh = rows.map(function (r) { return cellH(r, size, false); }); }
      var one = size * 1.4 / 72;
      while (rows.length > 1 && rows.length * one > room()) { rows = rows.slice(0, -1); rh = rh.slice(0, rows.length); }
      if (rows.length < tb.rows.length) c.warn.push('표가 길어 앞 ' + rows.length + '행만 넣었습니다.');
      if (sum(rh) > room()) {
        // 한 줄짜리 행은 그대로 두고, 남는 높이를 긴 행들에 고르게 나눕니다(한 줄 아래로는 줄이지 않음)
        var rest = room(), longRows = rh.length;
        var sorted = rh.slice().sort(function (a, b) { return a - b; });
        for (var q = 0; q < sorted.length; q++) { var share = rest / (sorted.length - q); if (sorted[q] <= share) { rest -= sorted[q]; longRows--; } else break; }
        var capH = Math.max(one, rest / Math.max(1, longRows));
        rh = rh.map(function (h) { return Math.min(h, capH); });
      }
      rows.forEach(function (r, ri) {
        r.forEach(function (v, k) { c.text(xs[k], y, ws[k] - pad, rh[ri] + 0.005, v, { size: size, min: size, bold: k === 0, color: k === 0 ? P.ink : P.body, lh: 1.4, role: 'td', what: '표 칸' }); });
        y += rh[ri] + gapY;
        if (ri < rows.length - 1) c.line(M, y, CW, 0, { color: P.hair, lw: 0.75, role: 'sep' });
        y += gapY;
      });
    },
    compare: function (c) {
      var P = c.P, s = c.s;
      kicker(c, M, M, 8);
      var t = title(c, M, 1.15, 10.6, 1.3, 34, 20);
      var top = Math.max(2.7, 1.15 + (t ? t.usedH : 0) + 0.5), mid = SW / 2;
      var vl = c.line(mid, top - 0.05, 0, BOT - top + 0.05, { lw: 0.75, role: 'line' }), start = c.prims.length;
      [[s.left, M], [s.right, mid + 0.45]].forEach(function (pair, i) {
        var side = pair[0], x = pair[1];
        if (i === 1) c.bar(x, top - 0.02, 1.0);
        var hd = c.text(x, top + 0.2, 5.0, 0.95, side.head, { size: 24, min: 14, bold: true, color: P.ink, lh: 1.2, role: 'head', what: '비교 제목' });
        var y = top + 0.2 + (hd ? hd.usedH + 0.32 : 0);
        stackItems(c, x, y, bodyW(c, 5.0), 5.0, BOT, side.items, { what: '비교 항목' });
      });
      vl.h = Math.min(BOT - top + 0.05, Math.max(2.4, bottomOf(c, start) + 0.35 - top));
    },
    quote: function (c) {
      var P = c.P, s = c.s, q = s.quote;
      c.text(M - 0.08, 0.35, 2.4, 2.3, '\u201C', { size: 160, min: 120, bold: true, color: c.theme === 'navy' ? P.hair : P.hair, font: 'disp', lh: 1.0, role: 'mark', quiet: true });
      kicker(c, 8.85, M, 3.7);
      if (s.title) c.text(8.85, M + 0.42, 3.7, 0.95, s.title, { size: 14, min: 11, bold: true, color: P.ink, lh: 1.3, role: 'title', what: '제목' });
      var qt = c.text(M, 2.35, 10.6, 3.25, c.accentParas(q.text, s.accent), { size: 38, min: 20, bold: true, color: P.ink, lh: 1.3, role: 'quote', what: '인용문' });
      var y = 2.35 + (qt ? qt.usedH : 0) + 0.42;
      c.bar(M, y, 0.7);
      c.text(M, y + 0.2, 7.5, 0.35, q.by, { size: 12, color: P.mute, lh: 1.3, role: 'by', what: '출처' });
    },
    closing: function (c) {
      var P = c.P, s = c.s, acc = !!(s.accent && s.title.indexOf(s.accent) >= 0);
      kicker(c, M, M, 8.4);
      title(c, M, 1.6, 8.8, 3.25, 56, 26, { valign: 'bottom', lh: 1.1 });
      c.bar(M, 5.1, 1.4);
      body(c, M, 5.38, bodyW(c, 8.6), 1.45, s.body);
      if (s.items.length) {
        c.rect(10.15, M, RX - 10.15, 6.0, { line: P.ink, lw: 1.5, role: 'frame' });
        c.text(10.45, M + 0.3, RX - 10.15 - 0.55, 5.3, itemParas(c, s.items), { size: 12, min: 11, color: P.body, lh: 1.5, after: 10, role: 'items', what: '마무리 항목' });
        if (!acc) c.bar(10.45, 5.9, 0.05, 0.45, 'mark');
      } else frame(c, !acc);
    }
  };

  function assignNames(c) {
    var names = c.s.names || {}, seen = {}, morph = c.spec.morph;
    c.prims.forEach(function (p) {
      var r = p.role || p.t;
      seen[r] = (seen[r] || 0) + 1;
      var key = seen[r] > 1 ? r + seen[r] : r;
      p.name = own(names, key) ? names[key] : morph ? '!!' + key : key;
    });
  }
  function layoutSlide(spec, idx) {
    var c = new Ctx(spec, idx);
    (own(LAYOUT_FN, c.s.layout) ? LAYOUT_FN[c.s.layout] : LAYOUT_FN.text)(c);
    assignNames(c);
    return { bg: c.P.bg, prims: c.prims, warn: c.warn };
  }

  /* ------------------------------------------------------------------ */
  /* 뒷단 1: pptxgenjs (진짜 .pptx, 글은 모두 편집되는 글 상자)                 */
  /* ------------------------------------------------------------------ */

  function addPrim(pptx, ps, p) {
    var o = { x: round(p.x), y: round(p.y), w: round(p.w), h: round(p.h) };
    if (p.name) o.objectName = p.name;
    if (p.t === 'rect' || p.t === 'ellipse') {
      o.fill = p.fill ? { color: p.fill } : { type: 'none' };
      o.line = p.line ? { color: p.line, width: p.lw } : { type: 'none' };
      ps.addShape(p.t === 'rect' ? pptx.ShapeType.rect : pptx.ShapeType.ellipse, o);
    } else if (p.t === 'line') {
      o.line = { color: p.color, width: p.lw };
      ps.addShape(pptx.ShapeType.line, o);
    } else if (p.t === 'text') {
      var runs = [];
      p.paras.forEach(function (para, pi) {
        para.forEach(function (r, ri) {
          var ro = {};
          if (r.color) ro.color = r.color;
          if (r.bold != null) ro.bold = !!r.bold;
          if (ri === para.length - 1 && pi < p.paras.length - 1) ro.breakLine = true;
          runs.push({ text: r.text, options: ro });
        });
      });
      o.fontFace = p.font === 'disp' ? FONT_LATIN : FONT_KR;
      o.fontSize = p.size; o.bold = p.bold; o.color = p.color;
      o.align = p.align; o.valign = p.valign; o.inset = 0; o.wrap = true; o.lang = 'ko-KR';
      o.lineSpacing = round(p.size * p.lh, 1);
      if (p.after) o.paraSpaceAfter = p.after;
      if (p.spacing) o.charSpacing = p.spacing;
      if (p.rotate) o.rotate = p.rotate;
      ps.addText(runs, o);
    }
  }

  function buildDeck(spec, opts) {
    return Promise.resolve().then(function () {
      opts = opts || {};
      var v = normalizeDeck(spec, []);
      if (opts.theme && own(THEMES, opts.theme)) v.theme = opts.theme;
      if (opts.morph != null) v.morph = !!opts.morph;
      var Pptx = getPptxGen(), pptx = new Pptx();
      pptx.layout = 'LAYOUT_WIDE';
      pptx.title = v.title; pptx.subject = v.subtitle.replace(/\n/g, ' ').slice(0, 200);
      pptx.author = v.author || 'PptKit'; pptx.company = v.author || '';
      pptx.theme = { headFontFace: FONT_KR, bodyFontFace: FONT_KR };
      v.slides.forEach(function (s, i) {
        var L = layoutSlide(v, i), ps = pptx.addSlide();
        ps.background = { color: L.bg };
        L.prims.forEach(function (p) { addPrim(pptx, ps, p); });
        if (s.notes) ps.addNotes(s.notes);
      });
      return pptx.write({ outputType: 'arraybuffer' }).then(function (ab) {
        return v.morph ? morphPackage(ab) : ab;
      }).then(function (ab) { return new Blob([ab], { type: PPTX_MIME }); });
    });
  }

  // 만든 덱에 모핑 전환 넣기 (pptxgenjs 가 못 쓰므로 JSZip 으로 고침): 2..N 장
  function morphPackage(ab) {
    var JSZip = getJSZip();
    return JSZip.loadAsync(ab).then(function (zip) {
      return readPres(zip).then(function (pres) {
        var jobs = pres.slides.slice(1).map(function (sl) {
          return zip.file(sl.path).async('string').then(function (xml) {
            var nx = setMorphTransition(xml);
            if (nx) zip.file(sl.path, nx);
          });
        });
        return Promise.all(jobs);
      }).then(function () {
        return zip.generateAsync({ type: 'arraybuffer', compression: 'DEFLATE', compressionOptions: { level: 6 }, mimeType: PPTX_MIME });
      });
    });
  }

  /* ------------------------------------------------------------------ */
  /* 뒷단 2: HTML 미리보기 (같은 primitive, cqw 단위, 인라인 스타일만)          */
  /* ------------------------------------------------------------------ */

  function cq(inch) { return round(inch / SW * 100, 3) + 'cqw'; }
  var VALIGN = { top: 'flex-start', middle: 'center', bottom: 'flex-end' };
  function primHtml(p) {
    var box = 'position:absolute;left:' + cq(p.x) + ';top:' + cq(p.y) + ';width:' + cq(p.w) + ';height:' + cq(p.h) + ';';
    if (p.t === 'rect' || p.t === 'ellipse') {
      return '<div style="' + box + 'box-sizing:border-box;' + (p.fill ? 'background:#' + p.fill + ';' : '')
        + (p.line ? 'border:max(1px,' + cq(p.lw / 72) + ') solid #' + p.line + ';' : '') + (p.t === 'ellipse' ? 'border-radius:50%;' : '') + '"></div>';
    }
    if (p.t === 'line') {
      var th = 'max(1px,' + cq(p.lw / 72) + ')';
      return '<div style="position:absolute;left:' + cq(p.x) + ';top:' + cq(p.y) + ';width:' + (p.w ? cq(p.w) : th) + ';height:' + (p.h ? cq(p.h) : th) + ';background:#' + p.color + ';"></div>';
    }
    var st = box + 'display:flex;flex-direction:column;justify-content:' + (VALIGN[p.valign] || 'flex-start') + ';margin:0;padding:0;'
      + 'font-family:' + (p.font === 'disp' ? CSS_LATIN : CSS_KR) + ';font-size:' + cq(p.size / 72) + ';line-height:' + cq(p.size * p.lh / 72) + ';'
      + 'font-weight:' + (p.bold ? 700 : 400) + ';color:#' + p.color + ';text-align:' + (p.align === 'center' ? 'center' : p.align === 'right' ? 'right' : 'left') + ';'
      + 'white-space:pre-wrap;word-break:keep-all;overflow-wrap:anywhere;'
      + (p.spacing ? 'letter-spacing:' + cq(p.spacing / 72) + ';' : '') + (p.rotate ? 'transform:rotate(' + Number(p.rotate) + 'deg);' : '');
    var inner = p.paras.map(function (para, i) {
      var gap = i < p.paras.length - 1 && p.after ? 'margin:0 0 ' + cq(p.after / 72) + ' 0' : 'margin:0';
      return '<div style="' + gap + '">' + para.map(function (r) {
        var rs = (r.color ? 'color:#' + r.color + ';' : '') + (r.bold != null ? 'font-weight:' + (r.bold ? 700 : 400) + ';' : '');
        var t = htmlEsc(r.text).replace(/\n/g, '<br>');
        return rs ? '<span style="' + rs + '">' + t + '</span>' : t;
      }).join('') + '</div>';
    }).join('');
    return '<div style="' + st + '">' + inner + '</div>';
  }

  function previewHtml(spec) {
    var v = normalizeDeck(spec, []);
    var out = '<div class="ppt-prev" style="display:flex;flex-direction:column;gap:18px;">';
    v.slides.forEach(function (s, i) {
      var L = layoutSlide(v, i);
      out += '<div style="display:flex;flex-direction:column;gap:6px;">'
        + '<div style="container-type:inline-size;position:relative;width:100%;aspect-ratio:16/9;overflow:hidden;background:#' + L.bg + ';box-shadow:0 0 0 1px rgba(0,0,0,.12);">'
        + L.prims.map(primHtml).join('') + '</div>'
        + '<div style="font:12px/1.4 ' + CSS_KR + ';color:#8a8a8a;">' + (i + 1) + ' · ' + htmlEsc(s.layout) + (s.notes ? ' · 발표 노트 있음' : '') + '</div></div>';
    });
    return out + '</div>';
  }

  /* ------------------------------------------------------------------ */
  /* 작은 XML 읽개: 원문 위치(start/openEnd/closeStart/end)를 기억하므로        */
  /* 고칠 곳만 잘라 바꾸고 나머지 바이트는 그대로 둡니다.                       */
  /* ------------------------------------------------------------------ */

  var TAG_SRC = '<!--[\\s\\S]*?-->|<!\\[CDATA\\[[\\s\\S]*?\\]\\]>|<\\?[\\s\\S]*?\\?>|<!DOCTYPE[^>]*>|<\\/([A-Za-z_][\\w.:-]*)\\s*>|<([A-Za-z_][\\w.:-]*)((?:\\s+[^\\s=\\/>]+\\s*=\\s*(?:"[^"]*"|\'[^\']*\'))*)\\s*(\\/?)>';
  function parseXml(src) {
    var doc = { name: '#doc', attrRaw: '', start: 0, openEnd: 0, closeStart: src.length, end: src.length, children: [], parent: null };
    var stack = [doc], re = new RegExp(TAG_SRC, 'g'), m;
    while ((m = re.exec(src))) {
      var top = stack[stack.length - 1];
      if (m[1]) {
        var k = stack.length - 1;
        while (k > 0 && stack[k].name !== m[1]) k--;
        if (k === 0) continue;
        while (stack.length - 1 > k) { var orphan = stack.pop(); orphan.closeStart = orphan.end = m.index; }
        var el = stack.pop(); el.closeStart = m.index; el.end = m.index + m[0].length;
      } else if (m[2]) {
        var e2 = { name: m[2], attrRaw: m[3] || '', start: m.index, openEnd: m.index + m[0].length, children: [], parent: top };
        top.children.push(e2);
        if (m[4]) { e2.closeStart = e2.end = e2.openEnd; } else stack.push(e2);
      }
    }
    while (stack.length > 1) { var o = stack.pop(); o.closeStart = o.end = src.length; o.broken = true; doc.broken = true; }
    return doc;
  }
  function decodeEnt(s) {
    return String(s).replace(/&(#x[0-9a-fA-F]+|#\d+|lt|gt|amp|quot|apos);/g, function (all, e) {
      if (e === 'lt') return '<'; if (e === 'gt') return '>'; if (e === 'amp') return '&'; if (e === 'quot') return '"'; if (e === 'apos') return "'";
      var n = e.charAt(1) === 'x' ? parseInt(e.slice(2), 16) : parseInt(e.slice(1), 10);
      try { return n > 0 && n <= 0x10FFFF ? String.fromCodePoint(n) : ''; } catch (er) { return ''; }
    });
  }
  function attrs(el) {
    if (!el) return {};
    if (el._a) return el._a;
    var a = Object.create(null), re = /([^\s=]+)\s*=\s*(?:"([^"]*)"|'([^']*)')/g, m;
    while ((m = re.exec(el.attrRaw))) a[m[1]] = decodeEnt(m[2] != null ? m[2] : m[3]);
    el._a = a; return a;
  }
  function attr(el, n) { return attrs(el)[n]; }
  function kid(el, name) { if (!el) return null; for (var i = 0; i < el.children.length; i++) if (el.children[i].name === name) return el.children[i]; return null; }
  function kids(el, name) { return el ? el.children.filter(function (c) { return c.name === name; }) : []; }
  function chain(el) { for (var i = 1; i < arguments.length && el; i++) el = kid(el, arguments[i]); return el || null; }
  function xmlText(raw) {
    var out = '', re = /<!\[CDATA\[([\s\S]*?)\]\]>/g, last = 0, m;
    while ((m = re.exec(raw))) { out += decodeEnt(raw.slice(last, m.index).replace(/<[^>]*>/g, '')) + m[1]; last = re.lastIndex; }
    return out + decodeEnt(raw.slice(last).replace(/<[^>]*>/g, ''));
  }
  function inner(src, el) { return src.slice(el.openEnd, el.closeStart); }
  function outer(src, el) { return el ? src.slice(el.start, el.end) : ''; }
  // 겹치지 않는 [s,e) 구간들을 바꿉니다. 같은 자리에 넣기(s===e)가 먼저.
  function spliceAll(src, ops) {
    ops = ops.slice().sort(function (a, b) { return a.s - b.s || (a.e - a.s) - (b.e - b.s); });
    var out = '', pos = 0;
    ops.forEach(function (o) {
      if (o.s < pos) throw new Error('겹치는 고치기');
      out += src.slice(pos, o.s) + o.t; pos = o.e;
    });
    return out + src.slice(pos);
  }

  /* 패키지 관계(rels)와 슬라이드 순서 */
  function parseRels(xml) {
    var out = [];
    if (!xml) return out;
    var root = kid(parseXml(xml), 'Relationships');
    kids(root, 'Relationship').forEach(function (r) { var a = attrs(r); out.push({ id: a.Id || '', type: a.Type || '', target: a.Target || '', mode: a.TargetMode || '' }); });
    return out;
  }
  function relsPathOf(part) { var i = part.lastIndexOf('/'); return part.slice(0, i + 1) + '_rels/' + part.slice(i + 1) + '.rels'; }
  function resolveTarget(part, target) {
    var t = String(target || '');
    try { t = decodeURI(t); } catch (e) { /* keep */ }
    if (t.charAt(0) === '/') return t.slice(1);
    var segs = part.slice(0, Math.max(0, part.lastIndexOf('/'))).split('/').filter(Boolean);
    t.split('/').forEach(function (seg) { if (seg === '..') segs.pop(); else if (seg && seg !== '.') segs.push(seg); });
    return segs.join('/');
  }
  function zstr(zip, path) { var f = path && zip.file(path); return f ? f.async('string') : Promise.resolve(''); }
  function readPres(zip) {
    return zstr(zip, '_rels/.rels').then(function (rx) {
      var main = 'ppt/presentation.xml';
      parseRels(rx).forEach(function (r) { if (/\/officeDocument$/.test(r.type) && !r.mode) main = resolveTarget('', r.target); });
      if (!zip.file(main)) throw new Error('PPTX 파일이 아니거나 망가진 파일입니다 (presentation.xml 없음).');
      return Promise.all([zstr(zip, main), zstr(zip, relsPathOf(main))]).then(function (r) {
        var pres = kid(parseXml(r[0]), 'p:presentation'), rels = parseRels(r[1]), byId = Object.create(null);
        if (!pres) throw new Error('PPTX 파일이 아니거나 망가진 파일입니다 (presentation 없음).');
        rels.forEach(function (x) { byId[x.id] = x; });
        var slides = [], seen = Object.create(null);
        kids(kid(pres, 'p:sldIdLst'), 'p:sldId').forEach(function (sid) {
          var a = attrs(sid), rid = a['r:id'];
          if (!rid) Object.keys(a).forEach(function (k) { if (/:id$/.test(k) && byId[a[k]]) rid = a[k]; });
          var rel = byId[rid];
          if (!rel || rel.mode) return;
          var p = resolveTarget(main, rel.target);
          if (zip.file(p) && !seen[p]) { seen[p] = 1; slides.push({ path: p, rels: relsPathOf(p) }); }
        });
        var sz = kid(pres, 'p:sldSz'), nm = rels.filter(function (x) { return /\/notesMaster$/.test(x.type); })[0];
        return {
          main: main, mainRels: relsPathOf(main), slides: slides,
          cx: Number(attr(sz, 'cx')) || 12192000, cy: Number(attr(sz, 'cy')) || 6858000,
          notesMaster: nm && zip.file(resolveTarget(main, nm.target)) ? resolveTarget(main, nm.target) : null
        };
      });
    });
  }

  /* 모핑 전환 (buildDeck 와 applyEdits 가 같이 씀) */
  var NS = {
    a: 'http://schemas.openxmlformats.org/drawingml/2006/main',
    r: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    p: 'http://schemas.openxmlformats.org/presentationml/2006/main',
    mc: 'http://schemas.openxmlformats.org/markup-compatibility/2006',
    p14: 'http://schemas.microsoft.com/office/powerpoint/2010/main',
    rel: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    pkg: 'http://schemas.openxmlformats.org/package/2006/relationships'
  };
  var MORPH_XML = '<mc:AlternateContent xmlns:mc="' + NS.mc + '"><mc:Choice xmlns:p159="http://schemas.microsoft.com/office/powerpoint/2015/09/main" Requires="p159">'
    + '<p:transition spd="slow" p14:dur="1400"><p159:morph option="byObject"/></p:transition></mc:Choice>'
    + '<mc:Fallback><p:transition spd="slow"><p:fade/></p:transition></mc:Fallback></mc:AlternateContent>';
  // 슬라이드 XML 에 모핑 전환을 넣습니다(있던 전환은 뺌). p:clrMapOvr 바로 뒤(없으면 p:cSld 뒤), p:timing 앞.
  // 구조를 모르는 파일이면 null.
  function setMorphTransition(xml) {
    var doc = parseXml(xml), root = kid(doc, 'p:sld');
    if (!root || doc.broken || root.closeStart === root.openEnd) return null;
    var cSld = kid(root, 'p:cSld');
    if (!cSld) return null;
    var anchor = kid(root, 'p:clrMapOvr') || cSld, ops = [];
    root.children.forEach(function (ch) {
      if (ch.name === 'p:transition' || (ch.name === 'mc:AlternateContent' && /<p:transition\b/.test(outer(xml, ch)))) ops.push({ s: ch.start, e: ch.end, t: '' });
    });
    ops.push({ s: anchor.end, e: anchor.end, t: MORPH_XML });
    var open = xml.slice(root.start, root.openEnd), add = '';
    if (!/\sxmlns:mc=/.test(open)) add += ' xmlns:mc="' + NS.mc + '"';
    if (!/\sxmlns:p14=/.test(open)) add += ' xmlns:p14="' + NS.p14 + '"';
    if (add) ops.push({ s: root.start, e: root.openEnd, t: open.replace(/\s*>$/, add + '>') });
    return spliceAll(xml, ops);
  }

  /* ------------------------------------------------------------------ */
  /* inspect / outline: 올린 PPTX 읽기                                     */
  /* ------------------------------------------------------------------ */

  // 자리표시자(ph) 종류 -> title | body | other. 자리표시자가 아니면 text.
  function phKind(ph) {
    if (!ph) return 'text';
    var t = attr(ph, 'type') || 'obj';
    if (t === 'title' || t === 'ctrTitle') return 'title';
    if (t === 'body' || t === 'subTitle' || t === 'obj') return 'body';
    return 'other';
  }
  // 슬라이드(또는 노트 쪽) 도형 목록. 그룹 안, mc:AlternateContent 의 mc:Choice 안, 표 칸까지 문서 순서대로.
  // 표 칸 id 는 "표 id.r행c열"(1부터, 병합되어 가려진 칸은 뺌). 글이 들어갈 수 없는 도형은 txBody 가 null.
  function walkShapes(src, tree, out) {
    tree.children.forEach(function (el) {
      var c;
      if (el.name === 'p:sp') {
        var nv = kid(el, 'p:nvSpPr');
        c = kid(nv, 'p:cNvPr');
        var kind = phKind(chain(nv, 'p:nvPr', 'p:ph')), nm = attr(c, 'name') || '';
        if (kind === 'text' && /^(!!)?title$/i.test(nm)) kind = 'title'; // PptKit 이 만든 덱의 제목 상자
        out.push({ id: attr(c, 'id') || '', name: nm, kind: kind, cNvPr: c, txBody: kid(el, 'p:txBody'), obj: 'shape' });
      } else if (el.name === 'p:grpSp') {
        c = chain(el, 'p:nvGrpSpPr', 'p:cNvPr');
        out.push({ id: attr(c, 'id') || '', name: attr(c, 'name') || '', kind: 'other', cNvPr: c, txBody: null, obj: 'group' });
        walkShapes(src, el, out);
      } else if (el.name === 'mc:AlternateContent') {
        var ch = kid(el, 'mc:Choice');
        if (ch) walkShapes(src, ch, out);
      } else if (el.name === 'p:pic' || el.name === 'p:cxnSp') {
        c = chain(el, el.name === 'p:pic' ? 'p:nvPicPr' : 'p:nvCxnSpPr', 'p:cNvPr');
        out.push({ id: attr(c, 'id') || '', name: attr(c, 'name') || '', kind: 'other', cNvPr: c, txBody: null, obj: el.name === 'p:pic' ? 'picture' : 'line' });
      } else if (el.name === 'p:graphicFrame') {
        c = chain(el, 'p:nvGraphicFramePr', 'p:cNvPr');
        var fid = attr(c, 'id') || '', tbl = chain(el, 'a:graphic', 'a:graphicData', 'a:tbl');
        out.push({ id: fid, name: attr(c, 'name') || '', kind: tbl ? 'table' : 'other', cNvPr: c, txBody: null, obj: tbl ? 'table' : 'frame' });
        if (!tbl) return;
        kids(tbl, 'a:tr').forEach(function (tr, ri) {
          kids(tr, 'a:tc').forEach(function (tc, ci) {
            var a = attrs(tc);
            if (/^(1|true)$/.test(a.hMerge || '') || /^(1|true)$/.test(a.vMerge || '')) return;
            out.push({ id: fid + '.r' + (ri + 1) + 'c' + (ci + 1), name: attr(c, 'name') || '', kind: 'table', cNvPr: null, txBody: kid(tc, 'a:txBody'), obj: null, cell: true });
          });
        });
      }
    });
    return out;
  }
  function paraTextXml(src, p) {
    var t = '';
    p.children.forEach(function (ch) {
      if (ch.name === 'a:r' || ch.name === 'a:fld') { var at = kid(ch, 'a:t'); if (at) t += xmlText(inner(src, at)); }
      else if (ch.name === 'a:br') t += '\n';
    });
    return clean(t, 20000);
  }
  function parasOf(src, txBody) { return kids(txBody, 'a:p').map(function (p) { return paraTextXml(src, p); }); }
  function notesPathOf(slidePath, rels) {
    var r = rels.filter(function (x) { return /\/notesSlide$/.test(x.type) && !x.mode; })[0];
    return r ? resolveTarget(slidePath, r.target) : null;
  }
  function notesBodyOf(xml) {
    var doc = parseXml(xml), tree = chain(doc, 'p:notes', 'p:cSld', 'p:spTree');
    if (!tree || doc.broken) return null;
    return walkShapes(xml, tree, []).filter(function (x) { return x.kind === 'body' && x.txBody; })[0] || null;
  }
  function readSlide(zip, sl, n) {
    return Promise.all([zstr(zip, sl.path), zstr(zip, sl.rels)]).then(function (r) {
      var xml = r[0], doc = parseXml(xml), tree = chain(doc, 'p:sld', 'p:cSld', 'p:spTree');
      var shapes = [], objects = [];
      (tree ? walkShapes(xml, tree, []) : []).forEach(function (s) {
        var paras = s.txBody ? parasOf(xml, s.txBody) : [], text = paras.join('\n');
        if (text.trim()) shapes.push({ id: s.id, name: clean(s.name, LIM.name), kind: s.kind, text: text, paragraphs: paras });
        else if (s.obj && objects.length < 60) objects.push({ id: s.id, name: clean(s.name, LIM.name), kind: s.obj });
      });
      var t = shapes.filter(function (x) { return x.kind === 'title'; })[0];
      var np = notesPathOf(sl.path, parseRels(r[1]));
      return zstr(zip, np).then(function (nx) {
        var body = nx ? notesBodyOf(nx) : null;
        var notes = body ? parasOf(nx, body.txBody).join('\n').trim() : '';
        return { n: n, title: t ? t.text.replace(/\s*\n\s*/g, ' ') : '', notes: notes, shapes: shapes, objects: objects };
      });
    });
  }
  function inspect(buf) {
    return Promise.resolve().then(function () {
      return getJSZip().loadAsync(buf);
    }).then(function (zip) {
      return readPres(zip).then(function (pres) {
        return Promise.all(pres.slides.map(function (sl, i) { return readSlide(zip, sl, i + 1); })).then(function (slides) {
          return { slideCount: slides.length, size: { w: round(pres.cx / 914400), h: round(pres.cy / 914400), cx: pres.cx, cy: pres.cy }, slides: slides };
        });
      });
    });
  }

  // 모델에게 줄 개요. 문단은 " | " 로 잇고, 문단 안 줄바꿈은 \n 으로 보입니다(JSON 문자열 표기).
  function outline(info, opts) {
    var max = (opts && opts.maxChars) || 60000, per = 700, out = [], len = 0;
    var q = function (s, cap) { s = String(s); return s.length > cap ? JSON.stringify(s.slice(0, cap)) + ' …(이하 ' + (s.length - cap) + '자 생략)' : JSON.stringify(s); };
    var slides = info && Array.isArray(info.slides) ? info.slides : [];
    for (var i = 0; i < slides.length; i++) {
      var s = slides[i], lines = ['[슬라이드 ' + s.n + ']'];
      (s.shapes || []).forEach(function (sh) {
        var paras = Array.isArray(sh.paragraphs) ? sh.paragraphs : String(sh.text || '').split('\n');
        var nm = /^!!/.test(sh.name || '') ? ' (이름 ' + JSON.stringify(sh.name) + ')' : '';
        lines.push('#' + sh.id + ' ' + sh.kind + nm + ': ' + q(paras.join(' | '), per));
      });
      if (s.objects && s.objects.length) lines.push('#objects: ' + s.objects.slice(0, 20).map(function (o) { return '#' + o.id + ' ' + o.kind + ' ' + JSON.stringify(String(o.name || '').slice(0, 40)); }).join(', '));
      if (s.notes) lines.push('#notes: ' + q(String(s.notes).replace(/\n/g, ' | '), 800));
      var block = lines.join('\n');
      if (len + block.length + 2 > max) {
        out.push('…(슬라이드 ' + s.n + '~' + slides[slides.length - 1].n + ' 생략: 글이 너무 깁니다)');
        break;
      }
      out.push(block); len += block.length + 2;
    }
    return out.join('\n\n');
  }

  /* ------------------------------------------------------------------ */
  /* applyEdits: 올린 PPTX 의 글·노트·도형 이름만 고치고 나머지는 그대로          */
  /* ------------------------------------------------------------------ */

  var REL_T = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/';
  var CT_NOTES = 'application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml';
  var XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';

  function normId(v) { return str(v).trim().replace(/^#/, '').toLowerCase().replace(/\s+/g, ''); }
  function slideNo(v) { var m = /^\s*#?(\d{1,4})\s*$/.exec(str(v)); return m ? Number(m[1]) : 0; }
  function newParas(v, max) {
    var a = Array.isArray(v) ? v : [v];
    return a.slice(0, LIM.paras).map(function (t) { return clean(str(t), max).replace(/\s+$/, ''); });
  }
  function renameTag(xml, from, to) {
    return xml.replace(new RegExp('^<' + from + '(?=[\\s/>])'), function () { return '<' + to; })
      .replace(new RegExp('</' + from + '>$'), function () { return '</' + to + '>'; });
  }
  // tpl 문단의 a:pPr 와 첫 글 조각의 a:rPr 를 그대로 쓰는 새 문단. \n 은 <a:br/>.
  function buildPara(src, tpl, text) {
    var pPr = '', rPr = '', end = '';
    if (tpl) {
      pPr = outer(src, kid(tpl, 'a:pPr'));
      var r = kid(tpl, 'a:r') || kid(tpl, 'a:fld'), rp = r && kid(r, 'a:rPr'), ep = kid(tpl, 'a:endParaRPr');
      if (rp) rPr = outer(src, rp); else if (ep) rPr = renameTag(outer(src, ep), 'a:endParaRPr', 'a:rPr');
      end = ep ? outer(src, ep) : rPr ? renameTag(rPr, 'a:rPr', 'a:endParaRPr') : '';
    }
    var body = '';
    text.split('\n').forEach(function (seg, i) {
      if (i > 0) body += rPr ? '<a:br>' + rPr + '</a:br>' : '<a:br/>';
      if (seg) body += '<a:r>' + rPr + '<a:t>' + xmlEsc(seg) + '</a:t></a:r>';
    });
    return '<a:p>' + pPr + body + end + '</a:p>';
  }
  // 글 상자(txBody)의 문단을 바꾸는 고치기 하나. 같으면 null. 글이 같은 문단은 원문 그대로 둡니다.
  function rebuildParas(src, txBody, paras) {
    var ps = kids(txBody, 'a:p'), before = ps.map(function (p) { return paraTextXml(src, p); });
    if (before.length === paras.length && before.every(function (t, i) { return t === paras[i]; })) return null;
    var pieces = paras.map(function (t, i) {
      if (i < ps.length && before[i] === t) return outer(src, ps[i]);
      return buildPara(src, ps.length ? ps[Math.min(i, ps.length - 1)] : null, t);
    });
    if (!pieces.length) pieces = [buildPara(src, ps[0] || null, '')];
    var op;
    if (ps.length) op = { s: ps[0].start, e: ps[ps.length - 1].end, t: pieces.join('') };
    else if (txBody.closeStart === txBody.openEnd) op = { s: txBody.start, e: txBody.end, t: '<' + txBody.name + '><a:bodyPr/><a:lstStyle/>' + pieces.join('') + '</' + txBody.name + '>' };
    else op = { s: txBody.closeStart, e: txBody.closeStart, t: pieces.join('') };
    return { op: op, before: before.join('\n'), after: paras.join('\n') };
  }
  function relTarget(from, to) {
    var a = from.split('/').slice(0, -1), b = to.split('/'), i = 0;
    while (i < a.length && i < b.length - 1 && a[i] === b[i]) i++;
    var up = []; for (var k = i; k < a.length; k++) up.push('..');
    return up.concat(b.slice(i)).join('/');
  }
  function notesXml(paras) {
    var ps = (paras.length ? paras : ['']).map(function (t) {
      var rp = '<a:rPr lang="ko-KR" altLang="en-US" dirty="0"/>', body = '';
      t.split('\n').forEach(function (seg, i) { if (i > 0) body += '<a:br>' + rp + '</a:br>'; if (seg) body += '<a:r>' + rp + '<a:t>' + xmlEsc(seg) + '</a:t></a:r>'; });
      return '<a:p>' + body + '<a:endParaRPr lang="ko-KR" altLang="en-US" dirty="0"/></a:p>';
    }).join('');
    return XML_DECL + '<p:notes xmlns:a="' + NS.a + '" xmlns:r="' + NS.r + '" xmlns:p="' + NS.p + '"><p:cSld><p:spTree>'
      + '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>'
      + '<p:sp><p:nvSpPr><p:cNvPr id="2" name="Slide Image Placeholder 1"/><p:cNvSpPr><a:spLocks noGrp="1" noRot="1" noChangeAspect="1"/></p:cNvSpPr><p:nvPr><p:ph type="sldImg"/></p:nvPr></p:nvSpPr><p:spPr/></p:sp>'
      + '<p:sp><p:nvSpPr><p:cNvPr id="3" name="Notes Placeholder 2"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr><p:spPr/>'
      + '<p:txBody><a:bodyPr/><a:lstStyle/>' + ps + '</p:txBody></p:sp></p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:notes>';
  }

  function applyEdits(buf, edits, opts) {
    opts = opts || {};
    var st = { zip: null, pres: null, mod: Object.create(null), skipped: [], changed: [], applied: 0, morphed: [] };
    var skip = function (e, reason) { st.skipped.push({ edit: e, reason: reason }); };
    var get = function (path) { return path in st.mod ? Promise.resolve(st.mod[path]) : zstr(st.zip, path); };
    return Promise.resolve().then(function () {
      return getJSZip().loadAsync(buf);
    }).then(function (zip) {
      st.zip = zip;
      return readPres(zip);
    }).then(function (pres) {
      st.pres = pres;
      var list = Array.isArray(edits) ? edits : [];
      if (!Array.isArray(edits) && edits != null) skip(null, 'edits 가 배열이 아닙니다');
      if (list.length > LIM.edits) { skip({ count: list.length - LIM.edits }, '고칠 곳이 너무 많아 앞 ' + LIM.edits + '곳만 씁니다'); list = list.slice(0, LIM.edits); }
      var bySlide = Object.create(null);
      list.forEach(function (e) {
        if (!isObj(e)) return skip(e, '형식이 맞지 않는 고치기입니다');
        var n = slideNo(e.slide);
        if (!n || n > pres.slides.length) return skip(e, '없는 슬라이드 번호입니다 (' + text1(e.slide, 12) + ')');
        var hasShape = normId(e.shape) !== '';
        if (e.notes == null && !(hasShape && (e.text != null || e.name != null))) return skip(e, '알 수 없는 고치기 형식입니다 (text, notes, name 중 하나가 필요)');
        (bySlide[n] = bySlide[n] || []).push(e);
      });
      var nums = Object.keys(bySlide).map(Number).sort(function (a, b) { return a - b; });
      return nums.reduce(function (p, n) { return p.then(function () { return editSlide(st, n, bySlide[n], skip, get); }); }, Promise.resolve());
    }).then(function () {
      if (!opts.morph) return;
      return st.pres.slides.slice(1).reduce(function (p, sl, i) {
        return p.then(function () {
          return get(sl.path).then(function (xml) {
            var nx = setMorphTransition(xml);
            if (nx) { st.mod[sl.path] = nx; st.morphed.push(i + 2); } else skip({ slide: i + 2, morph: true }, '이 슬라이드에는 모핑 전환을 넣지 못했습니다');
          });
        });
      }, Promise.resolve());
    }).then(function () {
      return writeZip(st);
    }).then(function (blob) {
      return { blob: blob, applied: st.applied, skipped: st.skipped, changed: st.changed, morphed: st.morphed };
    });
  }

  function editSlide(st, n, list, skip, get) {
    var sl = st.pres.slides[n - 1];
    return get(sl.path).then(function (xml) {
      var doc = parseXml(xml), tree = chain(doc, 'p:sld', 'p:cSld', 'p:spTree');
      var shapeEdits = list.filter(function (e) { return e.notes == null; }), noteEdits = list.filter(function (e) { return e.notes != null; });
      if (!tree || doc.broken) { shapeEdits.forEach(function (e) { skip(e, '이 슬라이드 구조를 읽지 못했습니다'); }); shapeEdits = []; }
      var byId = Object.create(null);
      if (tree) walkShapes(xml, tree, []).forEach(function (en) { if (en.id && !(normId(en.id) in byId)) byId[normId(en.id)] = en; });
      var textMap = Object.create(null), nameMap = Object.create(null), order = [];
      shapeEdits.forEach(function (e) {
        var id = normId(e.shape), en = byId[id];
        if (!en) return skip(e, '이 슬라이드에 없는 도형 id 입니다 (#' + text1(e.shape, 24) + ')');
        if (order.indexOf(id) < 0) {
          if (order.length >= LIM.shapes) return skip(e, '한 장에서 고칠 도형은 ' + LIM.shapes + '개까지입니다');
          order.push(id);
        }
        if (e.text != null) { if (textMap[id]) skip(textMap[id], '같은 도형을 다시 고쳐 마지막 것만 썼습니다'); textMap[id] = e; }
        if (e.name != null) { if (nameMap[id]) skip(nameMap[id], '같은 도형 이름을 다시 바꿔 마지막 것만 썼습니다'); nameMap[id] = e; }
      });
      var ops = [], changes = [];
      order.forEach(function (id) {
        var en = byId[id], e = textMap[id], ne = nameMap[id];
        if (e) {
          if (!en.txBody) skip(e, '글을 넣을 수 없는 도형입니다 (그림·선·묶음 등)');
          else {
            var r = rebuildParas(xml, en.txBody, newParas(e.text, en.kind === 'title' ? LIM.title : LIM.para));
            if (!r) skip(e, '이미 같은 글이라 바꾸지 않았습니다');
            else { ops.push(r.op); changes.push({ slide: n, shape: en.id, field: 'text', before: r.before, after: r.after }); }
          }
        }
        if (ne) {
          var c = en.cNvPr, nm = text1(ne.name, LIM.name), before = c ? attr(c, 'name') || '' : '';
          if (!c) skip(ne, '표 칸에는 이름이 없습니다. 표 id 로 바꾸세요');
          else if (!nm) skip(ne, '빈 이름입니다');
          else if (nm === before) skip(ne, '이미 같은 이름입니다');
          else {
            var open = xml.slice(c.start, c.openEnd), val = xmlEsc(nm);
            var nx = /\sname\s*=\s*("[^"]*"|'[^']*')/.test(open)
              ? open.replace(/(\sname\s*=\s*)("[^"]*"|'[^']*')/, function (all, pre) { return pre + '"' + val + '"'; })
              : open.replace(/\s*(\/?)>$/, function (all, sl2) { return ' name="' + val + '"' + sl2 + '>'; });
            ops.push({ s: c.start, e: c.openEnd, t: nx });
            changes.push({ slide: n, shape: en.id, field: 'name', before: before, after: nm });
          }
        }
      });
      if (ops.length) {
        st.mod[sl.path] = spliceAll(xml, ops);
        st.changed = st.changed.concat(changes);
        st.applied += changes.length;
      }
      if (!noteEdits.length) return;
      noteEdits.slice(0, -1).forEach(function (e) { skip(e, '같은 장의 노트를 다시 고쳐 마지막 것만 썼습니다'); });
      return editNotes(st, n, sl, noteEdits[noteEdits.length - 1], skip, get);
    });
  }

  var KEEP_SEP = '── 원본 메모 ──';
  // keepOriginal: true 이면 새 글 아래에 "── 원본 메모 ──" 와 원래 노트를 붙입니다(여러 번 돌려도 원본은 한 번만).
  function notesParas(e, old) {
    var t = textN(e.notes, LIM.notes);
    if (e.keepOriginal === true && old.trim()) {
      var k = old.lastIndexOf(KEEP_SEP), orig = (k >= 0 ? old.slice(k + KEEP_SEP.length) : old).replace(/^\s*\n/, '').replace(/\s+$/, '');
      if (orig) t = (t ? t + '\n' : '') + KEEP_SEP + '\n' + orig;
    }
    return t ? t.split('\n') : [''];
  }
  function editNotes(st, n, sl, e, skip, get) {
    var paras = notesParas(e, '');
    return get(sl.rels).then(function (relsXml) {
      var rels = parseRels(relsXml), np = notesPathOf(sl.path, rels);
      if (np && (np in st.mod || st.zip.file(np))) {
        return get(np).then(function (xml) {
          var body = notesBodyOf(xml);
          if (!body) return skip(e, '노트 쪽에 노트 글 상자가 없습니다 (no notes page)');
          paras = notesParas(e, parasOf(xml, body.txBody).join('\n'));
          var r = rebuildParas(xml, body.txBody, paras);
          if (!r) return skip(e, '이미 같은 노트라 바꾸지 않았습니다');
          st.mod[np] = spliceAll(xml, [r.op]);
          st.changed.push({ slide: n, shape: 'notes', field: 'notes', before: r.before, after: r.after }); st.applied++;
        });
      }
      if (paras.join('') === '') return skip(e, '빈 노트라 만들지 않았습니다');
      if (!st.pres.notesMaster || !relsXml || !/<\/Relationships>\s*$/.test(relsXml)) return skip(e, '이 파일에는 발표자 노트 틀이 없어 노트를 만들 수 없습니다 (no notes page)');
      return get('[Content_Types].xml').then(function (ct) {
        if (!/<\/Types>\s*$/.test(ct)) return skip(e, '파일 형식 목록을 읽지 못해 노트를 만들 수 없습니다 (no notes page)');
        var k = 1;
        while (st.zip.file('ppt/notesSlides/notesSlide' + k + '.xml') || ('ppt/notesSlides/notesSlide' + k + '.xml') in st.mod) k++;
        var path = 'ppt/notesSlides/notesSlide' + k + '.xml', rp = 'ppt/notesSlides/_rels/notesSlide' + k + '.xml.rels';
        var maxId = 0;
        rels.forEach(function (x) { var m = /^rId(\d+)$/.exec(x.id); if (m) maxId = Math.max(maxId, Number(m[1])); });
        st.mod[path] = notesXml(paras);
        st.mod[rp] = XML_DECL + '<Relationships xmlns="' + NS.pkg + '"><Relationship Id="rId1" Type="' + REL_T + 'notesMaster" Target="' + xmlEsc(relTarget(path, st.pres.notesMaster)) + '"/>'
          + '<Relationship Id="rId2" Type="' + REL_T + 'slide" Target="' + xmlEsc(relTarget(path, sl.path)) + '"/></Relationships>';
        st.mod[sl.rels] = relsXml.replace(/<\/Relationships>\s*$/, function () {
          return '<Relationship Id="rId' + (maxId + 1) + '" Type="' + REL_T + 'notesSlide" Target="' + xmlEsc(relTarget(sl.path, path)) + '"/></Relationships>';
        });
        st.mod['[Content_Types].xml'] = ct.replace(/<\/Types>\s*$/, function () { return '<Override PartName="/' + path + '" ContentType="' + CT_NOTES + '"/></Types>'; });
        st.changed.push({ slide: n, shape: 'notes', field: 'notes', before: '', after: paras.join('\n') }); st.applied++;
      });
    });
  }

  function isMedia(name) { return /^ppt\/media\//i.test(name) || /\.(png|jpe?g|gif|bmp|tiff?|webp|emf|wmf|mp4|m4v|mov|wmv|avi|mp3|m4a|wav|wma|fntdata|ttf|otf|odttf)$/i.test(name); }
  // 새 zip: 원래 순서대로, 손대지 않은 항목은 바이트 그대로(그림·소리는 STORE, 나머지는 DEFLATE)
  function writeZip(st) {
    var JSZip = getJSZip(), out = new JSZip(), zip = st.zip, names = Object.keys(zip.files);
    return Promise.all(names.map(function (name) {
      var f = zip.files[name];
      if (f.dir || name in st.mod) return null;
      return f.async('uint8array');
    })).then(function (datas) {
      names.forEach(function (name, i) {
        var f = zip.files[name];
        if (f.dir) { out.file(name, null, { dir: true, date: f.date }); return; }
        var data = name in st.mod ? st.mod[name] : datas[i];
        out.file(name, data, { compression: isMedia(name) ? 'STORE' : 'DEFLATE', date: f.date, createFolders: false, binary: typeof data !== 'string' });
      });
      Object.keys(st.mod).forEach(function (name) { if (!zip.files[name]) out.file(name, st.mod[name], { compression: 'DEFLATE', createFolders: false }); });
      return out.generateAsync({ type: 'blob', mimeType: PPTX_MIME, compression: 'DEFLATE', compressionOptions: { level: 6 } });
    });
  }

  /* ------------------------------------------------------------------ */
  /* 내보내기                                                             */
  /* ------------------------------------------------------------------ */

  var PptKit = {
    VERSION: VERSION,
    THEMES: THEMES,
    LAYOUTS: LAYOUTS.slice(),
    parseModelJson: parseModelJson,
    validateDeck: validateDeck,
    previewHtml: previewHtml,
    buildDeck: buildDeck,
    inspect: inspect,
    outline: outline,
    applyEdits: applyEdits,
    _internal: {
      layoutSlide: layoutSlide, normalizeDeck: normalizeDeck, measure: measure, toParas: toParas,
      setMorphTransition: setMorphTransition, parseXml: parseXml, readPres: readPres, walkShapes: walkShapes,
      MORPH_XML: MORPH_XML, LIM: LIM, PALETTES: PALETTES, countAccent: countAccent
    }
  };
  return PptKit;
});
