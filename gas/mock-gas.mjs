// A small in-memory stand-in for the Google Apps Script services the recipes use, so the .gs sources run under plain node.
// The fake class-lab server validates every request with the REAL validators from functions/ (jev.js, limits.js):
// a recipe that sends a question shape or an attachment the server would reject fails here, not in the classroom.
import vm from 'node:vm';
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { validateJev } from '../functions/api/jev.js';
import { validateMessages, validateAttachments } from '../functions/_lib/limits.js';

const here = path.dirname(fileURLToPath(import.meta.url));
export const readGs = (name) => fs.readFileSync(path.join(here, name), 'utf8');

const colNum = (s) => [...s].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0);
function a1(ref) {
  const m = /^([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?$/.exec(ref);
  if (!m) throw new Error('bad A1 ' + ref);
  const c1 = colNum(m[1]), r1 = Number(m[2]);
  return [r1, c1, m[4] ? Number(m[4]) - r1 + 1 : 1, m[3] ? colNum(m[3]) - c1 + 1 : 1];
}

class Range {
  constructor(sh, r, c, nr, nc) { Object.assign(this, { sh, r, c, nr, nc }); }
  getValues() {
    return Array.from({ length: this.nr }, (_, i) => Array.from({ length: this.nc }, (_, j) => this.sh.cell(this.r + i, this.c + j)));
  }
  getValue() { return this.sh.cell(this.r, this.c); }
  setValues(v) { v.forEach((row, i) => row.forEach((x, j) => this.sh.put(this.r + i, this.c + j, x))); return this; }
  setValue(x) { this.sh.put(this.r, this.c, x); return this; }
  setFontWeight() { return this; }
  setNumberFormat(f) { for (let i = 0; i < this.nr; i++) for (let j = 0; j < this.nc; j++) this.sh.formats.set(`${this.r + i},${this.c + j}`, f); return this; }
}

class Sheet {
  constructor(name) { this.name = name; this.rows = []; this.formats = new Map(); }
  cell(r, c) { const v = (this.rows[r - 1] || [])[c - 1]; return v === undefined ? '' : v; }
  put(r, c, x) {
    while (this.rows.length < r) this.rows.push([]);
    const row = this.rows[r - 1];
    while (row.length < c) row.push('');
    // a text-formatted cell keeps its string; otherwise sheets turn a yyyy-mm-dd string into a date, like the real thing
    // a leading ' keeps the rest as plain text; any other text starting with = + - @ is a formula, and reads back as its result
    if (typeof x === 'string' && x.startsWith("'")) x = x.slice(1);
    else if (typeof x === 'string' && /^[=+\-@]/.test(x) && !/^[+-]?\d+(\.\d+)?$/.test(x)) x = '#FORMULA(' + x + ')';
    row[c - 1] = !this.formats.has(`${r},${c}`) && typeof x === 'string' && /^\d{4}-\d{2}-\d{2}$/.test(x) ? new Date(x + 'T00:00:00+09:00') : x;
  }
  getRange(r, c, nr = 1, nc = 1) {
    if (typeof r === 'string') [r, c, nr, nc] = a1(r);
    return new Range(this, r, c, nr, nc);
  }
  getLastRow() {
    for (let i = this.rows.length; i > 0; i--) if (this.rows[i - 1].some((x) => x !== '')) return i;
    return 0;
  }
  appendRow(arr) { const r = this.getLastRow() + 1; arr.forEach((x, j) => this.put(r, j + 1, x)); }
  setColumnWidth() {}
  setFrozenRows() {}
}

class Spreadsheet {
  constructor(id) { this.id = id; this.sheets = new Map(); this.nextSheetId = 100; }
  getId() { return this.id; }
  getSheetByName(n) { return this.sheets.get(n) || null; }
  insertSheet(n) {
    // like Sheets: a name that is already taken is an error, and every sheet gets its own id
    if (this.sheets.has(n)) throw new Error('A sheet with the name "' + n + '" already exists.');
    const s = new Sheet(n), id = this.nextSheetId++;
    s.getSheetId = () => id;
    s.setName = (nn) => { if (this.sheets.has(nn) && this.sheets.get(nn) !== s) throw new Error('A sheet with the name "' + nn + '" already exists.'); this.sheets.delete(s.name); s.name = nn; this.sheets.set(nn, s); return s; };
    this.sheets.set(n, s); return s;
  }
}

const items = () => { const it = {}; for (const m of ['setTitle', 'setBounds', 'setLabels', 'setRequired']) it[m] = () => it; return it; };

export function createEnv({ jev, ai, fetchExtra } = {}) {
  const log = { prompts: [], logged: [], alerts: [], mails: [], triggers: [], requests: [], menus: [], sleeps: 0 };
  const active = new Spreadsheet('SS1');
  const props = new Map([['CLASS_CODE', 'TESTCODE']]);
  const cache = new Map();
  const folders = new Map();
  const docs = new Map();
  let seq = 0;
  const ui = {
    ButtonSet: { OK_CANCEL: 1 }, Button: { OK: 'OK' },
    alert: (m) => log.alerts.push(String(m)),
    // tests queue answers in log.prompts; otherwise the class-code prompt gets the test code
    prompt: () => { const a = log.prompts.length ? log.prompts.shift() : ' TESTCODE '; return { getSelectedButton: () => (a === null ? 'CANCEL' : 'OK'), getResponseText: () => a || '' }; },
    createMenu: (name) => { const m = { name, items: [], addItem(l, fn) { m.items.push([l, fn]); return m; }, addToUi() { log.menus.push(m); return m; } }; return m; },
  };
  const response = (status, text) => ({ getResponseCode: () => status, getContentText: () => text, getBlob: () => ({ getContentType: () => 'image/jpeg', getBytes: () => Buffer.alloc(1000, 1) }) });

  const server = (url, body) => {
    log.requests.push({ url, body });
    const code = body.code;
    if (code !== 'TESTCODE') return [401, { error: '수업 코드가 맞지 않습니다.', code: 'bad_code' }];
    if (url.endsWith('/api/jev')) {
      if (body.check === true) return [200, { ok: true }];
      if (log.jevFail) return [log.jevFail, { error: 'Jev down', code: 'upstream', retryAfter: 1 }];
      const v = validateJev(body);
      if (!v.ok) return [v.status, { error: v.message, code: v.code }];
      const answers = jev ? jev(v.state, v.questions) : defaultAnswers(v.questions);
      return [200, { ok: true, model: 'jev-latest', answers }];
    }
    if (url.endsWith('/api/ai')) {
      const m = validateMessages(body.messages);
      if (!m.ok) return [m.status, { error: m.message, code: m.code }];
      const att = validateAttachments(body.files, body.images, body.messages);
      if (!att.ok) return [att.status, { error: att.message, code: att.code }];
      if (body.ex !== 'gas-run') return [400, { error: 'ex must be gas-run for scripts', code: 'bad_request' }];
      const text = ai ? ai({ messages: m.messages, files: att.files, images: att.images }) : 'OK';
      const events = [{ meta: { model: 'x' } }, ...text.match(/[\s\S]{1,40}/g).map((d) => ({ delta: d })), { done: true }];
      return [200, events.map((e) => 'data: ' + JSON.stringify(e)).join('\n\n') + '\n\n'];
    }
    return [404, { error: 'no route' }];
  };

  const mailApp = { sendEmail: (to, subject, body, opts) => log.mails.push({ to, subject, body, html: opts && opts.htmlBody }), getRemainingDailyQuota: () => 100 };
  const fakeFolder = (name) => {
    const id = 'F' + ++seq;
    const f = { id, name, files: [], getId: () => id, getUrl: () => 'https://drive.google.com/drive/folders/' + id, getFiles: () => { let i = 0; return { hasNext: () => i < f.files.length, next: () => f.files[i++] }; } };
    folders.set(id, f);
    return f;
  };
  const ctx = {
    Logger: { log: (m) => { log.logged.push(String(m)); } },
    console, Date, Math, JSON, Object, Array, String, Number, Buffer, URLSearchParams,
    SpreadsheetApp: { getActive: () => active, getActiveSpreadsheet: () => active, openById: (id) => { if (id !== active.id) throw new Error('no such spreadsheet'); return active; }, getUi: () => ui, flush: () => {} },
    PropertiesService: { getScriptProperties: () => ({ getProperty: (k) => (props.has(k) ? props.get(k) : null), setProperty: (k, v) => { props.set(k, String(v)); } }) },
    CacheService: { getScriptCache: () => ({ get: (k) => (cache.has(k) ? cache.get(k) : null), put: (k, v) => { cache.set(k, v); }, remove: (k) => { cache.delete(k); } }) },
    LockService: { getScriptLock: () => ({ waitLock() {}, tryLock: () => true, releaseLock() {} }) },
    Utilities: {
      sleep: () => { log.sleeps++; },
      getUuid: () => crypto.randomUUID(),
      base64Encode: (b) => Buffer.from(b).toString('base64'),
      formatDate: (d, tz, fmt) => {
        const k = new Date(d.getTime() + 9 * 3600e3), p = (n) => String(n).padStart(2, '0');
        const map = { yyyy: k.getUTCFullYear(), MM: p(k.getUTCMonth() + 1), dd: p(k.getUTCDate()), HH: p(k.getUTCHours()), mm: p(k.getUTCMinutes()), ss: p(k.getUTCSeconds()) };
        return fmt.replace(/yyyy|MM|dd|HH|mm|ss/g, (t) => map[t]);
      },
    },
    UrlFetchApp: {
      fetch: (url, opts = {}) => {
        if (fetchExtra && !/\/api\/(ai|jev)$/.test(url)) return fetchExtra(url, opts);
        const [status, payload] = server(url, JSON.parse(opts.payload));
        return response(status, typeof payload === 'string' ? payload : JSON.stringify(payload));
      },
    },
    HtmlService: { XFrameOptionsMode: { ALLOWALL: 'ALLOWALL' }, createHtmlOutput: (html) => { const o = { html, title: '', xframe: '', setTitle(t) { o.title = t; return o; }, addMetaTag() { return o; }, setXFrameOptionsMode(m) { o.xframe = m; return o; }, getContent: () => html }; return o; } },
    ScriptApp: {
      WeekDay: { MONDAY: 'MONDAY' },
      // the real getUrl() can return /dev or a deployment id that does not exist; recipes must not use it
      getService: () => ({ getUrl: () => { throw new Error('getUrl() is unreliable; use webUrl_()'); } }),
      getOAuthToken: () => 'tok',
      getProjectTriggers: () => log.triggers,
      deleteTrigger: (t) => { log.triggers.splice(log.triggers.indexOf(t), 1); },
      newTrigger: (fn) => {
        const t = { fn, getHandlerFunction: () => fn, spec: [] };
        const chain = new Proxy(t, { get: (o, k) => (k in o ? o[k] : k === 'create' ? () => { log.triggers.push(t); return t; } : (...a) => { t.spec.push([k, ...a]); return chain; }) });
        return chain;
      },
    },
    MailApp: mailApp,
    Session: { getEffectiveUser: () => ({ getEmail: () => 'prof@example.com' }), getScriptTimeZone: () => 'Asia/Seoul' },
    DriveApp: {
      createFolder: (n) => fakeFolder(n),
      getFolderById: (id) => folders.get(id),
      getFileById: (id) => ({ moveTo: (folder) => { folder.files.push({ docId: id }); } }),
    },
    DocumentApp: {
      ParagraphHeading: { HEADING1: 'H1', HEADING2: 'H2', NORMAL: 'P' },
      create: (name) => {
        const id = 'D' + ++seq, lines = [], images = [];
        const para = (t, kind) => { const l = { t, kind }; lines.push(l); const api = new Proxy({}, { get: (_, k) => (k === 'setHeading' ? (h) => { l.kind = h; return api; } : () => api) }); return api; };
        const doc = { id, name, lines, images, getId: () => id, getUrl: () => 'https://docs.google.com/document/d/' + id, saveAndClose() {},
          getBody: () => ({ clear: () => { lines.length = 0; images.length = 0; }, appendImage: (b) => { images.push(b); return { getWidth: () => 800, getHeight: () => 600, setWidth() { return this; }, setHeight() { return this; } }; }, appendParagraph: (t) => para(t, 'P'), appendListItem: (t) => para(t, 'LI'), getText: () => lines.map((l) => l.t).join('\n') }) };
        docs.set(id, doc);
        return doc;
      },
      openById: (id) => docs.get(id),
    },
    FormApp: {
      DestinationType: { SPREADSHEET: 'SPREADSHEET' },
      create: (title) => {
        const form = { id: 'FORM' + ++seq, title, questions: [] };
        const add = (type) => () => { const it = items(); const t = it.setTitle; it.setTitle = (x) => { form.questions.push([type, x]); return it; }; return it; };
        Object.assign(form, { setDescription: () => form, addScaleItem: add('scale'), addParagraphTextItem: add('paragraph'), setDestination: () => form, getId: () => form.id, getPublishedUrl: () => 'https://forms.gle/' + form.id });
        return form;
      },
      openById: () => ({ getPublishedUrl: () => 'https://forms.gle/x' }),
    },
  };
  ctx.globalThis = ctx;
  vm.createContext(ctx);
  return { ctx, log, active, props, cache, folders, docs, run: (src, file = 'recipe.gs') => vm.runInContext(src, ctx, { filename: file }) };
}

// Runs the browser script of a served page against a stub DOM. google.script.run calls are executed synchronously on the
// server functions of `env`, so the page logic is exercised end to end (a call to a function that does not exist throws).
export function runPage(env, html, { values = {} } = {}) {
  html = (env.ctx.TOAST_JS || '') + html; // page_ puts the shared toast in front of every page
  // every inline <script> block, in order (src-only scripts are skipped, as they load third-party code)
  const script = [...html.matchAll(/<script>([\s\S]*?)<\/script>/g)].map((m) => m[1]).join(';\n');
  const els = new Map(), calls = [], timers = [], store = new Map();
  let tid = 0;
  const el = (sel) => {
    if (!els.has(sel)) {
      const e = { sel, innerHTML: '', value: values[sel] ?? '', hidden: false, disabled: false, className: '', style: {}, children: [], onclick: null, oninput: null };
      let text = '';
      Object.defineProperty(e, 'textContent', { get: () => text, set: (v) => { text = String(v); e.children.length = 0; } }); // like the DOM: assigning text drops the children
      e.appendChild = (c) => { e.children.push(c); return c; };
      e.scrollIntoView = () => {};
      els.set(sel, e);
    }
    return els.get(sel);
  };
  const clone = (r) => (r && typeof r === 'object' ? JSON.parse(JSON.stringify(r)) : r);
  const run = (ok, bad) => new Proxy({}, {
    get: (_, name) => {
      if (name === 'withSuccessHandler') return (f) => run(f, bad);
      if (name === 'withFailureHandler') return (f) => run(ok, f);
      return (...args) => {
        if (typeof env.ctx[name] !== 'function') throw new Error('page calls a server function that does not exist: ' + name);
        calls.push(name);
        let r;
        try { r = clone(env.ctx[name](...args)); } catch (e) { if (bad) { bad(e); return; } throw e; }
        if (ok) ok(r);
      };
    },
  });
  const sandbox = {
    google: { script: { run: run() } },
    document: { querySelector: el, createTextNode: (t) => ({ textContent: String(t), children: [] }), createElement: () => {
      const e = { className: '', style: { setProperty() {} }, children: [], scrollIntoView() {} };
      let own = '';
      // like the DOM: reading joins the children's text, assigning drops the children
      Object.defineProperty(e, 'textContent', { get: () => own + e.children.map((c) => c.textContent).join(''), set: (v) => { own = String(v); e.children.length = 0; } });
      e.appendChild = (c) => { e.children.push(c); return c; };
      e.insertBefore = (c, ref) => { const k = ref ? e.children.indexOf(ref) : -1; if (k < 0) e.children.push(c); else e.children.splice(k, 0, c); return c; };
      e.removeChild = (c) => { const k = e.children.indexOf(c); if (k >= 0) e.children.splice(k, 1); return c; };
      Object.defineProperty(e, 'firstChild', { get: () => e.children[0] || null });
      return e;
    } },
    localStorage: { getItem: (k) => store.get(k) ?? null, setItem: (k, v) => store.set(k, String(v)) },
    location: { search: '' }, Math, console,
    setTimeout: (f, ms) => { const t = [f, ms]; t.id = ++tid; timers.push(t); return t.id; },
    clearTimeout: (id) => { const k = timers.findIndex((t) => t.id === id); if (k >= 0) timers.splice(k, 1); },
  };
  vm.runInNewContext(script, sandbox, { filename: 'page.js' });
  return {
    el, calls, timers,
    set: (sel, v) => { el(sel).value = v; },
    click: (sel) => { const e = el(sel); if (typeof e.onclick !== 'function') throw new Error('no click handler on ' + sel); e.onclick(); },
    texts: (sel) => [el(sel).textContent, ...el(sel).children.map((c) => c.textContent)].filter(Boolean),
  };
}

export function defaultAnswers(questions) {
  const out = {};
  for (const [name, q] of Object.entries(questions)) {
    if (q.type === 'noul') out[name] = { type: 'noul', p: 0.1 };
    else if (q.type === 'choice') { const k = Object.keys(q.criteria)[0]; out[name] = { type: 'choice', choice: k, confidence: 0.9, probs: { [k]: 0.9 } }; }
    else out[name] = { type: 'score', score: 2, level: 2, label: q.criteria[2] ?? q.criteria.at(-1), confidence: 0.9, probs: {} };
  }
  return out;
}

export const choice = (choice, confidence = 0.9) => ({ type: 'choice', choice, confidence, probs: { [choice]: confidence } });
export const noul = (p) => ({ type: 'noul', p });
export const score = (level, labels = []) => ({ type: 'score', score: level, level, label: labels[level] ?? null, confidence: 0.9, probs: {} });
