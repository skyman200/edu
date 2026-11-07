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
  constructor(id) { this.id = id; this.sheets = new Map(); }
  getId() { return this.id; }
  getSheetByName(n) { return this.sheets.get(n) || null; }
  insertSheet(n) { const s = new Sheet(n); this.sheets.set(n, s); return s; }
}

const items = () => { const it = {}; for (const m of ['setTitle', 'setBounds', 'setLabels', 'setRequired']) it[m] = () => it; return it; };

export function createEnv({ jev, ai, fetchExtra } = {}) {
  const log = { alerts: [], mails: [], triggers: [], requests: [], menus: [], sleeps: 0 };
  const active = new Spreadsheet('SS1');
  const props = new Map([['CLASS_CODE', 'TESTCODE']]);
  const cache = new Map();
  const folders = new Map();
  const docs = new Map();
  let seq = 0;
  const ui = {
    ButtonSet: { OK_CANCEL: 1 }, Button: { OK: 'OK' },
    alert: (m) => log.alerts.push(String(m)),
    prompt: () => ({ getSelectedButton: () => 'OK', getResponseText: () => ' TESTCODE ' }),
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

  const mailApp = { sendEmail: (to, subject, body) => log.mails.push({ to, subject, body }), getRemainingDailyQuota: () => 100 };
  const fakeFolder = (name) => {
    const id = 'F' + ++seq;
    const f = { id, name, files: [], getId: () => id, getUrl: () => 'https://drive.google.com/drive/folders/' + id, getFiles: () => { let i = 0; return { hasNext: () => i < f.files.length, next: () => f.files[i++] }; } };
    folders.set(id, f);
    return f;
  };
  const ctx = {
    console, Date, Math, JSON, Object, Array, String, Number, Buffer, URLSearchParams,
    SpreadsheetApp: { getActive: () => active, getActiveSpreadsheet: () => active, openById: (id) => { if (id !== active.id) throw new Error('no such spreadsheet'); return active; }, getUi: () => ui },
    PropertiesService: { getScriptProperties: () => ({ getProperty: (k) => (props.has(k) ? props.get(k) : null), setProperty: (k, v) => { props.set(k, String(v)); } }) },
    CacheService: { getScriptCache: () => ({ get: (k) => (cache.has(k) ? cache.get(k) : null), put: (k, v) => { cache.set(k, v); } }) },
    LockService: { getScriptLock: () => ({ waitLock() {}, tryLock: () => true, releaseLock() {} }) },
    Utilities: {
      sleep: () => { log.sleeps++; },
      getUuid: () => crypto.randomUUID(),
      base64Encode: (b) => Buffer.from(b).toString('base64'),
      formatDate: (d, tz, fmt) => {
        const k = new Date(d.getTime() + 9 * 3600e3), p = (n) => String(n).padStart(2, '0');
        const map = { yyyy: k.getUTCFullYear(), MM: p(k.getUTCMonth() + 1), dd: p(k.getUTCDate()), HH: p(k.getUTCHours()), mm: p(k.getUTCMinutes()) };
        return fmt.replace(/yyyy|MM|dd|HH|mm/g, (t) => map[t]);
      },
    },
    UrlFetchApp: {
      fetch: (url, opts = {}) => {
        if (fetchExtra && !/\/api\/(ai|jev)$/.test(url)) return fetchExtra(url, opts);
        const [status, payload] = server(url, JSON.parse(opts.payload));
        return response(status, typeof payload === 'string' ? payload : JSON.stringify(payload));
      },
    },
    HtmlService: { createHtmlOutput: (html) => { const o = { html, title: '', setTitle(t) { o.title = t; return o; }, addMetaTag() { return o; }, getContent: () => html }; return o; } },
    ScriptApp: {
      WeekDay: { MONDAY: 'MONDAY' },
      getService: () => ({ getUrl: () => 'https://script.google.com/macros/s/X/exec' }),
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
      ParagraphHeading: { HEADING1: 'H1', HEADING2: 'H2' },
      create: (name) => {
        const id = 'D' + ++seq, lines = [];
        const para = (t, kind) => { const l = { t, kind }; lines.push(l); return { setHeading: (h) => { l.kind = h; return l; } }; };
        const doc = { id, name, lines, getId: () => id, getUrl: () => 'https://docs.google.com/document/d/' + id, saveAndClose() {},
          getBody: () => ({ clear: () => { lines.length = 0; }, appendParagraph: (t) => para(t, 'P'), appendListItem: (t) => para(t, 'LI'), getText: () => lines.map((l) => l.t).join('\n') }) };
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
  const script = /<script>([\s\S]*)<\/script>/.exec(html)[1];
  const els = new Map(), calls = [], timers = [], store = new Map();
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
    document: { querySelector: el, createElement: () => ({ textContent: '', className: '', style: {}, children: [], appendChild(c) { this.children.push(c); return c; }, scrollIntoView() {} }) },
    localStorage: { getItem: (k) => store.get(k) ?? null, setItem: (k, v) => store.set(k, String(v)) },
    location: { search: '' }, setTimeout: (f, ms) => { timers.push([f, ms]); return timers.length; }, Math, console,
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
