// /api/live: the workshop participation board, kept in Cloudflare D1 (binding LIVE_DB).
// What it holds, per room: one-line answers from phones (kind worry | take), each professor's score for the AI-written
// answer in mission 1 (kind score, one row per browser), and the apps from mission 2 (kind app, with a public link).
//
// POST JSON (all ops need the class code, same AI_CLASS_CODE as /api/ai):
//   { code, room, op:'check' }                                      -> { ok:true }
//   { code, room, op:'post',  kind:'worry'|'take', text }           -> { ok:true, id }
//   { code, room, op:'score', client, num:0..100, type, cmp }       -> { ok:true }   (one row per client, re-sending updates it)
//   { code, room, op:'app',   title, dept, name, desc, html[, token, edit] } -> { ok:true, token, edit }  (token+edit updates)
//   { code, room, op:'appdel', token, edit }                        -> { ok:true }
//   { code, room, op:'feed' }                                       -> { ok:true, texts, scores, apps, now }
// GET ?app=<token> (no code: the link a professor hands to students) -> { ok:true, title, dept, name, desc, html, updated }
// Errors: JSON { error:<Korean>, code }.
import { isAllowedOrigin, jsonResponse, errorResponse, preflightResponse } from '../_lib/cors.js';
import { readJsonBody, parseCodes, matchCode, clientKey, createStore, badCodeGate, readLimits, isKillSwitchOff } from '../_lib/limits.js';

export const LIVE = {
  maxBody: 600 * 1024,
  maxText: 300,
  maxTitle: 60,
  maxDept: 40,
  maxName: 30,
  maxDesc: 200,
  maxType: 30,
  maxHtml: 400_000,
  maxTextsPerRoom: 800,
  maxAppsPerRoom: 300,
  maxScoresPerRoom: 400,
  feedLimit: 500,
};

const MSG = {
  method: '허용되지 않은 요청 방식입니다.',
  origin: '허용되지 않은 사이트에서 보낸 요청입니다.',
  noDb: '참여판 저장소가 아직 연결되지 않았습니다. 강사에게 알려 주세요.',
  disabled: '참여판이 지금 꺼져 있습니다. 강사에게 문의해 주세요.',
  noCode: '수업 코드를 입력해 주세요.',
  badCode: '수업 코드가 맞지 않습니다. 강사가 알려 준 코드를 확인해 주세요.',
  bad: '요청 형식이 올바르지 않습니다.',
  empty: '내용을 적어 주세요.',
  full: '참여판이 가득 찼습니다. 강사에게 알려 주세요.',
  noApp: '앱을 찾지 못했습니다. 링크를 다시 확인해 주세요.',
  noEdit: '이 브라우저에서 올린 앱만 바꾸거나 내릴 수 있습니다.',
  bigHtml: '앱 코드가 너무 큽니다(40만 자 초과). 사진이나 긴 자료를 빼고 다시 올려 주세요.',
  notHtml: '올릴 코드가 HTML이 아닙니다. 실행 칸에서 돌아가는 코드를 올려 주세요.',
  internal: '참여판 요청을 처리하지 못했습니다. 잠시 후 다시 시도해 주세요.',
};

const SCHEMA = [
  `CREATE TABLE IF NOT EXISTS posts (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    room TEXT NOT NULL,
    kind TEXT NOT NULL,
    client TEXT NOT NULL DEFAULT '',
    token TEXT,
    edit TEXT,
    title TEXT NOT NULL DEFAULT '',
    body TEXT NOT NULL DEFAULT '',
    num REAL,
    meta TEXT NOT NULL DEFAULT '',
    html TEXT,
    created INTEGER NOT NULL,
    updated INTEGER NOT NULL
  )`,
  'CREATE INDEX IF NOT EXISTS posts_room ON posts(room, kind, id)',
  'CREATE UNIQUE INDEX IF NOT EXISTS posts_token ON posts(token)',
];
export const SCHEMA_SQL = SCHEMA.map((s) => s.replace(/\s+/g, ' ') + ';').join('\n');

let schemaReady = null;
function ensureSchema(db) {
  if (!schemaReady) schemaReady = db.batch(SCHEMA.map((s) => db.prepare(s))).catch((e) => { schemaReady = null; throw e; });
  return schemaReady;
}

const KINDS_TEXT = new Set(['worry', 'take']);
const CMP = new Set(['ai', 'same', 'student']);

// ---------- small validators (exported for tests) ----------
export function cleanText(v, max) {
  if (typeof v !== 'string') return '';
  return v.replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F\u007F]/g, '').replace(/\r\n?/g, '\n').trim().slice(0, max);
}
export function cleanRoom(v) {
  const r = typeof v === 'string' ? v.trim().toLowerCase() : '';
  return /^[a-z0-9-]{1,24}$/.test(r) ? r : 'main';
}
export function cleanClient(v) {
  return typeof v === 'string' && /^[a-z0-9]{8,40}$/i.test(v) ? v : '';
}
export function cleanToken(v) {
  return typeof v === 'string' && /^[a-z0-9]{6,32}$/i.test(v) ? v : '';
}
export function looksLikeHtml(s) {
  return /<(!doctype|html|body|div|main|section|script|style|button|canvas|h1)\b/i.test(s);
}
function randomId(n) {
  const a = 'abcdefghijkmnpqrstuvwxyz23456789';
  const bytes = crypto.getRandomValues(new Uint8Array(n));
  let s = '';
  for (const b of bytes) s += a[b % a.length];
  return s;
}

export async function onRequest(context) {
  const origin = context.request.headers.get('Origin');
  try {
    return await handle(context, origin);
  } catch (e) {
    console.error('[live] internal error', e && (e.message || e.name));
    return errorResponse(503, 'internal', MSG.internal, origin);
  }
}

async function handle(context, origin) {
  const { request, env } = context;
  if (request.method === 'OPTIONS') return preflightResponse(origin);
  if (origin != null && !isAllowedOrigin(origin)) return errorResponse(403, 'bad_request', MSG.origin, origin);
  const db = env.LIVE_DB;
  if (!db) return errorResponse(503, 'no_db', MSG.noDb, origin);
  if (String(env.LIVE_ENABLED ?? '').trim() === '0') return errorResponse(503, 'disabled', MSG.disabled, origin);

  if (request.method === 'GET') {
    const token = cleanToken(new URL(request.url).searchParams.get('app'));
    if (!token) return errorResponse(400, 'bad_request', MSG.bad, origin);
    await ensureSchema(db);
    const row = await db.prepare("SELECT title, body, meta, html, updated FROM posts WHERE token = ? AND kind = 'app'").bind(token).first();
    if (!row) return errorResponse(404, 'no_app', MSG.noApp, origin);
    const meta = safeJson(row.meta);
    return jsonResponse(200, { ok: true, title: row.title, desc: row.body, dept: meta.dept || '', name: meta.name || '', html: row.html, updated: row.updated }, origin);
  }
  if (request.method !== 'POST') {
    const r = errorResponse(405, 'bad_request', MSG.method, origin);
    r.headers.set('Allow', 'GET, POST, OPTIONS');
    return r;
  }

  const codesRaw = String(env.AI_CLASS_CODE ?? '');
  if (isKillSwitchOff({ AI_ENABLED: env.LIVE_ENABLED }) || parseCodes(codesRaw).length === 0) return errorResponse(503, 'disabled', MSG.disabled, origin);

  const parsed = await readJsonBody(request, LIVE.maxBody);
  if (!parsed.ok) return errorResponse(parsed.status, parsed.code, parsed.message, origin);
  const body = parsed.body;
  if (typeof body.code !== 'string' || !body.code.trim()) return errorResponse(401, 'bad_code', MSG.noCode, origin);
  if (body.code.length > 64) return errorResponse(401, 'bad_code', MSG.badCode, origin);

  const store = createStore(env, typeof context.waitUntil === 'function' ? context.waitUntil.bind(context) : undefined);
  const ip = clientKey(request);
  const gate = await badCodeGate(store, ip, Date.now(), readLimits(env).badCode10m);
  if (gate.locked) return errorResponse(gate.locked.status, gate.locked.code, gate.locked.message, origin, gate.locked.retryAfter);
  if (!(await matchCode(body.code, codesRaw))) {
    gate.recordWrong();
    return errorResponse(401, 'bad_code', MSG.badCode, origin);
  }

  const room = cleanRoom(body.room);
  const op = typeof body.op === 'string' ? body.op : '';
  if (op === 'check') return jsonResponse(200, { ok: true }, origin);
  await ensureSchema(db);
  const now = Date.now();

  if (op === 'post') {
    const kind = KINDS_TEXT.has(body.kind) ? body.kind : '';
    const text = cleanText(body.text, LIVE.maxText);
    if (!kind) return errorResponse(400, 'bad_request', MSG.bad, origin);
    if (!text) return errorResponse(400, 'empty', MSG.empty, origin);
    const n = await count(db, room, "kind IN ('worry','take')");
    if (n >= LIVE.maxTextsPerRoom) return errorResponse(429, 'full', MSG.full, origin);
    const r = await db.prepare('INSERT INTO posts (room, kind, client, body, created, updated) VALUES (?, ?, ?, ?, ?, ?)')
      .bind(room, kind, cleanClient(body.client), text, now, now).run();
    return jsonResponse(200, { ok: true, id: r.meta && r.meta.last_row_id }, origin);
  }

  if (op === 'score') {
    const client = cleanClient(body.client);
    const num = Number(body.num);
    if (!client || !Number.isFinite(num) || num < 0 || num > 100) return errorResponse(400, 'bad_request', MSG.bad, origin);
    const meta = JSON.stringify({ type: cleanText(body.type, LIVE.maxType), cmp: CMP.has(body.cmp) ? body.cmp : '' });
    const prev = await db.prepare("SELECT id FROM posts WHERE room = ? AND kind = 'score' AND client = ?").bind(room, client).first();
    if (prev) {
      await db.prepare('UPDATE posts SET num = ?, meta = ?, updated = ? WHERE id = ?').bind(Math.round(num), meta, now, prev.id).run();
    } else {
      if ((await count(db, room, "kind = 'score'")) >= LIVE.maxScoresPerRoom) return errorResponse(429, 'full', MSG.full, origin);
      await db.prepare("INSERT INTO posts (room, kind, client, num, meta, created, updated) VALUES (?, 'score', ?, ?, ?, ?, ?)")
        .bind(room, client, Math.round(num), meta, now, now).run();
    }
    return jsonResponse(200, { ok: true }, origin);
  }

  if (op === 'app') {
    const html = typeof body.html === 'string' ? body.html.trim() : '';
    if (html.length > LIVE.maxHtml) return errorResponse(413, 'too_long', MSG.bigHtml, origin);
    if (!html || !looksLikeHtml(html)) return errorResponse(400, 'not_html', MSG.notHtml, origin);
    const title = cleanText(body.title, LIVE.maxTitle) || '수업 앱';
    const desc = cleanText(body.desc, LIVE.maxDesc);
    const meta = JSON.stringify({ dept: cleanText(body.dept, LIVE.maxDept), name: cleanText(body.name, LIVE.maxName) });
    const token = cleanToken(body.token), edit = cleanToken(body.edit);
    if (token && edit) {
      const r = await db.prepare("UPDATE posts SET title = ?, body = ?, meta = ?, html = ?, updated = ? WHERE token = ? AND edit = ? AND kind = 'app'")
        .bind(title, desc, meta, html, now, token, edit).run();
      if (r.meta && r.meta.changes) return jsonResponse(200, { ok: true, token, edit, updated: true }, origin);
      // the app was taken down or the key is stale: fall through and post it as a new one
    }
    if ((await count(db, room, "kind = 'app'")) >= LIVE.maxAppsPerRoom) return errorResponse(429, 'full', MSG.full, origin);
    const t = randomId(10), e = randomId(24);
    await db.prepare("INSERT INTO posts (room, kind, client, token, edit, title, body, meta, html, created, updated) VALUES (?, 'app', ?, ?, ?, ?, ?, ?, ?, ?, ?)")
      .bind(room, cleanClient(body.client), t, e, title, desc, meta, html, now, now).run();
    return jsonResponse(200, { ok: true, token: t, edit: e }, origin);
  }

  if (op === 'appdel') {
    const token = cleanToken(body.token), edit = cleanToken(body.edit);
    if (!token || !edit) return errorResponse(400, 'bad_request', MSG.bad, origin);
    const r = await db.prepare("DELETE FROM posts WHERE token = ? AND edit = ? AND kind = 'app'").bind(token, edit).run();
    if (!(r.meta && r.meta.changes)) return errorResponse(403, 'no_edit', MSG.noEdit, origin);
    return jsonResponse(200, { ok: true }, origin);
  }

  if (op === 'feed') {
    const [texts, scores, apps] = await db.batch([
      db.prepare("SELECT id, kind, body, created FROM posts WHERE room = ? AND kind IN ('worry','take') ORDER BY id DESC LIMIT ?").bind(room, LIVE.feedLimit),
      db.prepare("SELECT id, num, meta, updated FROM posts WHERE room = ? AND kind = 'score' ORDER BY id DESC LIMIT ?").bind(room, LIVE.feedLimit),
      db.prepare("SELECT id, token, title, body, meta, created, updated, length(html) AS size FROM posts WHERE room = ? AND kind = 'app' ORDER BY id DESC LIMIT ?").bind(room, LIVE.feedLimit),
    ]);
    return jsonResponse(200, {
      ok: true,
      now,
      texts: texts.results.map((r) => ({ id: r.id, kind: r.kind, text: r.body, at: r.created })),
      scores: scores.results.map((r) => { const m = safeJson(r.meta); return { id: r.id, num: r.num, type: m.type || '', cmp: m.cmp || '', at: r.updated }; }),
      apps: apps.results.map((r) => { const m = safeJson(r.meta); return { id: r.id, token: r.token, title: r.title, desc: r.body, dept: m.dept || '', name: m.name || '', at: r.created, updated: r.updated, size: r.size }; }),
    }, origin);
  }

  return errorResponse(400, 'bad_request', MSG.bad, origin);
}

async function count(db, room, where) {
  const r = await db.prepare(`SELECT COUNT(*) AS n FROM posts WHERE room = ? AND ${where}`).bind(room).first();
  return (r && r.n) || 0;
}

function safeJson(s) {
  try { const o = JSON.parse(s || '{}'); return o && typeof o === 'object' ? o : {}; } catch { return {}; }
}
