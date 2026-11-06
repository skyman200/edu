// Request validation, access-code check, and abuse limits for /api/ai.
// Counters live in KV (env.AI_KV) when bound and always in a per-isolate Map as well.
// KV is eventually consistent and rate limited per key, so every cap is coarse: KV errors
// never fail a request, they only leave the in-memory count in charge.

export const SIZE = {
  maxMessages: 24,
  maxContent: 12000,
  maxTotal: 40000, // messages only; attachments have their own caps below
  maxCode: 64,
  maxFiles: 3, // documents, as text the browser extracted (PDF, DOCX, PPTX, HWPX, TXT)
  maxFileChars: 100000,
  maxFilesTotal: 150000,
  maxName: 80, // longer attachment names are cut
  maxImages: 4, // data:image/(png|jpeg|webp);base64 URLs, downscaled in the browser
  maxImageUrl: 1_400_000,
  maxImagesTotal: 4_000_000,
  maxBodyBytes: 6 * 1024 * 1024, // 150,000 Korean chars are ~450 KB of UTF-8, images up to ~4 MB
};

// Defaults assume ~60 professors sharing one venue NAT address and one class code.
export const DEFAULT_LIMITS = {
  ip10m: 1500, // model requests per IP per 10 minutes (the whole room shares one venue IP, so this is a room-wide cap)
  inflight: 150, // concurrent answers per IP (room-wide)
  badCode10m: 1000, // wrong codes per IP per 10 minutes, then every request from that IP is refused (high: one NAT IP, typos happen)
  codeDay: 2500, // model requests per class code per KST day
  globalDay: 3000, // model requests in total per KST day
  globalMin: 100, // model requests in total per minute; the AI ONE gateway allows 120/min per member
};

const ENV_NAMES = {
  ip10m: 'AI_LIMIT_IP_10M',
  inflight: 'AI_LIMIT_INFLIGHT',
  badCode10m: 'AI_LIMIT_BADCODE_10M',
  codeDay: 'AI_LIMIT_CODE_DAY',
  globalDay: 'AI_LIMIT_GLOBAL_DAY',
  globalMin: 'AI_LIMIT_GLOBAL_MIN',
};

export const FLIGHT_TTL_MS = 320_000; // longer than the 300 s hard stream cap
const WINDOW_MIN = 10;
const WINDOW_TTL = (WINDOW_MIN + 1) * 60;
const DAY_TTL = 2 * 86400;
const KST_OFFSET_MS = 9 * 3600_000;
const MEM_MAX = 5000;

export function readLimits(env = {}) {
  const out = {};
  for (const [k, name] of Object.entries(ENV_NAMES)) {
    const n = Number(String(env[name] ?? '').trim());
    out[k] = Number.isInteger(n) && n > 0 ? n : DEFAULT_LIMITS[k];
  }
  return out;
}

export function isKillSwitchOff(env = {}) {
  const v = String(env.AI_ENABLED ?? '').trim().toLowerCase();
  return v === '0' || v === 'false' || v === 'off' || v === 'no';
}

// ---------- request body ----------

function fail(status, code, message) {
  return { ok: false, status, code, message };
}

export async function readJsonBody(request, maxBytes = SIZE.maxBodyBytes) {
  const tooBig = fail(413, 'too_long', '보낸 내용이 너무 큽니다. 내용을 줄여 다시 보내 주세요.');
  const len = Number(request.headers.get('Content-Length'));
  if (Number.isFinite(len) && len > maxBytes) return tooBig;
  if (!request.body) return fail(400, 'bad_request', '요청 내용이 비어 있습니다.');
  const reader = request.body.getReader();
  const chunks = [];
  let size = 0;
  for (;;) {
    const { value, done } = await reader.read();
    if (done) break;
    size += value.byteLength;
    if (size > maxBytes) {
      reader.cancel().catch(() => {});
      return tooBig;
    }
    chunks.push(value);
  }
  const bytes = new Uint8Array(size);
  let off = 0;
  for (const c of chunks) {
    bytes.set(c, off);
    off += c.byteLength;
  }
  try {
    const body = JSON.parse(new TextDecoder().decode(bytes));
    if (!body || typeof body !== 'object' || Array.isArray(body)) throw new Error('not an object');
    return { ok: true, body };
  } catch {
    return fail(400, 'bad_request', '요청 형식이 올바르지 않습니다.');
  }
}

// Accepts only {role:'user'|'assistant', content:string}. Everything else the client sends
// (model, system, max_tokens, tools...) is ignored by the caller. Empty assistant turns and
// leading assistant turns are dropped; the conversation must end with a user turn.
export function validateMessages(messages) {
  if (!Array.isArray(messages) || messages.length === 0) return fail(400, 'bad_request', '보낼 질문이 없습니다.');
  if (messages.length > SIZE.maxMessages) {
    return fail(413, 'too_long', `대화가 너무 길어졌습니다(${SIZE.maxMessages}개 초과). 새 대화로 다시 시작해 주세요.`);
  }
  let total = 0;
  const clean = [];
  for (const m of messages) {
    if (!m || typeof m !== 'object' || (m.role !== 'user' && m.role !== 'assistant')) {
      return fail(400, 'bad_request', '허용되지 않은 메시지 형식입니다. 질문과 답만 보낼 수 있습니다.');
    }
    if (typeof m.content !== 'string') return fail(400, 'bad_request', '메시지 내용은 글자여야 합니다.');
    if (m.content.length > SIZE.maxContent) {
      return fail(413, 'too_long', '한 번에 보낼 수 있는 글은 12,000자까지입니다. 내용을 나눠서 보내 주세요.');
    }
    total += m.content.length;
    if (total > SIZE.maxTotal) {
      return fail(413, 'too_long', '대화 전체가 40,000자를 넘었습니다. 새 대화로 다시 시작해 주세요.');
    }
    const empty = m.content.trim().length === 0;
    if (m.role === 'user' && empty) return fail(400, 'bad_request', '빈 질문은 보낼 수 없습니다.');
    if (m.role === 'assistant' && (empty || clean.length === 0)) continue;
    clean.push({ role: m.role, content: m.content });
  }
  if (clean.length === 0 || clean[clean.length - 1].role !== 'user') {
    return fail(400, 'bad_request', '마지막 메시지는 질문이어야 합니다.');
  }
  return { ok: true, messages: clean };
}

// Attachment names reach the prompt inside «» markers: control characters and line breaks become
// spaces, invisible format characters and the marker quotes are dropped.
const NAME_SPACE = /[\u0000-\u001f\u007f-\u009f\u2028\u2029]/g;
const NAME_DROP = /[\u200b-\u200f\u202a-\u202e\u2066-\u2069\ufeff«»]/g;
const IMAGE_PREFIXES = ['data:image/png;base64,', 'data:image/jpeg;base64,', 'data:image/webp;base64,'];
const BASE64_TAIL = /[A-Za-z0-9+/]+={0,2}$/y; // sticky: one anchored attempt at lastIndex, linear time

function cleanName(name, fallback) {
  let s = typeof name === 'string' ? name.slice(0, 400) : '';
  s = s.replace(NAME_SPACE, ' ').replace(NAME_DROP, '').replace(/\s+/g, ' ').trim();
  if (s.length > SIZE.maxName) s = s.slice(0, SIZE.maxName).replace(/[\ud800-\udbff]$/, '').trim();
  return s || fallback;
}

function isImageDataUrl(url) {
  const prefix = IMAGE_PREFIXES.find((p) => url.startsWith(p));
  if (!prefix || (url.length - prefix.length) % 4 !== 0) return false;
  BASE64_TAIL.lastIndex = prefix.length;
  return BASE64_TAIL.test(url);
}

const isList = (a) => a == null || Array.isArray(a);
const badItem = (x, key) => !x || typeof x !== 'object' || typeof x[key] !== 'string' || (x.name != null && typeof x.name !== 'string');

// Optional attachments: files [{name, text}] and images [{name, url}]. Called after validateMessages.
// Returns { ok:true, files, images } with cleaned names (empty lists when none) or a rejection.
// Attachments go into the first user turn, so with attachments the client's first message must be a question.
export function validateAttachments(files, images, messages) {
  if (!isList(files)) return fail(400, 'bad_request', '첨부 파일 형식이 올바르지 않습니다.');
  if (!isList(images)) return fail(400, 'bad_request', '첨부 이미지 형식이 올바르지 않습니다.');
  const fileList = files || [];
  const imageList = images || [];
  if (fileList.length > SIZE.maxFiles) return fail(413, 'too_long', `파일은 한 번에 ${SIZE.maxFiles}개까지 첨부할 수 있습니다.`);
  if (imageList.length > SIZE.maxImages) return fail(413, 'too_long', `이미지는 한 번에 ${SIZE.maxImages}장까지 첨부할 수 있습니다.`);

  const cleanFiles = [];
  let fileTotal = 0;
  for (const f of fileList) {
    if (badItem(f, 'text')) return fail(400, 'bad_request', '첨부 파일 형식이 올바르지 않습니다.');
    const name = cleanName(f.name, `파일 ${cleanFiles.length + 1}`);
    if (f.text.length > SIZE.maxFileChars) {
      return fail(413, 'too_long', `«${name}» 파일의 글이 100,000자를 넘습니다. 필요한 부분만 남겨 다시 첨부해 주세요.`);
    }
    fileTotal += f.text.length;
    if (fileTotal > SIZE.maxFilesTotal) {
      return fail(413, 'too_long', '첨부 파일의 글은 모두 합쳐 150,000자까지 보낼 수 있습니다. 파일 수를 줄여 다시 보내 주세요.');
    }
    if (!/\S/.test(f.text)) {
      return fail(400, 'bad_request', `«${name}» 파일에서 읽을 수 있는 글자를 찾지 못했습니다. 스캔한 이미지 PDF라면 글자를 고를 수 있는 파일로 바꿔 주세요.`);
    }
    cleanFiles.push({ name, text: f.text });
  }

  const cleanImages = [];
  let imageTotal = 0;
  for (const im of imageList) {
    if (badItem(im, 'url')) return fail(400, 'bad_request', '첨부 이미지 형식이 올바르지 않습니다.');
    const name = cleanName(im.name, `이미지 ${cleanImages.length + 1}`);
    if (im.url.length > SIZE.maxImageUrl) return fail(413, 'too_long', `«${name}» 이미지가 너무 큽니다. 더 작게 줄여 다시 첨부해 주세요.`);
    imageTotal += im.url.length;
    if (imageTotal > SIZE.maxImagesTotal) {
      return fail(413, 'too_long', '첨부 이미지를 모두 합친 크기가 너무 큽니다. 이미지 수를 줄여 다시 보내 주세요.');
    }
    if (!isImageDataUrl(im.url)) {
      return fail(400, 'bad_request', `«${name}» 이미지를 읽을 수 없습니다. PNG, JPEG, WebP 이미지만 첨부할 수 있습니다.`);
    }
    cleanImages.push({ name, url: im.url });
  }

  if ((cleanFiles.length || cleanImages.length) && messages?.[0]?.role !== 'user') {
    return fail(400, 'bad_request', '첨부 자료는 첫 질문과 함께 보내야 합니다.');
  }
  return { ok: true, files: cleanFiles, images: cleanImages };
}

// ---------- access code ----------

const enc = new TextEncoder();
let codeCache = { raw: null, digests: [] };

function normCode(s) {
  return String(s).normalize('NFC').trim().toLowerCase();
}

async function sha256(s) {
  return new Uint8Array(await crypto.subtle.digest('SHA-256', enc.encode(s)));
}

function timingSafeEqual(a, b) {
  let diff = a.length ^ b.length;
  for (let i = 0; i < a.length; i++) diff |= a[i] ^ (b[i] ?? 0);
  return diff === 0;
}

export function parseCodes(raw) {
  return String(raw ?? '')
    .split(',')
    .map(normCode)
    .filter(Boolean);
}

// Returns a short non-reversible id of the matched code (used as the per-code counter key), or null.
// Compares the input digest against every configured digest without early exit.
export async function matchCode(input, raw) {
  if (codeCache.raw !== raw) {
    codeCache = { raw, digests: await Promise.all(parseCodes(raw).map(sha256)) };
  }
  const d = await sha256(normCode(input));
  let hit = false;
  for (const cd of codeCache.digests) hit = timingSafeEqual(d, cd) | hit;
  if (!hit) return null;
  return Array.from(d.subarray(0, 8), (b) => b.toString(16).padStart(2, '0')).join('');
}

// ---------- client identity ----------

// IPv4 as is; IPv6 collapsed to its /64 so one host cannot rotate addresses inside its prefix.
export function clientKey(request) {
  const raw = String(request.headers.get('CF-Connecting-IP') || '').trim().toLowerCase();
  if (!raw) return 'unknown';
  if (!raw.includes(':')) return raw.slice(0, 45);
  const ip = raw.split('%')[0];
  const lastColon = ip.lastIndexOf(':');
  if (ip.includes('.')) return ip.slice(lastColon + 1).slice(0, 45); // IPv4-mapped
  const [head, tail] = ip.split('::');
  const h = head ? head.split(':') : [];
  let groups = h;
  if (tail !== undefined) {
    const t = tail ? tail.split(':') : [];
    groups = [...h, ...Array(Math.max(0, 8 - h.length - t.length)).fill('0'), ...t];
  }
  const prefix = groups.slice(0, 4).map((g) => g.replace(/^0+(?=.)/, '') || '0');
  while (prefix.length < 4) prefix.push('0');
  return `${prefix.join(':')}::/64`;
}

// ---------- counter store ----------

const mem = new Map(); // key -> { v, exp }
const flights = new Map(); // ip -> Map(id -> exp)
let kvWarned = false;

function memGet(key, now) {
  const e = mem.get(key);
  if (!e) return undefined;
  if (e.exp <= now) {
    mem.delete(key);
    return undefined;
  }
  return e.v;
}

function memSet(key, v, ttlSec, now) {
  mem.set(key, { v, exp: now + ttlSec * 1000 });
  if (mem.size > MEM_MAX) {
    for (const [k, e] of mem) if (e.exp <= now) mem.delete(k);
    let extra = mem.size - Math.floor(MEM_MAX * 0.8);
    for (const k of mem.keys()) {
      if (extra-- <= 0) break;
      mem.delete(k);
    }
  }
}

function toInt(v) {
  const n = parseInt(v, 10);
  return Number.isFinite(n) && n > 0 ? n : 0;
}

function warnKv() {
  if (!kvWarned) {
    kvWarned = true;
    console.warn('[ai] KV error; using in-memory limits only');
  }
}

export function createStore(env = {}, waitUntil) {
  const b = env.AI_KV;
  const kv = b && typeof b.get === 'function' && typeof b.put === 'function' ? b : null;
  const later = (p) => {
    const q = Promise.resolve(p).catch(warnKv);
    if (typeof waitUntil === 'function') waitUntil(q);
  };
  return {
    hasKv: Boolean(kv),
    async get(key) {
      if (!kv) return null;
      try {
        return await kv.get(key);
      } catch {
        warnKv();
        return null;
      }
    },
    put(key, value, ttlSec) {
      if (kv) later(kv.put(key, value, { expirationTtl: Math.max(60, Math.ceil(ttlSec)) }));
    },
    // Read-modify-write in the background: fn(oldValue) -> newValue.
    update(key, fn, ttlSec) {
      if (!kv) return;
      later(
        (async () => {
          const next = fn(await kv.get(key));
          await kv.put(key, next, { expirationTtl: Math.max(60, Math.ceil(ttlSec)) });
        })(),
      );
    },
  };
}

async function readCounts(store, keys, now) {
  const kvVals = await Promise.all(keys.map((k) => store.get(k)));
  return keys.map((k, i) => Math.max(toInt(memGet(k, now)), toInt(kvVals[i])));
}

function bump(store, key, base, ttlSec, now) {
  const v = Math.max(toInt(memGet(key, now)), base) + 1;
  memSet(key, v, ttlSec, now);
  store.put(key, String(v), ttlSec);
}

function minuteOf(now) {
  return Math.floor(now / 60000);
}

function windowKeys(prefix, ip, minute) {
  const keys = [];
  for (let m = minute - WINDOW_MIN + 1; m <= minute; m++) keys.push(`${prefix}:${ip}:${m}`);
  return keys;
}

function sum(a) {
  return a.reduce((x, y) => x + y, 0);
}

// Seconds until enough old minute buckets leave the window for the total to drop below the limit.
function windowRetryAfter(counts, minute, now, limit) {
  let total = sum(counts);
  for (let i = 0; i < counts.length; i++) {
    total -= counts[i];
    if (total < limit) return Math.max(1, (minute - WINDOW_MIN + 1 + i + WINDOW_MIN) * 60 - now / 1000);
  }
  return 60;
}

export function kstDay(now) {
  return Math.floor((now + KST_OFFSET_MS) / 86400_000);
}

function untilKstMidnight(now) {
  return Math.max(1, ((kstDay(now) + 1) * 86400_000 - (now + KST_OFFSET_MS)) / 1000);
}

function reject(status, code, message, retryAfter) {
  return { ok: false, status, code, message, retryAfter };
}

const MSG_BADCODE_LOCK = '잘못된 수업 코드가 너무 많이 입력되었습니다. 10분쯤 뒤에 다시 시도해 주세요.';
const MSG_RATE = 'AI 호출 한도에 도달했습니다. 잠시 후 다시 시도해 주세요.';
const MSG_GLOBAL_MIN = '지금 AI를 쓰는 사람이 많습니다. 1분쯤 뒤에 다시 시도해 주세요.';
const MSG_INFLIGHT = '지금 만들고 있는 답이 너무 많습니다. 앞의 답이 끝난 뒤 다시 시도해 주세요.';
const MSG_CODE_DAY = '오늘 이 수업 코드로 쓸 수 있는 AI 호출을 모두 썼습니다. 내일 다시 이용해 주세요.';
const MSG_GLOBAL_DAY = '오늘 준비된 AI 호출을 모두 썼습니다. 내일 다시 이용해 주세요.';

// Wrong-code lockout. While locked, every request from the IP is refused (also with a right code),
// otherwise the lockout would still answer "right"/"wrong" to a guesser.
export async function badCodeGate(store, ip, now, limit) {
  const minute = minuteOf(now);
  const keys = windowKeys('w', ip, minute);
  const memCounts = keys.map((k) => toInt(memGet(k, now)));
  if (sum(memCounts) >= limit) {
    return { locked: reject(429, 'rate', MSG_BADCODE_LOCK, windowRetryAfter(memCounts, minute, now, limit)) };
  }
  const counts = await readCounts(store, keys, now);
  if (sum(counts) >= limit) {
    return { locked: reject(429, 'rate', MSG_BADCODE_LOCK, windowRetryAfter(counts, minute, now, limit)) };
  }
  return {
    locked: null,
    recordWrong() {
      bump(store, keys[keys.length - 1], counts[counts.length - 1], WINDOW_TTL, now);
    },
  };
}

function liveMemFlights(ip, now) {
  const m = flights.get(ip);
  if (!m) return 0;
  for (const [id, exp] of m) if (exp <= now) m.delete(id);
  if (m.size === 0) flights.delete(ip);
  return m.size;
}

function parseFlights(raw, now) {
  try {
    const a = JSON.parse(raw || '[]');
    return Array.isArray(a) ? a.filter((e) => Array.isArray(e) && typeof e[0] === 'string' && e[1] > now) : [];
  } catch {
    return [];
  }
}

// Checks every cap, then counts the request and takes an in-flight slot.
// Returns { ok:true, release } or a rejection { ok:false, status, code, message, retryAfter }.
export async function admit(store, { ip, codeId, now, limits }) {
  const minute = minuteOf(now);
  const day = kstDay(now);
  const ipKeys = windowKeys('r', ip, minute);
  const codeKey = `c:${codeId}:${day}`;
  const globalKey = `g:${day}`;
  const minuteKey = `m:${minute}`;
  const flightKey = `f:${ip}`;

  const evaluate = (ipCounts, codeCount, globalCount, minuteCount, flightCount) => {
    if (sum(ipCounts) >= limits.ip10m) {
      return reject(429, 'rate', MSG_RATE, windowRetryAfter(ipCounts, minute, now, limits.ip10m));
    }
    if (flightCount >= limits.inflight) return reject(429, 'rate', MSG_INFLIGHT, 15);
    if (minuteCount >= limits.globalMin) return reject(429, 'rate', MSG_GLOBAL_MIN, (minute + 1) * 60 - now / 1000);
    if (codeCount >= limits.codeDay) return reject(429, 'budget', MSG_CODE_DAY, untilKstMidnight(now));
    if (globalCount >= limits.globalDay) return reject(429, 'budget', MSG_GLOBAL_DAY, untilKstMidnight(now));
    return null;
  };

  // This isolate already knows a cap is hit: refuse without spending KV operations.
  const memIp = ipKeys.map((k) => toInt(memGet(k, now)));
  const early = evaluate(
    memIp,
    toInt(memGet(codeKey, now)),
    toInt(memGet(globalKey, now)),
    toInt(memGet(minuteKey, now)),
    liveMemFlights(ip, now),
  );
  if (early) return early;

  const [ipCounts, shared, flightRaw] = await Promise.all([
    readCounts(store, ipKeys, now),
    readCounts(store, [codeKey, globalKey, minuteKey], now),
    store.get(flightKey),
  ]);
  const kvFlights = parseFlights(flightRaw, now);
  const flightCount = Math.max(liveMemFlights(ip, now), kvFlights.length);
  const denied = evaluate(ipCounts, shared[0], shared[1], shared[2], flightCount);
  if (denied) return denied;

  bump(store, ipKeys[ipKeys.length - 1], ipCounts[ipCounts.length - 1], WINDOW_TTL, now);
  bump(store, codeKey, shared[0], DAY_TTL, now);
  bump(store, globalKey, shared[1], DAY_TTL, now);
  bump(store, minuteKey, shared[2], 120, now);

  const id = crypto.randomUUID();
  if (!flights.has(ip)) flights.set(ip, new Map());
  flights.get(ip).set(id, now + FLIGHT_TTL_MS);
  const flightTtl = FLIGHT_TTL_MS / 1000 + 60;
  if (store.hasKv) store.put(flightKey, JSON.stringify([...kvFlights, [id, now + FLIGHT_TTL_MS]]), flightTtl);

  let released = false;
  const release = () => {
    if (released) return;
    released = true;
    const m = flights.get(ip);
    if (m) {
      m.delete(id);
      if (m.size === 0) flights.delete(ip);
    }
    store.update(flightKey, (raw) => JSON.stringify(parseFlights(raw, Date.now()).filter((e) => e[0] !== id)), flightTtl);
  };
  return { ok: true, release };
}

// Test hook: clears the per-isolate counters.
export function _resetMemory() {
  mem.clear();
  flights.clear();
  codeCache = { raw: null, digests: [] };
  kvWarned = false;
}
