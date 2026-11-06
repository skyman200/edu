// CORS allow-list and response helpers shared by /api/ai.

const EXACT_ORIGINS = new Set([
  'https://kimcook-edu.pages.dev',
  'https://skyman200.github.io',
  'http://localhost:8788',
  'http://127.0.0.1:8788',
]);
// Cloudflare Pages preview and branch aliases: https://<label>.kimcook-edu.pages.dev
const PREVIEW_ORIGIN = /^https:\/\/[a-z0-9-]{1,63}\.kimcook-edu\.pages\.dev$/;

export function isAllowedOrigin(origin) {
  return typeof origin === 'string' && (EXACT_ORIGINS.has(origin) || PREVIEW_ORIGIN.test(origin));
}

// Headers every response carries. The Origin is reflected only when allowed.
export function baseHeaders(origin, extra) {
  const h = new Headers(extra);
  h.set('Cache-Control', 'no-store');
  h.set('X-Robots-Tag', 'noindex');
  h.set('X-Content-Type-Options', 'nosniff');
  h.set('Vary', 'Origin');
  if (isAllowedOrigin(origin)) {
    h.set('Access-Control-Allow-Origin', origin);
    h.set('Access-Control-Expose-Headers', 'Retry-After');
  }
  return h;
}

export function jsonResponse(status, body, origin, extra) {
  const h = baseHeaders(origin, extra);
  h.set('Content-Type', 'application/json; charset=utf-8');
  return new Response(JSON.stringify(body), { status, headers: h });
}

export function errorResponse(status, code, message, origin, retryAfterSec) {
  const extra = {};
  const body = { error: message, code };
  if (retryAfterSec) {
    const s = Math.max(1, Math.ceil(retryAfterSec));
    extra['Retry-After'] = String(s);
    body.retryAfter = s;
  }
  return jsonResponse(status, body, origin, extra);
}

export function preflightResponse(origin) {
  if (origin != null && !isAllowedOrigin(origin)) {
    return errorResponse(403, 'bad_request', '허용되지 않은 사이트에서 보낸 요청입니다.', origin);
  }
  const h = baseHeaders(origin, { Allow: 'POST, OPTIONS' });
  if (origin != null) {
    h.set('Access-Control-Allow-Methods', 'POST, OPTIONS');
    h.set('Access-Control-Allow-Headers', 'Content-Type');
    h.set('Access-Control-Max-Age', '600');
  }
  return new Response(null, { status: 204, headers: h });
}
