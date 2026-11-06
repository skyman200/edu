// Result "views": the browser asks for a designed document (doc), a JSON slide deck (deck) or JSON edits of an uploaded deck (edit).
// The server owns the prompts and the model choice; the client only names the view.
import { DOC_PROMPT, DOC_HINTS } from './view-doc.js';
import { DECK_PROMPT, EDIT_PROMPT, DECK_MAX_TOKENS, EDIT_MAX_TOKENS } from './view-ppt.js';

const BAL_WHY = '를 골랐습니다';
export const VIEWS = {
  doc: { maxTokens: 8000, why: `표지와 카드·표·타임라인으로 디자인된 문서를 한 번에 짜야 해서 Claude Sonnet 5.5${BAL_WHY}` },
  deck: { maxTokens: DECK_MAX_TOKENS, why: `슬라이드 구성과 발표 노트를 정확한 형식으로 짜야 해서 Claude Sonnet 5.5${BAL_WHY}` },
  edit: { maxTokens: EDIT_MAX_TOKENS, why: `올린 PPT를 읽고 고칠 곳만 정확히 짚어야 해서 Claude Sonnet 5.5${BAL_WHY}` },
};

const PROMPTS = {
  doc: (ex) => `${DOC_PROMPT}${DOC_HINTS[ex] ? `\n\n[이번 실습에 맞는 구성]\n${DOC_HINTS[ex]}` : ''}`,
  deck: () => DECK_PROMPT,
  edit: () => EDIT_PROMPT,
};

export function normalizeView(v) {
  return typeof v === 'string' && Object.hasOwn(VIEWS, v) ? v : null;
}

// Same model whatever the exercise: the view decides length and format, so it overrides the exercise's own tier.
export function applyView(route, view, tiers, labels) {
  if (!view) return route;
  const t = tiers.balanced, v = VIEWS[view];
  return { ...route, model: t.model, fallbacks: [...t.fallbacks], tier: 'balanced', label: labels[t.model], maxTokens: v.maxTokens, why: v.why };
}

export function viewPrompt(view, ex) {
  const f = view && PROMPTS[view];
  return f ? f(ex) : '';
}
