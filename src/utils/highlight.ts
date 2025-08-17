/*
 * Utilities for building and highlighting snippets for Discord messages
 */

function escapeMarkdown(s: string): string {
  return s.replace(/([*_`~|>])/g, '\\$1');
}

export function tokenizeQuery(query: string): string[] {
  return (query || '')
    .split(/\s+/)
    .map(t => t.trim())
    .filter(Boolean)
    .slice(0, 8); // cap tokens
}

export function buildSnippet(text: string, terms: string[], maxLen = 240): string {
  const src = text || '';
  if (!src) return '';
  const lower = src.toLowerCase();
  const needles = terms.map(t => t.toLowerCase()).filter(Boolean);
  if (needles.length === 0) return escapeMarkdown(src.slice(0, maxLen));

  // find first hit
  let hitIdx = -1;
  let hitLen = 0;
  for (const n of needles) {
    const i = lower.indexOf(n);
    if (i >= 0 && (hitIdx === -1 || i < hitIdx)) {
      hitIdx = i;
      hitLen = n.length;
    }
  }
  if (hitIdx === -1) return escapeMarkdown(src.slice(0, maxLen));

  const pad = Math.max(0, Math.floor((maxLen - hitLen) / 2));
  const start = Math.max(0, hitIdx - pad);
  const end = Math.min(src.length, hitIdx + hitLen + pad);
  const prefix = start > 0 ? '…' : '';
  const suffix = end < src.length ? '…' : '';
  const slice = src.slice(start, end).replace(/\s+/g, ' ').trim();
  return prefix + escapeMarkdown(slice) + suffix;
}

export function highlightSnippet(snippet: string, terms: string[]): string {
  let out = snippet || '';
  for (const t of terms.filter(Boolean)) {
    const esc = t.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    const re = new RegExp(`(${esc})`, 'gi');
    out = out.replace(re, '**$1**');
  }
  return out;
}
