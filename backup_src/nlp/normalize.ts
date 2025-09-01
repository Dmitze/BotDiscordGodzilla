import logger from '@/utils/logger';

export type NormalizeOptions = {
  maxLength?: number; // truncate length
  preserveUrls?: boolean;
  preserveMentions?: boolean;
};

const URL_RE = /(https?:\/\/[^\s)]+)|(www\.[^\s)]+)/gi;
const MENTION_RE = /<@!?\d+>|<@&\d+>|<#\d+>|@[A-Za-z0-9_]+/g;
const PUNCT_RE = /[!"#$%&'()*+,\-./:;<=>?@[\\\]^_`{|}~]/g; // hyphen left to collapse logic

export function normalizeText(input: string, opts: NormalizeOptions = {}): string {
  try {
    const maxLength = opts.maxLength ?? 768;
    if (!input) return '';

    let text = String(input);

    // trim
    text = text.trim();

    // optionally extract URLs and mentions placeholders
    const urls: string[] = [];
    const mentions: string[] = [];

    if (opts.preserveUrls !== false) {
      text = text.replace(URL_RE, (m) => {
        urls.push(m);
        return ` __URL_${urls.length - 1}__ `;
      });
    }

    if (opts.preserveMentions !== false) {
      text = text.replace(MENTION_RE, (m) => {
        mentions.push(m);
        return ` __MENTION_${mentions.length - 1}__ `;
      });
    }

    // lower-case
    text = text.toLowerCase();

    // remove punctuation (keep placeholder tokens)
    text = text.replace(PUNCT_RE, ' ');

    // collapse spaces
    text = text.replace(/\s+/g, ' ').trim();

    // restore placeholders
    if (urls.length) {
      text = text.replace(/__url_(\d+)__/gi, (_m, idx) => urls[Number(idx)] ?? '');
    }
    if (mentions.length) {
      text = text.replace(/__mention_(\d+)__/gi, (_m, idx) => mentions[Number(idx)] ?? '');
    }

    // limit length
    if (text.length > maxLength) {
      text = text.slice(0, maxLength);
    }

    return text;
  } catch (e) {
    logger.debug('normalize_text_failed', { error: e instanceof Error ? e.message : String(e) });
    return String(input ?? '').slice(0, opts.maxLength ?? 768);
  }
}
