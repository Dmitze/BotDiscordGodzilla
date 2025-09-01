import type { APIEmbed, BaseMessageOptions, InteractionReplyOptions } from 'discord.js';
import { securityConfig } from '../config/security';

const EMAIL_REGEX = /([a-zA-Z0-9._%+-]{1,2})[a-zA-Z0-9._%+-]*(@[A-Za-z0-9.-]+\.[A-Za-z]{2,})/g;
// Simplified phone: sequences of digits (optionally with separators) of length >= 7
const PHONE_REGEX = /(?:(?<!\d)(?:\+?\d[\s-]?){7,}(?!\d))/g;

export function maskText(input: string): string {
  const { pii } = securityConfig;
  if (!pii.master) return input;
  let out = input;
  if (pii.email) {
    out = out.replace(EMAIL_REGEX, (_m, p1, p2) => `${p1}***${p2}`);
  }
  if (pii.phone) {
    out = out.replace(PHONE_REGEX, (m) => {
      // Keep last 2 digits visible
      const digits = m.replace(/\D/g, '');
      if (digits.length < 7) return m;
      const keep = digits.slice(-2);
      const masked = '*'.repeat(Math.max(0, digits.length - 2));
      // Reconstruct as masked+last2
      return masked + keep;
    });
  }
  return out;
}

function maskEmbed(embed: APIEmbed): APIEmbed {
  const copy: APIEmbed = { ...embed };
  if (copy.title) copy.title = maskText(copy.title);
  if (copy.description) copy.description = maskText(copy.description);
  if (copy.footer?.text) copy.footer = { ...copy.footer, text: maskText(copy.footer.text) };
  if (copy.author?.name) copy.author = { ...copy.author, name: maskText(copy.author.name) };
  if (Array.isArray(copy.fields)) {
    copy.fields = copy.fields.map((f) => ({
      ...f,
      name: f.name ? maskText(f.name) : f.name,
      value: f.value ? maskText(f.value) : f.value,
    }));
  }
  return copy;
}

export function maskReplyOptions<T extends string | InteractionReplyOptions | BaseMessageOptions>(
  content: T
): T {
  if (typeof content === 'string') {
    return maskText(content) as T;
  }
  const opts: any = { ...content };
  if (typeof opts.content === 'string') opts.content = maskText(opts.content);
  if (Array.isArray(opts.embeds)) {
    opts.embeds = opts.embeds.map((e: APIEmbed) => maskEmbed(e));
  }
  return opts as T;
}
