// Simple redaction helpers to avoid PII/URLs leaking into logs

const URL_REGEX = /(https?:\/\/[^\s)]+)|((drive|docs|sheets|google)\.[^\s)]+)/gi;
const EMAIL_REGEX = /[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}/g;

export function redactString(input: string): string {
  return input
    .replace(EMAIL_REGEX, '[REDACTED_EMAIL]')
    .replace(URL_REGEX, '[REDACTED_URL]');
}

export function redact<T>(value: T): T {
  if (typeof value === 'string') return redactString(value) as unknown as T;
  if (Array.isArray(value)) return (value.map((v) => redact(v)) as unknown) as T;
  if (value && typeof value === 'object') {
    const out: Record<string, unknown> = {};
    for (const [k, v] of Object.entries(value)) out[k] = redact(v as unknown as string);
    return out as T;
  }
  return value;
}
