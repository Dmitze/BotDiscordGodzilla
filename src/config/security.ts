import { z } from 'zod';

// Helpers to parse env
const bool = (v: string | undefined, d: boolean) => {
  if (v == null) return d;
  return /^(1|true|yes|on)$/i.test(v);
};
const num = (v: string | undefined, d: number) => {
  const n = Number(v);
  return Number.isFinite(n) ? n : d;
};
const jsonArr = (v: string | undefined, d: string[]) => {
  if (!v) return d;
  try {
    const parsed = JSON.parse(v);
    return Array.isArray(parsed) ? parsed.map(String) : d;
  } catch {
    // Support comma-separated fallback
    return v.split(',').map((s) => s.trim()).filter(Boolean);
  }
};

// Defaults
const DEFAULT_MIME_ALLOWLIST = [
  'application/pdf',
  'application/vnd.openxmlformats-officedocument.wordprocessingml.document', // docx
  'application/msword', // doc
  'text/plain',
  'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', // xlsx
  'application/vnd.ms-excel', // xls
  'image/png',
  'image/jpeg',
];
const DEFAULT_MAX_BYTES = 25 * 1024 * 1024; // 25 MB
const DEFAULT_PII_EMAIL = true;
const DEFAULT_PII_PHONE = true;
const DEFAULT_PII_MASTER = true;
const DEFAULT_COMPONENT_TTL = 15 * 60 * 1000; // 15 minutes

export const SecurityConfigSchema = z.object({
  mime: z.object({
    allowlist: z.array(z.string()).nonempty(),
  }),
  file: z.object({
    maxBytes: z.number().int().positive(),
  }),
  pii: z.object({
    master: z.boolean(),
    email: z.boolean(),
    phone: z.boolean(),
  }),
  components: z.object({
    hmacKey: z.string().min(16),
    ttlMs: z.number().int().positive(),
  }),
});

export type SecurityConfig = z.infer<typeof SecurityConfigSchema>;

const raw: SecurityConfig = {
  mime: {
    allowlist: jsonArr(process.env['SECURITY_MIME_ALLOWLIST'], DEFAULT_MIME_ALLOWLIST),
  },
  file: {
    maxBytes: num(process.env['SECURITY_MAX_BYTES'], DEFAULT_MAX_BYTES),
  },
  pii: {
    master: bool(process.env['SECURITY_PII_MASTER'], DEFAULT_PII_MASTER),
    email: bool(process.env['SECURITY_PII_EMAIL'], DEFAULT_PII_EMAIL),
    phone: bool(process.env['SECURITY_PII_PHONE'], DEFAULT_PII_PHONE),
  },
  components: {
    hmacKey: process.env['COMPONENT_HMAC_KEY'] || 'change_me_please_min16chars',
    ttlMs: num(process.env['COMPONENT_TTL_MS'], DEFAULT_COMPONENT_TTL),
  },
};

export const securityConfig: SecurityConfig = SecurityConfigSchema.parse(raw);

export const isMimeAllowed = (mime: string): boolean =>
  securityConfig.mime.allowlist.includes(mime);

export const withinSizeLimit = (sizeBytes: number): boolean =>
  sizeBytes <= securityConfig.file.maxBytes;
