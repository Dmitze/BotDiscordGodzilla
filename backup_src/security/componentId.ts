import crypto from 'crypto';
import { securityConfig } from '../config/security';

// Base64url helpers
const b64url = {
  encode: (buf: Buffer) => buf.toString('base64').replace(/=/g, '').replace(/\+/g, '-').replace(/\//g, '_'),
  decode: (str: string) => Buffer.from(str.replace(/-/g, '+').replace(/_/g, '/'), 'base64'),
};

export type ComponentPayload = Record<string, any> & { t?: number };

// Field shorteners to reduce payload length in customId
const shortenMap: Record<string, string> = {
  kind: 'k',
  sid: 's',
  page: 'p',
  ts: 't',
  action: 'a',
  documentId: 'd',
  pageSize: 'ps',
  format: 'f',
};

const expandMap: Record<string, string> = Object.fromEntries(
  Object.entries(shortenMap).map(([full, short]) => [short, full])
);

// Value shorteners (for specific fields like kind)
const kindShortMap: Record<string, string> = {
  filesrch: 'fs',
  filetxt: 'ft',
  srch: 's',
};
const kindExpandMap: Record<string, string> = Object.fromEntries(
  Object.entries(kindShortMap).map(([full, short]) => [short, full])
);

function toCompactPayload(obj: Record<string, any>, expMs: number): Record<string, any> {
  const out: Record<string, any> = { e: expMs };
  for (const [k, v] of Object.entries(obj)) {
    const short = shortenMap[k] || k;
    if (short === 'k' && typeof v === 'string') {
      out[short] = kindShortMap[v] || v;
    } else {
      out[short] = v;
    }
  }
  return out;
}

function fromCompactPayload<T extends Record<string, any>>(obj: Record<string, any>): T & { exp?: number } {
  const out: Record<string, any> = {};
  for (const [k, v] of Object.entries(obj)) {
    if (k === 'e') {
      (out as any)['exp'] = v as number;
      continue;
    }
    const full = expandMap[k] || k;
    if (full === 'kind' && typeof v === 'string') {
      out[full] = kindExpandMap[v] || v;
    } else {
      out[full] = v;
    }
  }
  return out as T & { exp?: number };
}

export function signComponentId(payload: ComponentPayload, ttlMs = securityConfig.components.ttlMs): string {
  // Use seconds to reduce payload size
  const expSec = Math.floor((Date.now() + ttlMs) / 1000);
  // Compact 3-part format: 'c'.<encBody>.<encSig>
  const compactBody = toCompactPayload(payload, expSec);
  const encBody = b64url.encode(Buffer.from(JSON.stringify(compactBody)));
  const prefix = 'c';
  const toSign = `${prefix}.${encBody}`;
  // Truncate HMAC to 12 bytes (96-bit) to keep IDs under Discord's 100-char limit
  const fullSig = crypto.createHmac('sha256', securityConfig.components.hmacKey).update(toSign).digest();
  const sig = fullSig.subarray(0, 12);
  const encSig = b64url.encode(sig);
  return `${toSign}.${encSig}`;
}

export function verifyComponentId<T extends ComponentPayload = ComponentPayload>(customId: string): { valid: boolean; payload?: T; reason?: string } {
  try {
    const parts = customId.split('.');
    if (parts.length !== 3) return { valid: false, reason: 'format' };
    const [p1, p2, p3] = parts as [string, string, string];

    // Compact format: 'c'.<encBody>.<encSig>
    if (p1 === 'c') {
      const toSign = `c.${p2}`;
      // Expect truncated signature (12 bytes)
      const expectedFull = crypto.createHmac('sha256', securityConfig.components.hmacKey).update(toSign).digest();
      const expected = expectedFull.subarray(0, 12);
      const given = b64url.decode(p3);
      if (given.length !== expected.length || !crypto.timingSafeEqual(expected, given)) return { valid: false, reason: 'signature' };

      const bodyRaw = b64url.decode(p2).toString('utf8');
      const compact = JSON.parse(bodyRaw) as Record<string, any>;
      const expanded = fromCompactPayload<T>(compact); // includes exp
      // exp is stored in seconds
      if (typeof expanded.exp === 'number' && Math.floor(Date.now() / 1000) > expanded.exp) {
        return { valid: false, reason: 'expired' };
      }
      if ('exp' in expanded) delete (expanded as any).exp;
      return { valid: true, payload: expanded as T };
    }

    // Legacy JWT-like format: <encHeader>.<encBody>.<encSig>
    const toSign = `${p1}.${p2}`;
    const expected = crypto.createHmac('sha256', securityConfig.components.hmacKey).update(toSign).digest();
    const given = b64url.decode(p3);
    if (!crypto.timingSafeEqual(expected, given)) return { valid: false, reason: 'signature' };

    const bodyRaw = b64url.decode(p2).toString('utf8');
    const payload = JSON.parse(bodyRaw) as T & { exp?: number };
    if (typeof payload.exp === 'number' && Date.now() > payload.exp) {
      return { valid: false, reason: 'expired' };
    }
    if ('exp' in payload) delete (payload as any).exp;
    return { valid: true, payload };
  } catch (e) {
    return { valid: false, reason: 'error' };
  }
}
