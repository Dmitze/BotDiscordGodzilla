import crypto from 'crypto';
import { securityConfig } from '../config/security';

// Base64url helpers
const b64url = {
  encode: (buf: Buffer) => buf.toString('base64').replace(/=/g, '').replace(/\+/g, '-').replace(/\//g, '_'),
  decode: (str: string) => Buffer.from(str.replace(/-/g, '+').replace(/_/g, '/'), 'base64'),
};

export type ComponentPayload = Record<string, any> & { t?: number };

export function signComponentId(payload: ComponentPayload, ttlMs = securityConfig.components.ttlMs): string {
  const exp = Date.now() + ttlMs;
  const header = { alg: 'HS256', typ: 'CID' };
  const body = { ...payload, exp };
  const encHeader = b64url.encode(Buffer.from(JSON.stringify(header)));
  const encBody = b64url.encode(Buffer.from(JSON.stringify(body)));
  const toSign = `${encHeader}.${encBody}`;
  const sig = crypto.createHmac('sha256', securityConfig.components.hmacKey).update(toSign).digest();
  const encSig = b64url.encode(sig);
  return `${toSign}.${encSig}`;
}

export function verifyComponentId<T extends ComponentPayload = ComponentPayload>(customId: string): { valid: boolean; payload?: T; reason?: string } {
  try {
    const parts = customId.split('.');
    if (parts.length !== 3) return { valid: false, reason: 'format' };
    const [encHeader, encBody, encSig] = parts as [string, string, string];
    const toSign = `${encHeader}.${encBody}`;
    const expected = crypto.createHmac('sha256', securityConfig.components.hmacKey).update(toSign).digest();
    const given = b64url.decode(encSig);
    if (!crypto.timingSafeEqual(expected, given)) return { valid: false, reason: 'signature' };

    const bodyRaw = b64url.decode(encBody).toString('utf8');
    const payload = JSON.parse(bodyRaw) as T & { exp?: number };
    if (typeof payload.exp === 'number' && Date.now() > payload.exp) {
      return { valid: false, reason: 'expired' };
    }
    // Hide exp from downstream usage
    if ('exp' in payload) delete (payload as any).exp;
    return { valid: true, payload };
  } catch (e) {
    return { valid: false, reason: 'error' };
  }
}
