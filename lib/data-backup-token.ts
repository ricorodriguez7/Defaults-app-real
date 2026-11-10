import { createHmac, timingSafeEqual } from 'node:crypto';

export function createDataBackupToken(email: string, issuedAt: number) {
  const secret = process.env.AUTH_SECRET;
  if (!secret) throw new Error('AUTH_SECRET is required to enable protected data exports.');
  const payload = Buffer.from(`${email}:${issuedAt}`).toString('base64url');
  const signature = createHmac('sha256', secret).update(payload).digest('base64url');
  return `${payload}.${signature}`;
}

export function verifyDataBackupToken(token: unknown, email: string) {
  if (typeof token !== 'string') return false;
  const [payload, signature, extra] = token.split('.');
  if (!payload || !signature || extra) return false;
  const secret = process.env.AUTH_SECRET;
  if (!secret) return false;
  const expected = createHmac('sha256', secret).update(payload).digest();
  let actual: Buffer;
  try {
    actual = Buffer.from(signature, 'base64url');
  } catch {
    return false;
  }
  if (actual.length !== expected.length || !timingSafeEqual(actual, expected)) return false;
  let decoded = '';
  try {
    decoded = Buffer.from(payload, 'base64url').toString('utf8');
  } catch {
    return false;
  }
  const separator = decoded.lastIndexOf(':');
  const tokenEmail = decoded.slice(0, separator);
  const issuedAt = Number(decoded.slice(separator + 1));
  return separator > 0 && tokenEmail === email &&
    Number.isFinite(issuedAt) && Date.now() - issuedAt >= 0 &&
    Date.now() - issuedAt <= 30 * 60 * 1000;
}
