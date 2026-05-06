import { randomBytes } from 'node:crypto';
import { encryptToken, decryptToken } from '@/lib/crypto/token';

type Payload = { userId: string; purpose: 'ms' | 'monday'; nonce?: string; callbackToken?: string };

export function encodeState(p: Omit<Payload, 'nonce'>): string {
  const full: Payload = { ...p, nonce: randomBytes(16).toString('hex') };
  const blob = encryptToken(JSON.stringify(full));
  return Buffer.from(blob).toString('base64url');
}

export function decodeState(s: string): Required<Payload> {
  const blob = Buffer.from(s, 'base64url');
  const json = decryptToken(blob);
  const parsed = JSON.parse(json) as Required<Payload>;
  if (!parsed.userId || !parsed.nonce || !parsed.purpose) {
    throw new Error('invalid state payload');
  }
  return parsed;
}
