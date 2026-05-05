import { describe, it, expect, beforeAll } from 'vitest';
import { randomBytes } from 'node:crypto';
import { encodeState, decodeState } from '@/lib/auth/oauth-state';

beforeAll(() => {
  process.env.TOKEN_ENC_KEY = randomBytes(32).toString('base64');
});

describe('oauth-state', () => {
  it('round-trips userId + nonce', () => {
    const s = encodeState({ userId: 'u-123', purpose: 'monday' });
    const out = decodeState(s);
    expect(out.userId).toBe('u-123');
    expect(out.purpose).toBe('monday');
    expect(out.nonce).toMatch(/^[a-f0-9]{32}$/);
  });

  it('rejects tampered state', () => {
    const s = encodeState({ userId: 'u-123', purpose: 'monday' });
    const tampered = s.slice(0, -2) + (s.endsWith('A') ? 'B' : 'A');
    expect(() => decodeState(tampered)).toThrow();
  });

  it('produces base64url output (no slashes/pluses/equals)', () => {
    const s = encodeState({ userId: 'u-123', purpose: 'ms' });
    expect(s).toMatch(/^[A-Za-z0-9\-_]+$/);
  });
});
