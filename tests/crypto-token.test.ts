import { describe, it, expect, beforeAll } from 'vitest';
import { randomBytes } from 'node:crypto';
import { encryptToken, decryptToken } from '@/lib/crypto/token';

describe('token encryption', () => {
  beforeAll(() => {
    process.env.TOKEN_ENC_KEY = randomBytes(32).toString('base64');
  });

  it('round-trips a string through encrypt/decrypt', () => {
    const plaintext = 'ya29.a0AeXRPp7-not-a-real-token';
    const cipher = encryptToken(plaintext);
    expect(cipher).toBeInstanceOf(Buffer);
    expect(cipher.length).toBeGreaterThan(28);
    expect(decryptToken(cipher)).toBe(plaintext);
  });

  it('produces different ciphertext on repeated calls (random IV)', () => {
    const a = encryptToken('hello');
    const b = encryptToken('hello');
    expect(a.equals(b)).toBe(false);
  });

  it('fails to decrypt tampered data', () => {
    const cipher = encryptToken('hello');
    cipher[cipher.length - 1] ^= 0x01;
    expect(() => decryptToken(cipher)).toThrow();
  });

  it('throws if key is missing', () => {
    const saved = process.env.TOKEN_ENC_KEY;
    delete process.env.TOKEN_ENC_KEY;
    expect(() => encryptToken('x')).toThrow(/TOKEN_ENC_KEY/);
    process.env.TOKEN_ENC_KEY = saved;
  });
});
