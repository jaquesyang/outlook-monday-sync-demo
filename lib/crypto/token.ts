import { createCipheriv, createDecipheriv, randomBytes } from 'node:crypto';

const ALG = 'aes-256-gcm';
const IV_LEN = 12;
const TAG_LEN = 16;

function key(): Buffer {
  const b64 = process.env.TOKEN_ENC_KEY;
  if (!b64) throw new Error('TOKEN_ENC_KEY env var not set');
  const k = Buffer.from(b64, 'base64');
  if (k.length !== 32) throw new Error('TOKEN_ENC_KEY must decode to 32 bytes');
  return k;
}

/** Encrypt a UTF-8 string. Returns: iv (12) || tag (16) || ciphertext. */
export function encryptToken(plaintext: string): Uint8Array<ArrayBuffer> {
  const iv = randomBytes(IV_LEN);
  const cipher = createCipheriv(ALG, key(), iv);
  const ct = Buffer.concat([cipher.update(plaintext, 'utf8'), cipher.final()]);
  const tag = cipher.getAuthTag();
  const buf = Buffer.concat([iv, tag, ct]);
  return Uint8Array.from(buf).slice() as Uint8Array<ArrayBuffer>;
}

export function decryptToken(blob: Uint8Array): string {
  const buf = Buffer.from(blob);
  const iv = buf.subarray(0, IV_LEN);
  const tag = buf.subarray(IV_LEN, IV_LEN + TAG_LEN);
  const ct = buf.subarray(IV_LEN + TAG_LEN);
  const decipher = createDecipheriv(ALG, key(), iv);
  decipher.setAuthTag(tag);
  return Buffer.concat([decipher.update(ct), decipher.final()]).toString('utf8');
}
