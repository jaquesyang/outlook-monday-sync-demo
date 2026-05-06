/** Temporary in-memory mapping for OAuth callback tokens.
 *  Outlook desktop iframe and popup do not share cookies,
 *  so we use a one-time token passed through OAuth state instead.
 */

const STORE = new Map<string, { userId: string; expiresAt: number }>();
const TTL_MS = 5 * 60 * 1000; // 5 minutes

function gc() {
  const now = Date.now();
  for (const [k, v] of STORE) {
    if (v.expiresAt < now) STORE.delete(k);
  }
}

export function setCallbackToken(token: string, userId: string) {
  gc();
  STORE.set(token, { userId, expiresAt: Date.now() + TTL_MS });
}

export function consumeCallbackToken(token: string): string | null {
  gc();
  const entry = STORE.get(token);
  if (!entry) return null;
  STORE.delete(token);
  return entry.userId;
}
