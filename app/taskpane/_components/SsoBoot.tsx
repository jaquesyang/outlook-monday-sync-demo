'use client';

import { useEffect } from 'react';

export function SsoBoot() {
  useEffect(() => {
    const Office = (window as unknown as { Office?: { onReady?: (fn: () => void) => void; auth?: { getAccessToken: (opts: object) => Promise<string> } } }).Office;
    if (!Office?.onReady) return;
    Office.onReady(async () => {
      try {
        const idToken = await Office.auth!.getAccessToken({ allowSignInPrompt: true });
        await fetch('/api/session', {
          method: 'POST',
          credentials: 'include',
          headers: { 'content-type': 'application/json' },
          body: JSON.stringify({ idToken }),
        });
      } catch (e) {
        console.warn('SSO unavailable:', e);
      }
    });
  }, []);
  return null;
}
