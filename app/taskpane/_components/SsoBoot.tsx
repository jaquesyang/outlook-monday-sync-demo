'use client';

import { useEffect, useRef } from 'react';

export function SsoBoot() {
  const mounted = useRef(false);

  useEffect(() => {
    mounted.current = true;
    const Office = (window as unknown as { Office?: { onReady?: (fn: () => void) => void; auth?: { getAccessToken: (opts: object) => Promise<string> } } }).Office;
    if (!Office?.onReady) return;
    Office.onReady(async () => {
      if (!mounted.current) return;
      try {
        if (!Office.auth?.getAccessToken) {
          console.log('SSO not supported in this Outlook version');
          return;
        }
        const idToken = await Office.auth.getAccessToken({ allowSignInPrompt: true });
        if (!mounted.current) return;
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
    return () => { mounted.current = false; };
  }, []);

  return null;
}
