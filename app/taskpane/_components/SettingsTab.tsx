'use client';

import { useEffect, useRef, useState } from 'react';
import { ConnectionStatus } from './ConnectionStatus';

type StatusPayload = {
  microsoft: { connected: boolean; email?: string };
  monday: { connected: boolean; account?: string };
};

export function SettingsTab() {
  const [status, setStatus] = useState<StatusPayload | null>(null);
  const callbackTokenRef = useRef<string>('');

  async function refresh() {
    const token = callbackTokenRef.current;
    const url = token ? `/api/status?callbackToken=${encodeURIComponent(token)}` : '/api/status';
    const r = await fetch(url, { credentials: 'include' });
    if (r.ok) setStatus(await r.json());
  }

  useEffect(() => {
    let cancelled = false;
    fetch('/api/status', { credentials: 'include' })
      .then((r) => (r.ok ? r.json() : null))
      .then((data) => {
        if (!cancelled && data) setStatus(data);
      });

    const handler = (e: MessageEvent) => {
      if (e.data?.type === 'oauth-success') {
        refresh();
      }
    };
    window.addEventListener('message', handler);
    return () => {
      cancelled = true;
      window.removeEventListener('message', handler);
    };
  }, []);

  function openPopup(url: string) {
    const token = crypto.randomUUID();
    callbackTokenRef.current = token;
    const u = new URL(url, window.location.href);
    u.searchParams.set('callbackToken', token);
    const w = window.open(u.toString(), '_blank', 'width=520,height=640');
    const t = setInterval(() => {
      if (w?.closed) {
        clearInterval(t);
        refresh();
      }
    }, 800);
  }

  return (
    <div className="space-y-3 p-4">
      <ConnectionStatus
        label="Microsoft Graph"
        variant={status?.microsoft.connected ? 'green' : 'red'}
        detail={status?.microsoft.email ?? 'Not connected'}
        action={
          status?.microsoft.connected
            ? undefined
            : { label: 'Connect', onClick: () => openPopup('/api/auth/microsoft') }
        }
      />
      <ConnectionStatus
        label="monday.com"
        variant={status?.monday.connected ? 'green' : 'red'}
        detail={status?.monday.account ?? 'Not connected'}
        action={
          status?.monday.connected
            ? undefined
            : { label: 'Connect', onClick: () => openPopup('/api/auth/monday') }
        }
      />
    </div>
  );
}
