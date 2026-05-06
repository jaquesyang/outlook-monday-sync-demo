'use client';

import { useState } from 'react';

type SyncResult = {
  totalMondayItems: number;
  totalOutlookEvents: number;
  mappings: number;
  mondayToOutlook: { created: number; updated: number; failed: number };
  outlookToMonday: { created: number; updated: number; failed: number };
  conflicts: number;
};

export function BoardsTab() {
  const [loading, setLoading] = useState(false);
  const [result, setResult] = useState<SyncResult | null>(null);
  const [error, setError] = useState('');

  async function handleSync() {
    setLoading(true);
    setError('');
    setResult(null);
    try {
      const r = await fetch('/api/sync', {
        method: 'POST',
        credentials: 'include',
      });
      if (!r.ok) {
        const j = await r.json().catch(() => ({ error: r.statusText }));
        throw new Error(j.error || r.statusText);
      }
      setResult(await r.json());
    } catch (err) {
      setError(err instanceof Error ? err.message : String(err));
    } finally {
      setLoading(false);
    }
  }

  return (
    <div className="space-y-3 p-4">
      <button
        onClick={handleSync}
        disabled={loading}
        className="w-full rounded bg-zinc-900 px-3 py-2 text-sm font-medium text-white disabled:opacity-50"
      >
        {loading ? 'Syncing...' : 'Sync Calendar'}
      </button>

      {error && (
        <div className="rounded bg-red-50 p-2 text-xs text-red-700">
          {error}
        </div>
      )}

      {result && (
        <div className="rounded bg-green-50 p-2 text-xs text-green-800 space-y-1">
          <div>Monday items: {result.totalMondayItems}</div>
          <div>Outlook events: {result.totalOutlookEvents}</div>
          <div>Mappings: {result.mappings}</div>
          {result.conflicts > 0 && <div className="text-orange-600">Conflicts resolved: {result.conflicts}</div>}
          <div className="border-t border-green-200 pt-1 mt-1">
            <div className="font-medium">monday → Outlook</div>
            <div>Created: {result.mondayToOutlook.created}</div>
            <div>Updated: {result.mondayToOutlook.updated}</div>
            {result.mondayToOutlook.failed > 0 && <div className="text-red-600">Failed: {result.mondayToOutlook.failed}</div>}
          </div>
          <div className="border-t border-green-200 pt-1">
            <div className="font-medium">Outlook → monday</div>
            <div>Created: {result.outlookToMonday.created}</div>
            <div>Updated: {result.outlookToMonday.updated}</div>
            {result.outlookToMonday.failed > 0 && <div className="text-red-600">Failed: {result.outlookToMonday.failed}</div>}
          </div>
        </div>
      )}
    </div>
  );
}
