'use client';

import { useState } from 'react';
import { SettingsTab } from './_components/SettingsTab';
import { BoardsTab } from './_components/BoardsTab';

const TABS = ['Status', 'Boards', 'Settings', 'Log'] as const;

export default function TaskpanePage() {
  const [tab, setTab] = useState<(typeof TABS)[number]>('Settings');

  return (
    <div className="flex h-screen flex-col">
      <header className="border-b border-zinc-200 px-4 py-2 text-xs font-medium">
        Outlook ↔ monday Sync
      </header>
      <nav className="flex border-b border-zinc-200">
        {TABS.map((t) => (
          <button
            key={t}
            onClick={() => setTab(t)}
            className={
              t === tab
                ? 'flex-1 border-b-2 border-zinc-900 px-2 py-2 text-xs font-semibold'
                : 'flex-1 px-2 py-2 text-xs text-zinc-500'
            }
          >
            {t}
          </button>
        ))}
      </nav>
      <main className="flex-1 overflow-y-auto">
        {tab === 'Settings' && <SettingsTab />}
        {tab === 'Boards' && <BoardsTab />}
        {tab !== 'Settings' && tab !== 'Boards' && (
          <div className="p-4 text-zinc-500">Coming in a later phase.</div>
        )}
      </main>
    </div>
  );
}
