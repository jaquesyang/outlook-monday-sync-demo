import Script from 'next/script';
import { SsoBoot } from './_components/SsoBoot';

export const metadata = { title: 'Outlook ↔ monday Sync' };

export default function TaskpaneLayout({ children }: { children: React.ReactNode }) {
  return (
    <>
      <Script
        src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js"
        strategy="beforeInteractive"
      />
      <div className="min-h-screen bg-white text-zinc-900 text-sm">
        <SsoBoot />
        {children}
      </div>
    </>
  );
}
