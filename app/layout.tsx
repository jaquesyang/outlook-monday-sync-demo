import type { Metadata } from "next";
import { Geist, Geist_Mono } from "next/font/google";
import "./globals.css";

const geistSans = Geist({
  variable: "--font-geist-sans",
  subsets: ["latin"],
});

const geistMono = Geist_Mono({
  variable: "--font-geist-mono",
  subsets: ["latin"],
});

export const metadata: Metadata = {
  title: "Outlook ↔ monday.com Calendar Sync",
  description: "Two-way calendar sync between Outlook and monday.com.",
};

export default function RootLayout({
  children,
}: Readonly<{
  children: React.ReactNode;
}>) {
  return (
    <html
      lang="en"
      className={`${geistSans.variable} ${geistMono.variable} h-full antialiased`}
    >
      <head>
        <script
          dangerouslySetInnerHTML={{
            __html: `
              (function(){
                console.log('[polyfill] start — history.replaceState is:', typeof window.history?.replaceState);
                try {
                  var h = window.history;
                  // Method 1: patch instance
                  if (h) {
                    try { h.replaceState = function(){}; console.log('[polyfill] method1 ok'); } catch(e1){ console.log('[polyfill] method1 fail:', e1.message); }
                    try { h.pushState = function(){}; } catch(e2){}
                  }
                  // Method 2: patch prototype
                  try {
                    var proto = window.History && window.History.prototype;
                    if (proto) {
                      proto.replaceState = function(){};
                      proto.pushState = function(){};
                      console.log('[polyfill] method2 ok');
                    }
                  } catch(e3){ console.log('[polyfill] method2 fail:', e3.message); }
                  // Method 3: intercept getter
                  try {
                    var desc = Object.getOwnPropertyDescriptor(window, 'history');
                    if (desc && desc.get) {
                      var orig = desc.get;
                      Object.defineProperty(window, 'history', {
                        get: function(){
                          var cur = orig.call(window);
                          if (cur) {
                            try { cur.replaceState = function(){}; } catch(e){}
                            try { cur.pushState = function(){}; } catch(e){}
                          }
                          return cur;
                        },
                        configurable: true
                      });
                      console.log('[polyfill] method3 ok');
                    }
                  } catch(e4){ console.log('[polyfill] method3 fail:', e4.message); }
                  // Method 4: Proxy fallback
                  try {
                    if (h && typeof h.replaceState !== 'function') {
                      var proxy = new Proxy(h, {
                        get: function(t, p){
                          if (p === 'replaceState' || p === 'pushState') return function(){};
                          return t[p];
                        }
                      });
                      Object.defineProperty(window, 'history', {
                        get: function(){ return proxy; },
                        configurable: true
                      });
                      console.log('[polyfill] method4 ok');
                    }
                  } catch(e5){ console.log('[polyfill] method4 fail:', e5.message); }
                  console.log('[polyfill] end — history.replaceState is:', typeof window.history?.replaceState);
                } catch(e) { console.error('[polyfill] outer catch:', e); }
              })();
            `,
          }}
        />
      </head>
      <body className="min-h-full flex flex-col">{children}</body>
    </html>
  );
}
