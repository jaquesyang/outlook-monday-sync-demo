import { NextRequest, NextResponse } from 'next/server';
import { encryptToken, decryptToken } from '@/lib/crypto/token';
import { randomUUID } from 'node:crypto';

const COOKIE = 'oms.sid';
const MAX_AGE = 60 * 60 * 24 * 30; // 30 days

export type Session = {
  userId: string;
  /** call before responding to set/refresh the cookie */
  applyTo(res: NextResponse): void;
};

function read(req: NextRequest): string | null {
  const blob = req.cookies.get(COOKIE)?.value;
  if (!blob) return null;
  try {
    const json = decryptToken(Buffer.from(blob, 'base64url'));
    const parsed = JSON.parse(json) as { userId: string };
    return parsed.userId ?? null;
  } catch {
    return null;
  }
}

/** Write a fresh session cookie for the given userId. */
export function setSessionCookie(res: NextResponse, userId: string) {
  const blob = Buffer.from(encryptToken(JSON.stringify({ userId }))).toString('base64url');
  res.cookies.set(COOKIE, blob, {
    httpOnly: true,
    secure: true,
    sameSite: 'none',
    path: '/',
    maxAge: MAX_AGE,
  });
}

export async function getOrInitSession(req: NextRequest): Promise<Session> {
  const existing = read(req);
  const userId = existing ?? randomUUID();
  return {
    userId,
    applyTo(res) {
      setSessionCookie(res, userId);
    },
  };
}

export async function requireSession(req: NextRequest): Promise<{ userId: string }> {
  const id = read(req);
  if (!id) throw new Response('unauthenticated', { status: 401 });
  return { userId: id };
}
