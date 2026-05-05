import { NextRequest, NextResponse } from 'next/server';
import { encodeState } from '@/lib/auth/oauth-state';
import { buildAuthorizeUrl } from '@/lib/monday/oauth';
import { getOrInitSession } from '@/lib/auth/session';

export async function GET(req: NextRequest) {
  const session = await getOrInitSession(req);
  const redirectUri = `${process.env.APP_BASE_URL}/api/auth/monday/callback`;
  const state = encodeState({ userId: session.userId, purpose: 'monday' });
  const url = buildAuthorizeUrl({ state, redirectUri });
  const res = NextResponse.redirect(url);
  session.applyTo(res);
  return res;
}
