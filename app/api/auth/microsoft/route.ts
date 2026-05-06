import { NextRequest, NextResponse } from 'next/server';
import { encodeState } from '@/lib/auth/oauth-state';
import { buildAuthorizeUrl } from '@/lib/ms/oauth';
import { getOrInitSession } from '@/lib/auth/session';
import { getAppBaseUrl } from '@/lib/env';

export async function GET(req: NextRequest) {
  const session = await getOrInitSession(req);
  const callbackToken = req.nextUrl.searchParams.get('callbackToken') ?? undefined;
  const redirectUri = `${getAppBaseUrl()}/api/auth/microsoft/callback`;
  const state = encodeState({ userId: session.userId, purpose: 'ms', callbackToken });
  const url = buildAuthorizeUrl({ state, redirectUri });
  const res = NextResponse.redirect(url);
  session.applyTo(res);
  return res;
}
