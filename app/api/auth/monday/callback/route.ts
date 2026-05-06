import { NextRequest, NextResponse } from 'next/server';
import { decodeState } from '@/lib/auth/oauth-state';
import { exchangeCodeForToken, fetchMondayMe } from '@/lib/monday/oauth';
import { encryptToken } from '@/lib/crypto/token';
import { prisma } from '@/lib/db/client';
import { setCallbackToken } from '@/lib/auth/callback-tokens';
import { getAppBaseUrl } from '@/lib/env';

export async function GET(req: NextRequest) {
  const code = req.nextUrl.searchParams.get('code');
  const state = req.nextUrl.searchParams.get('state');
  if (!code || !state) return NextResponse.json({ error: 'missing code/state' }, { status: 400 });

  const { userId, purpose, callbackToken } = decodeState(state);
  if (purpose !== 'monday') return NextResponse.json({ error: 'wrong purpose' }, { status: 400 });

  const redirectUri = `${getAppBaseUrl()}/api/auth/monday/callback`;
  const tok = await exchangeCodeForToken({ code, redirectUri });
  const me = await fetchMondayMe(tok.access_token);

  // The user must already exist (Microsoft connect happens first).
  const user = await prisma.user.findUnique({ where: { id: userId } });
  if (!user) return NextResponse.json({ error: 'connect Microsoft first' }, { status: 412 });

  const refreshEnc = tok.refresh_token
    ? encryptToken(tok.refresh_token)
    : null;
  const expiresAt =
    typeof tok.expires_in === 'number' && !isNaN(tok.expires_in)
      ? new Date(Date.now() + tok.expires_in * 1000)
      : null;

  await prisma.mondayAccount.upsert({
    where: { userId },
    update: {
      mondayUserId: me.mondayUserId,
      mondayAccountId: me.mondayAccountId,
      accessTokenEnc: encryptToken(tok.access_token),
      refreshTokenEnc: refreshEnc,
      expiresAt,
    },
    create: {
      userId,
      mondayUserId: me.mondayUserId,
      mondayAccountId: me.mondayAccountId,
      accessTokenEnc: encryptToken(tok.access_token),
      refreshTokenEnc: refreshEnc,
      expiresAt,
    },
  });

  if (callbackToken) setCallbackToken(callbackToken, user.id);

  return new NextResponse(
    `<!doctype html><meta charset="utf-8"><title>Connected</title>
     <body style="font-family:system-ui;padding:2rem">
       <h2>monday connected ✓</h2>
       <p>You can close this window.</p>
       <script>
         if (window.opener) {
           window.opener.postMessage({ type: 'oauth-success', provider: 'monday' }, '*');
         }
         window.close();
       </script>
     </body>`,
    { headers: { 'content-type': 'text/html; charset=utf-8' } },
  );
}
