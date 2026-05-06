import { NextRequest, NextResponse } from 'next/server';
import { decodeState } from '@/lib/auth/oauth-state';
import { exchangeCodeForToken } from '@/lib/ms/oauth';
import { encryptToken } from '@/lib/crypto/token';
import { prisma } from '@/lib/db/client';
import { setSessionCookie } from '@/lib/auth/session';
import { setCallbackToken } from '@/lib/auth/callback-tokens';
import { decodeJwt } from 'jose';

export async function GET(req: NextRequest) {
  const code = req.nextUrl.searchParams.get('code');
  const state = req.nextUrl.searchParams.get('state');
  if (!code || !state) return NextResponse.json({ error: 'missing code/state' }, { status: 400 });

  const { userId, purpose, callbackToken } = decodeState(state);
  if (purpose !== 'ms') return NextResponse.json({ error: 'wrong purpose' }, { status: 400 });

  const redirectUri = `${process.env.APP_BASE_URL}/api/auth/microsoft/callback`;
  const tok = await exchangeCodeForToken({ code, redirectUri });

  // ID token gives us tenantId/userId/email without an extra Graph call.
  const idClaims = decodeJwt(tok.id_token) as {
    tid: string; oid: string; preferred_username: string;
  };

  const user = await prisma.user.upsert({
    where: { msTenantId_msUserId: { msTenantId: idClaims.tid, msUserId: idClaims.oid } },
    update: { msUserEmail: idClaims.preferred_username },
    create: {
      id: userId,
      msTenantId: idClaims.tid,
      msUserId: idClaims.oid,
      msUserEmail: idClaims.preferred_username,
    },
  });

  await prisma.msAccount.upsert({
    where: { userId: user.id },
    update: {
      accessTokenEnc: encryptToken(tok.access_token),
      refreshTokenEnc: encryptToken(tok.refresh_token),
      expiresAt: new Date(Date.now() + tok.expires_in * 1000),
    },
    create: {
      userId: user.id,
      accessTokenEnc: encryptToken(tok.access_token),
      refreshTokenEnc: encryptToken(tok.refresh_token),
      expiresAt: new Date(Date.now() + tok.expires_in * 1000),
    },
  });

  if (callbackToken) setCallbackToken(callbackToken, user.id);

  const res = new NextResponse(
    `<!doctype html><meta charset="utf-8"><title>Connected</title>
     <body style="font-family:system-ui;padding:2rem">
       <h2>Microsoft connected ✓</h2>
       <p>You can close this window.</p>
       <script>
         if (window.opener) {
           window.opener.postMessage({ type: 'oauth-success', provider: 'microsoft' }, '*');
         }
         window.close();
       </script>
     </body>`,
    { headers: { 'content-type': 'text/html; charset=utf-8' } },
  );
  // Bind the cookie to the persisted user (in case SSO was unavailable
  // and the cookie was carrying a fresh random uuid that doesn't match
  // the upserted record's id).
  setSessionCookie(res, user.id);
  return res;
}
