import { createRemoteJWKSet, jwtVerify, type JWTPayload } from 'jose';

const JWKS = createRemoteJWKSet(
  new URL('https://login.microsoftonline.com/common/discovery/v2.0/keys'),
);

export type MsIdentity = { tenantId: string; userId: string; email: string };

export async function verifyOfficeSsoToken(token: string): Promise<MsIdentity> {
  const audience = `api://${new URL(process.env.APP_BASE_URL!).host}/${process.env.MS_GRAPH_CLIENT_ID}`;
  const { payload } = await jwtVerify(token, JWKS, {
    audience,
    issuer: /^https:\/\/login\.microsoftonline\.com\/[^/]+\/v2\.0$/ as unknown as string,
  });
  return extractIdentity(payload);
}

export function extractIdentity(p: JWTPayload & {
  tid?: string;
  oid?: string;
  preferred_username?: string;
}): MsIdentity {
  if (!p.tid || !p.oid || !p.preferred_username) {
    throw new Error('SSO token missing required claims');
  }
  return {
    tenantId: p.tid,
    userId: p.oid,
    email: p.preferred_username,
  };
}
