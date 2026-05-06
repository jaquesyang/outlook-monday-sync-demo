const AUTH = 'https://auth.monday.com/oauth2/authorize';
const TOKEN = 'https://auth.monday.com/oauth2/token';

export const MONDAY_SCOPES = [
  'boards:read',
  'boards:write',
  'webhooks:read',
  'webhooks:write',
  'me:read',
  'users:read',
];

export function buildAuthorizeUrl(opts: { state: string; redirectUri: string }) {
  const u = new URL(AUTH);
  u.searchParams.set('client_id', process.env.MONDAY_CLIENT_ID!);
  u.searchParams.set('response_type', 'code');
  u.searchParams.set('redirect_uri', opts.redirectUri);
  u.searchParams.set('state', opts.state);
  u.searchParams.set('scope', MONDAY_SCOPES.join(' '));
  return u.toString();
}

export type MondayTokenResponse = {
  access_token: string;
  refresh_token?: string;
  token_type: string;
  expires_in: number;
  scope: string;
};

export async function exchangeCodeForToken(opts: {
  code: string;
  redirectUri: string;
}): Promise<MondayTokenResponse> {
  const body = new URLSearchParams({
    client_id: process.env.MONDAY_CLIENT_ID!,
    client_secret: process.env.MONDAY_CLIENT_SECRET!,
    grant_type: 'authorization_code',
    code: opts.code,
    redirect_uri: opts.redirectUri,
  });
  const r = await fetch(TOKEN, {
    method: 'POST',
    headers: { 'content-type': 'application/x-www-form-urlencoded' },
    body,
  });
  const rawText = await r.text();
  if (!r.ok) throw new Error(`monday token exchange failed: ${r.status} ${rawText}`);
  const payload = JSON.parse(rawText) as MondayTokenResponse;
  if (!payload.access_token) {
    throw new Error(`monday token response missing access_token: ${rawText}`);
  }
  return payload;
}

/** Query the `me` endpoint to learn the monday user/account ids. */
export async function fetchMondayMe(accessToken: string) {
  const r = await fetch('https://api.monday.com/v2', {
    method: 'POST',
    headers: {
      'content-type': 'application/json',
      authorization: accessToken,
    },
    body: JSON.stringify({
      query: '{ me { id name email account { id } } }',
    }),
  });
  if (!r.ok) throw new Error(`monday me query failed: ${r.status}`);
  const j = (await r.json()) as {
    data: { me: { id: string; name: string; email: string; account: { id: string } } };
  };
  return {
    mondayUserId: BigInt(j.data.me.id),
    mondayAccountId: BigInt(j.data.me.account.id),
    name: j.data.me.name,
    email: j.data.me.email,
  };
}
