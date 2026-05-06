const AUTH = 'https://login.microsoftonline.com/common/oauth2/v2.0/authorize';
const TOKEN = 'https://login.microsoftonline.com/common/oauth2/v2.0/token';

export const MS_SCOPES = [
  'openid',
  'profile',
  'offline_access',
  'Calendars.ReadWrite',
  'User.Read',
];

export function buildAuthorizeUrl(opts: { state: string; redirectUri: string }) {
  const u = new URL(AUTH);
  u.searchParams.set('client_id', process.env.MS_GRAPH_CLIENT_ID!);
  u.searchParams.set('response_type', 'code');
  u.searchParams.set('response_mode', 'query');
  u.searchParams.set('redirect_uri', opts.redirectUri);
  u.searchParams.set('scope', MS_SCOPES.join(' '));
  u.searchParams.set('state', opts.state);
  u.searchParams.set('prompt', 'select_account');
  return u.toString();
}

export type MsTokenResponse = {
  access_token: string;
  refresh_token: string;
  expires_in: number;
  id_token: string;
};

export async function exchangeCodeForToken(opts: {
  code: string;
  redirectUri: string;
}): Promise<MsTokenResponse> {
  const body = new URLSearchParams({
    client_id: process.env.MS_GRAPH_CLIENT_ID!,
    client_secret: process.env.MS_GRAPH_CLIENT_SECRET!,
    grant_type: 'authorization_code',
    code: opts.code,
    redirect_uri: opts.redirectUri,
    scope: MS_SCOPES.join(' '),
  });
  const r = await fetch(TOKEN, {
    method: 'POST',
    headers: { 'content-type': 'application/x-www-form-urlencoded' },
    body,
  });
  if (!r.ok) throw new Error(`MS token exchange failed: ${r.status} ${await r.text()}`);
  return (await r.json()) as MsTokenResponse;
}

export async function refreshAccessToken(refreshToken: string): Promise<MsTokenResponse> {
  const body = new URLSearchParams({
    client_id: process.env.MS_GRAPH_CLIENT_ID!,
    client_secret: process.env.MS_GRAPH_CLIENT_SECRET!,
    grant_type: 'refresh_token',
    refresh_token: refreshToken,
    scope: MS_SCOPES.join(' '),
  });
  const r = await fetch(TOKEN, {
    method: 'POST',
    headers: { 'content-type': 'application/x-www-form-urlencoded' },
    body,
  });
  if (!r.ok) throw new Error(`MS token refresh failed: ${r.status} ${await r.text()}`);
  return (await r.json()) as MsTokenResponse;
}
