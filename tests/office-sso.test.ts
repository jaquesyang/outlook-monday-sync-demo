import { describe, it, expect } from 'vitest';
import { extractIdentity } from '@/lib/auth/office-sso';

describe('office-sso extractIdentity', () => {
  it('maps tid + oid + preferred_username to MsIdentity', () => {
    const claims = {
      tid: 'tenant-xyz',
      oid: 'user-abc',
      preferred_username: 'alice@contoso.com',
      aud: 'api://localhost:3000/<app-id>',
      iss: 'https://login.microsoftonline.com/<tid>/v2.0',
      exp: Math.floor(Date.now() / 1000) + 3600,
    };
    expect(extractIdentity(claims)).toEqual({
      tenantId: 'tenant-xyz',
      userId: 'user-abc',
      email: 'alice@contoso.com',
    });
  });

  it('throws if any required claim is missing', () => {
    expect(() => extractIdentity({ tid: 'x' } as never)).toThrow();
    expect(() => extractIdentity({ tid: 'x', oid: 'y' } as never)).toThrow();
  });
});
