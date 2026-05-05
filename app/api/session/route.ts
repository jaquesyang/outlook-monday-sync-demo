import { NextRequest, NextResponse } from 'next/server';
import { verifyOfficeSsoToken } from '@/lib/auth/office-sso';
import { prisma } from '@/lib/db/client';
import { setSessionCookie } from '@/lib/auth/session';

export async function POST(req: NextRequest) {
  const { idToken } = (await req.json()) as { idToken?: string };
  if (!idToken) return NextResponse.json({ error: 'missing idToken' }, { status: 400 });

  const id = await verifyOfficeSsoToken(idToken);
  const user = await prisma.user.upsert({
    where: { msTenantId_msUserId: { msTenantId: id.tenantId, msUserId: id.userId } },
    update: { msUserEmail: id.email },
    create: {
      msTenantId: id.tenantId,
      msUserId: id.userId,
      msUserEmail: id.email,
    },
  });

  const res = NextResponse.json({ ok: true, userId: user.id, email: user.msUserEmail });
  setSessionCookie(res, user.id);
  return res;
}
