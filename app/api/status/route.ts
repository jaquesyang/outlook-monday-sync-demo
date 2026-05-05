import { NextRequest, NextResponse } from 'next/server';
import { prisma } from '@/lib/db/client';
import { requireSession } from '@/lib/auth/session';

export async function GET(req: NextRequest) {
  let session;
  try {
    session = await requireSession(req);
  } catch {
    return NextResponse.json({
      microsoft: { connected: false },
      monday: { connected: false },
    });
  }

  const [user, ms, monday] = await Promise.all([
    prisma.user.findUnique({ where: { id: session.userId } }),
    prisma.msAccount.findUnique({ where: { userId: session.userId } }),
    prisma.mondayAccount.findUnique({ where: { userId: session.userId } }),
  ]);

  return NextResponse.json({
    microsoft: {
      connected: Boolean(ms),
      email: user?.msUserEmail,
    },
    monday: {
      connected: Boolean(monday),
      account: monday ? `account ${monday.mondayAccountId.toString()}` : undefined,
    },
  });
}
