import { NextRequest, NextResponse } from 'next/server';
import { prisma } from '@/lib/db/client';
import { requireSession } from '@/lib/auth/session';
import { consumeCallbackToken } from '@/lib/auth/callback-tokens';

export async function GET(req: NextRequest) {
  let userId: string | null = null;

  // 1) Try cookie-based session first
  try {
    const session = await requireSession(req);
    userId = session.userId;
  } catch {
    // no cookie session
  }

  // 2) Fall back to one-time callback token (Outlook desktop iframe/popup isolation)
  if (!userId) {
    const token = req.nextUrl.searchParams.get('callbackToken');
    if (token) userId = consumeCallbackToken(token);
  }

  if (!userId) {
    return NextResponse.json({
      microsoft: { connected: false },
      monday: { connected: false },
    });
  }

  const [user, ms, monday] = await Promise.all([
    prisma.user.findUnique({ where: { id: userId } }),
    prisma.msAccount.findUnique({ where: { userId } }),
    prisma.mondayAccount.findUnique({ where: { userId } }),
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
