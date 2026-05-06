import { PrismaClient } from '@prisma/client';

declare global {
  var __prisma: PrismaClient | undefined;
}

// Vercel Supabase integration provides POSTGRES_PRISMA_URL with ?pgbouncer=true,
// making it safe to use the connection pooler at runtime.
// Fallback chain: Prisma URL → non-pooling direct URL → pooled URL.
const dbUrl =
  process.env.POSTGRES_PRISMA_URL ||
  process.env.POSTGRES_URL_NON_POOLING ||
  process.env.POSTGRES_URL;

export const prisma =
  global.__prisma ??
  new PrismaClient({
    datasources: { db: { url: dbUrl } },
    log: ['error', 'warn'],
  });

if (process.env.NODE_ENV !== 'production') global.__prisma = prisma;

/** Per-user query helper. Always pass through this when reading user data. */
export function forUser(userId: string) {
  return {
    msAccount: () => prisma.msAccount.findUnique({ where: { userId } }),
    mondayAccount: () => prisma.mondayAccount.findUnique({ where: { userId } }),
    boards: () => prisma.boardSubscription.findMany({ where: { userId } }),
    mappings: () => prisma.eventMapping.findMany({ where: { userId, deletedAt: null } }),
    log: (limit = 50) =>
      prisma.syncLog.findMany({
        where: { userId },
        orderBy: { occurredAt: 'desc' },
        take: limit,
      }),
  };
}
