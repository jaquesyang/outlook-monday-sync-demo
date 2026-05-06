import { PrismaClient } from '@prisma/client';

declare global {
  var __prisma: PrismaClient | undefined;
}

// Use non-pooling direct URL at runtime to avoid PgBouncer prepared-statement errors (42P05).
// POSTGRES_URL (pooled) is kept in schema for Prisma Migrate / Introspect via directUrl.
const dbUrl = process.env.POSTGRES_URL_NON_POOLING || process.env.POSTGRES_URL;

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
