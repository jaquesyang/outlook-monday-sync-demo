import { PrismaClient } from '@prisma/client';
import { PrismaPg } from '@prisma/adapter-pg';
import { Pool } from 'pg';

declare global {
  // eslint-disable-next-line no-var
  var __prisma: PrismaClient | undefined;
}

function createPrisma() {
  const pool = new Pool({ connectionString: process.env.DATABASE_URL });
  const adapter = new PrismaPg(pool);
  return new PrismaClient({ adapter, log: ['error', 'warn'] });
}

export const prisma = global.__prisma ?? createPrisma();

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
