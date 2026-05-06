import { type VercelConfig } from '@vercel/config/v1';

export const config: VercelConfig = {
  framework: 'nextjs',
  buildCommand: 'prisma generate && prisma migrate deploy && ([ -f .env ] || touch .env) && next build --turbopack',
  crons: [{ path: '/api/cron/refresh', schedule: '0 7 * * *' }],
};
