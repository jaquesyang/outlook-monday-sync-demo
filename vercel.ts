import { type VercelConfig } from '@vercel/config/v1';

export const config: VercelConfig = {
  framework: 'nextjs',
  buildCommand: 'prisma generate && (prisma migrate deploy || echo "migrate skipped — run manually on Supabase") && ([ -f .env ] || touch .env) && next build --turbopack',
  crons: [{ path: '/api/cron/refresh', schedule: '0 7 * * *' }],
};
