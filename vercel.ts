import { type VercelConfig } from '@vercel/config/v1';

export const config: VercelConfig = {
  framework: 'nextjs',
  buildCommand: '[ -f .env ] || touch .env && next build --turbopack',
  crons: [{ path: '/api/cron/refresh', schedule: '0 7 * * *' }],
};
