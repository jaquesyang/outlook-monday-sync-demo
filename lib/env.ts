export function getAppBaseUrl(): string {
  if (process.env.VERCEL_ENV === "production") {
    return `https://${process.env.VERCEL_PROJECT_PRODUCTION_URL}`;
  }
  if (process.env.VERCEL_URL) {
    return `https://${process.env.VERCEL_URL}`;
  }
  if (process.env.APP_BASE_URL) {
    return process.env.APP_BASE_URL;
  }
  return "http://localhost:3000";
}
