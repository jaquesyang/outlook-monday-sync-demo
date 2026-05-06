import type { NextConfig } from "next";

const nextConfig: NextConfig = {
  allowedDevOrigins: ['puerto-station-birth-sword.trycloudflare.com'],
  webpack: (config) => {
    config.watchOptions = {
      ...config.watchOptions,
      poll: 1000,
      aggregateTimeout: 300,
    };
    return config;
  },
};

export default nextConfig;
