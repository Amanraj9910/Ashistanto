import type { NextConfig } from 'next';

const backendUrl =
  process.env.BACKEND_URL?.replace(/\/$/, '') ||
  process.env.NEXT_PUBLIC_API_URL?.replace(/\/$/, '') ||
  'http://127.0.0.1:3000';

const nextConfig: NextConfig = {
  reactStrictMode: true,
  async rewrites() {
    return [
      {
        source: '/api/:path*',
        destination: `${backendUrl}/api/:path*`
      },
      {
        source: '/auth/login',
        destination: `${backendUrl}/auth/login`
      },
      {
        source: '/auth/redirect',
        destination: `${backendUrl}/auth/redirect`
      },
      {
        source: '/auth/callback',
        destination: `${backendUrl}/auth/callback`
      },
      {
        source: '/auth/logout',
        destination: `${backendUrl}/auth/logout`
      },
      {
        source: '/auth/user',
        destination: `${backendUrl}/auth/user`
      },
      {
        source: '/auth/session-token/:path*',
        destination: `${backendUrl}/auth/session-token/:path*`
      }
    ];
  }
};

export default nextConfig;
