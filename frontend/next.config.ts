import type { NextConfig } from 'next';

// The UI is exported to static HTML and served by the Express app in ../server.js on the
// SAME ORIGIN as the API. That makes every /api/* and /auth/* fetch a plain relative request,
// so the dev-only rewrites() layer this file used to carry is unnecessary (and is ignored
// under output: 'export' anyway).
const nextConfig: NextConfig = {
  reactStrictMode: true,
  output: 'export',
  trailingSlash: true,
  images: { unoptimized: true },
  outputFileTracingRoot: __dirname
};

export default nextConfig;
