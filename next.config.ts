import type { NextConfig } from "next";

const nextConfig: NextConfig = {
  /* config options here */
  reactStrictMode: true,
  api: {
    bodyParser: {
      sizeLimit: '50mb', // 기본값: 4.5MB → 10MB로 증가
    },
  },
};

export default nextConfig;
