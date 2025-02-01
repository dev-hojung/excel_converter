import type { NextConfig } from "next";

const nextConfig: NextConfig = {
  /* config options here */
  reactStrictMode: true,
  api: {
    bodyParser: false, // Next.js의 기본 bodyParser를 사용하지 않도록 설정
  },
};

export default nextConfig;
