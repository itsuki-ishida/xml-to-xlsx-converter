import type { NextConfig } from "next";

const nextConfig: NextConfig = {
  output: "export",
  transpilePackages: ["xlsx-js-style"],
};

export default nextConfig;
