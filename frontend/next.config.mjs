/** @type {import('next').NextConfig} */
const isGhPages = process.env.DEPLOY_TARGET === "gh-pages";
const basePath = isGhPages ? "/ptit-oil-price-tracker" : "";

const nextConfig = {
  reactStrictMode: true,
  output: "export",
  images: {
    unoptimized: true,
  },
  env: {
    NEXT_PUBLIC_BASE_PATH: basePath,
  },
  basePath,
  assetPrefix: isGhPages ? "/ptit-oil-price-tracker/" : "",
  trailingSlash: true,
};

export default nextConfig;
