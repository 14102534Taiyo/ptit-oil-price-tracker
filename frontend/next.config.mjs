/** @type {import('next').NextConfig} */
const isGhPages = process.env.DEPLOY_TARGET === 'gh-pages';
const repoName = process.env.GITHUB_REPOSITORY ? process.env.GITHUB_REPOSITORY.split('/')[1] : 'ptit-oil-price-tracker';

const nextConfig = {
  output: 'export',
  images: {
    unoptimized: true,
  },
  basePath: isGhPages ? `/${repoName}` : '',
  assetPrefix: isGhPages ? `/${repoName}/` : '',
};

export default nextConfig;
