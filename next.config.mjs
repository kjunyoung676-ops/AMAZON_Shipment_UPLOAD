/** @type {import('next').NextConfig} */
const nextConfig = {
  typescript: {
    ignoreBuildErrors: true,
  },
  images: {
    unoptimized: true,
  },
  webpack: (config) => {
    config.resolve.alias['canvas'] = false
    return config
  },
  experimental: {
    turbo: {
      resolveAlias: {
        canvas: './canvas-mock.js',
      },
    },
  },
}
export default nextConfig
