import type { NextConfig } from "next";

const nextConfig: NextConfig = {
  // Статическая сборка для Yandex Object Storage
  output: "export",
  // Отключаем оптимизацию изображений (не работает в статике)
  images: {
    unoptimized: true,
  },
  // Для статического хостинга — все маршруты как index.html
  trailingSlash: true,
  typescript: {
    ignoreBuildErrors: true,
  },
  reactStrictMode: false,
};

export default nextConfig;
