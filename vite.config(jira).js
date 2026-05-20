import { defineConfig } from 'vite';
import react from '@vitejs/plugin-react';
import dotenv from 'dotenv';

dotenv.config();

export default defineConfig({
  plugins: [react()],
  server: {
    proxy: {
      // 1. Proxy Snifox (Sudah ada, biarkan saja)
      '/api': {
        target: 'https://core.snifoxai.com',
        changeOrigin: true,
        rewrite: path => path.replace(/^\/api/, '/v1'),
        configure: proxy => {
          proxy.on('proxyReq', proxyReq => {
            const key = process.env.SNIFOX_API_KEY;
            if (key) {
              proxyReq.setHeader('Authorization', `Bearer ${key}`);
            }
          });
        },
      },
      // 2. TAMBAHKAN PROXY JIRA DI SINI
      '/jira-api': {
        target: 'https://zandrake417.atlassian.net', // Domain Jira Anda
        changeOrigin: true,
        rewrite: path => path.replace(/^\/jira-api/, '')
      }
    },
  },
});