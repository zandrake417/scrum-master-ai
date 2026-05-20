import { defineConfig } from 'vite';
import react from '@vitejs/plugin-react';
import dotenv from 'dotenv';

dotenv.config();

export default defineConfig({
  plugins: [react()],
  server: {
    proxy: {
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
      // 🟢 TAMBAHKAN PROXY TRELLO
      '/api-trello': {
        target: 'https://api.trello.com',
        changeOrigin: true,
        rewrite: path => path.replace(/^\/api-trello/, ''),
      },
      // 🟢 TAMBAHKAN PROXY JIRA (Ganti 'domain-anda' dengan subdomain Jira Anda)
      '/api-jira': {
        target: 'https://zandrake417.atlassian.net', 
        changeOrigin: true,
        rewrite: path => path.replace(/^\/api-jira/, ''),
      },
    },
  },
});