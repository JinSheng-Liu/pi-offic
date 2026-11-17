import { defineConfig } from 'vite'
import vue from '@vitejs/plugin-vue'
import tailwindcss from '@tailwindcss/vite'
import { fileURLToPath, URL } from 'node:url'
import { createApiMiddleware } from './server/api.mjs'

export default defineConfig({
  resolve: { alias: { '@': fileURLToPath(new URL('./src', import.meta.url)) } },
  server: {
    watch: { ignored: ['**/data/**', '**/output/**', '**/tmp/**'] },
  },
  plugins: [vue(), tailwindcss(), {
    name: 'pi-docx-local-api',
    configureServer(server) {
      server.middlewares.use(createApiMiddleware())
    },
  }],
  optimizeDeps: {
    esbuildOptions: {
      target: 'esnext',
    },
  },
  build: {
    target: 'esnext',
  },
})
