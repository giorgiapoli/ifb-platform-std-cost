import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'

export default defineConfig({
  plugins: [react()],
  base: '/ifb-platform-std-cost/',
  build: {
    outDir: 'docs',
    emptyOutDir: true,
    copyPublicDir: true,
  }
})