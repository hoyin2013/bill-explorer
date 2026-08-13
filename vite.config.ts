import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'
import { resolve } from 'path'

export default defineConfig({
  plugins: [react()],
  // 关键：使用相对路径，否则 file:// 协议下 /assets/... 会被解析到文件系统根目录，找不到文件导致白屏
  base: './',
  resolve: {
    alias: {
      '@': resolve(__dirname, './src/renderer'),
    },
  },
  server: {
    port: 5173,
  },
  build: {
    outDir: 'dist',
    rollupOptions: {
      // 单页构建：主窗口（小票识图已内嵌为右侧面板）
      input: {
        main: resolve(__dirname, 'index.html'),
      },
    },
  },
})
