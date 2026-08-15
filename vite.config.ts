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
    // Univer 自带 React 运行时，与主工程 React 去重，避免「两个 React 实例」导致 hooks/事件失效
    dedupe: ['react', 'react-dom'],
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
