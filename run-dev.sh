#!/bin/zsh
# run-dev.sh - 一键启动账单查找器开发模式
# 用法: cd ~/Desktop/personal/git/xiujuan/bill-explorer && ./run-dev.sh

set -e

# 防止环境里残留 ELECTRON_RUN_AS_NODE=1（会令 Electron 以纯 Node 模式运行、无法显示窗口）
export ELECTRON_RUN_AS_NODE=

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
cd "$SCRIPT_DIR"

echo "📦 安装 Electron（从官方 npm）..."
rm -rf node_modules/electron
npm install electron@28.3.3 --save-dev --os=darwin --cpu=arm64 2>&1 | tail -1

BIN="node_modules/electron/dist/Electron.app/Contents/MacOS/Electron"
if [ ! -f "$BIN" ]; then
  echo "❌ Electron 二进制下载失败，尝试重新下载..."
  node node_modules/electron/install.js
fi

echo "🔏 签名 Electron（绕过 Gatekeeper）..."
xattr -d com.apple.quarantine node_modules/electron/dist/Electron.app 2>/dev/null || true
xattr -d com.apple.provenance node_modules/electron/dist/Electron.app 2>/dev/null || true
codesign --sign - --force --deep node_modules/electron/dist/Electron.app 2>&1 | tail -1

echo "🚀 启动 Vite dev server..."
npx vite --host 127.0.0.1 --port 5173 &
VITE_PID=$!

sleep 4
if ! kill -0 $VITE_PID 2>/dev/null; then
  echo "❌ Vite 启动失败"
  exit 1
fi

echo "🚀 启动 Electron..."
VITE_DEV_SERVER_URL=http://localhost:5173 "$BIN" . &
ELEC_PID=$!

echo ""
echo "✅ 已启动！"
echo "   Vite PID:     $VITE_PID (http://127.0.0.1:5173/)"
echo "   Electron PID: $ELEC_PID"
echo ""
echo "按 Ctrl+C 停止。"
wait $ELEC_PID $VITE_PID