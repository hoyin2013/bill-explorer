# BillExplorer 打包文档

本文档详细说明 BillExplorer（账单查找器）从源码到生成可分发安装包的完整流程。

---

## 1. 环境要求

| 项目 | 要求 |
|------|------|
| Node.js | >= 18（推荐 20 LTS） |
| npm | >= 9 |
| 操作系统 | macOS 12+ / Windows 10+ / Linux |

> **跨平台打包限制**：electron-builder 不支持交叉编译。macOS 版 DMG 必须在 macOS 上打包，Windows 版 NSIS/Portable 必须在 Windows 上打包，Linux 版 AppImage 必须在 Linux 上打包。
>
> **没有 Windows 电脑？** 用 GitHub Actions 免费在云端 Windows 环境打包，见 [第 9 章](#9-用-github-actions-在云端打包-windows-安装包)。

### macOS 额外准备

首次在 macOS 上运行 Electron 开发环境时，需要去除隔离属性并重新签名：

```bash
xattr -d com.apple.quarantine node_modules/electron/dist/Electron.app 2>/dev/null || true
xattr -d com.apple.provenance node_modules/electron/dist/Electron.app 2>/dev/null || true
codesign --sign - --force --deep node_modules/electron/dist/Electron.app
```

项目已提供一键脚本 `run-dev.sh`，会自动完成上述操作。

---

## 2. 项目构建架构

```
源码                         构建工具         产物目录
─────────────────────────────────────────────────────────
src/renderer/**/*.tsx  ──→  tsc + vite  ──→  dist/           （渲染进程：HTML/CSS/JS）
src/main/main.ts       ──→  esbuild      ──→  dist-electron/main.js
src/main/preload.ts    ──→  esbuild      ──→  dist-electron/preload.js
dist/ + dist-electron/ ──→  electron-builder ──→  release/    （安装包）
```

### 2.1 三套 TypeScript 配置

| 文件 | 用途 | include |
|------|------|---------|
| `tsconfig.json` | 渲染进程（React） | `src/renderer` |
| `tsconfig.electron.json` | 主进程（Electron） | `src/main` |
| `tsconfig.node.json` | Vite 配置本身 | `vite.config.ts` |

三者均设置 `noEmit: true`，类型检查由 `tsc` 负责，实际产物由 Vite / esbuild 生成。

### 2.2 Vite 配置要点

```ts
// vite.config.ts
base: './'  // 关键：相对路径，避免 file:// 协议下白屏
```

生产模式 Electron 通过 `file://` 加载 `dist/index.html`，如果 `base` 为默认的 `/`，资源路径会解析到文件系统根目录导致白屏。

### 2.3 主进程构建命令

```bash
# 打包 main.ts → dist-electron/main.js
esbuild src/main/main.ts \
  --bundle \
  --platform=node \
  --external:electron \
  --format=cjs \
  --outfile=dist-electron/main.js

# 打包 preload.ts → dist-electron/preload.js
esbuild src/main/preload.ts \
  --bundle \
  --platform=node \
  --external:electron \
  --format=cjs \
  --outfile=dist-electron/preload.js
```

关键参数说明：
- `--external:electron`：electron 模块由运行时提供，不打包进去
- `--format=cjs`：Electron 主进程使用 CommonJS
- `--platform=node`：目标为 Node.js 环境

---

## 3. npm scripts 速查

| 命令 | 说明 |
|------|------|
| `npm run dev` | 启动 Vite dev server（仅渲染进程，端口 5173） |
| `./run-dev.sh` | 一键启动开发模式（Vite + Electron，macOS 专用） |
| `npm run build:renderer` | `tsc` 类型检查 + `vite build` → `dist/` |
| `npm run build:main` | esbuild 打包主进程 → `dist-electron/main.js` |
| `npm run build:preload` | esbuild 打包预加载脚本 → `dist-electron/preload.js` |
| `npm run build:electron` | `build:main` + `build:preload` |
| `npm run build` | `build:renderer` + `build:electron`（全量构建） |
| `npm run electron:preview` | `build` + `electron .`（本地预览构建产物） |
| `npm run electron:build` | `build` + `electron-builder`（打包为安装包） |
| `npm run electron:build:win` | 仅打 Windows NSIS + Portable |
| `npm run electron:build:portable` | 仅打 Windows 便携版 |
| `npm run electron:build:mac` | 仅打 macOS DMG |

---

## 4. 打包完整步骤

### 步骤 1：安装依赖

```bash
cd bill-explorer
npm install
```

### 步骤 2：全量构建源码

```bash
npm run build
```

执行后产物：
- `dist/index.html` — Vite 打包的渲染进程入口
- `dist/assets/*.js` / `*.css` — 渲染进程资源
- `dist-electron/main.js` — 主进程
- `dist-electron/preload.js` — 预加载脚本

### 步骤 3（可选）：本地预览构建产物

```bash
npm run electron:preview
```

验证构建产物在 Electron 中能否正常运行。此时 Electron 加载 `dist/index.html`（file:// 协议），不走 Vite dev server。

### 步骤 4：打包为安装包

#### macOS

```bash
npm run electron:build:mac
```

产物：
```
release/
├── BillExplorer-1.0.0-arm64.dmg        # macOS 安装镜像（Apple Silicon）
├── BillExplorer-1.0.0-arm64.dmg.blockmap  # 增量更新用
└── builder-debug.yml                    # 构建调试信息
```

> 如需 Intel 芯片版本，需在 Intel Mac 上执行相同命令，产物文件名含 `x64`。

#### Windows

```bash
npm run electron:build:win
```

产物：
```
release/
├── BillExplorer Setup 1.0.0.exe    # NSIS 安装包
└── BillExplorer 1.0.0.exe          # 便携版（Portable）
```

> 如仅需便携版：`npm run electron:build:portable`

#### Linux

```bash
npm run electron:build
```

产物：
```
release/
└── BillExplorer-1.0.0.AppImage    # Linux AppImage
```

---

## 5. electron-builder 配置详解

`package.json` 中的 `build` 字段：

```json
{
  "build": {
    "appId": "com.bill-explorer.app",
    "productName": "BillExplorer",
    "directories": {
      "output": "release"
    },
    "files": [
      "dist/**/*",
      "dist-electron/**/*",
      "package.json"
    ],
    "win": {
      "target": ["nsis", "portable"]
    },
    "mac": {
      "target": "dmg"
    },
    "linux": {
      "target": "AppImage"
    }
  }
}
```

| 字段 | 说明 |
|------|------|
| `appId` | 应用唯一标识，用于注册表/系统关联 |
| `productName` | 安装包显示名称 |
| `directories.output` | 打包输出目录（`release/`） |
| `files` | 打入 asar 的文件白名单：渲染产物 + 主进程产物 + package.json |
| `win.target` | Windows 打两种格式：NSIS 安装包 + Portable 便携版 |
| `mac.target` | macOS 打 DMG 格式 |
| `linux.target` | Linux 打 AppImage 格式 |

### 打包文件清单

electron-builder 最终打入 `app.asar` 的内容包括：

- `dist/` — 渲染进程全部产物
- `dist-electron/main.js` + `dist-electron/preload.js` — 主进程产物
- `package.json` — 只包含 `name`、`version`、`main`、`dependencies` 等必要字段
- `node_modules/` 中 `dependencies` 依赖（非 devDependencies）

以下文件会被自动排除：`release/`、源码 `src/`、`.git/`、测试文件、lock 文件等。

---

## 6. 常见问题排查

### 6.1 白屏（生产模式）

**现象**：Electron 启动后窗口空白，控制台报资源 404。

**原因**：Vite `base` 配置为 `/`（绝对路径），`file://` 协议下资源路径解析到文件系统根目录。

**解决**：确认 `vite.config.ts` 中 `base: './'`。

### 6.2 Electron 二进制下载失败

**现象**：`npm install electron` 超时或失败。

**解决**：

```bash
# 设置镜像
export ELECTRON_MIRROR=https://npmmirror.com/mirrors/electron/

# 或在项目根目录 .npmrc 中添加
# electron_mirror=https://npmmirror.com/mirrors/electron/
```

然后重新安装：

```bash
rm -rf node_modules/electron
npm install electron@28.3.3 --save-dev
```

如果二进制仍未下载，手动触发：

```bash
node node_modules/electron/install.js
```

### 6.3 macOS 无法打开 Electron（损坏提示）

**原因**：Gatekeeper 隔离属性未清除。

**解决**：

```bash
xattr -d com.apple.quarantine node_modules/electron/dist/Electron.app 2>/dev/null || true
xattr -d com.apple.provenance node_modules/electron/dist/Electron.app 2>/dev/null || true
codesign --sign - --force --deep node_modules/electron/dist/Electron.app
```

### 6.4 打出的 DMG/EXE 体积过大

**原因**：`node_modules` 中包含了不必要的 devDependencies。

**解决**：确认 `package.json` 的 `dependencies` 只包含运行时依赖：

```json
"dependencies": {
  "electron-store": "^8.2.0",
  "exceljs": "^4.4.0",
  "pinyin-pro": "^3.24.0",
  "react": "^18.3.1",
  "react-dom": "^18.3.1"
}
```

`electron`、`esbuild`、`vite`、`typescript` 等构建工具必须在 `devDependencies` 中。

### 6.5 打包时 electron-builder 报 "cannot execute cause=exit status 1"

**原因**：常见于 macOS 上打 Windows 包时缺少 Wine，或 Windows 上打 DMG 时缺少相关工具。

**解决**：在目标平台原生环境中打包，不要交叉编译。

---

## 7. 版本管理

更新版本号时，修改 `package.json` 中的 `version` 字段：

```json
{
  "version": "1.1.0"
}
```

electron-builder 会自动将版本号写入安装包文件名，例如 `BillExplorer-1.1.0-arm64.dmg`。

---

## 8. 产物分发

| 平台 | 文件 | 使用方式 |
|------|------|----------|
| macOS | `BillExplorer-x.x.x-arm64.dmg` | 双击挂载，拖入 Applications |
| Windows (NSIS) | `BillExplorer Setup x.x.x.exe` | 双击安装，开始菜单启动 |
| Windows (Portable) | `BillExplorer x.x.x.exe` | 单文件运行，无需安装 |
| Linux | `BillExplorer-x.x.x.AppImage` | `chmod +x` 后直接运行 |

---

## 9. 用 GitHub Actions 在云端打包 Windows 安装包

没有 Windows 电脑时，可以用 GitHub Actions 免费借用微软的 Windows 云服务器完成打包。项目已内置工作流文件 `.github/workflows/build-windows.yml`。

### 9.1 工作流做了什么

```
push tag / 手动触发
      │
      ▼
windows-latest 云服务器（免费，微软提供）
      │
      ├─ checkout 拉取源码
      ├─ 安装 Node 20 + npm ci 安装依赖
      ├─ npm run electron:build:win   ← 与本地 Windows 完全相同的命令
      │     └─ 产出 release/*.exe（NSIS 安装包 + Portable 便携版）
      ├─ 上传为 Actions artifact（可下载）
      └─ 若由 tag 触发 → 自动发布到 GitHub Release
```

### 9.2 前提：把项目推到 GitHub

当前仓库 remote 指向 Gitee，GitHub Actions 只在 GitHub 上运行。需要先把代码推到 GitHub：

```bash
# 1. 在 github.com 上新建一个仓库（Public/Private 均可），不要勾选初始化 README
# 2. 在项目根目录执行：
git remote add github https://github.com/<你的用户名>/<仓库名>.git
git push -u github main
```

推送后 `.github/workflows/build-windows.yml` 会一并上传（首次推送后 Actions 页面会显示该工作流）。

### 9.3 手动触发打包

1. 打开仓库页面 → **Actions** 标签
2. 左侧选择 **Build Windows Installer**
3. 点击 **Run workflow** → 绿色按钮确认
4. 等待约 5~10 分钟，黄色圆点 = 运行中，绿色对勾 = 成功

### 9.4 下载安装包

两种方式任选：

**方式 A：Actions artifact（每次构建都有）**

构建成功后，进入该次运行的页面，底部 **Artifacts** 区域点击 **BillExplorer-windows** 下载 zip，解压得到 `.exe` 文件。

**方式 B：GitHub Release（打 tag 时自动发布）**

```bash
# 打 tag 并推送，工作流会自动触发并创建 Release
git tag v1.0.0
git push github v1.0.0
```

构建成功后，仓库 **Releases** 页面会出现 `v1.0.0`，下方附带了安装包，可让别人直接下载。

### 9.5 注意事项

| 事项 | 说明 |
|------|------|
| 免费额度 | 公开仓库无限时长；私有仓库免费版每月 2000 分钟 |
| 代码签名 | 未配置证书，Windows 首次运行会显示"未知发布者"提示，点"更多信息 → 仍要运行"即可 |
| 版本号 | 打包版本取自 `package.json` 的 `version`，改版号后重新触发即可 |
| package-lock.json | `npm ci` 依赖它，必须随代码一起提交（已在仓库中） |
| 触发方式 | `workflow_dispatch`（手动）+ `push tags v*`（自动）双支持 |
