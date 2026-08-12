# 账单录入器 (BillExplorer)

基于 **Electron + React + TypeScript + Vite** 的本地桌面工具，用于快速定位、打开本地大量账单 Excel 文件。纯本地运行，无云端请求，不读取 Excel 表格内部单元格内容，只操作文件系统。

## 功能

- **选择文件夹** → 递归扫描目录下全部 `.xlsx` / `.xls` 文件
- **实时搜索**（输入即过滤，无需搜索按钮）：
  - 原始文件名模糊检索（数字、字母、中文，大小写不敏感）
  - 中文拼音首字母缩写（输入 `wm` 匹配"王梅"）
  - 全拼音检索（输入 `wangmei` 匹配"王梅"）
  - 路径片段匹配
- **单击**列表条目，调用系统默认程序打开 Excel
- 工作目录持久保存，软件重启自动加载
- 生僻字拼音解析异常容错，不崩溃

## 项目结构

```
bill-explorer/
├── index.html                 # Vite 入口 HTML
├── package.json
├── vite.config.ts             # Vite 配置
├── tsconfig.json              # 渲染进程 TS 配置
├── tsconfig.node.json         # vite.config.ts 用
├── tsconfig.electron.json     # 主进程 TS 配置
├── src/
│   ├── main/
│   │   ├── main.ts            # Electron 主进程：窗口、IPC、持久化
│   │   ├── file-service.ts    # 文件扫描 + 拼音索引预计算 + 打开文件
│   │   └── preload.ts         # 预加载脚本：暴露安全 IPC API
│   └── renderer/
│       ├── main.tsx           # React 入口
│       ├── App.tsx            # 主页面：状态管理、IPC 调用
│       ├── index.css          # 全局样式
│       ├── types.ts           # 类型声明（ElectronAPI、FileEntry）
│       ├── utils.ts           # 内存过滤函数 filterFiles
│       └── components/
│           ├── DirectoryBar.tsx   # 工作目录区域
│           ├── ResultList.tsx     # 结果列表
│           └── ErrorMessage.tsx   # 错误提示
├── dist/                      # Vite 构建产物（渲染进程）
├── dist-electron/             # 主进程构建产物
└── release/                   # electron-builder 打包输出
```

## 安装依赖

```bash
cd bill-explorer
npm install
```

## 本地开发运行

```bash
npm run dev
```

启动后 Vite 会在 `http://localhost:5173` 启动 dev server，Electron 主进程通过 `VITE_DEV_SERVER_URL` 加载。开发时建议另开一个终端运行：

```bash
npx electron .
```

主进程会检测到 `VITE_DEV_SERVER_URL`（需在 electron 环境中设置该环境变量，或使用 `electron-vite` 工具）。最简单的开发方式是直接：

```bash
# 终端 1：启动 Vite dev server
npm run dev

# 终端 2：启动 Electron
VITE_DEV_SERVER_URL=http://localhost:5173 npx electron .
```

## 构建打包

### 1. 构建全部产物

```bash
npm run build
```

此命令依次执行：
- `npm run build:renderer` — TypeScript 编译 + Vite 打包渲染进程 → `dist/`
- `npm run build:main` — esbuild 打包主进程（external 依赖）→ `dist-electron/main.js`

### 2. 本地预览构建产物

```bash
npm run electron:preview
```

### 3. 打包为安装包

```bash
npm run electron:build
```

- Windows：生成 `release/BillExplorer Setup x.x.x.exe`（NSIS 安装包）
- macOS：生成 `release/BillExplorer x.x.x.dmg`
- Linux：生成 `release/BillExplorer-x.x.x.AppImage`

> Windows 打包需要在本机安装 Node.js，macOS 打包需要在 macOS 上进行。

## 注意事项

1. **文件数量很大时首次扫描会有短暂耗时**：扫描是递归遍历全部子目录，几万级文件可能需要数秒到数十秒，期间 UI 会显示"扫描中"状态。
2. **索引仅保存在内存**：拼音全拼、首字母缩写在扫描完成后预计算并存储在内存中，**不做数据库持久索引**。关闭软件或点击【刷新】后索引会重新生成。
3. **不读取 Excel 内部内容**：所有检索仅基于文件名和文件路径，绝不打开或读取 `.xlsx` / `.xls` 文件内容。
4. **生僻字容错**：遇到 `pinyin-pro` 无法解析的生僻汉字，会跳过该字符，不影响其他字段的匹配，程序不会崩溃。
5. **目录权限**：遇到权限不足的子目录会跳过并继续，不影响其他文件扫描；目录不存在会给出明确错误提示。
6. **单击触发打开**：为降低误操作门槛，列表条目单击即调用系统默认程序打开，无需双击。

## 技术栈

| 用途 | 技术 |
|------|------|
| 桌面框架 | Electron 31 |
| 构建工具 | Vite 5 + esbuild |
| 前端框架 | React 18 + TypeScript |
| 中文拼音 | pinyin-pro（处理全拼 + 首字母缩写） |
| 配置持久化 | electron-store（仅保存工作目录） |
| 打包工具 | electron-builder |
