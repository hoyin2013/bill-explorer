# 项目长期约定（bill-explorer 账单录入器）

## 交互约定（2026-08-14 确定）
- 网格中「日期」列是**普通文本列**（type:'text'），**无日历控件**；用户直接输入或 Ctrl+V 粘贴，支持 2026/8/11、8-11、20260811 等格式（`parseDateText` 规范化）。保存到 Excel 时 `saveSheet` 仍把文本转成真正的 Excel 日期（numFmt yyyy-mm-dd），数据不丢。
- **单击单元格不再复制**。复制走 Ctrl+C / 右键菜单「复制」/ Ctrl+X 剪切（剪贴板读不到时用 `lastCopyRef` 兜底）。
- 打开 Excel 走主进程 `loadSheet`（IPC），用 `ws.eachRow` 单遍遍历读取；异常超大文件（rowCount>SAFE_MAX_ROWS=100_000）直接拦截弹窗，不读数据。
- **识图面板「序号」列已移除**：AI 识别结果表（`ImageWindow` 的 `RESULT_COLS`）不再显示序号列（该列是自动编号、对录入无意义）。
- **识图图片旋转持久化为全局默认方向**：旋转任一图片（总览视图 ↺/↻）即通过 `set-image-rotation` 把角度（0/90/180/270）存进 electron-store 的 `imageRotation`；之后每张图片 `loadPreview` 自动套用 `defaultRotateRef`，检测（`detectOnly`→`rotatePreviewForDetect`）也按此方向切图。单张视图的 `singleRotate` 仅本地显示、不持久化（裁剪图已含方向）。

## 架构要点
- 渲染进程 `src/renderer/*`（tsconfig.json include），主进程 `src/main/*`（esbuild 打包，忽略 tsc 类型告警）。
- `loadSheet`/`saveSheet`/`appendMemo`/`updateMemo`/`restoreBackup`/`listBackups` 在 `src/main/excel-memo.ts`，经 `src/main/preload.ts` 暴露为 `window.electronAPI`。
- 固定 9 列账单表头：序号/日期/货品名称/单位/数量/单价/金额/调货人/备注（`HEADER` 常量）。
- 保存前自动备份到同级 `.billbackups/`（最多 5 份，时间戳精确到毫秒）。
