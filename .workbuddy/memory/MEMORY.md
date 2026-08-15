# 项目长期约定（bill-explorer 账单录入器）

## 交互约定（2026-08-14 确定）
- 网格中「日期」列是**普通文本列**（type:'text'），**无日历控件**；用户直接输入或 Ctrl+V 粘贴，支持 2026/8/11、8-11、20260811 等格式（`parseDateText` 规范化）。保存到 Excel 时 `saveSheet` 仍把文本转成真正的 Excel 日期（numFmt yyyy-mm-dd），数据不丢。
- **ExcelJS 日期陷阱（2026-08-15 修复）**：写 Excel 日期必须用 `new Date(Date.UTC(y, mo-1, d))`（UTC 午夜）。用 `new Date(y, mo-1, d)`（本地午夜）会被 ExcelJS 按 UTC 偏移成带小数序列号，真实 Excel 里既**差一天**又**带时分秒**。`toExcelDate()`（excel-memo.ts）已改为 UTC 午夜并容忍 `2026/8/15`、`2026-08-15 14:30`、`2026年8月15日`、`20260815` 等写法（只取年月日）；读回 `fmtDateLocal()` 用 UTC getter 取整，保证往返一致。
- **单击单元格不再复制**。复制走 Ctrl+C / 右键菜单「复制」/ Ctrl+X 剪切（剪贴板读不到时用 `lastCopyRef` 兜底）。
- 打开 Excel 走主进程 `loadSheet`（IPC），用 `ws.eachRow` 单遍遍历读取；异常超大文件（rowCount>SAFE_MAX_ROWS=100_000）直接拦截弹窗，不读数据。
- **识图面板「序号」列已移除**：AI 识别结果表（`ImageWindow` 的 `RESULT_COLS`）不再显示序号列（该列是自动编号、对录入无意义）。
- **识图图片旋转持久化为全局默认方向**：旋转任一图片（总览视图 ↺/↻）即通过 `set-image-rotation` 把角度（0/90/180/270）存进 electron-store 的 `imageRotation`；之后每张图片 `loadPreview` 自动套用 `defaultRotateRef`，检测（`detectOnly`→`rotatePreviewForDetect`）也按此方向切图。单张视图的 `singleRotate` 仅本地显示、不持久化（裁剪图已含方向）。
- **识图裁剪图不做自动旋转（noRotate:true）**：`detectTickets`（`src/main/detection.ts`，含 ONNX 主路径、`detectFromFile`、Python 回退三处）一律 `noRotate:true`。即每张小票 crop = 用户手动转好的整图的原始矩形切片，**不再逐张 `bestRotation` 自动旋转**——逐张自动旋转对近方形/文字少的小票常猜错角度，导致单张视图"横七竖八"。方向以用户手动整图旋转为准；个别小票若仍需微调，用单张视图 ↺/↻（`singleRotate`，仅显示不改 crop）。

## 架构要点
- 渲染进程 `src/renderer/*`（tsconfig.json include），主进程 `src/main/*`（esbuild 打包，忽略 tsc 类型告警）。
- `loadSheet`/`saveSheet`/`appendMemo`/`updateMemo`/`restoreBackup`/`listBackups` 在 `src/main/excel-memo.ts`，经 `src/main/preload.ts` 暴露为 `window.electronAPI`。
- 固定 9 列账单表头：序号/日期/货品名称/单位/数量/单价/金额/调货人/备注（`HEADER` 常量）。
- 保存前自动备份到同级 `.billbackups/`（最多 5 份，时间戳精确到毫秒）。

## 表格渲染引擎：Univer 混合接入（2026-08-15 起的架构）
- **渲染进程改用 Univer（Canvas 虚拟化表格）** 替代自研 `SheetGrid`，彻底解决 3000+ 行卡顿。文件读写仍由 **ExcelJS** 负责（保留日期UTC、空行占位、大文件拦截、自动备份全部约定），Univer 只做渲染+编辑。
- **关键文件**：`src/renderer/univerAdapter.ts`（`string[][]` ↔ Univer `IWorkbookData` 互转 + OCR 记录→9列映射）、`src/renderer/components/UniverSheet.tsx`（挂载 Univer、灌数据/取回、OCR 填入、保存/恢复）、`App.tsx` 用 `UniverSheet` 替换 `SheetGrid`。
- **数据流**：打开 `api.loadSheet`→`rowsToWorkbookData`→`createWorkbook`；保存 `getSnapshot`→`workbookDataToRows`→去尾空行→`api.saveSheet`。首行作为表头（保存时剥离，由 saveSheet 重写）。
- **OCR「填入到激活行」**：`UniverSheet` 监听 `apply-recognized-rows`，取 `getActiveRange().getRow()` 为起始行，从首列 `setValues` 写入并推进激活格。
- **Univer 原生替代了原自研逻辑**：矩形框选、复制/粘贴、拖拽填充序列（序号1/2/3）、单元格编辑、滚动虚拟化——无需再维护。`SheetGrid.tsx` 已不再被引用（保留文件备查）。
- **弃用「建议框」（拼音补全）**，待后续按需补回（Univer 有数据验证下拉可作临时替代）。
- 依赖 `@univerjs/presets` + `@univerjs/preset-sheets-core` @0.25.1；`vite.config.ts` 已 `resolve.dedupe:['react','react-dom']`。
- **日期列是「真实 Excel 日期列」**：Univer 会自动把输入的日期识别成序列号（serial），所以 `univerAdapter.ts` 的 `rowsToWorkbookData` 把第 2 列（0 基索引 1）建成 `{v:序列号, s:'bill-date'}` 且 `columnData` 整列默认 `bill-date` 样式（`{n:{pattern:'yyyy-mm-dd'}}`），`workbookDataToRows` 用 `normalizeDateValue` 把序列号/各种文本写法归一化成 `yyyy-mm-dd` 文本再交给 `saveSheet`。**千万不要把日期列当成纯文本、也不要去掉列 `numFmt`**，否则编辑器又会显示成裸数字（与 Excel 不一致）。序列号纪元 `1899-12-30`，与 ExcelJS `dateToSerial` 一致（已验证 2026-08-15 两端都是 46249）。
- **CSS 铁律（踩过坑，勿犯）**：`.univer-container` 的撑满规则**必须**写成精确选择器 `.univer-container > div[data-u-comp='workbench-layout']`，**绝不能用 `.univer-container > div` 通配**。Univer 会往同一容器插入两个 `position:fixed` 的隐藏输入法/选区容器 `div#univer-doc-selection-container-*`（DOCS_NORMAL = 单元格编辑器，DOCS_FORMULA_BAR = 编辑栏），通配规则的 `width/height:100%` 对 fixed 元素解析为**整个视口**，加上它 `activate()` 时会置 `z-index:1000`，就变成一层全屏透明遮罩，把顶栏/侧栏/表头按钮全部锁死（只有 Univer 编辑区还能用），且**只有点过单元格之后才复现**（未交互时它被推到视口外 -998,-896）。因此 index.css 常驻兜底规则：`div[id^='univer-doc-selection-container-'] { position:absolute !important; width:0 !important; height:0 !important }`（absolute 同时修正 IME 候选框位置，因为 Univer 内部按「相对父容器」算 left/top）；另给 `.app-top`/`.app-side`/`.memo-header` 加 `z-index:1001`（>Univer 1000，<Univer 弹层 1020）作二重保险。
- **滚动铁律（踩过坑，勿犯）**：Univer `UniverSheetsCorePreset` 虽注册了 `SHEET_UI_PLUGIN`，但渲染层**没有注册 `SheetsScrollRenderController`/`SheetScrollManagerService`**，因此 `sheet.command.scroll-to-cell` 永远失败（找不到控制器）。要让视图滚到某单元格，只能用 **`sheet.command.scroll-view`**（参数 `{ sheetViewStartRow, sheetViewStartColumn, offsetX:0, offsetY:0 }`，按行索引滚动，走 `SetScrollOperation`→`SheetScrollManagerService`，自动解析 unitId/sheetId）。且 **createWorkbook 之后 render 骨架要几百毫秒才就绪**，立刻执行会 `getRenderById(...).with(...)` 抛错——必须在 render 就绪后执行（用 `requestAnimationFrame`+`setTimeout` 重试即可，`setActiveAndScroll` 已封装）。注意：`setActiveRange` 只改数据层选区（光标会跳），不会自动滚视图。
- **金额列是「活公式」**（2026-08-15 改）：金额(列6, Excel=G) = 数量(列4,E)*单价(列5,F)，写成 Univer 公式 `=E{row+1}*F{row+1}`（`row`=Univer 0 基行，+1 才是对应的 Excel 行号）。`UniverSheet.recomputeAmount` 在数量/单价为正、金额空/0/非数字时写公式；金额列已有合理数值（手填或公式结果）则不覆盖。公式引擎由 core preset 的 `UniverSheetsFormulaPlugin`+`UniverFormulaEnginePlugin` 提供（公式是真计算）。**两个铁律**：① 数量/单价/金额必须存成数字（`{v:Number}`），否则 `=E*F` 变「文本*文本」→ `#VALUE!`；`rowsToWorkbookData` 已对这三列数值化，OCR 填入矩阵也数值化。② 公式存于 Univer `f` 字段（含 `=` 前缀）、ExcelJS 存于 `{formula,result}`，读取金额**必须读 `f`/`formula` 而非 `v`/`result`**（否则 `workbookDataToRows`/`loadSheet.cellToText` 会丢公式退化成静态值）；save 时以 `=` 开头的值写成 `cell.value={ formula: raw.slice(1) }`，Excel 打开自动计算。OCR 识别出金额则保留数值、否则留空等 `SheetValueChanged` 监听自动写公式。

## 识图：检测 vs 内容识别 是两套链路（重要，迁移必看）
- **小票检测（框出每张票）** 走本地 ONNX 模型：`models/ticket_detect.onnx` + `scripts/detect_onnx.cjs` + `onnxruntime-node`（原生模块，平台相关）。这些文件都在 **app/项目目录内**，随 app 拷贝走，所以换机器**检测通常仍能用**。
- **内容识别（OCR 出人名/商品/金额）** 走**远程 OpenAI 兼容 chat/completions API**（`src/main/ai-service.ts` 的 `recognizeReceipt`/`recognizeSingleCrop`/`recognizeTicketsWithDetection`）。需要 `settings.aiConfig`（baseURL / apiKey / model / temperature / fastMode），这些配置存在 **electron-store 的 `config.json`**，位于**用户目录、不在 app 包内**：
  - macOS：`~/Library/Application Support/bill-explorer/config.json`
  - Windows：`%APPDATA%\bill-explorer\config.json`
  - Linux：`~/.config/bill-explorer/config.json`
- **「换机器后小票还能框、但内容识别不到」的根因**：app 拷过去了（检测模型在包内），但 `config.json` 没拷（AI 配置在用户目录）→ 新机器 `aiConfig` 为空 → `recognizeReceipt` 的 `isValidConfig` 为 false → 返回 `{error, message:'AI 接口未配置：请在设置中填写…'}`。
- **迁移修复**：① 在新机器设置里重填 AI 配置（最稳）；② 或把旧机器 `config.json` 复制到新机器同路径后重启（注意 `workDir` 等路径在新机器可能要重选）。若填好仍失败，看识图窗口状态栏报错：`请求 AI 接口失败/超时`=网络或 key 被 IP 白名单拦；`HTTP 4xx`= key/model 错。

