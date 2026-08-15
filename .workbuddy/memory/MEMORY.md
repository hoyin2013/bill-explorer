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
- **金额列 = 数量×单价，但支持「手工金额」（2026-08-15 最终定稿）**：金额列绝大多数行自动算（=数量×单价），少数行（整单总价/运费/折扣）用户需手填一个≠乘积的金额且不被覆盖。实现：抽纯函数到 `src/renderer/univerAdapter.ts`——`classifyAmount(q,p,amt)` 返回 'auto'|'pending'|'manual'；`buildAutoAmountRows(rows)` 加载/初始化时按数据把“金额==乘积或待填”的行加入 `autoRows` 集合（手工金额行不加）；`recomputeAmount(ws,row,autoRows)` 在 数量>0&单价>0 时：若 **row 在 autoRows**（自动行，权威判定，**不靠比较当前金额==乘积**）或金额 pending → 写 `q*p`；否则（手工金额）不覆盖；数量/单价缺失则不动。`onAmountChanged(ws,row,autoRows)` 在**金额列被编辑**时按“当前值==乘积”重分类（手填≠乘积→移出集合；改成==乘积→重新纳入）。`UniverSheet.tsx` 加载后 `autoAmountRows.current = buildAutoAmountRows(res.rows)`；`SheetValueChanged` 监听：数量/单价变化→`recomputeAmount`、金额变化→`onAmountChanged`；OCR 填入后把数量单价为正的行加入集合。**铁律**：recomputeAmount 用 autoRows 集合判定自动行，绝不能用“当前金额==乘积”判断（否则改数量时旧金额≠新乘积会被误判手工行而不更新——这是测试抓出的真实 bug）。
- **金额读取铁律**：`loadSheet.cellToText`（src/main/excel-memo.ts）对公式单元格**优先返回 Excel 缓存的 `result`**（仅当无缓存才退回公式串），保证编辑框与 Excel 显示一致、不重算成 0。存盘金额走普通数值单元格；数量/单价/金额须数值化避免文本。OCR 金额留空、由监听自动算 `q*p`。
- **金额列禁止千分位逗号（2026-08-15 当晚修）**：`saveSheet`（src/main/excel-memo.ts）原本给 数量/单价/金额 设 `numFmt='#,##0'`，导致 Excel 里 1000 显示成 `1,000`、与编辑界面(无格式 1000)不一致。现改为 `numFmt='General'`（ExcelJS 归一化为 undefined，即默认无千分位），三种列都统一，重新保存也能清除旧文件残留的逗号格式。

## 识图：检测 vs 内容识别 是两套链路（重要，迁移必看）
- **小票检测（框出每张票）** 走本地 ONNX 模型：`models/ticket_detect.onnx` + `scripts/detect_onnx.cjs` + `onnxruntime-node`（原生模块，平台相关）。这些文件都在 **app/项目目录内**，随 app 拷贝走，所以换机器**检测通常仍能用**。
- **内容识别（OCR 出人名/商品/金额）** 走**远程 OpenAI 兼容 API**。`src/main/ai-service.ts` 的 `recognizeReceipt` 会根据 `baseURL` 自动路由：以 `/responses` 结尾时走 **Responses API**（适配火山方舟 Ark）；否则走传统 **chat/completions**。`recognizeSingleCrop`/`recognizeTicketsWithDetection` 复用该能力。需要 `settings.aiConfig`（baseURL / apiKey / model / temperature / fastMode），存在 **electron-store 的 `config.json`**，位于**用户目录、不在 app 包内**：
  - **批量为「整图单次调用」省 token（2026-08-15 晚改）**：`recognizeTicketsWithDetection` 先用本地 ONNX 框出 N 张小票（boxes/crops 供 UI），但识别不再逐张裁剪各调一次，而是把每张小票的边界框（x,y,w,h + 原图尺寸）拼进提示词（`buildDetectPrompt`），**对整图只发 1 次 `recognizeReceipt`（returnRaw:true）**，模型按坐标逐框定位并一次性返回全部小票数组，再映射回 `det.boxes[i]`。相比逐张 N 次调用，省下 N-1 次请求的图片/提示词 token 与往返；代价是整图下小票被统一缩放、单票分辨率略降（pic/1.jpg 20 张实测精度可接受）。`recognizeReceipt` 新增 `opts.returnRaw` 透传原始 {name,date,items} 数组（`doRequest`/`doRequestResponses` 的 ok 返回附带 `parsed`，`normalizeRows` 已 `export`）。**token 实测（pic/1.jpg/20票）：整图单次 4,459 tokens vs 逐张 20 次 38,074 tokens → 逐张是整图 8.5 倍，整图省约 88%**。根因：逐张重复发送 SINGLE_TICKET_PROMPT(~900 tok)×20 + 每张图独立被视觉编码器下采样各占 token 预算；Ark doubao-seed 对大图会内部下采样到固定分辨率，故整图 input 反而很低。
  - macOS：`~/Library/Application Support/bill-explorer/config.json`
  - Windows：`%APPDATA%\bill-explorer\config.json`
  - Linux：`~/.config/bill-explorer/config.json`
- **「换机器后小票还能框、但内容识别不到」的根因**：app 拷过去了（检测模型在包内），但 `config.json` 没拷（AI 配置在用户目录）→ 新机器 `aiConfig` 为空 → `recognizeReceipt` 的 `isValidConfig` 为 false → 返回 `{error, message:'AI 接口未配置：请在设置中填写…'}`。
- **迁移修复**：① 在新机器设置里重填 AI 配置（最稳）；② 或把旧机器 `config.json` 复制到新机器同路径后重启（注意 `workDir` 等路径在新机器可能要重选）。若填好仍失败，看识图窗口状态栏报错：`请求 AI 接口失败/超时`=网络或 key 被 IP 白名单拦；`HTTP 4xx`= key/model 错。


- **内容识别（识图第二条链路）推荐模型（2026-08-15 结论）**：场景=中文手写小票+看图+结构化抽取(JSON)。硬约束：必须支持 vision、OpenAI 兼容 chat/completions、中文手写 OCR 强、稳定输出 JSON。
  - **首选：阿里云百炼 Qwen-VL（`qwen-vl-max-latest` 或带日期的快照如 `qwen-vl-max-0919`）**：官方定位即「多语言文字+手写体识别」「发票/表单/表格结构化输出」，与本项目小票场景 100% 吻合；国内节点低延迟；原生 OpenAI 兼容，`ai-service.ts` 已对百炼做 `enable_thinking:false` 兼容（fastMode 关思考提速省 token）。baseURL=`https://dashscope.aliyuncs.com/compatible-mode/v1`，约 ¥0.02/千tokens。
  - **次选（快+便宜）：Gemini 2.5 Flash / 2.0 Flash**（OpenRouter 或 Google OpenAI 兼容接口），手写 OCR 强、极速极省，国内直连需代理。
  - **最高精度不计成本：GPT-5.5 / GPT-4o**（OpenAI 官方 OpenAI 兼容），通用强但中文手写略逊 Qwen、价高。
  - **高吞吐低成本：豆包 vision / 智谱 GLM-4V / Qwen-VL 小模型**（国内、便宜，量大适用）。
  - **设置要点**：temperature 用低值（代码默认 0.2，OCR 要稳定）；fastMode 保持开启（尤其 Qwen3/推理模型）；单票 crop 已压到 1280px 省 token；每张小票一次调用、并发 3，量大时优先选便宜且够准的模型而非最贵旗舰；纯文本模型即使 OpenAI 兼容也不能用（无 vision）。
  - **当前项目默认接入（2026-08-15 晚更新）**：火山方舟 Responses API + `doubao-seed-evolving`。该模型默认开启深度思考，Ark Responses API 必须显式传 `thinking: { type: 'disabled' }` 才能直出 JSON；`ai-service.ts` 已在 fastMode=true（默认）时自动设置。
  - **识别结果缓存（2026-08-15 收尾）**：所有视觉 AI 调用经 `recognizeReceipt` 唯一入口，已加内容级缓存 `src/main/recogCache.ts`（内存 Map + 磁盘 `userData/recog-cache`，key=图片内容sha256 + 模型/提示词指纹，不含 apiKey）。命中则跳过 HTTP 请求；换模型/改提示词/图片内容变化才失效。`force` 贯穿可主动重识别；UI 有「清除识别缓存」按钮（IPC `clear-recog-cache`）。即「同一张图只请求一次，重开复用缓存」。
