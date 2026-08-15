# 项目长期约定（bill-explorer 账单录入器）

## 交互约定
- 网格「日期」列是**普通文本列**（无日历控件），支持 `2026/8/11`、`8-11`、`20260811` 等（`parseDateText` 规范化）；保存时 `saveSheet` 转成真实 Excel 日期（numFmt `yyyy-mm-dd`）。
- **ExcelJS 日期陷阱**：写日期必须 `new Date(Date.UTC(y, mo-1, d))`（UTC 午夜），否则被按 UTC 偏移成带小数序列号，真实 Excel 里差一天+带时分秒。`toExcelDate()`(excel-memo.ts) 已容忍 `2026/8/15`、`2026-08-15 14:30`、`2026年8月15日`、`20260815`；`fmtDateLocal()` 用 UTC getter 取整，往返一致。
- 复制走 Ctrl+C / 右键「复制」/ Ctrl+X（剪贴板读不到用 `lastCopyRef` 兜底），**单击单元格不再复制**。
- 打开 Excel 走主进程 `loadSheet`（`ws.eachRow` 单遍读）；超大文件（rowCount>SAFE_MAX_ROWS=100_000）直接拦截弹窗不读。
- **识图结果表（`ImageWindow` 的 `RESULT_COLS`）不显示序号列**（自动编号、无意义）。
- **识图图片旋转持久化为全局默认方向**：总览 ↺/↻ 即经 `set-image-rotation` 存 electron-store `imageRotation`；`loadPreview` 套用 `defaultRotateRef`，`detectOnly`→`rotatePreviewForDetect` 按此切图。单张 `singleRotate` 仅本地显示、不持久化。
- **识图裁剪图不做自动旋转（noRotate:true）**：`detectTickets`(src/main/detection.ts 三处) 一律 `noRotate:true`，crop=用户手动转好的整图原始矩形切片；个别微调用单张 ↺/↻（`singleRotate`，仅显示）。
- **识图默认隐藏左侧图片列表**：`listCollapsed` 初值 `true`，点「☰ 显示列表」展开。
- **识图结果表列宽可拖动**：`<colgroup>` + th 右边界拖拽手柄（`startColResize`/`colWidths`），像 Excel 调列宽，`index.css .col-resizer` 兜底。
- **识图日期联动（2026-08-16）**：编辑某张小票日期 → 从该张起（含）所有小票每行日期一起改，该张之前不变；首改下标存 `dateAnchorStartRef`，换图重检测(`detectOnly`)复位。识别基准 `recogDateRef` 与手动值 `dateAnchorRef` 分离；`recognizeTicketByIdx` 对联动范围内票套用手动值，使「重新识别此张」(force) 不冲掉修正（即缓存）。

## 架构要点
- 渲染 `src/renderer/*`(tsconfig include)；主进程 `src/main/*`(esbuild 打包，忽略 tsc 告警)。
- `loadSheet`/`saveSheet`/`appendMemo`/`updateMemo`/`restoreBackup`/`listBackups` 在 `src/main/excel-memo.ts`，经 `src/main/preload.ts` 暴露 `window.electronAPI`。
- 固定 9 列账单表头：序号/日期/货品名称/单位/数量/单价/金额/调货人/备注（`HEADER`）。保存前自动备份同级 `.billbackups/`（最多 5 份，毫秒时间戳）。

## Univer 混合渲染（2026-08-15 起）
- 渲染进程用 **Univer**（Canvas 虚拟化）替代自研 `SheetGrid`，解 3000+ 行卡顿；文件读写仍 **ExcelJS**（保留日期UTC/空行占位/大文件拦截/自动备份）。`SheetGrid.tsx` 已不被引用（备查）。
- 关键文件：`univerAdapter.ts`(string[][]]↔IWorkbookData + OCR→9列映射)、`components/UniverSheet.tsx`(挂载/灌取/OCR填入/存读)、`App.tsx` 用 `UniverSheet` 替换 `SheetGrid`。
- 数据流：开 `api.loadSheet`→`rowsToWorkbookData`→`createWorkbook`；存 `getSnapshot`→`workbookDataToRows`→去尾空行→`api.saveSheet`。首行作表头（保存剥离、saveSheet 重写）。
- OCR「填入到激活行」：`UniverSheet` 监听 `apply-recognized-rows`，取 `getActiveRange().getRow()` 起始行 `setValues` 并推进激活格。
- 弃用「建议框」（拼音补全），待补回（Univer 数据验证下拉可临时替代）。依赖 `@univerjs/presets`+`@univerjs/preset-sheets-core`@0.25.1；`vite.config.ts` 已 `resolve.dedupe:['react','react-dom']`。
- **日期列是真实 Excel 日期列**：`rowsToWorkbookData` 把第 2 列(0基1)建成 `{v:序列号,s:'bill-date'}`，`columnData` 整列 `bill-date` 样式 `{n:{pattern:'yyyy-mm-dd'}}`；`workbookDataToRows` 用 `normalizeDateValue` 归一化再给 `saveSheet`。**绝不可当纯文本或去掉列 numFmt**（否则显示裸数字）。序列号纪元 `1899-12-30`，与 ExcelJS `dateToSerial` 一致。
- **CSS 铁律**：`.univer-container` 撑满必须精确选择器 `.univer-container > div[data-u-comp='workbench-layout']`，**绝不用 `.univer-container > div` 通配**（会让 Univer 注入的 fixed 选区容器 `div#univer-doc-selection-container-*` 变全屏透明遮罩锁死顶栏/侧栏）。index.css 常驻兜底：`div[id^='univer-doc-selection-container-']{position:absolute!important;width:0!important;height:0!important}`；`.app-top`/`.app-side`/`.memo-header` 加 `z-index:1001` 二重保险。
- **滚动铁律**：渲染层未注册 `SheetsScrollRenderController`/`SheetScrollManagerService`，`scroll-to-cell` 永远失败。滚到某格只能用 `sheet.command.scroll-view`(`{sheetViewStartRow,sheetViewStartColumn,offsetX:0,offsetY:0}`，走 `SetScrollOperation`→`SheetScrollManagerService`)。createWorkbook 后 render 几百毫秒才就绪，须 `requestAnimationFrame`+`setTimeout` 重试（`setActiveAndScroll` 已封装）。`setActiveRange` 只改数据层选区、不滚视图。
- **金额列 = 数量×单价，支持「手工金额」**：`univerAdapter.ts` 纯函数 `classifyAmount`(auto/pending/manual)、`buildAutoAmountRows`、`recomputeAmount`(数量>0&单价>0 时：row 在 autoRows 集合或金额 pending→写 `q*p`；手工金额不覆盖)、`onAmountChanged`(金额被编辑时按当前值==乘积重分类)。`UniverSheet.tsx` 加载后 `autoAmountRows.current=buildAutoAmountRows(...)`；`SheetValueChanged`：数量/单价变→`recomputeAmount`、金额变→`onAmountChanged`；OCR 填入后数量单价为正的行加入集合。**铁律：`recomputeAmount` 用 autoRows 集合判定，绝不能用「当前金额==乘积」判断**（改数量时旧金额≠新乘积会被误判手工行而不更新——真实 bug）。
- **金额读取铁律**：`loadSheet.cellToText` 对公式单元格**优先返回 Excel 缓存 `result`**（无缓存才退回公式串），编辑框与 Excel 一致、不重算成 0。存盘金额走普通数值；数量/单价/金额须数值化。OCR 金额留空、由监听自动算 `q*p`。
- **金额列禁止千分位逗号**：`saveSheet` 给数量/单价/金额设 `numFmt='General'`（非 `'#,##0'`），否则 Excel 显示 `1,000` 与编辑界面不一致。

## 识图：检测 vs 内容识别 两套链路
- **检测（框票）**走本地 ONNX：`models/ticket_detect.onnx` + `scripts/detect_onnx.cjs` + `onnxruntime-node`（原生模块，平台相关，随 app 拷贝走→换机器通常仍可用）。
- **内容识别（OCR）**走远程 OpenAI 兼容 API：`ai-service.ts` 的 `recognizeReceipt` 按 `baseURL` 路由——以 `/responses` 结尾走 Responses API(火山方舟 Ark)，否则 chat/completions。`settings.aiConfig`(baseURL/apiKey/model/temperature/fastMode) 在 electron-store `config.json`（用户目录，不在 app 包内）：macOS `~/Library/Application Support/bill-explorer/config.json`、Windows `%APPDATA%\bill-explorer\config.json`、Linux `~/.config/bill-explorer/config.json`。
- **整图单次调用省 token（2026-08-15 晚）**：`recognizeTicketsWithDetection` 先用 ONNX 框票，再把每票 bbox(x,y,w,h+原图尺寸) 拼进提示词，对整图只发 1 次 `recognizeReceipt(returnRaw:true)`，模型按坐标逐框返回全部票。实测 pic/1.jpg/20票：整图单次 4,459 tokens vs 逐张 38,074 tokens（省约 88%）。
- **「换机器能框但识别不到」根因**：app 拷了（检测模型在包内）但 `config.json` 没拷（AI 配置在用户目录）→ `isValidConfig` false → 报「AI 接口未配置」。修复：新机器重填设置，或复制旧 `config.json` 到同路径重启。失败看状态栏：`请求AI失败/超时`=网络或 key 被 IP 白名单拦；`HTTP 4xx`=key/model 错。
- **推荐模型**：首选阿里云百炼 Qwen-VL(`qwen-vl-max-latest`)；次选 Gemini 2.5/2.0 Flash（快+便宜，国内直连需代理）；不计成本 GPT-4o；高吞吐豆包/智谱 GLM-4V/Qwen-VL 小模型。硬约束：支持 vision、OpenAI 兼容、中文手写 OCR 强、稳定输出 JSON。当前默认火山方舟 Responses API + `doubao-seed-evolving`（fastMode 默认开 → 自动 `thinking:{type:'disabled'}` 直出 JSON）。temperature 用低值(默认 0.2)；单票 crop 压到 1280px 省 token；每次调用并发 3。
- **识别结果缓存**：所有视觉调用经 `recognizeReceipt` 唯一入口，内容级缓存 `src/main/recogCache.ts`（内存 Map + 磁盘 `userData/recog-cache`，key=图片内容sha256 + 模型/提示词指纹，不含 apiKey）；命中跳 HTTP；换模型/改提示词/图片变化才失效。`force` 可主动重识别；UI 有「清除识别缓存」(`clear-recog-cache`)。即「同一张图只请求一次，重开复用」。
