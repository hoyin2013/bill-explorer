// 单元格输入记忆（自动补全）控制器：基于 Univer 官方「编辑桥接服务」事件驱动。
// 设计参照 WPS/Excel 的“按列自动完成”——打字时下拉浮现本列曾出现过的、以及你以前输入过的内容。
//
// 为什么这样设计（以及为什么此前的版本反复失败）：
// 旧版每帧用 querySelector 轮询 DOM 里的单元格编辑器（[data-u-comp='editor']），
// 而 Univer 在 DOM 里同时存在一个屏幕外、零尺寸的“公式栏/隐藏占位”编辑器，
// querySelector 默认命中它，导致“永远读不到文本、永远不提示”。
// 本控制器改用官方状态源，彻底摆脱对 DOM 轮询的依赖：
//   - IEditorBridgeService.currentEditCell$ 在“进入/退出单元格编辑”时推送 {row,column,position,editorUnitId}
//   - IEditorService.getEditor(editorUnitId) 拿到 Slate 编辑器实例，订阅其 input$ 拿实时输入文本（content）
//   - 下拉定位用官方 position 反查真正“单元格编辑器”的屏幕矩形（排除占位壳）
//
// 回填（写入候选值）的实现——这是此前多版失败的核心，关键结论：
//   1) currentEditCell$.row 比真实 0 基行 +1（偏移），用它定位落盘格会写错行。
//   2) 在 keydown 里 replaceText 后，Univer 原生提交读的是另一份快照（时序问题），导致写不进候选。
//   3) 真正可靠的唯一注入点是 commandService.beforeCommandExecuted：
//      Univer 把编辑器内容写回单元格，走的是 _submitEdit → SetRangeValuesCommand。
//      我们在该命令“执行前”拦截，把它的 value 替换为候选值（字符串/数字 + 正确 CellValueType）。
//      于是落盘由 Univer 自己用 editCellState 的【正确 row】完成，既绕开 +1 偏移、也绕开 replaceText 时序。
//      键盘 Enter/Tab 直接放行原生提交（拦截器换值：Enter 仅当用户已用方向键主动选中过候选、Tab 始终回填）；
//      鼠标点选则显式触发关闭+提交（同一拦截器换值）。
//
// 本模块与渲染框架无关：宿主（React 组件 / 测试页）通过 AcHost 回调提供
//   computeItems（候选算法）、commit（仅记录历史，供后续候选）、render（渲染弹层）、container（坐标偏移）、univerAPI。
import { IEditorBridgeService, SetCellEditVisibleOperation } from '@univerjs/sheets-ui'
import { IEditorService } from '@univerjs/docs-ui'
import { ICommandService, CellValueType } from '@univerjs/core'
import { SetRangeValuesCommand } from '@univerjs/sheets'
import { COL_COUNT, DATE_COL_INDEX, type AcItem } from './autocomplete'

export interface AcViewState {
  open: boolean
  items: AcItem[]
  index: number
  pos: { left: number; top: number; width: number } | null
}

export interface AcHost {
  // 计算候选（复用宿主既有/已测的算法）
  computeItems: (text: string, col: number, row: number) => AcItem[]
  // 写入选中值后的「历史记录」回调（仅用于让自动补全学习该值；真正的单元格写入由 Univer 原生提交完成）。
  // 注意：此处不再传 row——currentEditCell$.row 存在 +1 偏移，落盘行由 Univer 自己用 editCellState 决定。
  commit: (item: AcItem, col: number) => void
  // 渲染当前弹层状态
  render: (state: AcViewState) => void
  // Univer 容器元素（用于坐标偏移）；可能为 null
  container: HTMLElement | null
  // Univer facade（用于执行“关闭编辑器”命令）
  univerAPI: { executeCommand: (id: string, params?: object) => unknown } | null
}

// attachAutocomplete 的返回值：宿主（React 组件 / 测试页）通过它驱动鼠标交互、并在卸载时释放。
export interface AcController {
  // 释放所有订阅与监听
  dispose: () => void
  // 设置高亮项（鼠标 hover 候选项时调用）
  setIndex: (i: number) => void
  // 选中并写入某候选（鼠标点按候选项时调用；传 null 仅关闭弹层）
  accept: (item: AcItem | null) => void
}

const noOpController: AcController = { dispose: () => {}, setIndex: () => {}, accept: () => {} }

interface IPosition {
  startX: number
  startY: number
  endX: number
  endY: number
}

// 挂载自动补全，返回卸载函数。
export function attachAutocomplete(univerInst: any, host: AcHost): AcController {
  let injector: any
  try {
    injector = univerInst.__getInjector()
  } catch {
    console.warn('[AC] 无法获取 Univer 注入器，自动补全不可用')
    return noOpController
  }
  let editorBridge: any
  let editorService: any
  let commandService: any
  try {
    editorBridge = injector.get(IEditorBridgeService)
    editorService = injector.get(IEditorService)
    commandService = injector.get(ICommandService)
  } catch (e) {
    console.warn('[AC] 注入编辑服务失败：', e)
    return noOpController
  }
  if (!editorBridge || !editorService || !commandService) return noOpController

  let editRow = -1
  let editCol = -1
  let editEditorId = ''
  let posRef: IPosition | null = null
  let lastKey = ''
  let editorSub: { dispose: () => void } | null = null
  let items: AcItem[] = []
  let index = 0
  // 用户是否已用方向键主动选择过候选项（Excel/WPS 语义）：
  // 只有此时 Enter 才“写入高亮候选”；否则 Enter 一律提交用户自己输入的内容，
  // 防止“输了 8.7、弹层开着、一按 Enter 就被换成上方最近的候选值”。
  let navigated = false
  let pos: { left: number; top: number; width: number } | null = null
  // 是否正处于“接受候选、交还 Univer 原生提交”的过程中。
  // 置位期间忽略编辑器 input$ 回调，避免任何文本变化又把弹层重新打开。
  let accepting = false
  // 待注入的候选：accept（鼠标）或在 onKey 里命中候选（键盘）时设置。
  // 而后由 beforeCommandExecuted 拦截 Univer 的“单元格提交命令”，在它读取编辑器快照之前
  // 把提交值替换为该候选——落盘由 Univer 用 editCellState 的正确 row 完成。
  let pendingAccept: AcItem | null = null
  // accept 时的列（arm 时捕获）：endEdit 会把 editCol 清零，而拦截器在 endEdit 之后才触发，
  // 故历史记录需用本变量而非 editCol，避免记到列 -1。
  let acceptCol = -1
  // 超时兜底：若 accept 后没有任何提交命令触发（极端情况），清空 pending，避免泄漏到后续操作。
  let pendingTimer: ReturnType<typeof setTimeout> | null = null
  const clearPendingTimer = () => {
    if (pendingTimer != null) {
      clearTimeout(pendingTimer)
      pendingTimer = null
    }
  }
  // accept 后的“重开抑制”：Univer 在 executeCommand(set-cell-edit-visible, visible:false) 的退出链里，
  // 提交候选值后会让编辑器“再次进入同一格编辑”（set-activate-cell-edit → 重新 emit currentEditCell$），
  // 并带着旧文本再次触发 recompute，把刚关掉的弹层又开出来（即“点一次能带入、弹层不消失、再点才消失”的根因）。
  // 因此 accept 后，对本格在短时间内（300ms）的 recompute 直接吞掉，弹层保持关闭；超时后恢复正常（用户主动在别处打字）。
  let acceptGuardRow = -1
  let acceptGuardCol = -1
  let acceptGuardTimer: ReturnType<typeof setTimeout> | null = null
  const clearAcceptGuard = () => {
    if (acceptGuardTimer != null) {
      clearTimeout(acceptGuardTimer)
      acceptGuardTimer = null
    }
    acceptGuardRow = -1
    acceptGuardCol = -1
  }
  const armAccept = (item: AcItem) => {
    pendingAccept = item
    acceptCol = editCol
    accepting = true
    clearPendingTimer()
    pendingTimer = setTimeout(() => {
      pendingAccept = null
    }, 1000)
  }

  const containerRect = () => {
    const r = host.container?.getBoundingClientRect()
    return { left: r?.left ?? 0, top: r?.top ?? 0 }
  }

  // 用官方 position 反查真正“单元格编辑器”的屏幕矩形（排除公式栏/隐藏占位壳）：
  // 只取视口内、有尺寸的 [data-u-comp='editor']；若仅一个直接用；多个则取离单元格坐标最近的。
  const findEditorRect = (): { left: number; top: number; width: number } | null => {
    if (!posRef) return null
    const { left: bl, top: bt } = containerRect()
    const ax = bl + posRef.startX
    const ay = bt + posRef.endY
    const cands: DOMRect[] = []
    document.querySelectorAll("[data-u-comp='editor']").forEach((el) => {
      const r = (el as HTMLElement).getBoundingClientRect()
      if (r.width <= 1 || r.height <= 1) return
      if (r.left < -50 || r.top < -50 || r.left > window.innerWidth + 50 || r.top > window.innerHeight + 50) return
      cands.push(r)
    })
    if (cands.length === 1) {
      return { left: cands[0].left, top: cands[0].top + cands[0].height + 2, width: cands[0].width }
    }
    if (cands.length > 1) {
      let best = cands[0]
      let bestD = Infinity
      for (const r of cands) {
        const d = Math.hypot(r.left - ax, r.top - ay)
        if (d < bestD) { bestD = d; best = r }
      }
      return { left: best.left, top: best.top + best.height + 2, width: best.width }
    }
    return { left: ax, top: ay + 2, width: Math.max(60, posRef.endX - posRef.startX) }
  }

  const closePopup = () => {
    items = []
    index = 0
    navigated = false
    pos = null
    host.render({ open: false, items: [], index: 0, pos: null })
  }

  const recompute = (text: string) => {
    const col = editCol
    const row = editRow
    if (col < 0 || col >= COL_COUNT || row < 0) {
      closePopup()
      return
    }
    // accept 后的重开抑制：本格刚被接受，Univer 退出链会让它“再次进入同一格编辑”并带旧文本重算，
    // 把刚关掉的弹层又开出来。此处直接关掉并保持关闭，避免“点一次不消失、再点才消失”。
    if (acceptGuardRow === row && acceptGuardCol === col) {
      closePopup()
      return
    }
    const t = (text || '').replace(/\s+/g, ' ').trim()
    if (!t) {
      closePopup()
      return
    }
    const next = host.computeItems(t, col, row)
    if (!next.length) {
      closePopup()
      return
    }
    items = next
    index = 0
    navigated = false
    pos = findEditorRect() || { left: 120, top: 120, width: 120 }
    host.render({ open: true, items, index, pos })
  }

  const handleText = (text: string) => {
    if (accepting) return
    const key = editCol + ':' + editRow + ':' + text
    if (key === lastKey) return
    lastKey = key
    recompute(text)
  }

  const readInitialText = (): string => {
    const editor = editorService.getEditor(editEditorId)
    try {
      const d = editor?.getDocumentData?.()
      return (d?.body?.dataStream || '').replace(/[\r\n]+$/g, '').replace(/[\r\n]+/g, ' ').trim()
    } catch {
      return ''
    }
  }

  const attachEditor = () => {
    if (editorSub) { try { editorSub.dispose() } catch {} editorSub = null }
    const editor = editorService.getEditor(editEditorId)
    if (!editor) {
      // 编辑器实例可能晚一帧才注册，稍后重试一次
      window.setTimeout(() => { if (editEditorId && editRow >= 0) attachEditor() }, 30)
      return
    }
    if (typeof editor.input$?.subscribe === 'function') {
      editorSub = editor.input$.subscribe((ev: { content?: string }) => handleText(ev?.content ?? ''))
    }
    handleText(readInitialText())
  }

  const endEdit = () => {
    // 注意：此处【不能】清空 pendingAccept。
    // 因为 SetCellEditVisibleOperation 的处理顺序是先 _exitInput（触发 visible$ → 本函数），
    // 之后才 _submitEdit 发起 SetRangeValuesCommand；拦截器在那之后才消费 pendingAccept。
    // 清空交给拦截器（消费后）与 pendingTimer（兜底），避免被提前清零导致换值失效。
    clearPendingTimer()
    if (editorSub) { try { editorSub.dispose() } catch {} editorSub = null }
    editRow = -1
    editCol = -1
    editEditorId = ''
    posRef = null
    lastKey = ''
    accepting = false
    closePopup()
  }

  const accept = (item: AcItem | null) => {
    if (!item) {
      pendingAccept = null
      clearPendingTimer()
      closePopup()
      return
    }
    // 鼠标路径：标记待注入候选 + 关闭弹层，然后显式触发“关闭并提交”。
    // 键盘路径不调用本函数（直接由 onKey 命中候选后放行原生 Enter/Tab），
    // 二者最终都经由 beforeCommandExecuted 把提交值替换为候选。
    // 先捕获当前编辑格坐标。
    const guardRow = editRow
    const guardCol = editCol
    // ⚠️ 必须【先】置位“重开抑制”，再 executeCommand：executeCommand 会【同步】触发
    // 退出链（endEdit → currentEditCell$ 重发 → recompute），若等到命令返回后才置位，
    // recompute 已经先跑过（无视抑制）把刚关掉的弹层又打开——这正是“点一次能带入、
    // 弹层不消失、再点才消失”的真正根因。抑制窗口必须提前覆盖整段命令执行期。
    acceptGuardRow = guardRow
    acceptGuardCol = guardCol
    if (acceptGuardTimer != null) clearTimeout(acceptGuardTimer)
    acceptGuardTimer = setTimeout(() => {
      acceptGuardRow = -1
      acceptGuardCol = -1
      acceptGuardTimer = null
    }, 300)
    armAccept(item)
    closePopup()
    if (host.univerAPI) {
      try {
    // keycode 绝不能传 13(Enter)/9(Tab)。若用 Enter 退出，Univer 会“提交 → 下移一格 →
    // set-activate-cell-edit 重新进入下一格编辑”，再次触发 currentEditCell$ → recompute。
    // 传 0 让 Univer 只“提交 + 刷新选区”，不移动、不重新进入编辑。
    host.univerAPI.executeCommand(SetCellEditVisibleOperation.id, {
      visible: false,
      keycode: 0,
      eventType: 2,
    })
      } catch {
        /* noop */
      }
    }
  }

  // 命令级拦截（与 keydown 监听顺序无关）：Univer 在 Enter/Tab 提交编辑器、或鼠标 accept 触发关闭时，
  // 都会先发 set-cell-edit-visible 命令。在其“执行前”且弹层仍开着、又是“提交型”关闭（Enter/Tab）时，
  // 把当前高亮候选记为待注入——随后该命令内部的 SetRangeValuesCommand 会由 beforeSub 换值落盘。
  // 这样无论 onKey 与 Univer 的 keydown 监听谁先执行，候选都能被注入（这是键盘 Enter 路径此前失败的根因：
  // Univer 的 window-capture keydown 监听注册早于本控制器，会先提交“旧文本”）。
  const beforeVisSub = commandService.beforeCommandExecuted((cmd: { id?: string; params?: any }) => {
    if (cmd.id !== SetCellEditVisibleOperation.id) return
    if (!cmd.params || cmd.params.visible !== false) return
    const kc = cmd.params.keycode
    // 仅 Enter(13)/Tab(9) 这类“提交并关闭”才注入候选；ESC(27) 或其它关闭不注入。
    if (kc !== 13 && kc !== 9) return
    if (!items.length) return
    // Excel/WPS 语义：Enter 只有用方向键【主动选择过】候选时才回填高亮候选；
    // 直接 Enter（未选过候选）提交的是用户自己输入的内容——防止“输入 8.7 被换成上方最近候选 170”。
    // Tab 则始终回填高亮候选（与 Excel Tab 自动完成一致）。
    if (kc === 13 && !navigated) return
    const idx = index >= 0 && index < items.length ? index : 0
    armAccept(items[idx])
  })

  // 拦截「单元格提交命令」：在 Univer 把编辑器内容写入单元格之前，把提交值替换为候选。
  // 这是唯一可靠的注入点——既绕开 currentEditCell$.row 的 +1 偏移（落盘由 Univer 用 editCellState 的正确 row），
  // 也绕开 replaceText 的快照同步时序问题（我们直接改的是提交命令的参数，与原生提交同帧生效）。
  const beforeSub = commandService.beforeCommandExecuted((cmd: { id?: string; params?: any }) => {
    if (!pendingAccept) return
    if (cmd.id !== SetRangeValuesCommand.id) return
    if (!cmd.params) return
    const item = pendingAccept
    const raw = item.raw != null ? item.raw : item.display
    const isNum = typeof raw === 'number'
    // 保留原有 value 的其它字段（如样式 s），仅覆盖 v / t。
    const prev = typeof cmd.params.value === 'object' && cmd.params.value ? cmd.params.value : {}
    cmd.params.value = {
      ...prev,
      v: raw,
      t: isNum ? CellValueType.NUMBER : CellValueType.STRING,
    }
    // 日期列回填的是日期序列号（number）：显式补上 yyyy-mm-dd 显示格式，
    // 否则单元格会按默认格式把 45866 显示成裸数字（列级默认样式对编辑提交写入的单元格不总是生效）。
    if (acceptCol === DATE_COL_INDEX && isNum) {
      cmd.params.value.s = { n: { pattern: 'yyyy-mm-dd' } }
    }
    // 记录历史，让自动补全后续能建议该值（acceptCol 而非 editCol：endEdit 可能已清零 editCol）。
    try { host.commit(item, acceptCol) } catch { /* noop */ }
    pendingAccept = null
    clearPendingTimer()
  })

  const sub = editorBridge.currentEditCell$.subscribe((state: any) => {
    if (!state) {
      endEdit()
      return
    }
    const row = state.row
    const col = state.column
    const editorId = state.editorUnitId
    const p = state.position
    if (row == null || col == null) {
      endEdit()
      return
    }
    const changed = editorId !== editEditorId
    editRow = row
    editCol = col
    editEditorId = editorId
    posRef = p || null
    accepting = false
    pendingAccept = null
    clearPendingTimer()
    // 仅在编辑器切换 / 尚未订阅时才重新挂载 input 订阅（避免每次输入都重挂）
    if (changed || !editorSub) {
      lastKey = '' // 强制重新计算（同一格重新进入也可能内容不同）
      attachEditor()
    }
  })

  const visSub = editorBridge.visible$.subscribe((v: any) => {
    if (v && v.visible === false) endEdit()
  })

  const onKey = (e: KeyboardEvent) => {
    if (e.key === 'ArrowDown') {
      if (!items.length) return
      e.preventDefault()
      e.stopPropagation()
      index = (index + 1) % items.length
      // 用户已主动用方向键选择候选 → 之后的 Enter 才回填高亮候选（Excel/WPS 语义）
      navigated = true
      host.render({ open: true, items, index, pos })
    } else if (e.key === 'ArrowUp') {
      if (!items.length) return
      e.preventDefault()
      e.stopPropagation()
      index = (index - 1 + items.length) % items.length
      navigated = true
      host.render({ open: true, items, index, pos })
    } else if (e.key === 'Enter' || e.key === 'Tab') {
      // 命中候选时不在此拦截：Univer 原生 Enter/Tab 会先发 set-cell-edit-visible 命令，
      // 由 beforeVisSub 在命令执行前把高亮候选记为待注入，再由 beforeSub 换值落盘。
      // 故此处直接放行（不 preventDefault），让原生提交走完；弹层由 endEdit 在命令结束后收起。
      // （Tab 始终回填高亮/首条候选，即 Excel/WPS“上方最近优先”；方向键已先行调整 index 并置 navigated。）
    } else if (e.key === 'Escape') {
      // 第一次 ESC 只关弹层、保留编辑（与 Excel/WPS 一致）；拦截默认，避免原生 ESC 取消整格编辑。
      // 弹层已关后 items 清空，下一轮 onKey 直接 early-return，原生 ESC 才会取消整格编辑。
      if (items.length) {
        e.preventDefault()
        e.stopPropagation()
        closePopup()
      }
    }
  }
  // 挂在 window 捕获阶段：比 Univer 在 document 上的 keydown 监听更早触发，
  // 才能在它“先提交旧文本”之前完成 replaceText / 拦截。
  window.addEventListener('keydown', onKey, true)

  const setIndex = (i: number) => {
    if (!items.length) return
    index = ((i % items.length) + items.length) % items.length
    host.render({ open: true, items, index, pos })
  }

  return {
    dispose: () => {
      try { sub.dispose() } catch {}
      try { visSub.dispose() } catch {}
      try { beforeVisSub.dispose() } catch {}
      try { beforeSub.dispose() } catch {}
      if (editorSub) { try { editorSub.dispose() } catch {} }
      clearPendingTimer()
      clearAcceptGuard()
      window.removeEventListener('keydown', onKey, true)
    },
    setIndex,
    accept,
  }
}
