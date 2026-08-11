import { existsSync, promises as fs } from 'fs'
import { join, dirname, extname } from 'path'
import { shell } from 'electron'
import { pinyin } from 'pinyin-pro'

// ============ 文件扫描：递归遍历目录，只收集 .xlsx / .xls 文件 ============
export async function scanDirectory(rootDir: string): Promise<{
  error?: boolean
  message?: string
  files?: FileEntry[]
}> {
  // 容错：目录不存在
  if (!existsSync(rootDir)) {
    return { error: true, message: '目录不存在或已删除，请重新选择。' }
  }

  const files: FileEntry[] = []
  const seen = new Set<string>() // 防循环引用
  await walk(rootDir)

  return { files }

  async function walk(dir: string) {
    let entries
    try {
      entries = await fs.readdir(dir, { withFileTypes: true })
    } catch {
      // 容错：某子目录权限不足时跳过，不崩溃
      return
    }
    for (const entry of entries) {
      const fullPath = join(dir, entry.name)
      // 跳过 . 开头的隐藏文件/目录
      if (entry.name.startsWith('.')) continue
      if (entry.isDirectory()) {
        await walk(fullPath)
      } else if (entry.isFile()) {
        const ext = extname(entry.name).toLowerCase()
        // 只收集 xlsx / xls
        if (ext === '.xlsx' || ext === '.xls') {
          if (!seen.has(fullPath)) {
            seen.add(fullPath)
            // 预生成拼音索引，全部转小写
            const index = buildIndex(entry.name, fullPath)
            files.push({
              fileName: entry.name,
              filePath: fullPath,
              ...index,
            })
          }
        }
      }
    }
  }
}

// ============ 为单条文件预生成搜索索引 ============
function buildIndex(fileName: string, filePath: string): FileIndex {
  // 取不带扩展名的文件名用于拼音（加 i 标志兼容 Windows 上的 .XLSX / .XLS 大写扩展名）
  const baseName = fileName.replace(/\.(xlsx|xls)$/i, '')
  const fileNameLower = fileName.toLowerCase()
  const pathLower = filePath.toLowerCase()

  // 容错：生僻字拼音解析失败时跳过该字符，不崩溃
  const pyFull = safePinyinFull(baseName)
  const pyInit = safePinyinInit(baseName)

  return {
    fileNameLower,
    pathLower,
    pyFull: pyFull.toLowerCase(),
    pyInit: pyInit.toLowerCase(),
  }
}

// 容错包装：获取汉字全拼
function safePinyinFull(text: string): string {
  try {
    // pinyin-pro v3：type='array' 返回每字的拼音数组，join 拼接成全拼
    return pinyin(text, { toneType: 'none', type: 'array' }).join('')
  } catch {
    return ''
  }
}

// 容错包装：获取汉字拼音首字母缩写
function safePinyinInit(text: string): string {
  try {
    // pinyin-pro v3 的 mode='initials' 不生效（返回全拼）；
    // 正确取首字母：取每字全拼的首字符。王金玉 → w/j/y → 'wjy'
    return pinyin(text, { toneType: 'none', type: 'array' })
      .map((seg) => seg[0])
      .join('')
  } catch {
    return ''
  }
}

// ============ 调用系统默认程序打开文件 ============
export function openFile(filePath: string): {
  error?: boolean
  message?: string
} {
  try {
    shell.openPath(filePath)
    return {}
  } catch (err) {
    const message = err instanceof Error ? err.message : '打开文件失败'
    return { error: true, message }
  }
}

// ============ 类型定义 ============
export interface FileIndex {
  fileNameLower: string
  pathLower: string
  pyFull: string      // 汉字全拼（全小写）
  pyInit: string      // 汉字拼音首字母缩写（全小写）
}

export interface FileEntry extends FileIndex {
  fileName: string
  filePath: string
}
