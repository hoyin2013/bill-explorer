import { FileEntry } from './types'

/**
 * 实时过滤：在内存索引中做匹配，任意命中即展示。
 * 匹配维度：文件名、文件路径、全拼、拼音首字母缩写，全部小写比较。
 */
export function filterFiles(files: FileEntry[], query: string): FileEntry[] {
  const q = query.trim().toLowerCase()
  if (!q) return files

  return files.filter((f) =>
    f.fileNameLower.includes(q) ||
    f.pathLower.includes(q) ||
    f.pyFull.includes(q) ||
    f.pyInit.includes(q),
  )
}
