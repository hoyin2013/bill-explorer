import { nativeImage } from 'electron'
import { readdirSync, statSync } from 'fs'
import { extname, join, basename } from 'path'

const IMAGE_EXTS = new Set(['.jpg', '.jpeg', '.png', '.webp', '.bmp', '.gif'])

export interface ImageItem {
  name: string
  path: string
}

export function listImages(dir: string): { error?: boolean; message?: string; images?: ImageItem[] } {
  if (!dir) {
    return { error: true, message: '图片目录未设置。' }
  }
  try {
    const entries = readdirSync(dir)
    const images: ImageItem[] = []
    for (const name of entries) {
      const lower = name.toLowerCase()
      if (!IMAGE_EXTS.has(extname(lower))) continue
      const full = join(dir, name)
      try {
        const st = statSync(full)
        if (st.isFile()) images.push({ name, path: full })
      } catch {
        // skip unreadable
      }
    }
    images.sort((a, b) => a.name.localeCompare(b.name, 'zh-CN'))
    return { images }
  } catch (err) {
    return { error: true, message: '读取图片目录失败：' + (err instanceof Error ? err.message : '未知错误') }
  }
}

export function readImageBase64(filePath: string, maxWidth = 1600): { error?: boolean; message?: string; base64?: string; mime?: string } {
  try {
    const img = nativeImage.createFromPath(filePath)
    if (img.isEmpty()) {
      return { error: true, message: '无法读取该图片文件。' }
    }
    const size = img.getSize()
    const resized = size.width > maxWidth ? img.resize({ width: maxWidth }) : img
    const buf = resized.toJPEG(85)
    return { base64: buf.toString('base64'), mime: 'image/jpeg' }
  } catch (err) {
    return { error: true, message: '读取图片失败：' + (err instanceof Error ? err.message : '未知错误') }
  }
}

export function imageFileName(filePath: string): string {
  return basename(filePath)
}
