import { nativeImage } from 'electron'
import { readdirSync, readFileSync } from 'fs'
import { extname, join, basename, isAbsolute } from 'path'
import convert from 'heic-convert'

const IMAGE_EXTS = new Set([
  '.jpg', '.jpeg', '.jpe', '.jfif',
  '.png', '.webp', '.bmp', '.gif', '.tif', '.tiff',
  // HEIC/HEIF：iPhone 等拍摄的默认格式，跨平台统一用 libheif 转码（见 readImageBase64）
  '.heic', '.heif',
])

// HEIC/HEIF：Electron 的 nativeImage 在 Windows/Linux 上无法解码，需走 libheif 转码分支
const HEIC_EXTS = new Set(['.heic', '.heif'])

// Windows 长路径（>260 字符）处理：给绝对路径加 \\?\ 前缀，规避长路径打开限制，
// 并兼容部分仍强制 MAX_PATH 的 Windows 环境。\\?\ 前缀要求反斜杠，故先归一化。
function toLongPath(p: string): string {
  if (process.platform !== 'win32' || !p) return p
  if (p.startsWith('\\\\?\\')) return p
  if (!isAbsolute(p)) return p
  return '\\\\?\\' + p.split('/').join('\\')
}

// 递归扫描时跳过这些目录名（系统/隐藏/缓存目录，避免误扫与卡顿）
const SKIP_DIRS = new Set(['node_modules', '.git', '$RECYCLE.BIN', 'System Volume Information'])

export interface ImageItem {
  name: string
  path: string
}

// 递归列出目录下所有图片（含子目录，最大深度 MAX_DEPTH），避免「图片在子文件夹里
// 却提示没有」的常见误报。用 withFileTypes 直接靠 dirent 判文件/目录，少一次 stat 调用。
const MAX_DEPTH = 5

export function listImages(dir: string): { error?: boolean; message?: string; images?: ImageItem[] } {
  if (!dir) {
    return { error: true, message: '图片目录未设置。' }
  }
  try {
    const images: ImageItem[] = []
    let totalFiles = 0
    const stack: Array<{ dir: string; depth: number }> = [{ dir, depth: 0 }]
    while (stack.length) {
      const { dir: cur, depth } = stack.pop()!
      let entries
      try {
        entries = readdirSync(toLongPath(cur), { withFileTypes: true })
      } catch {
        try {
          entries = readdirSync(cur, { withFileTypes: true })
        } catch {
          continue // 该层无法读取则跳过（无权限等）
        }
      }
      for (const d of entries) {
        if (d.isDirectory()) {
          if (depth < MAX_DEPTH && !SKIP_DIRS.has(d.name.toLowerCase())) {
            stack.push({ dir: join(cur, d.name), depth: depth + 1 })
          }
          continue
        }
        if (!d.isFile()) continue
        totalFiles++
        const lower = d.name.toLowerCase()
        if (!IMAGE_EXTS.has(extname(lower))) continue
        images.push({ name: d.name, path: join(cur, d.name) })
      }
    }
    images.sort((a, b) => a.name.localeCompare(b.name, 'zh-CN'))
    if (images.length === 0 && totalFiles > 0) {
      // 扫到了文件，但没有一个是支持的图片格式 —— 给一句可自查的提示
      return {
        images: [],
        message: `已递归扫描子目录，共 ${totalFiles} 个文件，但没有被识别为图片（支持 .jpg/.jpeg/.png/.webp/.bmp/.gif/.tif/.heic 等常见格式）。`,
      }
    }
    return { images }
  } catch (err) {
    return { error: true, message: '读取图片目录失败：' + (err instanceof Error ? err.message : '未知错误') }
  }
}

export async function readImageBase64(filePath: string, maxWidth = 1600): Promise<{ error?: boolean; message?: string; base64?: string; mime?: string }> {
  try {
    // HEIC/HEIF：nativeImage 在 Windows/Linux 上解不了，统一用 libheif 转成 JPEG 再走公共缩放/编码逻辑
    if (HEIC_EXTS.has(extname(filePath.toLowerCase()))) {
      const raw = readFileSync(toLongPath(filePath) || filePath)
      const jpegBuf = await convert({ buffer: raw, format: 'JPEG', quality: 0.92 })
      return encodeJpeg(jpegBuf, maxWidth)
    }
    // 其余格式走原生解码：先尝试长路径前缀；若 Electron 未识别该前缀导致读取失败，回退原始路径
    let img = nativeImage.createFromPath(toLongPath(filePath))
    if (img.isEmpty()) img = nativeImage.createFromPath(filePath)
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

// 把 JPEG/PNG 等已解码 buffer 按最长边缩放到 maxWidth 以内并重编码为 JPEG（降低 token 与推理耗时）
function encodeJpeg(buf: Buffer, maxWidth: number): { base64: string; mime: string } {
  const img = nativeImage.createFromBuffer(buf)
  const size = img.getSize()
  const resized = size.width > maxWidth ? img.resize({ width: maxWidth }) : img
  return { base64: resized.toJPEG(85).toString('base64'), mime: 'image/jpeg' }
}

// 把 base64 图片按最长边缩放到 maxDim 以内（JPEG），用于降低视觉模型的 token 消耗与推理耗时。
// 已在阈值内则原样返回（避免无意义的重编码）。视觉模型按像素量计费视觉 token，
// 单张小票裁剪图适度降采样通常不影响手写文字识别，却能显著省 token。
export function downscaleImageBase64(
  base64: string,
  maxDim = 1280,
  quality = 80,
): string {
  try {
    const buf = Buffer.from(base64, 'base64')
    const img = nativeImage.createFromBuffer(buf)
    if (img.isEmpty()) return base64
    const size = img.getSize()
    const longest = Math.max(size.width, size.height)
    if (longest <= maxDim) return base64
    const scale = maxDim / longest
    const resized = img.resize({
      width: Math.max(1, Math.round(size.width * scale)),
      height: Math.max(1, Math.round(size.height * scale)),
    })
    return resized.toJPEG(quality).toString('base64')
  } catch {
    return base64
  }
}

export function imageFileName(filePath: string): string {
  return basename(filePath)
}
