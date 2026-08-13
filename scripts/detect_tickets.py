#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
账单小票检测器（YOLOv8）
======================
从标准输入读取一张图片的 base64 编码（JPEG/PNG），用训练好的 YOLOv8 模型
（默认 models/ticket_detect.pt）检测出图中每一张小票的矩形边界框，
并将每张小票裁剪出来，以 base64 形式随结果一并返回。

输出（stdout，单行 JSON）：
  {
    "ok": true,
    "image_width":  W,
    "image_height": H,
    "boxes": [ {x, y, w, h, conf, cls, label}, ... ],   # 像素坐标（基于输入图）
    "crops": [ "base64 jpeg", ... ]                       # 与 boxes 一一对应
  }
若失败（缺依赖 / 模型缺失 / 推理异常），输出 {"ok": false, "message": "..."} 并退出 0，
以便 Node 侧解析后优雅回退。

用法（由 Electron 主进程调用）：
  python detect_tickets.py <model_path> [--conf 0.25] [--iou 0.45] [--no-crops]
图片经 stdin 传入；坐标与裁剪均基于传入的图片（即预览缩放后的图），
因此可直接对齐前端预览、并直接喂给视觉模型做单票识别。

注意：推理参数（conf / iou / imgsz / device）与训练后验证脚本 detect.py 保持一致，
其中 imgsz=640、iou=0.45 是该检测模型在训练阶段标定出的最佳组合，
改用其它值（如 imgsz=1280 或默认 iou=0.7）会导致漏检/多检，因此请勿随意改动。
"""

import sys
import io
import os
import json
import base64
import argparse


def fail(message):
    """统一失败输出：打印 JSON 到 stdout 并以退出码 0 结束，避免 Node 侧死等 stderr。"""
    sys.stdout.write(json.dumps({"ok": False, "message": str(message)}, ensure_ascii=False))
    sys.stdout.flush()
    sys.exit(0)


def _row_projection_variance(gray):
    """计算灰度图「逐行墨量方差」：文字正向排列时，每行墨量在文本行处出现明显峰谷，方差最大。
    用于在 0/90/180/270 四个候选角度中挑出文字最「正」的那一个。"""
    w, h = gray.size
    if h == 0 or w == 0:
        return 0
    px = list(gray.getdata())
    sums = [0] * h
    for i, v in enumerate(px):
        sums[i // w] += (255 - v)  # 墨（深色）= 高值
    mean = sum(sums) / h
    var = sum((s - mean) * (s - mean) for s in sums) / h
    return var


def best_rotation(img):
    """在 0/90/180/270 中选使行投影方差最大的角度（即文字行最水平的角度）。"""
    # 缩小后计算，省时又不显著影响方向判断
    small = img.copy()
    maxside = max(small.size)
    if maxside > 400:
        ratio = 400.0 / maxside
        small = small.resize((int(small.size[0] * ratio), int(small.size[1] * ratio)))
    gray = small.convert("L")
    best_angle = 0
    best_score = -1.0
    for angle in (0, 90, 180, 270):
        score = _row_projection_variance(gray.rotate(angle, expand=True))
        if score > best_score:
            best_score = score
            best_angle = angle
    return best_angle


def auto_rotate(img, do_rotate=True):
    """对小票裁剪图做最佳 90° 整转，使其文字尽量正向。返回 (旋转后图, 角度)。"""
    if not do_rotate:
        return img, 0
    angle = best_rotation(img)
    if angle == 0:
        return img, 0
    return img.rotate(angle, expand=True), angle


def main():
    parser = argparse.ArgumentParser(description="YOLOv8 小票检测器")
    parser.add_argument("model", nargs="?", default="models/ticket_detect.pt",
                        help="YOLOv8 模型路径（.pt）")
    parser.add_argument("--conf", type=float, default=0.25, help="置信度阈值（与 detect.py 一致）")
    parser.add_argument("--iou", type=float, default=0.45, help="NMS IoU 阈值（与 detect.py 一致）")
    parser.add_argument("--imgsz", type=int, default=640, help="推理输入尺寸（与训练/验证一致，勿改）")
    parser.add_argument("--device", default="0", help="推理设备：GPU 编号或 cpu（不可用自动回退）")
    parser.add_argument("--no-crops", action="store_true", help="不返回裁剪图（仅检测框）")
    parser.add_argument("--no-rotate", action="store_true", help="不对裁剪图做自动旋转校正")
    parser.add_argument("--pad", type=int, default=6, help="裁剪时四周外扩像素，避免切到边缘")
    args = parser.parse_args()

    # ---- 1. 读取 stdin 的 base64 图片 ----
    try:
        raw = sys.stdin.read().strip()
        if not raw:
            fail("未接收到图片数据（stdin 为空）")
        img_bytes = base64.b64decode(raw)
    except Exception as e:  # noqa: BLE001
        fail("图片解码失败：" + str(e))
        return

    try:
        from PIL import Image
    except Exception:  # noqa: BLE001
        fail("缺少依赖 Pillow，请先 pip install pillow")
        return

    try:
        img = Image.open(io.BytesIO(img_bytes)).convert("RGB")
    except Exception as e:  # noqa: BLE001
        fail("无法打开图片：" + str(e))
        return
    W, H = img.size

    # ---- 2. 加载模型（缺 ultralytics / 模型文件时给出明确信息）----
    if not os.path.exists(args.model):
        fail("模型文件不存在：" + args.model)
        return
    try:
        from ultralytics import YOLO
    except Exception:  # noqa: BLE001
        fail("未安装 ultralytics，请先 pip install ultralytics torch")
        return
    try:
        model = YOLO(args.model)
    except Exception as e:  # noqa: BLE001
        fail("加载模型失败：" + str(e))
        return

    # ---- 3. 设备选择：优先 GPU，不可用则回退 CPU（与 detect.py 保持一致）----
    device = args.device
    if device != "cpu":
        try:
            import torch
            if not torch.cuda.is_available():
                device = "cpu"
        except Exception:  # noqa: BLE001
            device = "cpu"

    # ---- 4. 推理（参数对齐 detect.py：conf/iou/imgsz/device）----
    try:
        results = model.predict(
            source=img,
            conf=args.conf,
            iou=args.iou,
            imgsz=args.imgsz,
            device=device,
            verbose=False,
        )
    except Exception as e:  # noqa: BLE001
        fail("推理失败：" + str(e))
        return

    boxes = []
    crops = []
    pad = args.pad
    names = getattr(model, "names", {}) or {}

    for r in results:
        xyxy = getattr(r.boxes, "xyxy", None)
        if xyxy is None:
            continue
        confs = getattr(r.boxes, "conf", None)
        clss = getattr(r.boxes, "cls", None)
        for i in range(len(xyxy)):
            x1, y1, x2, y2 = [float(v) for v in xyxy[i]]
            # 裁剪到图像范围内
            x1 = max(0.0, min(float(W), x1))
            y1 = max(0.0, min(float(H), y1))
            x2 = max(0.0, min(float(W), x2))
            y2 = max(0.0, min(float(H), y2))
            if x2 - x1 < 4 or y2 - y1 < 4:
                continue  # 过小的噪声框直接丢弃
            conf = float(confs[i]) if confs is not None else 0.0
            cls = int(clss[i]) if clss is not None else 0
            box_angle = 0
            if not args.no_crops:
                cx1 = max(0, int(round(x1)) - pad)
                cy1 = max(0, int(round(y1)) - pad)
                cx2 = min(W, int(round(x2)) + pad)
                cy2 = min(H, int(round(y2)) + pad)
                crop = img.crop((cx1, cy1, cx2, cy2))
                if not args.no_rotate:
                    crop, box_angle = auto_rotate(crop)
                buf = io.BytesIO()
                crop.save(buf, format="JPEG", quality=90)
                crops.append(base64.b64encode(buf.getvalue()).decode("ascii"))
            boxes.append({
                "x": int(round(x1)),
                "y": int(round(y1)),
                "w": int(round(x2 - x1)),
                "h": int(round(y2 - y1)),
                "conf": round(conf, 4),
                "cls": cls,
                "label": names.get(cls, str(cls)),
                "angle": box_angle,
            })

    # 按从上到下、从左到右排序，保证逐张识别顺序稳定
    boxes.sort(key=lambda b: (b["y"], b["x"]))
    if not args.no_crops:
        # 重新按排序后的 box 顺序对齐 crops
        crops = [crops[i] for i in sorted(range(len(boxes)), key=lambda k: (boxes[k]["y"], boxes[k]["x"]))]

    out = {
        "ok": True,
        "image_width": W,
        "image_height": H,
        "boxes": boxes,
        "crops": crops if not args.no_crops else [],
    }
    sys.stdout.write(json.dumps(out, ensure_ascii=False))
    sys.stdout.flush()


if __name__ == "__main__":
    main()
