#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
一次性把训练好的 YOLOv8 模型导出为 ONNX（仅此步需要 Python + ultralytics + torch）。
导出后运行时用 onnxruntime-node，完全不需要 Python / ultralytics / torch。

用法：
    python export_onnx.py models/ticket_detect.pt models/ticket_detect.onnx
    # 或省略参数，默认导出到 models/ticket_detect.onnx
"""
import sys
import os

try:
    from ultralytics import YOLO
except Exception as e:  # noqa: BLE001
    print("缺少依赖 ultralytics/torch，请先：pip install ultralytics torch")
    print("错误：", e)
    sys.exit(1)


def main():
    here = os.path.dirname(os.path.abspath(__file__))
    root = os.path.dirname(here)
    src = sys.argv[1] if len(sys.argv) > 1 else os.path.join(root, "models", "ticket_detect.pt")
    if not os.path.exists(src):
        print("源模型不存在：", src)
        sys.exit(1)

    model = YOLO(src)
    # dynamic=False 与 detect.py 推理参数一致（imgsz=640, iou=0.45, conf=0.25）
    exported = model.export(format="onnx", imgsz=640, dynamic=False, simplify=True, half=False)
    print("已导出 ONNX：", exported)


if __name__ == "__main__":
    main()
