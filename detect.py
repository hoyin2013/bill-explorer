"""
目标检测脚本

功能：
    - 使用训练好的模型进行小票检测
    - 支持单张图片、文件夹和视频输入
    - 输出检测结果（坐标、置信度）
    - 可视化检测框

使用方法：
    python scripts/detect.py --model runs/train/ticket_detector/weights/best.pt --source pic/
    python scripts/detect.py --model best.pt --source pic/example.jpg --conf 0.5
    python scripts/detect.py --model best.pt --source 0  # 实时摄像头
"""

import argparse
import os
import json
from pathlib import Path

from ultralytics import YOLO
import cv2
import numpy as np


def save_results(results, output_dir, source_name):
    """保存检测结果为JSON格式"""
    detections = []
    
    for result in results:
        if result.boxes is not None:
            for box in result.boxes:
                detection = {
                    "class": int(box.cls.item()),
                    "class_name": result.names[int(box.cls.item())],
                    "confidence": float(box.conf.item()),
                    "bbox": {
                        "x1": float(box.xyxy[0, 0].item()),
                        "y1": float(box.xyxy[0, 1].item()),
                        "x2": float(box.xyxy[0, 2].item()),
                        "y2": float(box.xyxy[0, 3].item()),
                        "width": float(box.xywh[0, 2].item()),
                        "height": float(box.xywh[0, 3].item()),
                    }
                }
                detections.append(detection)
    
    # 保存为JSON
    output_json = os.path.join(output_dir, f"{source_name}.json")
    with open(output_json, 'w', encoding='utf-8') as f:
        json.dump(detections, f, indent=2, ensure_ascii=False)
    
    return detections


def print_results(source_name, detections):
    """打印检测结果"""
    print(f"\n{'='*60}")
    print(f"文件: {source_name}")
    print(f"检测到 {len(detections)} 张小票")
    print(f"{'='*60}")
    
    for i, det in enumerate(detections, 1):
        print(f"\n小票 #{i}:")
        print(f"  类别: {det['class_name']}")
        print(f"  置信度: {det['confidence']:.2%}")
        print(f"  坐标 (x1, y1, x2, y2): ({det['bbox']['x1']:.0f}, {det['bbox']['y1']:.0f}, "
              f"{det['bbox']['x2']:.0f}, {det['bbox']['y2']:.0f})")
        print(f"  尺寸 (宽×高): {det['bbox']['width']:.0f}×{det['bbox']['height']:.0f}")


def main():
    parser = argparse.ArgumentParser(description="使用YOLO进行小票检测")
    parser.add_argument("--model", default="runs/train/ticket_detector/weights/best.pt",
                        help="模型权重文件路径")
    parser.add_argument("--source", default="pic/",
                        help="输入源（图片文件、文件夹、视频或摄像头）")
    parser.add_argument("--conf", type=float, default=0.25,
                        help="置信度阈值（默认: 0.25）")
    parser.add_argument("--iou", type=float, default=0.45,
                        help="NMS IoU阈值（默认: 0.45）")
    parser.add_argument("--imgsz", type=int, default=640,
                        help="输入图像大小（默认: 640）")
    parser.add_argument("--device", default="0",
                        help="GPU设备号（默认: 0）")
    parser.add_argument("--output", default="runs/detect",
                        help="输出目录（默认: runs/detect）")
    parser.add_argument("--save", action="store_true", default=True,
                        help="保存结果")
    parser.add_argument("--no-save", dest="save", action="store_false",
                        help="不保存结果")
    parser.add_argument("--save-txt", action="store_true",
                        help="保存结果为txt格式")
    parser.add_argument("--save-json", action="store_true", default=True,
                        help="保存结果为JSON格式")
    parser.add_argument("--visualize", action="store_true",
                        help="显示可视化结果")
    
    args = parser.parse_args()
    
    print("=" * 60)
    print("YOLO 小票检测")
    print("=" * 60)
    
    # 检查模型文件
    if not os.path.exists(args.model):
        print(f"❌ 错误: 模型文件不存在: {args.model}")
        return
    
    print(f"\n✓ 加载模型: {args.model}")
    model = YOLO(args.model)
    
    # 创建输出目录
    os.makedirs(args.output, exist_ok=True)
    
    print(f"✓ 输出目录: {args.output}")
    print(f"\n检测参数:")
    print(f"  置信度阈值: {args.conf}")
    print(f"  NMS阈值: {args.iou}")
    print(f"  输入尺寸: {args.imgsz}x{args.imgsz}")
    
    # 检查设备可用性
    device = args.device
    if device != "cpu":
        import torch
        if not torch.cuda.is_available():
            print(f"⚠️  GPU不可用，自动切换到CPU")
            device = "cpu"
    
    # 开始检测
    print(f"\n开始检测...")
    print("-" * 60)
    
    results = model.predict(
        source=args.source,
        conf=args.conf,
        iou=args.iou,
        imgsz=args.imgsz,
        device=device,
        save=args.save,
        save_txt=args.save_txt,
        project=args.output,
        name="predictions",
        verbose=True,
    )
    
    # 处理结果
    print("\n" + "=" * 60)
    print("检测完成！")
    print("=" * 60)
    
    # 统计总检测数
    total_detections = 0
    
    # 按文件保存结果
    for i, result in enumerate(results):
        source_name = Path(result.path).stem
        detections = save_results([result], args.output, source_name)
        total_detections += len(detections)
        print_results(source_name, detections)
    
    print(f"\n总计: {len(results)} 个文件，{total_detections} 张小票")
    print(f"\n✓ 结果已保存到: {args.output}/")
    print(f"  - 图片: predictions/ (带检测框的结果图)")
    print(f"  - JSON: *.json (检测坐标和置信度)")
    

if __name__ == "__main__":
    main()
