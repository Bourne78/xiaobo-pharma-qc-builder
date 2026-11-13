# -*- coding: utf-8 -*-
"""
OCR 管线：把扫描版中文 QC PDF（SOP / 质量标准）转成结构化文本。

本环境实测可用的组合（不要用 PyMuPDF，它会恒定 segfault exit 139）：
  - 光栅化：pypdfium2（pip install pypdfium2，无系统依赖）
  - 识别  ：rapidocr-onnxruntime 1.2.3（自带 PP-OCRv3 中文模型）
  - ⚠️ onnxruntime 必须用 1.20.1：系统版 1.31.1 在本机会 exit 139。
         把 1.20.1 装到隔离位置后，用 sys.path 插到最前再 import rapidocr。

用法：
  python ocr_pipeline.py <input.pdf> [--out ocr.json] [--scale 1.6] [--band 760] [--overlap 120]
  python ocr_pipeline.py <input.pdf> --rotate90   # 处理 90° 旋转横向表（质量标准 P3~P6）

输出 JSON：{ "pages": [ { "page": 1, "lines": [ {"text","conf","y"} ] } ] }
"""
import os
import sys
import json
import argparse

# ---- onnxruntime 1.20.1 隔离副本：把其 site-packages 插到 sys.path 最前 ----
# 安装：python -m venv .venv_ort120 && .venv_ort120\Scripts\activate &&
#       pip install "onnxruntime==1.20.1" "rapidocr-onnxruntime==1.2.3"
ORT120_PATH = os.environ.get("ORT120_PATH", "")   # 设为隔离 venv 的 Lib/site-packages
if ORT120_PATH and os.path.isdir(ORT120_PATH):
    sys.path.insert(0, ORT120_PATH)

import pypdfium2 as pdfium
from PIL import Image
from rapidocr_onnxruntime import RapidOCR


# 分块扫描参数：整页 OCR 会静默漏行（曾把 "3倍(1.5%)" 漏成 "0.6倍(0.3%)"），
# 必须分块重叠补扫。SCALE 放大提升小字识别率。
BAND = 760          # 每块高度(px @ scale=1.0 基准前的图高占比)
OVERLAP = 120       # 块间重叠(px)
SCALE = 1.6         # 渲染放大倍数


def render_page(path, idx, scale=SCALE):
    pdf = pdfium.PdfDocument(path)
    page = pdf[idx]
    if os.environ.get("ROTATE90"):
        # 90° 旋转横向表：rotate(-90, expand=True) 后按 x 聚类
        bitmap = page.render(scale=scale).to_pil().rotate(-90, expand=True)
    else:
        bitmap = page.render(scale=scale).to_pil()
    return bitmap


def band_ocr(ocr, img, band=BAND, overlap=OVERLAP):
    """把图按水平条带重叠切分后逐带 OCR，合并去重。"""
    W, H = img.size
    lines = []
    step = max(1, band - overlap)
    y = 0
    while y < H:
        y2 = min(H, y + band)
        crop = img.crop((0, y, W, y2))
        # RapidOCR 返回 ([ [x1,y1,x2,y2,x3,y3,x4,y4], ... ], [text,...], [conf,...])
        result, _ = ocr(crop)
        if not result:
            y += step
            continue
        boxes, texts, confs = result
        for box, text, conf in zip(boxes, texts, confs):
            # box 四点平均 y 作为行号，做重叠去重
            yc = sum([box[1], box[3], box[5], box[7]]) / 4.0 + y   # 还原到整图坐标
            lines.append({"text": text, "conf": float(conf), "y": yc})
        y += step
    return lines


def dedupe_near(lines, dy=45):
    """同 y±dy 内保留最长文本，去掉重复扫描行。"""
    lines.sort(key=lambda d: d["y"])
    out = []
    for ln in lines:
        dup = False
        for o in out:
            if abs(o["y"] - ln["y"]) <= dy:
                if len(ln["text"]) > len(o["text"]):
                    o["text"] = ln["text"]; o["conf"] = ln["conf"]
                dup = True
                break
        if not dup:
            out.append(ln)
    return out


def ocr_pdf(path, out=None, scale=SCALE, band=BAND, overlap=OVERLAP):
    ocr = RapidOCR()
    pdf = pdfium.PdfDocument(path)
    n = len(pdf)
    pages = []
    for i in range(n):
        img = render_page(path, i, scale=scale)
        lines = band_ocr(ocr, img, band=band, overlap=overlap)
        lines = dedupe_near(lines)
        pages.append({"page": i + 1, "lines": lines})
        print(f"[page {i+1}/{n}] {len(lines)} lines")
    result = {"source": path, "pages": pages}
    if out:
        with open(out, "w", encoding="utf-8") as f:
            json.dump(result, f, ensure_ascii=False, indent=2)
        print("wrote", out)
    return result


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("pdf")
    ap.add_argument("--out", default="ocr.json")
    ap.add_argument("--scale", type=float, default=SCALE)
    ap.add_argument("--band", type=int, default=BAND)
    ap.add_argument("--overlap", type=int, default=OVERLAP)
    ap.add_argument("--rotate90", action="store_true", help="处理 90° 旋转横向表")
    args = ap.parse_args()
    if args.rotate90:
        os.environ["ROTATE90"] = "1"
    ocr_pdf(args.pdf, out=args.out, scale=args.scale, band=args.band, overlap=args.overlap)


if __name__ == "__main__":
    main()
