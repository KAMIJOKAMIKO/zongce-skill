# -*- coding: utf-8 -*-
"""ocr_scan.py - 对扫描版 PDF 做 OCR 预筛，找出包含指定姓名的页码。

适用于 extract_pages.py 判定为"无文字层"的大型扫描 PDF（动辄几十上百页），
避免逐页人工查看。OCR 结果仅供定位候选页，最终必须人工查看页面图片确认。

用法：
  python ocr_scan.py --name 张三 --out hits.json file1.pdf file2.pdf ...
  python ocr_scan.py --name 张三 --dpi 150 --pages 31-68 file.pdf

输出 JSON：{文件: {"total_pages": N, "hits": [页码...], "snippets": {页码: 命中行}}}
"""
import argparse
import json
import os
import re
import unicodedata


def norm(s: str) -> str:
    s = unicodedata.normalize("NFKC", s or "")
    return re.sub(r"\s+", "", s)


def parse_pages(spec, total):
    if not spec:
        return list(range(total))
    pages = set()
    for part in spec.split(","):
        if "-" in part:
            a, b = part.split("-", 1)
            pages.update(range(int(a) - 1, int(b)))
        else:
            pages.add(int(part) - 1)
    return sorted(p for p in pages if 0 <= p < total)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("files", nargs="+")
    ap.add_argument("--name", required=True)
    ap.add_argument("--out", default="ocr_hits.json")
    ap.add_argument("--dpi", type=int, default=150)
    ap.add_argument("--pages", default="", help="页码范围，如 31-68（1-based），默认全部")
    args = ap.parse_args()

    import fitz
    import numpy as np
    from rapidocr_onnxruntime import RapidOCR

    engine = RapidOCR()
    target = norm(args.name)
    # 宽松全名匹配：允许字间插入 0-2 个字符（OCR 噪声 / 排版空格）
    if len(target) >= 2:
        loose = re.compile(".{0,2}".join(map(re.escape, target)))
    else:
        loose = re.compile(re.escape(target))
    # 部分匹配（防 OCR 单字误识别导致漏报）：姓+名首字、名尾两字
    partials = []
    if len(target) >= 3:
        partials = [target[:2], target[-2:]]
    elif len(target) == 2:
        partials = [target[0]]

    result = {}
    for f in args.files:
        f = os.path.abspath(f)
        doc = fitz.open(f)
        total = len(doc)
        pages = parse_pages(args.pages, total)
        hits, partial_hits, snippets = [], [], {}
        zoom = args.dpi / 72.0
        mat = fitz.Matrix(zoom, zoom)
        for i, p in enumerate(pages):
            pix = doc[p].get_pixmap(matrix=mat, alpha=False)
            img = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.height, pix.width, pix.n)
            ocr_out, _ = engine(img)
            lines = [box[1] for box in (ocr_out or [])]
            page_text = norm("".join(lines))
            if target in page_text or loose.search(page_text):
                hits.append(p + 1)
                snippets[str(p + 1)] = [l for l in lines if any(ch in l for ch in target[:1])][:5]
                tag = "  << HIT"
            elif any(pt in page_text for pt in partials):
                partial_hits.append(p + 1)
                snippets[str(p + 1)] = [l for l in lines
                                        if any(pt in norm(l) for pt in partials)][:5]
                tag = "  << PARTIAL"
            else:
                tag = ""
            print(f"[{os.path.basename(f)}] {i + 1}/{len(pages)} page {p + 1}{tag}", flush=True)
        result[f] = {"total_pages": total, "hits": hits,
                     "partial_hits": partial_hits, "snippets": snippets}
        doc.close()

    with open(args.out, "w", encoding="utf-8") as fp:
        json.dump(result, fp, ensure_ascii=False, indent=2)
    print(f"\nsaved: {args.out}")
    for f, r in result.items():
        print(f"- {os.path.basename(f)}: 共{r['total_pages']}页, "
              f"命中页={r['hits']}, 部分命中页={r['partial_hits']}")


if __name__ == "__main__":
    main()
