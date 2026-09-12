# -*- coding: utf-8 -*-
"""extract_pages.py - 从活动证明文件中定位包含指定姓名的页并导出为整页图片。

支持的输入：
  - 图片：.png .jpg .jpeg .bmp .webp（直接视为单页，需视觉核对）
  - PDF：优先用文字层搜索姓名；无文字层或文字层搜不到姓名时，渲染全部页面供视觉识别
  - Word：.docx .doc（先经 Word COM 转为 PDF，再按 PDF 流程处理，需本机安装 MS Word）

输出：
  <outdir>/<文件stem>/page_<页码>.png     每页一张整页图片
  <outdir>/manifest.json                 提取结果清单（供后续流程使用）

用法：
  python extract_pages.py --outdir OUT --name 张三 file1.pdf file2.docx img.png
  python extract_pages.py --outdir OUT --name 张三 --neighbors --dpi 200 proofs.pdf
  python extract_pages.py --render "proofs.pdf:2,3" --outdir OUT   # 补渲染指定页

选项：
  --neighbors   命中页的前后相邻页也一并渲染（用于"证明跨两页"的候选）
  --dpi N       渲染分辨率，默认 200
  --max-pages N 无文字层需全量渲染时的页数上限，默认 30
"""
import argparse
import json
import os
import re
import shutil
import sys
import unicodedata

IMG_EXTS = {".png", ".jpg", ".jpeg", ".bmp", ".webp"}
WORD_EXTS = {".docx", ".doc"}
VISUAL_CAP_DEFAULT = 30


def norm_text(s: str) -> str:
    """规范化文本：去空白、全角转半角，便于姓名匹配。"""
    s = unicodedata.normalize("NFKC", s or "")
    return re.sub(r"\s+", "", s)


def word_to_pdf(word_path: str, pdf_path: str) -> None:
    """用本机 MS Word 把 docx/doc 转成 PDF。"""
    try:
        from docx2pdf import convert
        convert(word_path, pdf_path)
        return
    except ImportError:
        pass
    import pythoncom  # type: ignore
    import win32com.client  # type: ignore

    pythoncom.CoInitialize()
    word = win32com.client.DispatchEx("Word.Application")
    word.Visible = False
    try:
        doc = word.Documents.Open(os.path.abspath(word_path), ReadOnly=True)
        doc.SaveAs(os.path.abspath(pdf_path), FileFormat=17)  # wdFormatPDF
        doc.Close(False)
    finally:
        word.Quit()


def render_pdf_pages(doc, page_indices, out_png_dir, stem, dpi):
    """把 PDF 指定页(0-based)渲染为 PNG，返回 [{page, image}]。"""
    import fitz  # PyMuPDF

    rendered = []
    zoom = dpi / 72.0
    mat = fitz.Matrix(zoom, zoom)
    for p in page_indices:
        pix = doc[p].get_pixmap(matrix=mat, alpha=False)
        img_path = os.path.join(out_png_dir, f"{stem}_page_{p + 1}.png")
        pix.save(img_path)
        rendered.append({"page": p + 1, "image": img_path})
    return rendered


def process_pdf(pdf_path, name, out_png_dir, stem, dpi, neighbors, max_pages):
    import fitz  # PyMuPDF

    doc = fitz.open(pdf_path)
    total = len(doc)
    target = norm_text(name)
    page_texts = [norm_text(doc[p].get_text()) for p in range(total)]
    has_text_layer = any(len(t) > 5 for t in page_texts)
    matched = [p for p in range(total) if target in page_texts[p]]

    entry = {
        "source_file": pdf_path,
        "kind": "pdf",
        "total_pages": total,
        "has_text_layer": has_text_layer,
        "matched_pages": [p + 1 for p in matched],
        "needs_visual": False,
        "rendered": [],
        "notes": "",
    }

    if matched:
        to_render = set(matched)
        if neighbors:
            for p in matched:
                if p - 1 >= 0:
                    to_render.add(p - 1)
                if p + 1 < total:
                    to_render.add(p + 1)
        entry["rendered"] = render_pdf_pages(doc, sorted(to_render), out_png_dir, stem, dpi)
        entry["notes"] = "文字层命中姓名；相邻页已渲染，需人工/视觉确认是否属于同一份证明"
    else:
        # 无文字层（扫描件）或有文字层但搜不到姓名 → 全量渲染供视觉识别
        cap = min(total, max_pages)
        entry["needs_visual"] = True
        entry["rendered"] = render_pdf_pages(doc, list(range(cap)), out_png_dir, stem, dpi)
        if not has_text_layer:
            entry["notes"] = "无文字层（疑似扫描件），已渲染全部页面，需逐页视觉查找姓名"
        else:
            entry["notes"] = "文字层未搜到姓名，已渲染全部页面，需逐页视觉核对（可能为图片型证明或姓名写法差异）"
        if total > max_pages:
            entry["notes"] += f"；文件共 {total} 页，仅渲染前 {max_pages} 页，可用 --render 补渲染"
    doc.close()
    return entry


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("files", nargs="*")
    ap.add_argument("--outdir", required=True)
    ap.add_argument("--name", default="")
    ap.add_argument("--dpi", type=int, default=200)
    ap.add_argument("--neighbors", action="store_true")
    ap.add_argument("--max-pages", type=int, default=VISUAL_CAP_DEFAULT)
    ap.add_argument("--render", default="", help='补渲染，格式 "文件.pdf:2,3"（页码1-based）')
    args = ap.parse_args()

    os.makedirs(args.outdir, exist_ok=True)

    # 补渲染模式
    if args.render:
        m = re.match(r"^(.+):([\d,]+)$", args.render)
        if not m:
            print("ERROR: --render 格式应为 文件.pdf:2,3", file=sys.stderr)
            sys.exit(2)
        pdf_path, pages = m.group(1), [int(x) - 1 for x in m.group(2).split(",")]
        import fitz
        stem = re.sub(r"[^\w一-鿿-]+", "_", os.path.splitext(os.path.basename(pdf_path))[0])
        png_dir = os.path.join(args.outdir, stem)
        os.makedirs(png_dir, exist_ok=True)
        doc = fitz.open(pdf_path)
        out = render_pdf_pages(doc, pages, png_dir, stem, args.dpi)
        print(json.dumps({"rendered": out}, ensure_ascii=False, indent=2))
        return

    if not args.name:
        print("ERROR: 必须通过 --name 提供姓名", file=sys.stderr)
        sys.exit(2)
    if not args.files:
        print("ERROR: 未提供任何证明文件", file=sys.stderr)
        sys.exit(2)

    manifest = {"name": args.name, "dpi": args.dpi, "files": []}
    for f in args.files:
        f = os.path.abspath(f)
        ext = os.path.splitext(f)[1].lower()
        stem = re.sub(r"[^\w一-鿿-]+", "_", os.path.splitext(os.path.basename(f))[0])
        png_dir = os.path.join(args.outdir, stem)
        os.makedirs(png_dir, exist_ok=True)
        try:
            if ext in IMG_EXTS:
                dst = os.path.join(png_dir, f"{stem}_page_1{ext if ext == '.png' else '.png'}")
                if ext == ".png":
                    shutil.copyfile(f, dst)
                else:
                    import fitz
                    pix = fitz.Pixmap(f)
                    if pix.alpha:
                        pix = fitz.Pixmap(fitz.csRGB, pix)
                    pix.save(dst)
                manifest["files"].append({
                    "source_file": f, "kind": "image", "total_pages": 1,
                    "has_text_layer": False, "matched_pages": [],
                    "needs_visual": True,
                    "rendered": [{"page": 1, "image": dst}],
                    "notes": "图片文件，需视觉确认是否包含姓名",
                })
            elif ext in WORD_EXTS:
                conv_dir = os.path.join(args.outdir, "_converted")
                os.makedirs(conv_dir, exist_ok=True)
                pdf_path = os.path.join(conv_dir, stem + ".pdf")
                word_to_pdf(f, pdf_path)
                entry = process_pdf(pdf_path, args.name, png_dir, stem, args.dpi,
                                    args.neighbors, args.max_pages)
                entry["source_file"] = f
                entry["kind"] = "word"
                manifest["files"].append(entry)
            elif ext == ".pdf":
                manifest["files"].append(
                    process_pdf(f, args.name, png_dir, stem, args.dpi,
                                args.neighbors, args.max_pages))
            else:
                manifest["files"].append({
                    "source_file": f, "kind": "unsupported", "total_pages": 0,
                    "has_text_layer": False, "matched_pages": [], "needs_visual": True,
                    "rendered": [], "notes": f"不支持的格式 {ext}",
                })
        except Exception as e:  # 单文件失败不中断整体
            manifest["files"].append({
                "source_file": f, "kind": "error", "total_pages": 0,
                "has_text_layer": False, "matched_pages": [], "needs_visual": True,
                "rendered": [], "notes": f"处理失败: {e}",
            })

    mpath = os.path.join(args.outdir, "manifest.json")
    with open(mpath, "w", encoding="utf-8") as fp:
        json.dump(manifest, fp, ensure_ascii=False, indent=2)

    # 控制台摘要
    print(f"manifest: {mpath}")
    for e in manifest["files"]:
        print(f"- {os.path.basename(e['source_file'])}: {e['kind']}, "
              f"共{e['total_pages']}页, 文字命中页={e['matched_pages']}, "
              f"需视觉识别={e['needs_visual']}, 已渲染={len(e['rendered'])}页 | {e['notes']}")


if __name__ == "__main__":
    main()
