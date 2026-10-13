"""Step 1: list the text lines of a drawing and render a gridded preview for choosing what to translate.

Usage: python extract.py <drawing.pdf|.png|.jpg> <work_dir> [--ocr]

Pages without a real text layer (fewer than 5 lines) are OCR'd with RapidOCR; --ocr forces OCR on every
page (use it when the text layer is offset from the drawn letters). OCR lines carry "ocr": true and a
"score"; overlay.py paints them white, since there is no text to delete.

Writes to <work_dir>:
  lines.json      every text line from the PDF text layer: id, page, bbox (PDF points), dir, size, text
  p<N>_grid.png   page render at 150 dpi with a coordinate grid every 50 pt (labels in PDF points)
                  and each text-layer line boxed in blue with its id
A page with no text layer (scan, or text drawn as strokes) gets only the grid. Read rects off it by hand.
Images are wrapped into a one-page PDF first (<work_dir>\\source.pdf) so the overlay step works the same.
"""
import json
import os
import sys

import pymupdf

SRC, WORK = sys.argv[1:3]
os.makedirs(WORK, exist_ok=True)

if SRC.lower().endswith(".pdf"):
    doc = pymupdf.open(SRC)
else:  # image: one page, 1 px = 1 pt at 150 dpi scale
    img = pymupdf.open(SRC)
    pdf_bytes = img.convert_to_pdf()
    doc = pymupdf.open("pdf", pdf_bytes)
    doc.save(os.path.join(WORK, "source.pdf"))
    print("image wrapped as", os.path.join(WORK, "source.pdf"), "- use it as the overlay input")

FORCE_OCR = "--ocr" in sys.argv
_ocr = None


def ocr_page(page, pno, start_id):
    """OCR a page with no text layer. Returns lines in the same format, marked "ocr": true.
    Two passes: the whole page at 150 dpi (catches bold, large and widely spaced text) and overlapping
    300 dpi tiles (catches small text that would be lost to downscaling). Duplicates keep the better score."""
    global _ocr
    try:
        if _ocr is None:
            from rapidocr_onnxruntime import RapidOCR
            _ocr = RapidOCR()
    except ImportError:
        print("  (rapidocr_onnxruntime not installed: no OCR, read rects off the grid)")
        return []
    import numpy as np
    found = []

    def add(quad, text, score, dx, dy, k):
        xs = [p[0] + dx for p in quad]
        ys = [p[1] + dy for p in quad]
        r = pymupdf.Rect(min(xs) * k, min(ys) * k, max(xs) * k, max(ys) * k)
        dup = next((f for f in found
                    if (f["r"] & r).get_area() > 0.5 * min(f["r"].get_area(), r.get_area())), None)
        if dup:
            if float(score) > dup["score"]:
                dup.update(r=r, text=text, score=float(score))
            return
        vertical = (max(ys) - min(ys)) > 2.5 * (max(xs) - min(xs)) and len(text) > 2
        found.append({"r": r, "text": text, "score": float(score), "vertical": vertical})

    for dpi, tile, over in ((150, None, 0), (300, 1600, 200)):
        pix = page.get_pixmap(dpi=dpi)
        img = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.height, pix.width, pix.n)[:, :, :3]
        tile = tile or max(pix.width, pix.height)
        for ty in range(0, pix.height, tile - over):
            for tx in range(0, pix.width, tile - over):
                res, _ = _ocr(np.ascontiguousarray(img[ty:ty + tile, tx:tx + tile]))
                for quad, text, score in res or []:
                    add(quad, text, score, tx, ty, 72 / dpi)
    out = []
    for i, f in enumerate(sorted(found, key=lambda f: (round(f["r"].y0), f["r"].x0))):
        r = f["r"]
        out.append({"id": start_id + i, "page": pno, "bbox": [round(v, 1) for v in r],
                    "dir": [0, -1] if f["vertical"] else [1, 0],
                    "size": round((r.width if f["vertical"] else r.height) * 0.85, 1),
                    "text": f["text"], "ocr": True, "score": round(f["score"], 2)})
    return out


lines, n = [], 0
for pno, page in enumerate(doc):
    if page.rotation:  # rotated page: take everything in the coordinates shown on screen
        page.remove_rotation()           # overlay.py does the same, so rects and bboxes match
    page_lines = []
    for b in page.get_text("dict")["blocks"]:
        for ln in b.get("lines", []):
            text = "".join(s["text"] for s in ln["spans"])
            if not text.strip():
                continue
            page_lines.append({"id": n, "page": pno, "bbox": [round(v, 1) for v in ln["bbox"]],
                               "dir": [round(v, 2) for v in ln["dir"]],
                               "size": round(ln["spans"][0]["size"], 1), "text": text.strip()})
            n += 1
    # no real text layer (scan, or CAD text drawn as strokes): OCR instead. A few stray words such as
    # a "Released for Manufacturing" stamp don't count as a text layer.
    if FORCE_OCR or len(page_lines) < 5:
        o = ocr_page(page, pno, n)
        if o:
            print(f"  page {pno + 1}: OCR found {len(o)} lines")
            page_lines += o
            n += len(o)
    lines += page_lines

    # gridded preview
    shape = page.new_shape()
    r = page.rect
    for x in range(0, int(r.width) + 1, 50):
        shape.draw_line((x, 0), (x, r.height))
    for y in range(0, int(r.height) + 1, 50):
        shape.draw_line((0, y), (r.width, y))
    shape.finish(color=(1, 0, 0), width=0.3, stroke_opacity=0.5)
    for ln in [l for l in lines if l["page"] == pno]:
        shape.draw_rect(pymupdf.Rect(ln["bbox"]))
    shape.finish(color=(0, 0, 1), width=0.4)
    shape.commit()
    for x in range(0, int(r.width) + 1, 50):
        page.insert_text((x + 1, 7), str(x), fontsize=5, color=(1, 0, 0))
    for y in range(50, int(r.height) + 1, 50):
        page.insert_text((1, y - 1), str(y), fontsize=5, color=(1, 0, 0))
    for ln in [l for l in lines if l["page"] == pno]:
        page.insert_text((ln["bbox"][0], ln["bbox"][1] + 0.5), str(ln["id"]), fontsize=4, color=(0, 0, 1))
    out = os.path.join(WORK, f"p{pno + 1}_grid.png")
    page.get_pixmap(dpi=150).save(out)
    print(f"page {pno + 1}: {r.width:.0f}x{r.height:.0f} pt, rotation {page.rotation}, "
          f"{sum(1 for l in lines if l['page'] == pno)} text lines -> {out}")

json.dump(lines, open(os.path.join(WORK, "lines.json"), "w", encoding="utf-8"), ensure_ascii=False, indent=0)
print("lines.json:", len(lines), "lines")
