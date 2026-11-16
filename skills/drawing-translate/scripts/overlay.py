"""Step 2: cover the chosen English text with white and write the Chinese in its place.

Usage: python overlay.py <source.pdf> <translations.json> <out.pdf> [preview_dir]

translations.json:
{
  "lines": "<work_dir>\\lines.json",          # optional, needed when items use "ids"
  "items": [
    {"page": 0, "ids": [101, 102], "zh": "T20 梅花槽"},                    # union of text-layer lines
    {"page": 0, "rect": [40, 337, 470, 470], "zh": "注：\\n1. ……", "size": 8},  # explicit rect, PDF points
    ...
  ]
}
Item options:
  size     max font size in pt (default: the original text size, else 9). It shrinks until the text fits.
  grow     [dx0, dy0, dx1, dy1] extra room added to the write box only (not the white-out), for long text
  align    "left" (default) | "center" | "right"
  rotate   0 / 90 / 270. Default: taken from the text-layer line direction.
  cover    true = also paint an "ids" item white (when stroked text sits under the text layer)
  review   "<short Chinese reason>": the translation is uncertain; listed in <out>_review.json
  en       the original English, for "rect" items that have a review (no text layer to quote from)
  note     true = review entry only, nothing drawn (e.g. English left on purpose that may matter):
           {"note": true, "page": 0, "en": "...", "zh": "", "review": "..."}

Only text is removed: line art and images stay. "ids" items delete the text-layer glyphs with no fill.
"rect" items are painted white (scanned or stroked text can only be covered), so keep those rects tight.
"""
import json
import os
import sys

import pymupdf

SRC, TR, OUT = sys.argv[1:4]
PREVIEW = sys.argv[4] if len(sys.argv) > 4 else None
FONTS = [r"C:\Windows\Fonts\Deng.ttf", r"C:\Windows\Fonts\simhei.ttf"]  # DengXian: proportional Latin/digits
FONTFILE = next((f for f in FONTS if os.path.exists(f)), None)
PAD = 0.8

spec = json.load(open(TR, encoding="utf-8-sig"))
lines = {}
if spec.get("lines"):
    lines = {l["id"]: l for l in json.load(open(spec["lines"], encoding="utf-8-sig"))}

doc = pymupdf.open(SRC)
items_by_page = {}
for it in spec["items"]:
    if not it.get("note"):  # note-only items are review entries, nothing to draw
        items_by_page.setdefault(it.get("page", 0), []).append(it)

ALIGN = {"left": 0, "center": 1, "right": 2}


def seg_hits(p, q, r):
    """Does segment p-q pass through rect r? (Liang-Barsky clip)"""
    t0, t1 = 0.0, 1.0
    dx, dy = q.x - p.x, q.y - p.y
    for d, e in ((-dx, p.x - r.x0), (dx, r.x1 - p.x), (-dy, p.y - r.y0), (dy, r.y1 - p.y)):
        if d == 0:
            if e < 0:
                return False
        else:
            t = e / d
            if d < 0:
                t0 = max(t0, t)
            else:
                t1 = min(t1, t)
            if t0 > t1:
                return False
    return True


def long_line_hits(d, r, min_len):
    """True if path d has a straight segment at least min_len long that crosses rect r."""
    if not d["rect"].intersects(r):
        return False
    for item in d["items"]:
        if item[0] == "l":
            segs = [(item[1], item[2])]
        elif item[0] == "re":
            q = item[1].quad
            segs = [(q.ul, q.ur), (q.ur, q.lr), (q.lr, q.ll), (q.ll, q.ul)]
        elif item[0] == "qu":
            q = item[1]
            segs = [(q.ul, q.ur), (q.ur, q.lr), (q.lr, q.ll), (q.ll, q.ul)]
        else:
            continue
        if any(abs(p - q) >= min_len and seg_hits(p, q, r) for p, q in segs):
            return True
    return False
for pno, items in items_by_page.items():
    page = doc[pno]
    if page.rotation:  # work in the coordinates shown on screen, same as extract.py
        page.remove_rotation()
    drawings = page.get_drawings()
    for d in drawings:  # a straight stroke (I, l, 1, -) has a zero-width rect, which counts as empty
        d["rect"] = d["rect"] + (-0.05, -0.05, 0.05, 0.05)
    small = [d for d in drawings if d["rect"].width < 20 and d["rect"].height < 20]
    longs = [d for d in drawings if max(d["rect"].width, d["rect"].height) >= 6]
    jobs, covers = [], []
    for it in items:
        if "ids" in it:
            ls = [lines[i] for i in it["ids"]]
            rect = pymupdf.Rect(ls[0]["bbox"])
            for l in ls[1:]:
                rect |= pymupdf.Rect(l["bbox"])
            size = it.get("size", max(l["size"] for l in ls))
            d = ls[0]["dir"]
            rot = it.get("rotate", 90 if d[1] < -0.5 else 270 if d[1] > 0.5 else 0)
        else:
            rect = pymupdf.Rect(it["rect"])
            size, rot = it.get("size", 9), it.get("rotate", 0)
        rect = rect + (-PAD, -PAD, PAD, PAD)
        if "ids" in it and any(l.get("ocr") for l in ls) and not it.get("cover"):
            # OCR lines: nothing to delete, paint each line box white (tight, so lines between stay)
            for l in ls:
                b0 = pymupdf.Rect(l["bbox"]) + (-0.5, -0.5, 0.5, 0.5)
                page.add_redact_annot(b0, fill=False)
                covers.append(b0)
        elif "ids" in it and not it.get("cover"):
            # text-layer lines are really deleted, one tight box per line, no fill: line art crossing
            # the area (boxes, leaders) stays; only paths lying wholly inside a line box (e.g. a ° drawn
            # as a little circle) go with the text
            boxes = [pymupdf.Rect(l["bbox"]) + (-PAD, -PAD, PAD, PAD) for l in ls]
            for l in ls:
                # delete through the middle of the line only, so a neighbouring line sitting
                # within a point above or below keeps its text
                b0 = pymupdf.Rect(l["bbox"])
                h = b0.height
                page.add_redact_annot(pymupdf.Rect(b0.x0 - PAD, b0.y0 + 0.3 * h, b0.x1 + PAD, b0.y1 - 0.3 * h),
                                      fill=False)
            # Stroked text: many CAD PDFs draw the visible words as line art and put an invisible text
            # layer on top. Cover every glyph-sized path of the line with a tiny white box: those wholly
            # inside the line box, plus those that stick out of it (the text-layer box often ends one
            # character short), as long as they sit in the same text band.
            for l, b in zip(ls, boxes):
                s = l["size"]
                reach = b + (-1.0 * s, 0, 1.0 * s, 0)  # text-layer boxes can end up to ~1 char short
                glyphs = [d["rect"] for d in small if d["rect"].width < s and d["rect"].height < s
                          and d["rect"].intersects(reach)]
                inside = [r for r in glyphs if b.contains(r)]
                if not inside:
                    continue
                y0, y1 = min(r.y0 for r in inside), max(r.y1 for r in inside)
                h = y1 - y0
                near = [d for d in longs if d["rect"].intersects(reach)]
                for r in glyphs:
                    if not b.contains(r):
                        if it.get("noauto") or not reach.contains(r):
                            continue
                        if r.y0 < y0 - 0.3 * h or r.y1 > y1 + 0.3 * h:
                            continue  # not in the text band
                        if r.height < 1.2 and r.width > 3.5:
                            continue  # horizontal leader stub, not a letter
                    c = r + (-0.4, -0.4, 0.4, 0.4)
                    # never paint across a long line (border, leader, box edge) that runs through the glyph
                    if any(long_line_hits(d, c, 3 * s) for d in near):
                        print(f"note page {pno + 1}: glyph at {[round(v) for v in r]} touches a long line, "
                              f"not covered ({l['text'][:20]!r})")
                        continue
                    covers.append(c)
        else:
            # explicit rects (scans, stroked text) can only be covered: remove any text inside, then
            # paint exactly the rect white (a redaction fill bleeds ~1.3 pt past the rect)
            exact = pymupdf.Rect(it["rect"]) if "rect" in it else rect
            page.add_redact_annot(exact, fill=False)
            covers.append(exact)
        jobs.append((rect, it, size, rot))
    page.apply_redactions(images=pymupdf.PDF_REDACT_IMAGE_NONE,
                          graphics=pymupdf.PDF_REDACT_LINE_ART_REMOVE_IF_COVERED)
    # glyph covers are painted directly in one go (one redaction per glyph is far too slow on
    # pages with tens of thousands of glyph paths)
    if covers:
        shape = page.new_shape()
        for c in covers:
            shape.draw_rect(c)
        shape.finish(color=None, fill=(1, 1, 1), width=0)
        shape.commit()
    if FONTFILE:
        page.insert_font(fontname="zh", fontfile=FONTFILE)
    for rect, it, size, rot in jobs:
        g = it.get("grow", [0, 0, 0, 0])
        box = rect + (-g[0], -g[1], g[2], g[3])
        fs = size
        while fs >= 3:
            # normal spacing first; tighter spacing before shrinking (single lines in short boxes)
            for lh in ((1.3,) if "\n" in it["zh"] else (1.3, 1.05)):
                rc = page.insert_textbox(box, it["zh"], fontname="zh" if FONTFILE else "china-s",
                                         fontsize=fs, rotate=rot, align=ALIGN[it.get("align", "left")],
                                         lineheight=lh)
                if rc >= 0:
                    break
            if rc >= 0:
                break
            fs -= 0.25
        else:
            print(f"WARNING page {pno + 1}: text does not fit even at 3 pt, use 'grow': {it['zh'][:30]!r}")
        if fs < size * 0.7:
            print(f"note page {pno + 1}: shrunk to {fs:.1f} pt (from {size}) : {it['zh'][:30]!r}")

doc.subset_fonts()
doc.save(OUT, garbage=3, deflate=True)
print("saved", OUT)

# translations that need a human check go to <out>_review.json, for the caller to collect
review = [{"page": it.get("page", 0) + 1,
           "en": it.get("en") or " / ".join(lines[i]["text"] for i in it.get("ids", [])) or "(see drawing)",
           "zh": it["zh"], "review": it["review"]}
          for it in spec["items"] if it.get("review")]
rpath = os.path.splitext(OUT)[0] + "_review.json"
if review:
    json.dump({"drawing": os.path.basename(SRC), "output": os.path.basename(OUT), "items": review},
              open(rpath, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
    print(f"{len(review)} item(s) to review -> {rpath}")
elif os.path.exists(rpath):
    os.remove(rpath)
if PREVIEW:
    os.makedirs(PREVIEW, exist_ok=True)
    for i, page in enumerate(pymupdf.open(OUT)):
        p = os.path.join(PREVIEW, f"{os.path.splitext(os.path.basename(OUT))[0]}_p{i + 1}.png")
        page.get_pixmap(dpi=200).save(p)
        print("preview", p)
