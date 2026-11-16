"""Write a red label (e.g. the inquiry line number and drawing number) into a blank area of every page.

Usage: python stamp.py <in.pdf> <out.pdf> --text "询价单序号 3｜图号 EX-2001-01" [--pages 0,1]

The blank area is found on a render of the page: the label goes where a box with a margin around it
touches no ink, using the largest font from 18 pt down to 9 pt that fits, preferring the top right,
then the top left, then anywhere else. If no blank box is found, the label goes in the top-right corner
on a white background and a warning is printed.
<out.pdf> must differ from <in.pdf>, so a re-run always starts from the unstamped file.
"""
import argparse
import os
import sys

import numpy as np
import pymupdf

FONTS = [r"C:\Windows\Fonts\Dengb.ttf", r"C:\Windows\Fonts\Deng.ttf", r"C:\Windows\Fonts\simhei.ttf"]
FONTFILE = next((f for f in FONTS if os.path.exists(f)), None)
RED = (0.85, 0, 0)
SIZES = [18, 16, 14, 12, 11, 10, 9]
MARGIN = 8      # pt of clear space kept around the label box
EDGE = 18       # pt kept from the sheet edge (printers cut the outer strip)
PAD = 4         # pt between the red frame and the text

ap = argparse.ArgumentParser()
ap.add_argument("src")
ap.add_argument("out")
ap.add_argument("--text", required=True)
ap.add_argument("--pages", help="0-based page numbers, comma separated (default: all)")
a = ap.parse_args()
if os.path.abspath(a.src) == os.path.abspath(a.out):
    sys.exit("out.pdf must differ from in.pdf (stamp the clean translation each time)")

font = pymupdf.Font(fontfile=FONTFILE) if FONTFILE else pymupdf.Font("china-s")
doc = pymupdf.open(a.src)
pages = [int(p) for p in a.pages.split(",")] if a.pages else range(len(doc))


def find_spot(ink, w, h):
    """Top-left (x, y) in pt of a w x h box (plus MARGIN) with no ink, or None. ink: bool array, 1 px = 1 pt."""
    H, W = ink.shape
    bw, bh = int(w + 2 * MARGIN) + 1, int(h + 2 * MARGIN) + 1
    if bw > W - 2 * EDGE or bh > H - 2 * EDGE:
        return None
    S = np.zeros((H + 1, W + 1), dtype=np.int32)
    S[1:, 1:] = ink.cumsum(0).cumsum(1)
    sums = S[bh:, bw:] - S[:-bh, bw:] - S[bh:, :-bw] + S[:-bh, :-bw]   # box sum for every top-left
    ys, xs = np.nonzero(sums == 0)
    ok = (xs >= EDGE) & (ys >= EDGE) & (xs + bw <= W - EDGE) & (ys + bh <= H - EDGE)
    xs, ys = xs[ok], ys[ok]
    if not len(xs):
        return None
    # prefer top right, then top left: cost = distance from the top, plus distance from the nearer side
    right = W - EDGE - (xs + bw)
    left = xs - EDGE
    cost = ys * 1.0 + np.minimum(right, left + 0.25 * W)
    i = int(np.argmin(cost))
    return xs[i] + MARGIN, ys[i] + MARGIN


for pno in pages:
    page = doc[pno]
    pix = page.get_pixmap(dpi=72, colorspace=pymupdf.csGRAY, annots=True)
    ink = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.height, pix.stride)[:, :pix.width] < 235
    ink = ink.copy()
    # the inside of a stroked frame (table cell, highlight box, notes box) counts as taken, except the sheet border
    area = page.rect.width * page.rect.height
    for d in page.get_drawings():
        if d.get("color") is None:
            continue
        for it in d["items"]:
            if it[0] not in ("re", "qu"):
                continue
            b = (it[1] if it[0] == "re" else it[1].rect) * page.rotation_matrix
            if b.width > 15 and b.height > 8 and b.width * b.height < 0.25 * area:
                ink[max(0, int(b.y0)):int(b.y1) + 1, max(0, int(b.x0)):int(b.x1) + 1] = True
    spot = None
    for fs in SIZES:
        tw = font.text_length(a.text, fontsize=fs)
        w, h = tw + 2 * PAD, fs * 1.25 + 2 * PAD
        spot = find_spot(ink, w, h)
        if spot:
            break
    fallback = spot is None
    if fallback:
        fs = 12
        w, h = font.text_length(a.text, fontsize=fs) + 2 * PAD, fs * 1.25 + 2 * PAD
        spot = (pix.width - EDGE - w, EDGE)
        print(f"WARN page {pno}: no blank area big enough, label put top right on a white background")
    x, y = spot
    vis = pymupdf.Rect(x, y, x + w, y + h)                 # in the rotated (on-screen) frame
    r = vis * page.derotation_matrix                        # back to unrotated page space
    shape = page.new_shape()
    shape.draw_rect(r)
    shape.finish(color=RED, fill=(1, 1, 1) if fallback else None, width=1.2)
    shape.commit()
    if FONTFILE:
        page.insert_font(fontname="stamp", fontfile=FONTFILE)
    # baseline point of the text in the on-screen frame, mapped to page space
    org = pymupdf.Point(x + PAD, y + PAD + fs * 0.98) * page.derotation_matrix
    page.insert_text(org, a.text, fontname="stamp" if FONTFILE else "china-s",
                     fontsize=fs, color=RED, rotate=page.rotation)
    print(f"page {pno}: label at ({x:.0f},{y:.0f}) {fs}pt{' FALLBACK' if fallback else ''}")

doc.subset_fonts()
doc.save(a.out, garbage=3, deflate=True)
print("saved", a.out)
