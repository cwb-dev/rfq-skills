---
name: drawing-translate
description: Translate a non-Chinese engineering drawing (PDF or image, including scans and faxes via OCR) into Chinese by writing the Chinese directly over the original words, on the drawing itself. Translate every requirement that can affect manufacturing, inspection, process, cost or quotation, including callouts, notes, dimension qualifiers, tolerances, material, finish, heat treatment and production-relevant title-block fields; numeric values, standard numbers and graphics stay unchanged. Produces a new PDF, PNG previews and a review list of uncertain translations for a colleague to check, and can add a red label (e.g. the inquiry line number) in a blank area. Use when the user says "translate this drawing", "把图纸翻译成中文", "图纸中文化", or when another skill needs a Chinese version of a drawing.
---

# Drawing translation (图纸翻译成中文)

Works on the drawing file only. It does not know about RFQs, 询价单 or clients, so any skill can call it with a file and get a Chinese file back.

Scripts are in `${CLAUDE_SKILL_DIR}/scripts/`. Run python without `-I`, because the packages live in the user site.

## Setup (once per machine)
Run:
```
python ${CLAUDE_SKILL_DIR}/scripts/check_deps.py
```
- If something is missing, run `python ${CLAUDE_SKILL_DIR}/scripts/check_deps.py --install`, or add `--cn` to install through the Tsinghua PyPI mirror (much faster in mainland China).
- Packages:
  - `pymupdf`: required.
  - `rapidocr_onnxruntime`: OCR for scans, faxes and CAD text drawn as lines. About 100 MB with onnxruntime and opencv. Runs offline after install.
  - `openpyxl`: the review spreadsheet.
  - `numpy`, `pillow`: the blank-area search for labels, and the pictures in spreadsheets.
- Font: DengXian or SimHei from `C:\Windows\Fonts`. Without them, the PDF built-in CJK font is used.

If `check_deps.py` reports something missing, tell the user and ask before installing: it downloads from PyPI.

## Inputs and outputs
- **Input:** a drawing PDF (any number of pages, rotated pages are fine), or a PNG/JPG.
- **Output,** saved next to the source unless the caller gives another path:
  - `<source name>_中文.pdf`;
  - `<name>_中文_p<N>.png` previews at 200 dpi;
  - `<name>_中文_review.json` when some translations need a human check (see **Review**).
  - optionally, a copy with a red label in a blank area (step 6), e.g. the inquiry line number for a 询价单.

  Never overwrite or change the source file.
- **Work files** (lines.json, grid renders, translations.json) go in the scratchpad: `<scratchpad>\drawing-translate\<source name>\`.

## What to translate
**Core principle: translate everything that can affect manufacturing, inspection, process, production cost or quotation.** Do not leave such information in English merely because it sits in a dimension, tolerance table, notes block or title block.

1. **Part-view and process callouts:** arrow/leader labels, head/drive/thread/point requirements, machining and forming instructions, `T20 TORX RECESS`, `10-32 TRI-LOBE`, `TYPE "CA" POINT`, `KNURL`, `FULL THREAD`, `SECTION A-A`, `DETAIL C`.
2. **All production notes and specifications:** material, finish, plating/coating/painting/adhesive/wax/passivation, heat treatment, hardness, salt spray, performance, welding, inspection, packaging, sampling, quality and change-control requirements.
3. **Dimension and tolerance language:** translate the words without changing numbers or geometry: `MIN` → `最小`, `MAX` → `最大`, `REF` → `参考`, `TYP` → `典型`, `FULL THREAD` → `全螺纹`, `THK` → `厚度`, general-tolerance tables, units, `DO NOT SCALE`, critical-characteristic instructions and inspection notes.
4. **Production-relevant title-block fields:** part description, MATERIAL, FINISH, HEAT TREATMENT, HARDNESS, units and similar fields. The factory must not need to interpret these English fields itself.

**Leave as is:**
- Numeric dimensions, tolerances, thread designations, standard numbers, part numbers, revision letters, drawing numbers and all graphics. Translate their descriptive words only.
- Purely administrative or non-production content such as drawn/checked by, dates, approval signatures, zone letters/numbers, confidentiality text, stamps, watermarks and fax headers. Revision history stays in English unless it changes the current production requirement; deleted/superseded instructions must not be carried into the current requirements.

When the English or its technical meaning is uncertain, still provide the best responsible Chinese translation on the drawing and mark it `review`; do not simply leave production/quotation information untranslated. Preserve the original source and make the uncertainty visible in the review list.

## Translation style
- **Accurate and complete.** This is a translation, not a summary: keep every requirement, number, unit, standard name and code exactly as written (`IFI`, `ASTM B633`, `400°F`, `HRC 40-44`, `.007 MIN.`, `60 in.lbs`).
- **Common Chinese fastener and machining terms:**

  | English | Chinese |
  |---|---|
  | TRI-LOBE | 三角牙 |
  | TORX / 6-LOBE RECESS | 梅花槽 |
  | PHILLIPS | 十字槽 |
  | SLOT | 一字槽 |
  | SQUARE DRIVE | 方槽 |
  | BELLEVILLE / CONICAL WASHER | 碟形垫圈 |
  | CROWN | 碟高 |
  | POINT | 尾 |
  | PAN HEAD | 盘头 |
  | TRUSS HEAD | 大扁头 |
  | HEX WASHER HEAD | 六角法兰头 |
  | KNURL | 滚花 |
  | CHAMFER | 倒角 |
  | UNDERCUT | 退刀槽 |
  | CASE HARDEN | 渗碳淬火 |
  | BAKE | 镀后/除氢语境：除氢烘烤；涂装/胶粘语境：固化或烘烤 |
  | PLATE | 电镀 |
  | THREAD-FORMING | 自攻成型 |
  | O.D. | 外径 |
  | THK | 厚度 |
  | GO / NO-GO | 通规 / 止规 |
  | TRI-MIC / REG-MIC | 三点千分尺 / 普通千分尺 |

- **NORMALIZE after plating:** when it refers to a plating or baking spec, it means the post-plating hydrogen-relief bake. Translate it as **除氢烘烤**, never 正火, and add a review note.
- **Note numbering:** keep it as is (1. 2. 3.). Start the notes block with `注：`.
- **Line breaks:** put a `\n` wherever the original starts a new item. The script fits the text into the original area and shrinks the font only if needed.

## Steps
1. **Extract.** Run:
   ```
   python ${CLAUDE_SKILL_DIR}/scripts/extract.py <drawing> <work_dir> [--ocr]
   ```
   - **What it writes:**
     - `lines.json`: every text line with id, page, bbox in PDF points, direction, size and text;
     - `p<N>_grid.png`: the page with a 50-pt red coordinate grid and each line boxed in blue with its id.
   - **Text layer present:** lines come from the PDF text layer.
   - **No text layer** (fewer than 5 lines: scans, faxes, CAD text drawn as strokes): the page is OCR'd automatically. It takes two passes, a full page at 150 dpi plus 300 dpi tiles, and the lines carry `"ocr": true` and a confidence `score`.
   - **`--ocr`:** forces OCR on every page. Use it when the text layer exists but is offset from the drawn letters.
   - **Images and rotation:** an image input is wrapped into `<work_dir>\source.pdf`; use that file as the overlay source. Rotated pages are handled in the coordinates you see on screen.
2. **Look at the whole drawing.** Read every `p<N>_grid.png` and zoom into the views and the notes block (render a clip with PyMuPDF at 300+ dpi) until each callout and note is legible.
   - Identify every production-, inspection-, cost- and quotation-relevant line using the rules above, including dimension language, tolerance tables and title-block material/finish fields.
   - **OCR text can be wrong.** Always translate from what you read on the drawing, not from the OCR string. Use OCR only for the position.
   - **Check that each box covers the whole phrase.** OCR sometimes drops the first or last letters. If a box is short, use an explicit `rect` for that item.
   - Text that neither the text layer nor OCR found also gets a `rect` `[x0, y0, x1, y1]` in PDF points, read off the grid and checked zoomed in.
3. **Write `translations.json`** in the work dir:
   ```json
   {"lines": "<work_dir>\\lines.json",
    "items": [
      {"page": 0, "ids": [74, 75], "zh": "T20 梅花槽"},
      {"page": 0, "ids": [117, 118, 119], "zh": "注：\n1. ……\n2. ……"},
      {"page": 0, "rect": [330, 282, 395, 304], "en": "RECESS PENETRATION GAGING DEPTH", "zh": "槽通规\n插入深度", "size": 8},
      {"page": 0, "ids": [86], "zh": "除氢烘烤：按 …", "review": "原文 NORMALIZE，按镀后除氢理解"},
      {"page": 0, "ids": [87], "en": "MATERIAL: STEEL", "zh": "材料：钢"}
    ]}
   ```
   - Use one item per callout and one item for the whole notes block (all its line ids together, so the Chinese reflows inside the original area).
   - Options:
     - `size`: max pt; default is the original size.
     - `grow` `[dx0, dy0, dx1, dy1]`: extra writing room into blank space.
     - `align`: left / center / right.
     - `rotate`: vertical text is detected automatically.
     - `cover: true`: paint the whole item area white.
     - `noauto: true`: don't auto-cover glyphs just outside the line box.
     - `review`, `en`, `note`: see **Review**.
4. **Overlay.** Run:
   ```
   python ${CLAUDE_SKILL_DIR}/scripts/overlay.py <source.pdf> <translations.json> <out.pdf> <preview_dir>
   ```
   - **Text-layer `ids` items:**
     - The English glyphs are deleted (only through the middle of each line, so neighbouring lines stay). No fill is used, so boxes and leader lines crossing the area stay intact.
     - Many CAD PDFs draw the visible letters as line art under an invisible text layer. Every glyph-sized path of the line, plus up to one character beyond the box in the same text band, gets a tiny white cover.
     - Leader stubs are skipped. So is any glyph crossed by a long straight line, which gets a "not covered" note instead.
     - If "ghost" letters are still visible (drawn letters offset from the text layer), set `cover: true` on that item.
   - **OCR `ids` items:** each line box is painted white exactly, then the Chinese is written.
   - **`rect` items:** exactly the rect is painted white; the rect is not enlarged.
   - **Text fitting:** text is written at the original size and shrinks until it fits. Line spacing is 1.3 for multi-line text; single lines may tighten to 1.05 first. The script warns when the text needed a lot of shrinking or doesn't fit.
5. **Check the preview.** Read every `_p<N>.png`, and zoom into each translated area. Look for:
   - leftover English;
   - leftover symbols (°, ±);
   - a leader, box, border or dimension line that disappeared;
   - Chinese overlapping other content;
   - text too small to read (under about 6 pt).

   Fix the item and run step 4 again; it always starts from the untouched source.
6. **Label (optional, when the caller gives a label text).** Write a red framed label, e.g. `询价单序号 3｜图号 EX-2001-01`, into a blank area of every page:
   ```
   python ${CLAUDE_SKILL_DIR}/scripts/stamp.py <name>_中文.pdf <out dir>\<name>_中文.pdf --text "<label>"
   ```
   - It finds a box with clear space all round that touches no lines, text, table cell or highlight box. It uses the largest font from 18 pt down to 9 pt that fits, preferring the top right, then the top left.
   - If no such box exists, the label goes in the top-right corner on a white background with a `WARN`. Check that page's preview.
   - The output must be a different file from the input, so a re-run always stamps the clean translation. Write it to another folder under the same name.
   - Re-render the preview of the stamped file and check that the label covers nothing.
7. **Report** to the caller: the output path, what was translated, what was left in English and why, and the number of review items.

## Review (复查机制)
Every translation that isn't certain goes on a review list for a colleague. When building translations.json, add `"review": "<short Chinese reason>"` to the item. For `rect` items, also add `"en": "<original text>"`. Mark an item for review when:
- the source text is unclear, partly unreadable, cut off at the sheet edge, or came from OCR with `score` < 0.8;
- a term has more than one possible meaning (NORMALIZE, DRIVER SIZE, LAST SCRATCH, FIRST SCRATCH);
- you expanded an abbreviation (CRS, THK, CRN);
- you copied an error in the original as is (wrong numbering, UNC on a fine thread, a typo);
- a production-relevant translation is uncertain. Put the best responsible Chinese translation on the drawing and list it for review; do not leave material, finish, dimensions, tolerances or process requirements in English just because they need confirmation. `note: true` is only for non-production administrative text or an illegible source that cannot be translated responsibly.

Before delivery, every production-relevant review item must have a non-empty Chinese `zh` translation on the drawing. Pure numeric values, list markers and administrative fields are not untranslated review items. An empty `zh` is allowed only for an illegible source that needs manual reading.

`overlay.py` writes them to `<out>_review.json`. To make the review spreadsheet for one or many drawings, run:
```
python ${CLAUDE_SKILL_DIR}/scripts/collect_review.py <out.xlsx> <folder with *_review.json | review files…> --title "<job name>"
```
It produces one row per item: drawing, page, original, translation, reason, plus yellow blank columns 审核意见 / 审核人. The calling skill decides where the xlsx goes (supplier-rfq puts it in the inquiry folder). When called directly, put it next to the outputs and mention it in the report.

## Notes
- The output is a normal PDF: callers can render it to PNG, embed it, or send it as is.
- Don't translate or change part numbers, revision letters, dimension values or customer names.
- Multi-page drawings: one `items` entry per area, with its `page` (0-based).
