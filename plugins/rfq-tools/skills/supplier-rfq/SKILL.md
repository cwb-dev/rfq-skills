---
name: supplier-rfq
description: Split a collected client RFQ into Chinese 询价单 (supplier inquiry sheets) grouped by product category or factory, built on the 询价单-模板.xlsx template, with one drawing sheet per drawing number (the drawing with its callouts and notes translated into Chinese, plus plain-Chinese notes on what drives the price), the drawings also placed under the table and notes on the 询价单 sheet with a red label giving each drawing's line numbers, and a review list of uncertain translations, ready to send to factories. Use when the user says "make 询价单 for RFQ 1234", "split this RFQ for suppliers", or "prepare supplier inquiry".
---

# Supplier 询价单

Needs an inquiry folder with an item list, e.g. `items.md` (one line per item: part/drawing no., description, quantity, material, finish, drawing file), plus `drawings\` with the client drawings. If the project has a `CLAUDE.md` with folder or naming rules, read it first.

Scripts: `${CLAUDE_SKILL_DIR}/scripts/` (this skill) and `${CLAUDE_PLUGIN_ROOT}/skills/drawing-translate/scripts/` (drawing translation). Run python without `-I` (packages are in the user site).

**Setup (once per machine):** run `python ${CLAUDE_PLUGIN_ROOT}/skills/drawing-translate/scripts/check_deps.py`. It covers everything both skills need: `pymupdf`, `openpyxl`, and `rapidocr_onnxruntime` for scanned drawings. If anything is missing, ask the user, then run it with `--install` (add `--cn` for the Tsinghua mirror in China).

## Steps
1. **Group the lines** from `items.md` (follow how past RFQs were split, if there are any):
   - by material: 碳钢 (carbon steel), 不锈钢 (stainless), 黄铜 (brass);
   - and/or by product type: 螺栓/螺钉, 螺母, 垫圈, 拉铆螺母, 木螺钉, 异型及特殊件 (special/drawing parts), 车加工件 (machined).

   Show the proposed grouping and ask which factory each group goes to. Use the project's supplier list if it has one.
2. **Build one Excel file per group** from the template `${CLAUDE_SKILL_DIR}/templates/inquiry-template.xlsx`. Never edit the template itself. Name each file `supplier_rfq\<no>询价单-<group or factory>.xlsx`.
   - **The quick way:** write a groups JSON in the scratchpad (rows and notes, in the format given at the top of `build_rfq.py`) and run:
     ```
     python ${CLAUDE_SKILL_DIR}/scripts/build_rfq.py <groups.json> <inquiry>\supplier_rfq <no>
     ```
     It fills the template so its fonts, borders and yellow headers carry over, and applies the layout below.
   - **Row 1 (template header):** `A1:G1` merged `询价单` | `H1` 日期 | `I1` `=TODAY()` | `K1` price basis, e.g. `价格：人民币含税含运上海，元/千件`. Use `元/件` for large or machined parts priced per piece.
   - **Row 2 (columns):** `序号 | 图纸号 | 描述 | 数量 | 材质 | 表面处理 | 含税运价格 | 单重 | 交期 | 模具费 | 包装：外贸纸箱+托盘包装`. The factory fills the yellow columns 含税运价格, 单重 and 交期.
   - **Rows 3+ (one row per line item):**
     - `A` = `=ROW()-2`.
     - `B` = the drawing or part number only, in Arial 10.5.
     - `C` = the Chinese part name, plus the thread or length on a second line.
     - `D` = the quantity as a number, formatted `#,##0`.
     - `E` / `F` = material and finish.
     - `G`–`J` stay blank, except that `J` can hold a 模具费 note.
     - `K` = per-row remarks: tolerances, torque, "second quantity tier", and the drawing file name when it doesn't start with the part number.
     - Never write "同上" in `B`; repeat the part number. Other cells may say 同上.
   - **Under the table:** a bold `说明` row, then numbered notes, each merged across `A:K`. Include:
     - the reply-by date;
     - units (inches) and that material and finish must follow the drawing;
     - that alternatives go in the remarks column;
     - RoHS/REACH;
     - the price unit and the 单重 unit (kg/千件);
     - anything the template has no column for, such as sample fee, MOQ or monthly capacity ("请在备注栏注明…").
   - **Formatting:** wrap text, centre-align short columns and left-align C/E/F/K, size row heights to fit the text, freeze panes at `A3`, and print landscape at one page wide. Widen columns as needed (roughly B 18, C 30, E 26, F 36, K 42).

   Write material and finish in plain Chinese, the way a Chinese factory would say it (see **Factory language** below), not as a translation of the client's spec codes.
3. **Add one drawing sheet per drawing number.** Do this for every line item whose drawing is in `drawings\`.
   1. **Read the whole drawing.** Render each PDF page with PyMuPDF. Zoom into the notes block, the title block and the dimension views until every note, tolerance and special-characteristic mark (SC / CC / diamond / shield symbols) can be read.
   2. **Check the standards behind it.** For each material or plating standard on the drawing, look up what it actually requires (web search), so you can state real values instead of the code.
   3. **Write a notes JSON** in the scratchpad, following `${CLAUDE_SKILL_DIR}/example_notes.json`. Sections run in this order, quote factors first:
      1. 生产工艺: process route, and the dies or fixtures needed (模具费 quoted separately). **Only for parts that need combined machining (复合加工, e.g. 车加工件, extrusion + CNC).** For screws, SEMS, studs, weld studs, nuts and other cold-headed fasteners (brass ones too), leave this section out, write no tooling or process hints anywhere, and let the factory decide.
      2. 材料: grade and key properties.
      3. 尺寸及公差: do **not** restate dimensions that are already on the drawing, and do not choose "重点尺寸" yourself. List only: dimensions the drawing itself marks as critical (SC/CC/★/"critical", inspection balloons); things that are wrong, missing or contradictory on the drawing (typos, hidden tables, missing pitch or length); and the general tolerance (未注公差) as the last item. The 备注 column on the 询价单 follows the same rule: tests, torque and requests, not dimensions.
      4. 表面处理: method, process steps, thickness, colour, salt spray, coverage.
      5. 包装.
      6. 质量及检测: performance tests, sample documents, traceability, change control.
      7. 环保: RoHS/REACH.
      8. 其他: how to read the drawing, tooling marking, anything else.

      Mark an item `true` (red) when it affects price or production difficulty.
   4. **Translate the drawing into Chinese** with the `drawing-translate` skill: follow its SKILL.md for every drawing in `drawings\`.
      - Write the outputs to `<scratchpad>\translated\`, never into `drawings\`.
      - Mark uncertain translations with `review`.
      - With many drawings, split them over a few parallel agents (about 6 drawings each). Tell each agent to follow drawing-translate's SKILL.md, to not edit its scripts, and to report review items.
      - A drawing with no callouts or notes is simply copied unchanged.
   5. **Label each translated drawing** with its line numbers on this 询价单, using step 6 of drawing-translate:
      ```
      python ${CLAUDE_PLUGIN_ROOT}/skills/drawing-translate/scripts/stamp.py <scratchpad>\translated\<name>_中文.pdf <scratchpad>\stamped\<name>_中文.pdf --text "询价单序号 3｜图号 EX-2001-01"
      ```
      - The number is the 序号 (row − 2) of every row with that drawing number, e.g. `询价单序号 7、8` for two quantity tiers.
      - When one drawing serves several part numbers, list them all.
      - A drawing used in two group files gets a separate stamped copy for each file.
      - Specification-only documents, such as plating specs, get no label.
      - Build the sheet from the stamped copy.
   6. **Build the sheet** with the stamped PDF by running:
      ```
      python ${CLAUDE_SKILL_DIR}/scripts/add_drawing_sheet.py <询价单.xlsx> <translated _中文.pdf> <notes.json> <scratchpad img dir>
      ```
      - **Layout:** the script adds a sheet named by the drawing number, with the notes table on the left and every drawing page as a full image on the right. Set the notes `intro` to say the image is the Chinese-annotated version and dimensions and the title block stay in English.
      - **Link:** it links the drawing number on the 询价单 sheet to that sheet (an internal link, which works in Excel and WPS) and adds "图纸要求中文说明见 sheet「…」" to that row's 备注.
      - **Rewriting the row:** its `main` block can rewrite that row's 材质 / 表面处理 / 备注 to the same plain wording.
      - **Re-running** replaces the sheet.
   7. **Finish each file** after all drawing sheets are in:
      ```
      python ${CLAUDE_SKILL_DIR}/scripts/fit_rows.py <all 询价单 xlsx>
      python ${CLAUDE_SKILL_DIR}/scripts/place_drawings.py <all 询价单 xlsx>
      ```
      - `fit_rows.py` sizes the item rows.
      - `place_drawings.py` puts every drawing on the 询价单 sheet under the 说明 block, in table order, so a factory sees the table, the notes and the drawings on one sheet.
        - Each drawing gets a linked header row `图纸 n：<no>（询价单序号 …）` and starts a new printed page.
        - Each page is fitted into a box as wide as columns A–K and at most 850 px tall, which is about one screen at 100 % zoom.
        - The pictures are anchored "move and size with cells" over rows sized to fit. This is not Excel 365's "place in cell", which older WPS versions can't show.
        - Pictures are stored at twice the display size as palette PNGs, so zooming in stays sharp and the file stays small.
        - The script rebuilds the whole block on every run, so run it again after any change to a drawing sheet.
      - In the 说明 notes, say that the drawings are under the notes, and that the red box on each drawing gives its 询价单序号.
   8. **Close the file first.** If the xlsx is open in Excel or WPS, the save fails, so ask the user to close it. Hidden WPS `et.exe` processes can keep the lock after the window closes.
   9. **Review list for a colleague.** Run:
      ```
      python ${CLAUDE_PLUGIN_ROOT}/skills/drawing-translate/scripts/collect_review.py <inquiry>\supplier_rfq\<no>_图纸翻译待审核.xlsx <scratchpad>\translated --title "<no> 图纸翻译"
      ```
      Tell the user how many items need review and the most important ones (e.g. NORMALIZE, unreadable or cut-off text, English left in the title block). The 询价单 should not go to factories before a colleague has checked this list.
4. **Keep the client out of it.** Remove the client name, client PN prefixes if the user asks, and any target prices. Reference drawings by file name only. In the Chinese text, refer to end-customer specs only as 客户规范, never by company name. The drawing images themselves stay as received.
5. **List** the drawings each factory needs from `drawings\`. These are the untouched originals, for factories that want the PDF. The sheets carry the Chinese-annotated images.
6. **Write a short WeChat or email message** in Chinese for each factory, asking for 含税单价, 单重 and 交期, with a reply-by date. Use the same plain wording as the sheets, and point to the drawings (图纸已标注中文，在询价单表格下方，红框为询价单序号). In the message file, note that the translation review list must be checked before sending.
7. **If the project keeps an inquiry tracker** (e.g. `INQUIRY_LOG.md`), set the status to `sourcing` and note which factories received which groups.

## Factory language (applies to every Chinese text a factory sees)
- **Plain shop-floor Chinese.** Don't list international standards or end-customer spec numbers (ASTM, ASME, BMS-, TS-, PPAP levels, EU directive numbers). Heavy spec wording makes factories think the part is demanding and pad the price. Say what is actually needed, e.g. "样品需全尺寸报告、材质证明、镀层测厚报告" rather than "PPAP 3 级".
- **Material:** if the grade isn't common in Chinese supply chains, give the nearest common Chinese grade (e.g. SAE 1022 → 22#/SWRCH22A, 12L14 → Y15Pb, C1010 → 10#). Add the key properties only when they matter (e.g. 导电率 for busbars).
- **Plating:** never just translate the standard code. Always state:
  - the method (电镀 / 化学镀 / 达克罗 / 磷化 …);
  - each layer's thickness;
  - the colour (蓝白 / 彩色 / 黑色 / 本色, 哑光 or 光亮);
  - the salt-spray hours, or "图纸无盐雾要求" when there are none;
  - other limits such as 除氢, 不加光亮剂, or 三价铬/无六价铬;
  - areas that must or must not be plated.

  Look up the standard to get these values. Don't write 外协 or ask whether plating is in-house; the factory decides that.
- **Drawing revision:** ignore revision status, "UNDER REVIEW" watermarks and the like. Quote the drawing as received.

Do not send anything. The user sends to factories.
