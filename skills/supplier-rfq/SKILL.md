---
name: supplier-rfq
description: Split a collected client RFQ into Chinese 询价单 (supplier inquiry sheets) grouped by product category or factory, built on the 询价单-模板.xlsx template, with one drawing sheet per drawing number (the drawing with its callouts and notes translated into Chinese, plus plain-Chinese notes on what drives the price), the drawings also placed under the table and notes on the 询价单 sheet with a red label giving each drawing's line numbers, and a review list of uncertain translations, ready to send to factories. Use when the user says "make 询价单 for RFQ 1234", "split this RFQ for suppliers", or "prepare supplier inquiry".
---

# Supplier 询价单

Needs an inquiry folder with an item list, e.g. `items.md` (one line per item: part/drawing no., description, quantity, material, finish, drawing file), plus the client drawings. If the project has a `CLAUDE.md` with folder or naming rules, read it first.

`<inquiry>` means the inquiry's folder. The client's drawings are in it, and the 询价单, the review list and the factory messages are written to it.

Scripts: `${CLAUDE_SKILL_DIR}/scripts/` (this skill) and `${CLAUDE_PLUGIN_ROOT}/skills/drawing-translate/scripts/` (drawing translation). Run python without `-I` (packages are in the user site).

**Setup (once per machine):** run `python ${CLAUDE_PLUGIN_ROOT}/skills/drawing-translate/scripts/check_deps.py`. It covers everything both skills need: `pymupdf`, `openpyxl`, and `rapidocr_onnxruntime` for scanned drawings. If anything is missing, ask the user, then run it with `--install` (add `--cn` for the Tsinghua mirror in China).

## Steps
1. **Group the lines** from `items.md` (follow how past RFQs were split, if there are any):
   - by material: 碳钢 (carbon steel), 不锈钢 (stainless), 黄铜 (brass);
   - and/or by product type: 螺栓/螺钉, 螺母, 垫圈, 拉铆螺母, 木螺钉, 异型及特殊件 (special/drawing parts), 车加工件 (machined).
   - Do not create a group solely for a finish or coating (for example cadmium plating). Keep parts with their actual product type/material factory, and put coating-specific requirements and environmental conflicts in the affected rows and notes.
   - For drawings belonging to the same product family, use consistent Chinese part names, process wording and translation terminology across groups. Do not merge rows or quantities: each drawing's thread, dimensions, material, finish and other drawing-specific requirements remain separate and take precedence.

   Show the proposed grouping and ask which factory each group goes to. Use the project's supplier list if it has one.
2. **Build one Excel file per group** from the template `${CLAUDE_SKILL_DIR}/templates/inquiry-template.xlsx`. Never edit the template itself. Name each file `<inquiry>\<no>询价单-<group or factory>.xlsx`.
   - **The quick way:** write a groups JSON in the scratchpad (rows and notes, in the format given at the top of `build_rfq.py`) and run:
     ```
     python ${CLAUDE_SKILL_DIR}/scripts/build_rfq.py <groups.json> <inquiry> <no>
     ```
     It fills the template so its fonts, borders and yellow headers carry over, and applies the layout below.
   - **Row 1 (template header):** `A1:G1` merged `询价单` | `H1` 日期 | `I1` `=TODAY()` | `K1` price basis, e.g. `价格：人民币含税含运上海，元/千件`. Use `元/件` for large or machined parts priced per piece.
   - **Row 2 (columns):** `序号 | 图纸号 | 描述 | 数量 | 材质 | 表面处理 | 含税运价格 | 单重 | 交期 | 模具费 | 包装：外贸纸箱+托盘包装`. The factory fills the yellow columns 含税运价格, 单重 and 交期.
   - **Rows 3+ (one row per line item):**
     - `A` = `=ROW()-2`.
     - `B` = the drawing or part number only, in Arial 10.5.
     - `C` = the Chinese part name, plus the thread or length on a second line.
     - `D` = the quantity as a number, formatted `#,##0`.
     - `E` / `F` = the **complete, factory-ready material and finish requirements**, resolved from the drawing and specification research before the workbook is built. Never write only `钢`, `不锈钢`, `镀锌`, `按图纸`, `详见图纸` or `同上` when the drawing contains more detail; the factory must not have to open each drawing to discover the grade or finish.
       - **材质 — keep familiar grades; align unfamiliar grades by chemistry.** If the original grade is already directly understood by Chinese factories, keep it as-is without adding a GB designation or a `对应国内/近似国内` phrase. This includes common stainless grades such as `304`, `304L`, `316`, `302`, `303` and `410`; write simply `304/304L不锈钢`, `316不锈钢`, `410不锈钢`, etc., plus any required condition or hardness. For a grade unfamiliar in the Chinese supply chain, compare its actual chemical-composition ranges with a familiar domestic grade, then write `<原牌号>对应国内<国内牌号>` when the chemistry corresponds, or `<原牌号>近似国内<国内牌号>` when it is only close. Never use `对应` merely because two grades are common industry substitutes: if the composition is not fully equivalent, use `近似`. Examples: `SAE 1010近似国内10#钢`, `C26000对应国内H70黄铜`, `ASTM A193 B7近似国内42CrMo`. Record the composition source used for the comparison in the specification research summary. If no suitable domestic grade exists, write `国内无成分一致的常用牌号` rather than forcing a match. Do not silently substitute a near grade; where the original grade controls, add `替换需确认`. Include required condition, hardness or standard grade (for example 12.9级, cold-headed) when specified. If no grade is given, write `牌号未注明` explicitly rather than sending the factory to find it.
       - **表面处理 — state all four parts in every row:** ① the process/nature (电镀、化学镀、热浸镀、磷化、发黑、钝化、喷漆、涂蜡等, plus colour/layers); ② whether trivalent chromium / hex-chrome-free / RoHS-REACH environmental compliance is required, or whether this is unstated or conflicts with the requested finish; ③ coating thickness, using µm (include the original inch value when useful); ④ salt-spray requirement. Write `厚度：图纸/表格未注明` and `盐雾：图纸/表格未注明` when absent — never invent values and never omit these fields. For `dichromate` without a chromium valence, say `三价/六价未注明，需确认环保`; do not assume it is trivalent. For a missing customer specification, keep the `missing specification` marker in the row.
       - **Do not omit secondary surface operations.** Scan the whole drawing and specifications for 喷漆/粉末喷涂、点胶/螺纹锁固胶/密封胶、钝化、涂蜡/润滑、封闭/面涂、除氢、染色、固化、遮蔽、阳极氧化、化学镀等, in addition to the main finish. Put their actual materials, colours, locations, thicknesses, temperatures, times and inspection requirements directly in the same `F` cell. Do not reduce these to a generic `等` or leave them only in the drawing sheet. Distinguish coating thickness from sheet/part thickness and hydrogen-relief bake from paint/adhesive curing.
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
3. **Add one drawing sheet per drawing number.** Do this for every line item whose drawing (PDF or image from the client) is in `<inquiry>`.
   1. **Read the whole drawing.** Render each PDF page with PyMuPDF. Zoom into the notes block, the title block and the dimension views until every note, tolerance and special-characteristic mark (SC / CC / diamond / shield symbols) can be read.
   2. **Resolve every customer specification; never quote a bare code.** This applies to material, plating, paint, wax, heat-treatment, inspection and other customer-specific requirements (for example `DS2193`, `DS2212` or a customer paint specification), not only public ASTM/SAE standards.
      - **Search the client's documents first.** Check the current inquiry's email, attachments, drawings and specification files. Also check the client's existing specification library or earlier supplied documents when available. Reuse a previous document only when the specification number, issuer and applicable revision match; record the actual file and page used.
      - **If the exact text is not supplied, research it on the web.** Search the exact code, spaced/hyphenated variants, issuer/customer name and process keywords. Prefer the customer's official document or supplier portal, the issuing standards body, then a credible public copy. Require an exact code and issuer match; a similarly numbered document, a search snippet or a similar public standard is not the specification. Do not upload client-confidential files to a search service.
      - **Save and review found material before using it.** Save the original public source flat in `<inquiry>` under a clear name such as `<spec>_<title>.pdf` (or `.html`), without overwriting client files. Record its source URL, document title, issuer, revision/date and access date in `<inquiry>\<no>_客户规范查找及翻译.md`. Show the user the saved source and a Chinese summary of the actual requirements for review **before** incorporating web-researched requirements into supplier sheets. Do not present unreviewed research as client-approved.
      - **If the exact specification cannot be found, write `missing specification`** in `items.md`, the affected 询价单 row's 备注, the drawing notes sheet and `<no>_发厂家消息及图纸清单.md`. State which part numbers and process it affects, and request the document and revision from the client. Never fill the gap with a nearby ASTM/SAE standard, factory habit or guessed thickness/colour/salt-spray value.
      - **Translate the resolved requirement into executable Chinese process language.** Give the actual process route, layer/coating thickness, colour, salt spray, hydrogen relief, masking/coverage, curing, wax/lubricant and inspection requirements as applicable. A code alone is not a translation. Keep the code for internal traceability, but the factory-facing notes must say what to do. Record unresolved fields explicitly instead of guessing.
      - The research summary must list every specification checked, client files/pages searched, web queries used, found source paths/URLs, the Chinese requirement summary, and each `missing specification`. Keep the summary in the inquiry folder even when nothing was found.
      - **Maintain an OEM/customer-spec missing register for client confirmation.** Scan every drawing and the client email for customer/internal codes (for example ASA 840, DS2193, EI-194, OEM color/coating codes). Separate these from public ASME/ASTM/SAE/ANSI/UL/MIL standards, which must not be mislabeled as missing customer spec sheets. Write `<inquiry>\<no>_缺失规范及客户确认清单.md` listing each missing code, every affected part number, the drawing evidence, missing process/inspection fields and research outcome. Include an English customer-confirmation email draft, but never send it unless the user asks. Also register missing detailed requirements when no specification number is given (for example cadmium thickness, salt spray, hydrogen relief, adhesion and inspection requirements).
      - **Include missing source drawings in the same register and customer-confirmation email.** For every referenced but absent drawing, list the part number, expected drawing file name/revision and the requirement it affects, and ask the client for the original drawing. Distinguish `missing drawing` from `missing specification`; do not silently quote a part without the drawing merely because its material and finish are described in the email table.
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
   4. **Run every drawing through `drawing-translate` automatically** by following its SKILL.md for every client drawing in `<inquiry>`. This is mandatory, including scans, faxes and CAD drawings with no usable text layer: run extraction and OCR rather than skipping the drawing or merely embedding the English original. Do not claim translation was performed when only a copy was made.
      - Write the outputs to `<scratchpad>\translated\`, never into `<inquiry>`, so the client's originals stay untouched and the inquiry folder doesn't fill with intermediate files.
      - Mark uncertain translations and low-confidence OCR with `review`; preserve the original source for comparison.
      - With many drawings, split them over a few parallel agents (about 6 drawings each). Tell each agent to follow drawing-translate's SKILL.md, to not edit its scripts, and to report review items.
      - Only after extraction/OCR confirms there is no production-, inspection-, cost- or quotation-relevant text to translate may a drawing be copied unchanged. Every resulting drawing, translated or confirmed unchanged, must still be stamped and added to its drawing sheet, then placed under the notes on the first 询价单 sheet by `place_drawings.py`.
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
      - **Layout:** the script adds a sheet named by the drawing number, with the notes table on the left and every drawing page as a full image on the right. Set the notes `intro` to say the image is the Chinese-annotated version: all production/quotation-relevant text, including dimension language, tolerances and title-block material/finish, is translated; numeric values, standard numbers and graphics stay unchanged.
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
   8. **Validate the first-sheet summary, OEM markers and drawing coverage.** Reopen every workbook and verify that every item row has a non-empty Chinese description, aligned material grade, and all four finish fields (process, environmental/chromium requirement, thickness, salt spray). Reject rows that merely say `按图纸` or force the factory to look up requirements. For every OEM/customer-specific code referenced on a drawing or in the client email, verify the affected first-sheet row carries either the resolved requirement or an explicit `missing specification` marker—never leave a bare code with no explanation. Verify missing source drawings are explicitly marked as missing and are included in the customer-confirmation email register. Confirm every available drawing has an OCR/extraction record, a linked drawing sheet and an image under the first-sheet notes.
   9. **Close the file first.** If the xlsx is open in Excel or WPS, the save fails, so ask the user to close it. Hidden WPS `et.exe` processes can keep the lock after the window closes.
   10. **Review list for a colleague.** Run:
      ```
      python ${CLAUDE_PLUGIN_ROOT}/skills/drawing-translate/scripts/collect_review.py <inquiry>\<no>_图纸翻译待审核.xlsx <scratchpad>\translated --title "<no> 图纸翻译"
      ```
      Tell the user how many items need review and the most important ones (e.g. NORMALIZE, unreadable or cut-off text, English left in the title block). The 询价单 should not go to factories before a colleague has checked this list.
4. **Keep the client out of it.** Remove the client name, client PN prefixes if the user asks, and any target prices. Reference drawings by file name only. In the Chinese text, refer to end-customer specs only as 客户规范, never by company name. The drawing images themselves stay as received.
5. **List** the client drawings in `<inquiry>` that each factory needs. These are the untouched originals, for factories that want the PDF. The sheets carry the Chinese-annotated images.
6. **Write a short WeChat or email message** in Chinese for each factory, asking for 含税单价, 单重 and 交期, with a reply-by date. Use the same plain wording as the sheets, and point to the drawings (图纸已标注中文，在询价单表格下方，红框为询价单序号). Save the messages as `<inquiry>\<no>_发厂家消息及图纸清单.md`, and note in that file that the translation review list must be checked before sending.
7. **If the project keeps an inquiry tracker** (e.g. `INQUIRY_LOG.md`), set the status to `sourcing` and note which factories received which groups.

## Factory language (applies to every Chinese text a factory sees)
- **Plain shop-floor Chinese.** Don't merely copy international standards or end-customer spec numbers (ASTM, ASME, BMS-, TS-, PPAP levels, EU directive numbers). Heavy spec wording makes factories think the part is demanding and pad the price. Resolve the specification as required above and state the actual requirements, e.g. "样品需全尺寸报告、材质证明、镀层测厚报告" rather than "PPAP 3 级". A specification code may remain in internal research notes or beside the translated requirement for traceability, but never use the code as a substitute for the process requirements.
- **Material:** compare chemistry before naming a domestic grade. Use `<原牌号>对应国内<国内牌号>` only for equivalent chemistry and `<原牌号>近似国内<国内牌号>` for a close substitute (e.g. SAE 1022近似国内22#钢, 12L14近似国内Y15Pb, C1010近似国内10#钢). Add the key properties only when they matter (e.g. 导电率 for busbars).
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
