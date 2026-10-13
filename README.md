# rfq-skills

Claude Code plugin with two skills for fastener / machined-part RFQs:

- **drawing-translate**: translates a non-Chinese engineering drawing (PDF, image, scan or fax) into Chinese **on the drawing itself**. Only the arrow callouts and the NOTES are translated; dimensions, title block and graphics stay untouched. Scanned drawings are read with OCR. Uncertain translations go into a review spreadsheet (待审核清单) for a colleague to check.
- **supplier-rfq**: splits a client RFQ into Chinese 询价单 for factories. One Excel file per product group, one sheet per drawing (the Chinese-annotated drawing plus plain-Chinese notes on what drives the price), and a WeChat/email message per factory.

## Install

In a terminal:

```
claude plugin marketplace add cwb-dev/rfq-skills
claude plugin install rfq-tools@rfq-skills
```

Or inside Claude Code: `/plugin marketplace add cwb-dev/rfq-skills`, then `/plugin install rfq-tools@rfq-skills`.

Update later with `claude plugin marketplace update rfq-skills` and `claude plugin update rfq-tools@rfq-skills` (or `/plugin` → Marketplaces → Update), then `/reload-plugins`.

## Python dependencies

You need Python 3.10+ on the machine. The packages are:
- `pymupdf`
- `openpyxl`
- `rapidocr_onnxruntime` (OCR, about 100 MB with onnxruntime and opencv)

The skills check them on first use, and Claude asks before installing anything. To install them yourself:

```
python -m pip install --user pymupdf openpyxl rapidocr_onnxruntime
```

In mainland China, add `-i https://pypi.tuna.tsinghua.edu.cn/simple` for a much faster download.

Fonts: DengXian or SimHei, which come with Windows. On other systems, the PDF built-in CJK font is used.

## Use

- "把这张图纸翻译成中文 <file>" → drawing-translate
- "make 询价单 for RFQ 1234" → supplier-rfq (it calls drawing-translate for every drawing)
