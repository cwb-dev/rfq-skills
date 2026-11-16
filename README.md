# rfq-skills

Model-independent skills for fastener and machined-part RFQs. The repository uses the common Agent Skills layout: each skill is a directory containing a `SKILL.md` with `name` and `description` frontmatter, plus the scripts and resources it needs. No plugin manifest, proprietary runtime, or model-specific path variable is required.

- **drawing-translate**: translates a non-Chinese engineering drawing (PDF, image, scan or fax) into Chinese **on the drawing itself**. Everything affecting manufacturing, inspection, process, cost or quotation is translated, including callouts, notes, dimension/tolerance language and title-block material/finish; numeric values, standard numbers and graphics stay unchanged. Scanned drawings are read with OCR. Uncertain translations go into a review spreadsheet (待审核清单) for a colleague to check.
- **supplier-rfq**: splits a client RFQ into Chinese 询价单 for factories. One Excel file per product group, one sheet per drawing (the Chinese-annotated drawing plus plain-Chinese notes on what drives the price), and a WeChat/email message per factory.

## Requirements

A language model alone does not load skills. Use an agent or client that can read `SKILL.md`, follow its instructions, run Python, and inspect images/PDFs. The exact skill-discovery mechanism belongs to the host, not this repository.

## Install

Clone or download this repository, then either:

- configure the host to discover skills from this repository's `skills/` directory; or
- copy both `skills/drawing-translate/` and `skills/supplier-rfq/` into the host's skills directory, preserving their full contents and sibling layout.

`supplier-rfq` uses `drawing-translate`, so install both even if you primarily need supplier inquiry sheets. If the host does not support automatic skill discovery, provide the relevant `SKILL.md` and its directory to the agent as task context.

No plugin installation or marketplace step is needed. Update by pulling or replacing the repository and refreshing the host's skills if required.

## Python dependencies

Python 3.10+ is required. From the repository root, install the dependencies with:

```sh
python -m pip install --user -r skills/drawing-translate/requirements.txt
```

Use `python3` instead of `python` if that is how Python is exposed on your system. In mainland China, adding `-i https://pypi.tuna.tsinghua.edu.cn/simple` can speed up the download.

To check the dependencies without installing anything:

```sh
python skills/drawing-translate/scripts/check_deps.py
```

The agent must ask before installing missing packages. DengXian or SimHei is used when available on Windows; otherwise, the PDF built-in CJK font is used.

## Use

Ask the agent in natural language:

- `把这张图纸翻译成中文 <file>` → `drawing-translate`
- `make 询价单 for RFQ 1234` → `supplier-rfq`, which also uses `drawing-translate` for every drawing

Command examples in the skills are written from the repository root with ordinary relative paths such as `skills/drawing-translate/scripts/extract.py`. If the skills are installed elsewhere, replace that prefix with the actual skill directory. Placeholders in angle brackets, such as `<work>` and `<inquiry>`, must be replaced with real paths; they are not environment variables. Quote paths containing spaces.

## Layout

```text
skills/
  drawing-translate/
    SKILL.md
    requirements.txt
    scripts/
  supplier-rfq/
    SKILL.md
    example_notes.json
    scripts/
    templates/
```
