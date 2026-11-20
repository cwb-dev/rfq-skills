# rfq-skills

Model-independent skills for fastener and machined-part RFQs. The repository uses the common Agent Skills layout: each skill is a directory containing a `SKILL.md` with `name` and `description` frontmatter, plus the scripts and resources it needs. No plugin manifest, proprietary runtime, or model-specific path variable is required.

- **drawing-translate**: translates a non-Chinese engineering drawing (PDF, image, scan or fax) into Chinese **on the drawing itself**. Everything affecting manufacturing, inspection, process, cost or quotation is translated, including callouts, notes, dimension/tolerance language and title-block material/finish; numeric values, standard numbers and graphics stay unchanged. Scanned drawings are read with OCR. Uncertain translations go into a review spreadsheet (待审核清单) for a colleague to check.
- **supplier-rfq**: splits a client RFQ into Chinese 询价单 for factories. One Excel file per product group, one sheet per drawing (the Chinese-annotated drawing plus plain-Chinese notes on what drives the price), and a WeChat/email message per factory.
- **proforma-invoice**: creates a Runsun (CNY/VTB) or CWB (USD/ANEXT) Proforma Invoice from the customer's order email, contract/PO and approved `Price to Client` values, with total validation and XLS/PDF output.
- **commercial-invoice**: creates a Commercial Invoice from one or more customer PIs and final packing lists from one or more factories. A line-level PI/factory allocation is required, and factory purchase prices must never be used as customer sales prices.
- **packing-list**: creates a Packing List governed by actual factory packing-list quantities; contracts are comparison references only. Quantity variance over 5% by default produces a user-facing warning without changing actual values. The user must explicitly confirm the physical pallet sequence and shipping mark before generation.

## Requirements

A language model alone does not load skills. Use an agent or client that can read `SKILL.md`, follow its instructions, run the required scripts, and inspect images/PDFs. The exact skill-discovery mechanism belongs to the host, not this repository.

The three invoice skills require **Windows, PowerShell and desktop Microsoft Excel** because they automate the bundled `.xls` templates through Excel COM. PyMuPDF (`python -m pip install pymupdf`) is used to render and inspect the generated PDF. Excel and these scripts are not runnable on macOS/Linux as-is. The RFQ skills' Python dependencies are listed separately below.

## Install

Clone or download this repository, then either:

- configure the host to discover skills from this repository's `skills/` directory; or
- copy each required `skills/<skill-name>/` directory into the host's skills directory, preserving its full contents and sibling layout.

`supplier-rfq` uses `drawing-translate`, so install both even if you primarily need supplier inquiry sheets. If the host does not support automatic skill discovery, provide the relevant `SKILL.md` and its directory to the agent as task context.

No plugin installation or marketplace step is needed. Update by pulling or replacing the repository and refreshing the host's skills if required.

### WorkBuddy

The five skills use the common `SKILL.md` layout and include `agent_created: true` plus bilingual descriptions. Copy the required skill directories to `~/.workbuddy/skills/` and refresh WorkBuddy, or import a ZIP whose root contains `SKILL.md` and all referenced resources. WorkBuddy can use the invoice skills **on Windows only when desktop Excel and PowerShell are available to the agent**; merely importing the skill does not provide Excel or access to company files/mail.

## First-use recommendation shown by the agent

Each invoice skill instructs the agent to give the user a short recommended workflow on its first invocation in a conversation:

- **PI**: provide the customer's PI-request email, the corresponding calculation file's `Price to Client` data, and the customer contract/PO. The email fixes the ordered parts and quantities; approved contract terms take precedence over a quote.
- **CI**: provide every customer PI/sales contract, every factory's actual packing list and the shipment details. The agent first prepares a line-level CI↔PI↔factory allocation. Customer prices come from the customer documents, not factory purchase contracts.
- **PL**: provide the actual factory packing list, matching customer PI/contract and shipment details. Actual packing-list quantities always govern; contract quantities are references only, and variance over 5% by default is reported to the user without altering the PL. Before generation, the agent shows the proposed physical pallet sequence and exact shipping-mark text for the user to mark, correct and explicitly confirm. Missing weights or dimensions are obtained from the source rather than estimated.

The agent then confirms the inputs, generates XLS + PDF, renders the PDF and checks totals and layout. No skill sends email automatically.

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
- `按客户这封要 PI 的邮件和对应计算文件做 PI` → `proforma-invoice`
- `按最终客户 PI 和实际装箱单做这次出货的 CI` → `commercial-invoice`
- `按工厂实际箱单做 PL，并核对箱数、毛净重和体积` → `packing-list`

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
  proforma-invoice/
    SKILL.md
    make_pi.ps1
    templates/
  commercial-invoice/
    SKILL.md
    make_ci.ps1
    templates/
  packing-list/
    SKILL.md
    make_pl.ps1
    templates/
```
