# cwb-skills

Model-independent Agent Skills for the CWB / Runsun fastener and machined-part workflow. The GitHub repository is `cwb-dev/rfq-skills`; the recommended local folder name is `cwb-skills`.

Each skill has a `SKILL.md` with `name` and `description` frontmatter plus the scripts, templates and assets it needs. No plugin manifest or model-specific runtime is required.

## Included skills

### Inquiry and quotation

- **new-inquiry** — collect an RFQ email and attachments, create the inquiry folder, extract line items and update the tracker.
- **mail-check** — incremental mail check, new-RFQ detection, tracker updates and daily digest.
- **inquiry-status** — report new, chased, overdue and recently quoted inquiries.
- **supplier-rfq** — split an RFQ into Chinese supplier inquiry workbooks with translated drawing sheets.
- **drawing-translate** — translate engineering-drawing text into Chinese on the drawing, with OCR and review output.
- **quote-compile** — compile factory quotes into cost/margin workbooks, comparisons, client forms and draft replies.
- **standard-stock-price** — query GPYH and Taodingding options and build a dropdown-linked stock-price workbook.

### Orders and shipment documents

- **purchase-contract** — create a Chinese Runsun factory purchase contract, with RMB validation and optional drawing annexes.
- **sample-order** — create a Chinese sampling agreement as DOCX plus combined PDF with drawing annexes.
- **proforma-invoice** — create a Runsun or CWB Proforma Invoice from confirmed customer evidence.
- **commercial-invoice** — create a Commercial Invoice with line-level PI/factory allocation.
- **packing-list** — create a Packing List governed by actual packing data, with explicit pallet-number and shipping-mark confirmation.

The public skills contain generic workflow instructions and fictitious examples. Private client profiles, mail accounts, network paths, order evidence, financial records and credentials are not part of this repository.

## Platform support

| Capability | Windows | macOS |
|---|---|---|
| RFQ collection, tracking and quotation workflows | Yes | Yes, with configured mail tools and shared storage |
| Drawing translation, supplier RFQ, sample order and stock workbook scripts | Yes | Yes, with Python dependencies |
| PI, CI, PL and purchase-contract generation | Yes, with PowerShell and desktop Excel | **Not natively**: the generators use Windows Excel COM |

On macOS, use a Windows environment for the four Excel COM document generators. Mounting the shared drive alone does not port these scripts. See [HANDOFF.md](HANDOFF.md) for migration and shared-drive setup.

## Install

Clone using the recommended folder name:

```sh
git clone git@github.com:cwb-dev/rfq-skills.git cwb-skills
```

Then either configure the agent to discover this repository's `skills/` directory, or copy/symlink each required `skills/<skill-name>` directory into the host's skills directory while preserving its contents.

For hosts using `~/.agents/skills`:

```sh
mkdir -p ~/.agents/skills
for skill in "$HOME/cwb-skills/skills"/*; do
  ln -sfn "$skill" "$HOME/.agents/skills/$(basename "$skill")"
done
```

`supplier-rfq` uses `drawing-translate`, so install both. Private client rules belong in the restored workspace, not in the public skill folders.

## Python dependencies

Python 3.10+ is recommended. From the repository root:

```sh
python -m pip install --user -r skills/drawing-translate/requirements.txt
python skills/drawing-translate/scripts/check_deps.py
```

Use `python3` if that is the executable name on the target machine. The sample-order generator additionally uses Pillow, and the stock workbook uses openpyxl. Install dependencies only with the user's approval.

## Use

Ask the agent naturally, for example:

- `收集这个 RFQ 并建立询价档案` → `new-inquiry`
- `把图纸翻译成中文` → `drawing-translate`
- `按产品类别制作供应商询价单` → `supplier-rfq`
- `汇总工厂报价并填写客户报价表` → `quote-compile`
- `按客户 PO 和工厂确认条款做采购合同` → `purchase-contract`
- `按客户订单和计算文件做 PI` → `proforma-invoice`
- `按最终 PI 和实际装箱单做 CI` → `commercial-invoice`
- `按实际箱单做 PL` → `packing-list`

No skill should send email, move/flag mail, place orders or overwrite formal source documents without explicit authorization.

## Restore a complete private workflow

The public repository is only the portable skill layer. A complete migration also needs the separately transferred private ZIP, direct access to the underlying shared file service, and locally re-entered credentials. Follow [HANDOFF.md](HANDOFF.md); do not upload the private ZIP or its contents back to this repository.
