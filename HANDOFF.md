# CWB 工作流迁移与恢复 Handoff

日期：2026-11-23

## 1. 迁移包如何配合

完整恢复需要两部分：

1. 公开仓库 `cwb-dev/rfq-skills`，本地目录名使用 `cwb-skills`：12 个脱敏 skills、脚本、模板和本说明。
2. 私密 ZIP：原始项目/用户 skills、工作区规则、客户档案和台账、邮件检查状态、财务脚本与已有试点数据、历史工作文件，以及私密版 handoff。

**私密 ZIP 不上传 GitHub。** ZIP 不包含邮箱密码、企业微信 secret、SSH 私钥、会话存档私钥或 `.env` 等凭据。ZIP 本身未加密，只通过加密磁盘或受控加密传输交给目标电脑。

网络盘上的原始询价、合同和出货证据不会重复塞进 ZIP；需要通过底层文件服务重新连接。Windows 本机 E: 上的财务源文件也不在 ZIP 内，需另行共享或复制。

## 2. 目标 Mac 推荐目录

```text
~/cwb-skills                 Git 仓库
~/cwb-workflow               原 F:\Claude 工作区
~/cwb-shared                 原 Y: 内容根目录的稳定链接
~/cwb-finance-sources        原 E: 财务源文件（按需挂载/复制）
~/.agents/skills             Kimi Code 等宿主的用户级 skills
```

建议使用以下路径映射，不要在 Mac 上创建名为 `Y:\...` 的伪路径：

| 来源 Windows | Mac |
|---|---|
| `F:\Claude` | `$HOME/cwb-workflow` |
| `Y:\` | `$HOME/cwb-shared` |
| `E:\ANEXT` | `$HOME/cwb-finance-sources/ANEXT` |
| `E:\银行对账单` | `$HOME/cwb-finance-sources/银行对账单` |
| `E:\企业微信表格*` | `$HOME/cwb-finance-sources/企业微信表格*` |

在 shell 和 agent 的工作环境中设置：

```bash
export CWB_WORKSPACE="$HOME/cwb-workflow"
export CWB_Y_ROOT="$HOME/cwb-shared"
export CWB_E_ROOT="$HOME/cwb-finance-sources"
```

私密 `CLAUDE.md` 和旧台账中的 Windows 路径是来源机器记录。恢复后应在 Mac 工作区的私有规则中补充上述映射；不要让 agent 在 Y: 全盘上搜索，仍只访问目标客户的单一文件夹。

## 3. 恢复仓库、工作区和 skills

```bash
git clone git@github.com:cwb-dev/rfq-skills.git "$HOME/cwb-skills"
mkdir -p "$HOME/cwb-transfer" "$HOME/cwb-workflow"
unzip <私密ZIP路径> -d "$HOME/cwb-transfer"
ditto "$HOME/cwb-transfer/workspace" "$HOME/cwb-workflow"
```

ZIP 中：

- `workspace/`：恢复到 `~/cwb-workflow`。
- `user-skills/`：来源机器用户级原始 skills，仅保留在本机。
- `HANDOFF.md`：含来源机器具体网络信息的私密版说明。

安装公开 skills 的推荐方式：

```bash
mkdir -p "$HOME/.agents/skills"
for skill in "$HOME/cwb-skills/skills"/*; do
  ln -sfn "$skill" "$HOME/.agents/skills/$(basename "$skill")"
done
```

如需逐字恢复来源机器的内部版 skills，可用 ZIP 的 `workspace/.claude/skills` 和 `user-skills` 覆盖本机安装，但不能把这些内部版本推回公开仓库。Claude Code 可直接使用工作区的 `.claude/skills`；其他宿主按其技能发现目录配置。

## 4. 在 macOS 连接原 Y: 服务

原 Y: 是 Windows 上的 RaiDrive 虚拟映射，不是可供 Mac 直接访问的 SMB 地址。**不要尝试挂载来源电脑的 RaiDrive 本地虚拟地址。** 必须取得底层 NAS/WebDAV/SMB 服务的真实地址、共享根目录和账户；这些信息不在公开仓库或凭据 ZIP 中。

### SMB

Finder → `前往` → `连接服务器`（Command-K）：

```text
smb://<服务器>/<共享名>
```

登录后通常挂载在 `/Volumes/<共享名>`。确认挂载根目录能直接看到原 Y: 下的 `询价`、`外贸合同`、`采购合同`、`出货文件` 等目录；若只看到上级目录，链接时应指向对应子目录。

### WebDAV

若底层服务提供 WebDAV，在 Finder 的“连接服务器”中使用服务提供的 HTTPS WebDAV URL：

```text
https://<服务器>/<WebDAV路径>
```

不要把密码写入 URL、脚本、shell 历史或 handoff。可将凭据保存在 macOS 钥匙串，并用“登录项目”重新连接。若服务只能通过 RaiDrive 访问，需要在该服务上启用直接 SMB/WebDAV，或使用服务官方 Mac 客户端；不能把 Windows 映射盘本身当成网络服务。

建立稳定链接：

```bash
ln -sfn "/Volumes/<实际挂载目录>" "$HOME/cwb-shared"
```

自动挂载未验证前不要运行批量询价或财务任务。每次恢复后先检查中文目录、文件名和读写权限；写入测试只在新建测试目录进行，不覆盖正式文件。

## 5. Python 与依赖

Mac 上安装 Git、Python 3.10+ 和 `uv`；不要复制 Windows 的可执行文件或 uv 缓存。建议使用虚拟环境：

```bash
python3 -m venv "$HOME/.venvs/cwb"
source "$HOME/.venvs/cwb/bin/activate"
python -m pip install -r "$HOME/cwb-skills/skills/drawing-translate/requirements.txt"
python -m pip install pillow xlrd pypdf cryptography
python "$HOME/cwb-skills/skills/drawing-translate/scripts/check_deps.py"
```

依赖安装需由用户确认。OCR 首次运行可能还需要下载运行库。财务脚本优先使用各脚本/README 中的 `uv run --with ...` 方式，避免依赖环境污染。

## 6. 263 邮箱 MCP

私密 ZIP 保留了来源机器的 `263MAIL_MCP_HANDOFF.md`，按该文件配置 Mac mini / Codex。核心启动参数：

```text
command: uvx
args: ["mcp-email-server@1.13.1", "stdio"]
startup/tool timeout: 60000 ms
```

由用户自行运行账户 UI/CLI 并输入密码或授权码。不要从 RaiDrive、MCP 配置或聊天中提取凭据。初始只开放收信/读取，保持发信关闭；验证时先列账户，再列少量 INBOX 元数据，正文读取使用 `mark_as_read=false`。

`Mail-Check/state.json` 已保留，可防止重复处理；但每日 9:00 的 Windows 计划任务不会自动迁移。先在 Mac 手动完成一次成功的 `/mail-check`，再用目标 agent 的调度器或 launchd 重建计划任务。

## 7. 企业微信与财务

- `CWB Finances` 的脚本、文档、已有数据库/Excel 和仪表盘数据在私密工作区。
- 企业微信 secret、会话存档 secret、回调密钥和 RSA 私钥均未打包，必须由用户在本机或云服务器重新配置。
- 云端回调、会话存档和机器人服务仍在原云服务器，不应在 Mac 上重复启动 collector，否则可能出现重复通知或游标竞争。
- 当前财务试点可在已有数据上重建：

```bash
cd "$HOME/cwb-workflow/CWB Finances"
uv run --with openpyxl python -I build_finance_pilot.py --data-dir pilot_data --output FINANCE_PILOT.db --xlsx FINANCE_PILOT.xlsx
python -I build_dashboard.py --db FINANCE_PILOT.db --output dashboard.html
open dashboard.html
```

- `build_ledger.py` 和专项 recon 脚本含来源机器的 E:/Y: 默认路径。Mac 上先改为上述环境变量/挂载路径，再重建；不要直接运行后把空目录或错误数据当成正式结果。
- 财务源文件必须另行连接。只有 ZIP 中的试点数据时，可以查看试点仪表盘，不能据此声称已恢复完整银行/订单证据链。

## 8. Excel 文档生成的限制

以下四个生成器使用 **Windows Excel COM**，不能因在 Mac 安装了 PowerShell 或 Excel for Mac 就直接运行：

- `proforma-invoice/make_pi.ps1`
- `commercial-invoice/make_ci.ps1`
- `packing-list/make_pl.ps1`
- `purchase-contract/make_pc.ps1`

在 Mac 上恢复这些工作需要 Windows 虚拟机或另一台 Windows 电脑，安装桌面版 Microsoft Excel，并让它访问同一共享目录。其他 Python skills 可在 Mac 原生运行。若要求全部原生 macOS，必须另行重写这四个生成器并重新验证 XLS/PDF 版式；本次迁移没有伪装完成这项移植。

## 9. 恢复后的验收顺序

1. `git -C ~/cwb-skills status` 正常，公开仓库不含私密 ZIP 内容。
2. agent 能看到 12 个 skills；`check_deps.py` 通过。
3. `~/cwb-shared/询价` 等中文目录可读，目标客户单目录可正常访问；测试新建/删除一个空测试文件。
4. 263 MCP 能列账户和少量邮件元数据，未修改邮件状态。
5. 对一个测试 RFQ 完成收集 → 供应商询价 → 报价汇总，输出到测试目录。
6. 打开已有财务仪表盘；如需重建总账，先验证 E:/Y: 源文件路径和金额，再运行。
7. 在 Windows 环境用虚构 spec 各生成一份 PI、CI、PL、采购合同 XLS+PDF，渲染检查；不得覆盖正式单据。
8. 最后再恢复每日邮件调度，并观察首轮 digest 是否重复。

## 10. 不包含或仍需人工提供

- 底层 Y: 文件服务的真实地址和账号。
- E: 财务源文件的 Mac 访问方式。
- 所有密码、secret、私钥和本机 MCP 认证状态。
- 原 Windows 计划任务和 RaiDrive 配置。
- 四个 Excel COM 文档生成器的原生 macOS 移植。
- 临时 scratchpad、缓存和测试输出；这些不作为正式工作流依据。
