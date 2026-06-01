# codex-skills

个人可分享的 Codex skills 仓库。每个 skill 都是一个独立文件夹，可以安装到 Codex 的 skills 目录中使用。

## Skills

| Skill | 说明 | 路径 |
|---|---|---|
| 整理发票 | 整理中文电子发票 PDF，提取字段，生成 Excel，汇总金额，并按规则重命名 PDF。支持默认字段、自定义字段、自定义命名规则和长编号文本保护。 | `skills/organize-invoices` |

## 安装方式

### 方式 1：使用 Codex Skill Installer

在 Codex 对话中运行：

```text
$skill-installer install https://github.com/603419608/codex-skills/tree/main/skills/organize-invoices
```

安装后重启 Codex，或新开一个 Codex 会话。

### 方式 2：手动安装

下载本仓库后，把 skill 文件夹复制到本机 Codex skills 目录。

Windows：

```text
C:\Users\<你的用户名>\.codex\skills\organize-invoices
```

macOS / Linux：

```text
~/.codex/skills/organize-invoices
```

最终结构应类似：

```text
~/.codex/skills/organize-invoices/SKILL.md
~/.codex/skills/organize-invoices/agents/openai.yaml
~/.codex/skills/organize-invoices/scripts/extract_invoices.py
~/.codex/skills/organize-invoices/scripts/patch_xlsx_text_columns.py
```

## 使用方式

默认整理发票：

```text
帮我整理一下这个文件夹路径下的发票：C:\Users\xxx\Desktop\发票\2026-5-28
```

手动指定 skill：

```text
$organize-invoices 帮我整理一下这个文件夹路径下的发票：C:\Users\xxx\Desktop\发票\2026-5-28
```

自定义字段和命名规则：

```text
$organize-invoices 帮我整理这个路径下的发票，收集开票时间、税号和金额，PDF 命名按 {开票日期}.pdf 保存：C:\Users\xxx\Desktop\发票
```

## organize-invoices 默认行为

默认 Excel 字段：

- 开票日期
- 税号
- 开票单位
- 金额

默认 PDF 命名：

```text
YYYY-MM-DD-开票单位.pdf
```

默认输出：

- Excel 保存到输入文件夹
- Excel 文件名按整理日期命名，例如 `2026-6-1-发票整理.xlsx`
- 最后一行汇总金额
- 检查重复销售方税号
- 长数字税号、发票号码按文本保存，避免科学计数法

## 安全规则

- 读取 PDF、提取字段、计算金额属于只读操作。
- 创建 Excel、覆盖 Excel、重命名 PDF、删除 PDF 前，应先给用户确认清单。
- 删除重复发票前，必须让用户确认具体保留和删除哪些文件。
- 不调用 Claude CLI 或任何 Claude 相关工具。

## 限制

- 主要适用于可复制文字的中文电子发票 PDF。
- 扫描版/图片版 PDF 可能需要 OCR，不在当前 skill 默认能力内。
- 特殊发票版式可能需要人工确认字段。

## 仓库结构

```text
codex-skills/
├─ README.md
├─ LICENSE
├─ .gitignore
└─ skills/
   └─ organize-invoices/
      ├─ SKILL.md
      ├─ agents/
      │  └─ openai.yaml
      └─ scripts/
         ├─ extract_invoices.py
         └─ patch_xlsx_text_columns.py
```

## License

MIT License
