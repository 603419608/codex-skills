---
name: organize-invoices
description: 整理中文电子发票 PDF 文件夹。Use when the user asks in Chinese or English to organize/整理/汇总/重命名 a folder of invoices or 发票, including extracting default or user-specified invoice fields, generating an Excel workbook named by the整理日期, applying user-specified or default PDF naming rules, renaming PDFs, checking duplicate seller tax IDs, deleting duplicates only after explicit confirmation, and reporting the total amount.
---

# Organize Invoices

## Core Rules

- Respect any active `AGENTS.md` instructions first.
- Never invoke Claude or any Claude CLI/tool.
- Treat all PDF parsing and folder listing as read-only.
- Before creating Excel files, renaming PDFs, deleting PDFs, or running any command that may write files, state the exact planned changes and wait for explicit user confirmation.
- If a required field is unclear, missing, conflicting, or ambiguous, ask the user instead of guessing.
- For duplicate tax IDs, use the seller/issuer tax ID, not the repeated buyer tax ID. Never delete duplicate PDFs without an explicit deletion plan and confirmation.
- For duplicate seller tax ID groups, do not choose which PDF to keep by default. List the duplicate group and ask the user to confirm the exact file(s) to keep and delete.

## Permission Model

- Proceed without asking for extra permission for read-only work when local rules allow it:
  - list the input folder;
  - read PDF text;
  - run the extractor script;
  - calculate totals;
  - prepare the Excel and rename plan.
- Ask for explicit confirmation before any file-changing work:
  - create or overwrite Excel files;
  - rename PDFs;
  - delete duplicate PDFs;
  - run scripts or commands that may write files.
- If the user wants a one-step workflow, still present the planned writes and wait for confirmation when active local instructions require it.

## Expected Inputs

- A folder path containing invoice PDFs.
- If the user omits the path, ask for it.
- If the user says “帮我整理一下这个文件夹路径下的发票：<path>”, use this skill automatically.

## Output Convention

- Save the Excel workbook into the input folder.
- Name the workbook with the current整理日期 using non-padded month/day:
  `YYYY-M-D-发票整理.xlsx`
- Include a sheet with these columns:
  `开票日期`, `税号`, `开票单位`, `金额`
- Sort rows by `开票日期` ascending.
- Add a final `合计` row or otherwise clearly report the total amount.
- Rename PDFs to:
  `YYYY-MM-DD-开票单位.pdf`

## Customization Rules

- If the user does not specify Excel fields, use the default fields:
  `开票日期`, `税号`, `开票单位`, `金额`.
- If the user explicitly specifies fields to collect, use the requested fields in the Excel output.
- Stable supported fields:
  - `开票日期` / `开票时间`
  - `税号` / `销售方税号` / `统一社会信用代码` / `纳税人识别号`
  - `开票单位` / `销售方名称`
  - `金额`
  - `发票号码`
  - `购买方名称`
  - `购买方税号`
  - `原文件名`
- For fields outside the supported list, try extraction only when the PDF text clearly contains the value. If it cannot be extracted reliably, ask the user before guessing.
- If the user does not specify a PDF naming rule, use the default:
  `YYYY-MM-DD-开票单位.pdf`.
- If the user explicitly specifies a PDF naming template, use it. Support templates such as:
  - `{开票日期}-{开票单位}.pdf`
  - `{开票日期}-{金额}-{开票单位}.pdf`
  - `{发票号码}-{开票单位}.pdf`
  - `{原文件名}-{开票单位}.pdf`
- If the user customizes only Excel fields, keep the default PDF naming rule.
- If the user customizes only PDF naming, keep the default Excel fields.
- If a requested field or naming placeholder cannot be reliably extracted, ask the user before guessing. Examples: `出行人`, `商品明细`, `备注完整内容`, `购方地址电话`, or any field that is absent or inconsistent across PDFs.

## Excel Formatting Requirements

- Format the workbook for readability, not just data correctness.
- Use enough column width and spacing so text is not clipped or crowded:
  - `开票日期`: wide enough for full `YYYY-MM-DD` dates.
  - `税号`: wide enough for 18-character tax IDs.
  - `开票单位`: wide enough for long Chinese company names.
  - `金额`: wide enough for currency values.
- Center all visible table cells horizontally and vertically.
- Make the header row bold, centered, and filled with a light blue background.
- Make the `合计` row bold and filled with a light orange/peach background.
- Display amount values with two decimal places, including the total.
- Store identifier columns as true text cells, not numbers. This applies to `税号`, `统一社会信用代码`, `纳税人识别号`, `发票号码`, and any other long numeric identifier.
- Prefer Excel `sharedStrings` or an equivalent broadly compatible text-cell representation for long identifiers. Do not rely only on number formats when the value is an all-digit long identifier.
- For all-digit long identifiers such as `914403006939641518`, verify they display in full and never as scientific notation like `9.14403E+17`.
- After exporting the workbook, run `scripts/patch_xlsx_text_columns.py` on identifier columns when available, then re-open or inspect the workbook to verify long identifiers still display in full.
- Use practical row heights and column widths instead of narrow defaults.
- After creating the workbook, inspect or render the visible table range and fix severe layout issues before finalizing. Avoid outputs where columns overlap visually, values are clipped, or the sheet resembles a compressed default export.

## Field Rules

- `开票日期`: invoice issue date, normalized as `YYYY-MM-DD`.
- `税号`: seller/issuer unified social credit code or taxpayer ID.
- `开票单位`: seller/issuer name.
- `金额`: invoice price-tax total, usually the small amount near `价税合计（小写）`; keep two decimals.
- If one buyer name/tax ID repeats across most PDFs, treat that repeated entity as the buyer and choose the other entity as the seller.
- If multiple records would produce the same target PDF name, stop and ask how to handle the collision.

## Workflow

1. Resolve and list the input folder. Confirm it contains PDFs.
2. Run `scripts/extract_invoices.py` as a read-only extraction pass:

   ```bash
   python scripts/extract_invoices.py "<folder-path>" --pretty
   ```

   Prefer the bundled Codex Python runtime when available. The script reads PDFs and prints JSON only.

3. Review the JSON:
   - `records`: extracted invoice rows and proposed PDF names.
   - `missing_required`: files missing date, seller tax ID, seller name, or amount.
   - `duplicate_seller_tax_groups`: seller tax IDs appearing in more than one PDF.
   - `rename_collisions`: proposed target PDF names used by multiple files.
   - `warnings`: extraction issues that may require user confirmation.

4. Apply the customization rules:
   - default fields and default PDF naming when the user did not specify alternatives;
   - user-specified fields and naming templates when provided.
5. If any required field is missing, any requested custom field cannot be extracted, any target name collides, or seller/buyer detection is ambiguous, ask the user a concise question and wait.
6. Present the write plan before changing files:
   - exact Excel path to create or overwrite;
   - exact PDF renames from original name to target name;
   - exact duplicate PDFs proposed for deletion, if any;
   - total amount.
7. After explicit confirmation:
   - create the Excel workbook using the spreadsheet tooling available in the environment;
   - apply the Excel formatting requirements above;
   - run `scripts/patch_xlsx_text_columns.py <xlsx-path>` when the workbook includes identifier columns such as `税号`, `统一社会信用代码`, `纳税人识别号`, or `发票号码`;
   - verify workbook contents, total, and visible layout;
   - rename PDFs only after checking that targets do not already exist;
   - delete duplicates only if the user confirmed the exact files to delete and the exact file to keep.
8. Final response:
   - link the Excel file;
   - report total amount;
   - report how many PDFs were renamed;
   - report duplicate deletions, or state that none were deleted.

## Safety Defaults

- If an output Excel already exists, ask before overwriting.
- If a rename target exists, stop and ask instead of overwriting.
- If the user previously asked for “相同税号删除”, still confirm the exact deletion list because deletion is destructive.
- Keep source PDFs untouched until after the Excel has been generated and verified.
