---
name: excel-translate
description: 将 Excel (.xlsx) 文件中的文本翻译成另一种语言，同时完整保留图片（含 wmf 格式）、公式、单元格样式与富文本格式。当用户提到翻译 xlsx 文件、把 Excel 转成其他语言、本地化工作簿、多语言 Excel、给 Excel 加英文/中文/日文版本等需求时，必须使用此 skill。即使用户没有明确说 "translate skill"，只要意图是「把 xlsx 里的文字改成另一种语言」就应该触发。
---

# Excel 翻译

把 `.xlsx` 文件中的文本翻译成目标语言，**完整保留**：图片（含 wmf）、公式、单元格样式、富文本 run 格式、命名空间与所有非文本内容。

## 为什么需要这个 skill

直接用 openpyxl 重写整份工作簿会**丢失 wmf 图片**（openpyxl 不支持该格式，会静默丢弃），还会把命名空间前缀（`r:`、`mc:` 等）改成 `ns0:`、`ns1:`，破坏 Excel 关系引用导致打开时报「文件已损坏」。本 skill 通过只编辑 xlsx 内部的 XML 文本节点来规避这些问题。

## 工作流程

翻译是一个三步流程，中间的翻译由你（Claude）直接完成：

1. **提取** — 运行 `extract.py` 把工作簿中所有唯一的文本值（去重后）输出为 `texts.json`
2. **翻译** — 你读取 `texts.json`，把每个值翻译成目标语言，写入 `translations.json`（键值对：原文 → 译文）
3. **应用** — 运行 `apply.py` 把翻译写回 xlsx，输出到用户指定的新文件

### Step 1：提取

```bash
python ~/.claude/skills/excel-translate/scripts/extract.py <input.xlsx>
```

会在当前工作目录生成 `texts.json`，格式：

```json
{
  "texts": ["主机", "弹簧", "按钮", ...]
}
```

去重逻辑：openpyxl 读取每个工作表的每个单元格，对 `cell.value` 做 `.strip()`，非空字符串加入 set。

### Step 2：翻译（你来完成）

读取 `texts.json`，把每条文本翻译成目标语言。**逐条翻译、保持完整性**，不要漏译、合并或拆分。翻译完成后写入 `translations.json`：

```json
{
  "主机": "Main unit",
  "弹簧": "Spring",
  "按钮": "Button"
}
```

注意：键必须是原文（包括原始的标点、空格、数字编号前缀如 `1-1.`、`步骤1：`），值是对应译文。原文里若有 `&`、`<`、`>` 等字符，按字面保留即可，apply.py 会处理 XML 转义。

### Step 3：应用

```bash
python ~/.claude/skills/excel-translate/scripts/apply.py <input.xlsx> translations.json <output.xlsx>
```

会读取 `translations.json` 作为翻译表，对原 xlsx 内部的 XML 做精准文本替换，输出到 `<output.xlsx>`。

## 触发示例

下面这些都是本 skill 应当触发的典型说法：

- 「将文件 `./excel/translate.xlsx` 由中文翻译成英文，新的文件保存到 `./translate-translated.xlsx` 中」
- 「把这个 xlsx 翻译成日语，输出到 `translated.xlsx`」
- 「帮我把这份 Excel 工作簿本地化成英文」
- 「translate this Excel file to Spanish」
- 「这个 SOP 表格里的中文换成英文版」

## 应用原则（apply.py 的关键设计）

- **只命中整段 `<t>` 内容**：正则匹配 `<t...>内容</t>`，仅当「内容」整体在翻译表中存在时才替换。这避免了部分匹配导致的语义破坏。
- **XML 转义**：原文与译文都用 `html.escape` 转义后再比对与写入，保证 `&`、`<`、`>` 在 XML 中正确。
- **富文本保留**：富文本单元格由多个 `<r><rPr>格式</rPr><t>片段</t></r>` 组成。如果某个 `<t>` 片段自身恰好等于翻译表中某个键，就替换该片段（保留 run 格式）；若是「整句跨多个 run 拼成」的情况，单片段不会命中，整句保留原文不翻译——这是为了不破坏格式而做的取舍。
- **图片 / 公式 / 样式 / drawing / 命名空间 / 关系文件**：在 zip 层面原样拷入新文件，完全不动。

## 限制

- 跨多个 run 的富文本整句无法被翻译（会保留原文）。如果用户大量依赖这种格式且需要全部翻译，可在反馈中说明，未来可考虑按比例分配译文到各 run 的扩展。
- 单元格中嵌入图片里的文字（如截图文字）不会被翻译——本 skill 只处理可编辑的文本节点。
- 翻译键是 `cell.value.strip()` 的结果。若原工作簿中存在肉眼难以察觉的前后空格差异，可能造成未命中。若发现某条没翻译，先检查空格。

## 调试提示

如果生成的文件仍有内容未被翻译，按以下顺序排查：

1. `extract.py` 是否覆盖了所有工作表？默认遍历所有 sheet。
2. `translations.json` 的键是否与 `texts.json` 完全一致（包括标点、空格）？diff 两个文件的键集合最稳妥。
3. 富文本单元格的整句如果跨多个 run，确实不会翻译——这是已知限制，不是 bug。
