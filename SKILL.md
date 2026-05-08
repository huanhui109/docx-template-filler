---
name: docx-template-filler
description: Fill Word document templates with content from source documents while preserving all formatting. Supports TWO template types: (1) Feasibility Analysis Reports (可行性分析报告), (2) Change/Release Documents (变更上线文档). Use when: user provides a .docx/.doc template and source documents; user needs to complete a template maintaining original formatting, styles, headings, and layout; user says things like "fill the template", "根据模板填写", "补充模板内容", "变更文档", "上线方案". Supports Chinese and English documents.
---

# Word Template Filler

Fill Word templates with extracted content while preserving **all** formatting.

## Fixed Template Paths

The skill includes fixed templates stored in the skill directory:

```
skills/docx-template-filler/
├── scripts/
│   ├── docx_utils.py       # Shared utilities (NEW)
│   ├── docx_parse.py       # Parse template → structure JSON
│   ├── docx_fill.py        # Fill template (font + numbering aware)
│   ├── docx_gen_change.py  # Generate Type B change doc
│   ├── docx_check.py       # Check completeness + clean punctuation (ENHANCED)
│   └── docx_meta.py        # Metadata, headers, indent, TOC, versions
└── templates/               # Fixed templates (NEW)
    ├── feasibility_template.docx   # 可行性分析报告模板
    └── change_template.docx        # 变更上线文档模板
```

### Template Usage

**Option 1: Use Fixed Templates (Recommended)**
```bash
# 可行性分析报告 - use built-in template
TEMPLATE_PATH=~/.openclaw/skills/docx-template-filler/templates/feasibility_template.docx

# 变更上线文档 - use built-in template  
TEMPLATE_PATH=~/.openclaw/skills/docx-template-filler/templates/change_template.docx
```

**Option 2: Use User-Provided Template**
```bash
python3 scripts/docx_parse.py <user_template.docx> --output /tmp/struct.json
```

---

## Prerequisites

```bash
pip3 install python-docx pdfplumber python-pptx
```

---

## Interactive Workflow

When user asks to fill a template, follow this workflow:

### Smart Template Resolution (Recommended)

Use the `resolve_template()` function for automatic template handling:

```python
from docx_utils import resolve_template, prompt_template_type_choice

# Auto-resolve template (user template takes priority, fallback to built-in)
result = resolve_template(
    user_template_path=user_provided_path,  # Optional: user's template
    template_type=requested_type,           # Optional: "feasibility" or "change"
    fallback_to_builtin=True                # Use built-in if no user template
)

if result["error"]:
    print(f"Error: {result['error']}")
else:
    template_path = result["template_path"]
    template_type = result["template_type"]
    source = result["source"]  # "user", "builtin", "auto_detected"
```

### Manual Template Selection

**If user provides a template file:**
- Auto-detect type using `detect_template_type()` function
- Or ask user to confirm

**If no template provided, ask user to choose:**
```python
from docx_utils import prompt_template_type_choice

# Show prompt to user
print(prompt_template_type_choice())
```
Or manually specify:
```
请选择文档类型：
1. **可行性分析报告** — 项目立项可行性分析报告
2. **变更上线文档** — 系统变更实施方案与操作指引

请输入数字(1或2)或文档类型名称(可行性/变更)
```

### Template Resolution Priority

1. **User provided template** → Analyze content, detect type, use as-is
2. **User specified type** → Use built-in template for that type
3. **Auto-detect** → Analyze content keywords, select best matching built-in
4. **Fallback** → Default to "变更上线文档" template

### Analyze User Template

```python
from docx_utils import analyze_user_template

analysis = analyze_user_template("/path/to/user_template.docx")

print(f"模板类型: {analysis['type']}")
print(f"章节数量: {len(analysis['sections'])}")
print(f"占位符: {analysis['placeholders']}")
```

### Get Template Path

**For fixed templates:**
```python
from docx_utils import get_template_path, TEMPLATE_TYPE_FEASIBILITY, TEMPLATE_TYPE_CHANGE

if template_type == "可行性分析报告" or template_type == "feasibility":
    template_path = get_template_path(TEMPLATE_TYPE_FEASIBILITY)
else:
    template_path = get_template_path(TEMPLATE_TYPE_CHANGE)
```

**For user templates:**
```python
template_path = user_provided_template_path
```

### Step 3: Parse and Fill

```bash
# 1. Parse template
python3 scripts/docx_parse.py <template.docx> --output /tmp/struct.json

# 2. Generate content from source documents (agent creates content.json)

# 3. Fill content
python3 scripts/docx_fill.py <template.docx> content.json <output.docx>

# 4. Clean punctuation + check
python3 scripts/docx_check.py <output.docx> /tmp/struct.json --clean <output.docx>

# 5. First-line indent (2 chars)
python3 scripts/docx_meta.py indent <output.docx> <output.docx> --chars 2

# 6. Mark TOC dirty
python3 scripts/docx_meta.py toc <output.docx> <output.docx>

# 7. Update metadata + header
python3 scripts/docx_meta.py update <output.docx> <output.docx> \
  --project-name "项目名称" --author "作者" --department "部门" \
  --revision-date "YYYY-MM-DD"

# 8. Version save
python3 scripts/docx_meta.py save <output.docx> <workspace> --name "文件名"
```

---

## Document Type A: 可行性分析报告

### Content Strategy
- Extract project background, requirements, technical analysis from source docs
- Fill each heading section with relevant paragraphs
- Maintain formal report tone
- Support rich tables for project plan, resources, benefits analysis

### Template Structure

```
可行性分析报告
├── 修订记录 (table: 版本/作者/日期/内容)
├── 文档信息 (table: 版本/编制人/需求部门/编制部门)
├── 引言
├── 可行性分析前提 (table: 数据说明/公网服务/权限管理)
├── 项目背景
├── 需求分析 (table: 序号/工作内容/开始时间/结束时间)
├── 技术方案
├── 项目计划 (table: 序号/工作内容/开始时间/结束时间)
├── 资源配置 (table: 序号/项目/数量/单价/总价/备注)
├── 效益分析 (table: 项目/金额/备注)
└── 结论与建议
```

### Content JSON Format
```json
{
  "template_type": "feasibility",
  "project_name": "项目名称",
  "metadata": {
    "version": "V1.0",
    "author": "项目经理",
    "department": "信息技术中心"
  },
  "sections": {
    "引言": {
      "content": ["引言内容..."]
    },
    "项目背景": {
      "content": ["背景内容..."]
    },
    "需求分析": {
      "content": ["需求描述..."]
    },
    "技术方案": {
      "content": ["技术实现方案..."]
    }
  }
}
```

---

## Document Type B: 变更上线文档

### Content Strategy
- **Detailed step-by-step operations** — every action must be explicit and sequential
- **Table format** for structured info (change plan, parameters, checklist)
- **Screenshot placeholders** — insert `[截图: XXX界面]` markers for user to add screenshots later
- **Foolproof language** — write so that any operator can follow without ambiguity
- **Rollback plan** — always include rollback steps
- **Pre/Post checks** — verification steps before and after change

### Template Structure

```
某某系统变更方案
├── 文档信息 (table: 版本/编制/校对/部门)
├── 部署架构图（系统上线需要，一般变更不需要）
├── 变更计划简要说明 (table: 涉及系统/变更项目/具体内容/实施人/复核人/时间)
├── 变更目的
├── 变更内容
│   ├── 变更前准备
│   └── 变更执行
├── 变更步骤
│   ├── 升级完成后检查
│   └── 变更回退
├── 参数设置（可选）(table: 系统参数)
├── 配置更新 (table: 项目/说明)
├── 软件及文档更新
├── 网络配置
├── 硬件配置
├── 变更后工作
└── 变更后检查
```

### Content JSON Format
```json
{
  "template_type": "change",
  "project_name": "FISP代销迁移直连",
  "metadata": {
    "version": "V1.0",
    "author": "张三",
    "department": "信息技术中心",
    "date": "2026-05-07"
  },
  "sections": {
    "变更目的": {
      "content": ["详细说明变更的背景和目标..."]
    },
    "变更内容": {
      "content": ["本次变更涉及以下内容：", "1、...", "2、..."]
    },
    "变更前准备": {
      "content": [
        "【环境准备】",
        "1、确认O32系统版本为V202302.20.000",
        "2、申请直连代销渠道前置机授权",
        "3、准备迁移脚本文件"
      ]
    },
    "变更步骤": {
      "content": [
        "步骤1：环境检查",
        "操作：登录O32系统，确认系统版本和当前运行状态",
        "预期：系统正常运行，版本为V202302.20.000",
        "[截图: O32系统版本信息界面]"
      ]
    },
    "升级完成后检查": {
      "content": [
        "【检查清单】",
        "☐ 产品登记界面【有效】状态不变",
        "☐ 基金开户界面FISP账号全部变为【注销】"
      ]
    },
    "变更回退": {
      "content": [
        "如迁移失败，执行以下回退步骤：",
        "1、使用备份脚本恢复T日迁移相关表数据",
        "2、恢复FISP渠道配置参数",
        "3、重启转换机任务"
      ]
    },
    "配置更新": {
      "table": {
        "headers": ["项目", "说明"],
        "rows": [
          ["O32系统渠道配置", "FISP代销渠道变更为直连代销渠道"],
          ["系统参数10123", "勾选303渠道"]
        ]
      }
    }
  }
}
```

### Three Section Formats

1. **Text paragraphs**: `"content": ["line1", "line2", ...]`
2. **Simple table**: `"table": { "headers": [...], "rows": [[...], ...] }`
3. **Structured steps table**: `"steps_table": { "headers": [...], "steps": [...] }`

---

## Common Operations

### Metadata + Header

```bash
python3 scripts/docx_meta.py update <output.docx> <output.docx> \
  --project-name "项目名称" \
  --title "文档标题" \
  --author "作者" \
  --department "部门" \
  --revision-date "YYYY-MM-DD"
```

Header auto-generates from template pattern: `某某系统某某功能项目` → `项目名系统项目名项目`

### Version Control

```bash
python3 scripts/docx_meta.py save <output.docx> ~/.openclaw/workspace --name "文件名"
```

Saves to `workspace/doc/`, keeps last 3 versions.

---

## Enhanced Checking

The skill now includes comprehensive document checking:

```bash
# Check only (no modification)
python3 scripts/docx_check.py <filled.docx> --check-only

# Check with structure validation
python3 scripts/docx_check.py <filled.docx> <structure.json>

# Clean punctuation
python3 scripts/docx_check.py <filled.docx> --clean <output.docx>

# Check against template (full format check)
python3 scripts/docx_check.py <filled.docx> <template.docx> --format-check
```

### Check Reports

- **Screenshot placeholders**: Warns if `[截图:xxx]` not filled
- **Unfilled placeholders**: Warns if template placeholders remain
- **Empty sections**: Errors if required sections are empty
- **Table completeness**: Checks tables have headers and rows
- **Template type detection**: Auto-detects feasibility vs change document

### Format Checking (Against Template)

When comparing filled document against template:

```python
from docx_utils import DocumentChecker
from docx import Document

# Load both documents
filled_doc = Document("filled.docx")
template_doc = Document("template.docx")

# Create checker and run comprehensive format check
checker = DocumentChecker(filled_doc)
issues = checker.check_format_comprehensive(template_doc)

for issue in issues:
    print(f"[{issue['severity']}] {issue['type']}: {issue['message']}")
```

#### Header/Footer Checking

- `header_missing`: Template has header but filled doc doesn't
- `footer_missing`: Template has footer but filled doc doesn't
- `first_page_setting_mismatch`: Different first page header/footer setting

#### Title Format Checking

- `missing_title`: Document missing title
- `title_font_name`: Font name mismatch
- `title_font_size`: Font size mismatch  
- `title_font_bold`: Bold setting mismatch
- `title_alignment`: Alignment mismatch

#### Paragraph Style Checking

- `missing_style`: Style from template not found in filled doc

---

## Error Handling

All scripts include proper error handling:

```python
from docx_utils import safe_load_docx

doc, error = safe_load_docx(template_path)
if error:
    print(f"ERROR: {error}")
    sys.exit(1)
```
