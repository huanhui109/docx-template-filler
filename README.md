# docx-template-filler

Word 文档模板填写技能 / Word Document Template Filler

自动填充 Word 文档模板，保留所有格式 / Fill Word templates while preserving all formatting.

---

## 如何加载到 OpenClaw / How to Load into OpenClaw

### 方式一：从 GitHub 安装 / Install from GitHub

```bash
# 使用 npx skills add 安装
npx skills add https://github.com/huanhui109/docx-template-filler

# 或者使用 openclaw 命令
openclaw skills add https://github.com/huanhui109/docx-template-filler
```

### 方式二：本地安装 / Local Installation

```bash
# 复制到 OpenClaw skills 目录
cp -r docx-template-filler ~/.openclaw/skills/

# 重启 OpenClaw 服务
openclaw restart
```

---

## 功能概述 / Features

- 支持两种文档类型：**可行性分析报告** 和 **变更上线文档**
- Supports two document types: Feasibility Analysis Reports and Change/Release Documents
- 保留文档原始格式（字体、段落、页眉页脚、编号等）
- Preserve original formatting (fonts, paragraphs, headers, footers, numbering)
- 自动检测模板类型 / Auto-detect template type
- 提供固定模板文件 / Built-in fixed templates
- 支持用户自定义模板 / Support user-defined templates
- 文档完整性检查 / Document completeness checking

---

## 支持的文档类型 / Supported Document Types

### 1. 可行性分析报告 / Feasibility Analysis Report

适用于项目立项可行性分析 / For project feasibility analysis:

- 修订记录 / Revision History
- 文档信息 / Document Information
- 引言 / Introduction
- 可行性分析前提 / Feasibility Analysis Prerequisites
- 项目背景 / Project Background
- 需求分析 / Requirements Analysis
- 技术方案 / Technical Solution
- 项目计划 / Project Plan
- 资源配置 / Resource Allocation
- 效益分析 / Benefit Analysis
- 结论与建议 / Conclusions and Recommendations

### 2. 变更上线文档 / Change/Release Document

适用于系统变更实施方案 / For system change implementation:

- 文档信息（版本、编制、校对、部门）
- Document Information (Version, Author, Reviewer, Department)
- 部署架构图 / Deployment Architecture
- 变更计划 / Change Plan
- 变更目的 / Change Purpose
- 变更内容 / Change Content
- 变更前准备 / Pre-change Preparation
- 变更执行 / Change Execution
- 变更步骤 / Change Steps
- 升级完成后检查 / Post-upgrade Verification
- 变更回退 / Rollback Plan
- 参数设置 / Parameter Settings
- 配置更新 / Configuration Updates
- 软件及文档更新 / Software and Documentation Updates
- 网络/硬件配置 / Network/Hardware Configuration
- 变更后工作 / Post-change Work
- 变更后检查 / Post-change Verification

---

## 环境要求 / Prerequisites

```bash
pip3 install python-docx pdfplumber python-pptx
```

---

## 快速开始 / Quick Start

### 使用方式 / Usage

#### 方式一：使用内置固定模板 / Use Built-in Templates

```python
from docx_utils import resolve_template, get_template_path

# 获取固定模板路径 / Get fixed template path
template_path = get_template_path("change")  # 变更模板 / Change template
# 或 / or
template_path = get_template_path("feasibility")  # 可行性分析模板 / Feasibility template
```

#### 方式二：智能模板解析（推荐）/ Smart Template Resolution (Recommended)

```python
from docx_utils import resolve_template

# 自动解析模板（用户模板优先，兜底内置模板）
# Auto-resolve template (user template takes priority, fallback to built-in)
result = resolve_template(
    user_template_path="/path/to/user_template.docx",  # 用户提供的模板
    template_type="feasibility",                        # 指定的类型
    fallback_to_builtin=True                           # 兜底使用内置模板
)

if result["error"]:
    print(f"错误 / Error: {result['error']}")
else:
    template_path = result["template_path"]
    template_type = result["template_type"]
```

#### 方式三：交互式选择 / Interactive Selection

```python
from docx_utils import prompt_template_type_choice

# 显示交互提示让用户选择 / Show interactive prompt
print(prompt_template_type_choice())
```

### 命令行脚本 / Command Line Scripts

```bash
# 1. 解析模板结构 / Parse template structure
python3 scripts/docx_parse.py <template.docx> --output /tmp/struct.json

# 2. 填写内容 / Fill content
python3 scripts/docx_fill.py <template.docx> content.json <output.docx>

# 3. 检查完整性 / Check completeness
python3 scripts/docx_check.py <output.docx> --check-only

# 4. 清理标点符号 / Clean punctuation
python3 scripts/docx_check.py <output.docx> --clean <output_clean.docx>

# 5. 首行缩进 / First-line indent
python3 scripts/docx_meta.py indent <output.docx> <output.docx> --chars 2

# 6. 更新目录 / Update TOC
python3 scripts/docx_meta.py toc <output.docx> <output.docx>

# 7. 更新元数据和页眉 / Update metadata and header
python3 scripts/docx_meta.py update <output.docx> <output.docx> \
  --project-name "项目名称 / Project Name" \
  --author "作者 / Author" \
  --department "部门 / Department" \
  --revision-date "2026-01-01"

# 8. 版本保存 / Version save
python3 scripts/docx_meta.py save <output.docx> <workspace> --name "文件名 / Filename"
```

---

## 文档检查 / Document Checking

### 基础检查 / Basic Checking

```python
from docx_utils import DocumentChecker
from docx import Document

doc = Document("filled.docx")
checker = DocumentChecker(doc)
issues = checker.check_all()

for issue in issues:
    print(f"[{issue['severity']}] {issue['type']}: {issue['message']}")
```

### 格式检查（与模板对比）/ Format Checking (Compare with Template)

```python
from docx_utils import DocumentChecker
from docx import Document

filled_doc = Document("filled.docx")
template_doc = Document("template.docx")

checker = DocumentChecker(filled_doc)
issues = checker.check_format_comprehensive(template_doc)

for issue in issues:
    print(f"[{issue['severity']}] {issue['type']}: {issue['message']}")
```

### 检查项 / Check Items

| 类型 / Type | 检查项 / Check Item | 严重程度 / Severity |
|-------------|---------------------|---------------------|
| 内容 / Content | 截图占位符未填写 / Screenshot placeholder not filled | warning |
| 内容 / Content | 未填写占位符 / Unfilled placeholder | warning |
| 内容 / Content | 空表格 / Empty table | error |
| 格式 / Format | 缺少页眉 / Missing header | error/warning |
| 格式 / Format | 缺少页脚 / Missing footer | error/warning |
| 格式 / Format | 标题字体不匹配 / Title font mismatch | warning |
| 格式 / Format | 标题字号不匹配 / Title font size mismatch | warning |
| 格式 / Format | 标题粗体设置不匹配 / Title bold mismatch | warning |
| 格式 / Format | 对齐方式不匹配 / Alignment mismatch | warning |

---

## 内容 JSON 格式 / Content JSON Format

### 可行性分析报告 / Feasibility Analysis Report

```json
{
  "template_type": "feasibility",
  "project_name": "项目名称 / Project Name",
  "metadata": {
    "version": "V1.0",
    "author": "项目经理 / Project Manager",
    "department": "信息技术中心 / IT Center"
  },
  "sections": {
    "引言 / Introduction": {
      "content": ["内容 / Content..."]
    },
    "项目背景 / Project Background": {
      "content": ["背景描述 / Background description..."]
    },
    "需求分析 / Requirements Analysis": {
      "content": ["需求说明 / Requirements description..."]
    },
    "技术方案 / Technical Solution": {
      "content": ["技术实现方案 / Technical implementation..."]
    }
  }
}
```

### 变更上线文档 / Change/Release Document

```json
{
  "template_type": "change",
  "project_name": "项目名称 / Project Name",
  "metadata": {
    "version": "V1.0",
    "author": "张三 / Zhang San",
    "department": "信息技术中心 / IT Center"
  },
  "sections": {
    "变更目的 / Change Purpose": {
      "content": ["变更目的说明 / Change purpose description..."]
    },
    "变更内容 / Change Content": {
      "content": ["本次变更涉及以下内容 / This change involves:", "1、...", "2、..."]
    },
    "变更前准备 / Pre-change Preparation": {
      "content": ["【环境准备 / Environment Prep】", "1、确认系统版本 / Confirm system version..."]
    },
    "变更步骤 / Change Steps": {
      "content": ["步骤1：环境检查 / Step 1: Environment Check", "操作 / Operation: ...", "预期 / Expected: ..."]
    },
    "变更回退 / Rollback": {
      "content": ["如迁移失败，执行以下回退步骤 / If migration fails, perform rollback:", "1、..."]
    },
    "配置更新 / Config Updates": {
      "table": {
        "headers": ["项目 / Item", "说明 / Description"],
        "rows": [["配置项1 / Config 1", "说明1 / Desc 1"], ["配置项2 / Config 2", "说明2 / Desc 2"]]
      }
    }
  }
}
```

---

## 目录结构 / Directory Structure

```
docx-template-filler/
├── scripts/
│   ├── docx_utils.py          # 共享工具模块 / Shared utilities
│   ├── docx_parse.py          # 解析模板结构 / Parse template structure
│   ├── docx_fill.py           # 填写内容（支持字体和编号）/ Fill content (font + numbering)
│   ├── docx_gen_change.py     # 生成变更文档 / Generate change document
│   ├── docx_check.py          # 检查完整性 / Check completeness
│   └── docx_meta.py           # 元数据、页眉、缩进、目录、版本 / Metadata, header, indent, TOC, version
├── templates/                  # 固定模板（用户模板不上传）/ Fixed templates (not uploaded)
│   ├── change_template.docx
│   └── feasibility_template.docx
├── SKILL.md                   # OpenClaw 技能定义 / OpenClaw skill definition
└── README.md                   # 本文件 / This file
```

---

## 注意事项 / Notes

1. **模板文件 / Template Files**: `templates/` 目录下的模板文件包含敏感信息，请勿上传到公共仓库 / Templates contain sensitive info, do not upload to public repo
2. **页眉页脚 / Headers/Footers**: 自动保留模板中的页眉页脚格式 / Automatically preserve header/footer
3. **字体格式 / Font Format**: 自动匹配模板中的字体格式 / Automatically match font format
4. **编号规则 / Numbering**: 自动保留模板中的列表编号格式 / Automatically preserve numbering

---

## 依赖 / Dependencies

- python-docx >= 0.8
- pdfplumber (用于从 PDF 提取内容 / For PDF extraction)
- python-pptx (用于从 PPT 提取内容 / For PPT extraction)

---

## 版本 / Version

v1.0.0 - 初始版本 / Initial Release
