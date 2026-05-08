# docx-template-filler

Word 文档模板填写技能 - 自动填充 Word 文档模板，保留所有格式。

## 功能概述

- 支持两种文档类型：**可行性分析报告** 和 **变更上线文档**
- 保留文档原始格式（字体、段落、页眉页脚、编号等）
- 自动检测模板类型
- 提供固定模板文件
- 支持用户自定义模板
- 文档完整性检查

## 支持的文档类型

### 1. 可行性分析报告

适用于项目立项可行性分析，包含：
- 修订记录
- 文档信息
- 引言
- 可行性分析前提
- 项目背景
- 需求分析
- 技术方案
- 项目计划
- 资源配置
- 效益分析
- 结论与建议

### 2. 变更上线文档

适用于系统变更实施方案，包含：
- 文档信息（版本、编制、校对、部门）
- 部署架构图
- 变更计划
- 变更目的
- 变更内容
- 变更前准备
- 变更执行
- 变更步骤
- 升级完成后检查
- 变更回退
- 参数设置
- 配置更新
- 软件及文档更新
- 网络/硬件配置
- 变更后工作
- 变更后检查

## 快速开始

### 环境要求

```bash
pip3 install python-docx pdfplumber python-pptx
```

### 使用方式

#### 方式一：使用内置固定模板

```python
from docx_utils import resolve_template, get_template_path

# 获取固定模板路径
template_path = get_template_path("change")  # 变更模板
# 或
template_path = get_template_path("feasibility")  # 可行性分析模板
```

#### 方式二：智能模板解析（推荐）

```python
from docx_utils import resolve_template

# 自动解析模板（用户模板优先，兜底内置模板）
result = resolve_template(
    user_template_path="/path/to/user_template.docx",  # 用户提供的模板
    template_type="feasibility",                        # 指定的类型
    fallback_to_builtin=True                           # 兜底使用内置模板
)

if result["error"]:
    print(f"错误: {result['error']}")
else:
    template_path = result["template_path"]
    template_type = result["template_type"]
```

#### 方式三：交互式选择

```python
from docx_utils import prompt_template_type_choice

# 显示交互提示让用户选择
print(prompt_template_type_choice())
```

### 命令行脚本

```bash
# 1. 解析模板结构
python3 scripts/docx_parse.py <template.docx> --output /tmp/struct.json

# 2. 填写内容
python3 scripts/docx_fill.py <template.docx> content.json <output.docx>

# 3. 检查完整性
python3 scripts/docx_check.py <output.docx> --check-only

# 4. 清理标点符号
python3 scripts/docx_check.py <output.docx> --clean <output_clean.docx>

# 5. 首行缩进
python3 scripts/docx_meta.py indent <output.docx> <output.docx> --chars 2

# 6. 更新目录
python3 scripts/docx_meta.py toc <output.docx> <output.docx>

# 7. 更新元数据和页眉
python3 scripts/docx_meta.py update <output.docx> <output.docx> \
  --project-name "项目名称" --author "作者" --department "部门" \
  --revision-date "2026-01-01"

# 8. 版本保存
python3 scripts/docx_meta.py save <output.docx> <workspace> --name "文件名"
```

## 文档检查

### 基础检查

```python
from docx_utils import DocumentChecker
from docx import Document

doc = Document("filled.docx")
checker = DocumentChecker(doc)
issues = checker.check_all()

for issue in issues:
    print(f"[{issue['severity']}] {issue['type']}: {issue['message']}")
```

### 格式检查（与模板对比）

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

### 检查项

| 类型 | 检查项 | 严重程度 |
|------|--------|----------|
| 内容 | 截图占位符未填写 | warning |
| 内容 | 未填写占位符 | warning |
| 内容 | 空表格 | error |
| 格式 | 缺少页眉 | error/warning |
| 格式 | 缺少页脚 | error/warning |
| 格式 | 标题字体不匹配 | warning |
| 格式 | 标题字号不匹配 | warning |
| 格式 | 标题粗体设置不匹配 | warning |
| 格式 | 对齐方式不匹配 | warning |

## 内容 JSON 格式

### 可行性分析报告

```json
{
  "template_type": "feasibility",
  "project_name": "某某系统某某功能项目",
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
      "content": ["背景描述..."]
    },
    "需求分析": {
      "content": ["需求说明..."]
    },
    "技术方案": {
      "content": ["技术实现方案..."]
    }
  }
}
```

### 变更上线文档

```json
{
  "template_type": "change",
  "project_name": "FISP代销迁移直连",
  "metadata": {
    "version": "V1.0",
    "author": "张三",
    "department": "信息技术中心"
  },
  "sections": {
    "变更目的": {
      "content": ["变更目的说明..."]
    },
    "变更内容": {
      "content": ["本次变更涉及以下内容：", "1、...", "2、..."]
    },
    "变更前准备": {
      "content": ["【环境准备】", "1、确认系统版本..."]
    },
    "变更步骤": {
      "content": ["步骤1：环境检查", "操作：...", "预期：..."]
    },
    "变更回退": {
      "content": ["如迁移失败，执行以下回退步骤：", "1、..."]
    },
    "配置更新": {
      "table": {
        "headers": ["项目", "说明"],
        "rows": [["配置项1", "说明1"], ["配置项2", "说明2"]]
      }
    }
  }
}
```

## 目录结构

```
docx-template-filler/
├── scripts/
│   ├── docx_utils.py          # 共享工具模块
│   ├── docx_parse.py          # 解析模板结构
│   ├── docx_fill.py           # 填写内容（支持字体和编号）
│   ├── docx_gen_change.py     # 生成变更文档
│   ├── docx_check.py          # 检查完整性
│   └── docx_meta.py           # 元数据、页眉、缩进、目录、版本
├── templates/                  # 固定模板（用户模板不上传）
│   ├── change_template.docx
│   └── feasibility_template.docx
├── SKILL.md                   # OpenClaw 技能定义
└── README.md                   # 本文件
```

## 注意事项

1. **模板文件**：`templates/` 目录下的模板文件包含敏感的公司模板，请勿上传到公共仓库
2. **页眉页脚**：自动保留模板中的页眉页脚格式
3. **字体格式**：自动匹配模板中的字体格式
4. **编号规则**：自动保留模板中的列表编号格式

## 依赖

- python-docx >= 0.8
- pdfplumber (用于从 PDF 提取内容)
- python-pptx (用于从 PPT 提取内容)
