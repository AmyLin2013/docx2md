[English](README.md) | [中文](README_zh.md)

# doc2md — 告别复杂依赖，用纯 Python 将 Word 精准转为 Markdown ⚡

零第三方转换依赖，直接解析 `.docx` XML，将 Word 文档精准转为结构化 Markdown。提供命令行与 Web 双界面，开箱即用。

## 功能特性

| 功能 | Word (.docx) |
|------|:-----------:|
| 标题识别 | ✅ 基于 outlineLvl 属性 |
| 自动编号 | ✅ 基于 numPr 属性 |
| 有序/无序列表 | ✅ |
| 加粗/斜体 | ✅ |
| 超链接 | ✅ |
| 图片提取 | ✅ 保存到 images/ |
| 表格 | ✅ 转为 Markdown 表格 |
| 段落智能合并 | ✅ |
| 去除封面页 | ✅ |
| 目录处理 | ✅ 4 种模式 |
| 摘要保留 | ✅ 论文场景 |
| 批量转换 | ✅ |
| Markdown 预览 | ✅ 渲染 + 源码 |
| 智能下载 | ✅ .md 或 .zip |

## 安装

```bash
# 推荐使用虚拟环境
python -m venv .venv
.venv\Scripts\activate      # Windows
# source .venv/bin/activate  # Linux/Mac

# 安装依赖
pip install -r requirements.txt

# 可选: 安装为命令行工具
pip install -e .
```

## 依赖说明

| 库 | 用途 |
|----|------|
| [Flask](https://flask.palletsprojects.com/) | Web 服务框架 |

Word 转换仅使用 Python 标准库（`xml.etree.ElementTree` + `zipfile`），直接解析 .docx XML，无额外依赖。

## 使用方法

### 命令行

```bash
# 单文件转换
doc2md document.docx               # → document.md

# 指定输出路径
doc2md document.docx -o output.md

# 批量转换
doc2md *.docx

# 输出到终端（不保存文件）
doc2md document.docx --stdout

# 不提取图片
doc2md document.docx --no-images

# 去掉封面页
doc2md document.docx --skip-cover

# 目录处理选项
doc2md document.docx --toc-mode toc_only              # 只去掉目录
doc2md document.docx --toc-mode before_toc             # 去掉目录及之前内容
doc2md paper.docx --toc-mode before_toc_keep_abstract   # 去掉目录及之前内容，保留摘要

# 组合使用
doc2md paper.docx --skip-cover --toc-mode before_toc_keep_abstract --no-images
```

### Python API

```python
from converter.word2md import convert_word_to_markdown

# Word → Markdown
md = convert_word_to_markdown("input.docx", "output.md")

# 去掉封面、去掉目录但保留摘要（论文场景）
md = convert_word_to_markdown(
    "paper.docx", "paper.md",
    skip_cover=True,
    toc_mode="before_toc_keep_abstract",
)

# 只移除目录页
md = convert_word_to_markdown("doc.docx", toc_mode="toc_only")

# 仅获取字符串，不保存文件
md = convert_word_to_markdown("input.docx")
print(md)
```

## Web 服务

### 启动服务

```bash
# 开发模式（使用默认的 uploads/ 和 converted/ 目录）
python -m converter.webapp

# 或使用 Flask CLI
flask --app converter.webapp run --port 5000

# 自定义临时文件目录
set DOC2MD_UPLOAD_DIR=d:\my_uploads      # 自定义上传文件目录
set DOC2MD_CONVERTED_DIR=d:\my_converted  # 自定义转换文件目录
python -m converter.webapp
```

启动后访问 http://localhost:5000 即可打开 Web 界面。

### Web 界面功能

- **拖拽/点击上传** .docx 文件（支持多文件批量）
- 根据文件格式自动显示对应的转换选项：
  - **Word**: 是否提取图片、去除封面页、目录处理方式（4 种模式）
- **Markdown 预览**：转换完成后直接在页面内预览，支持渲染视图和源码视图切换
- **一键复制**：复制 Markdown 源码到剪贴板
- **智能下载**：
  - 单文件且无图片 → 直接下载 `.md` 文件
  - 含图片或多文件 → 打包为 `.zip` 下载

### API 接口

```bash
# POST /convert - 上传文件并转换，返回 JSON（含预览内容 + 下载 ID）
curl -X POST http://localhost:5000/convert \
  -F "files=@document.docx" \
  -F "extract_images=true" \
  -F "skip_cover=false" \
  -F "toc_mode=none"
# 返回: { "id": "xxx", "files": [{"name": "document.md", "content": "..."}], "needs_zip": false }

# GET /download/<id> - 下载转换结果（自动返回 .md 或 .zip）
curl http://localhost:5000/download/xxx -o result.md

# POST /convert - Word 论文场景
curl -X POST http://localhost:5000/convert \
  -F "files=@paper.docx" \
  -F "skip_cover=true" \
  -F "toc_mode=before_toc_keep_abstract"

# GET /config - 查看当前配置（文件存储位置、活跃任务数）
curl http://localhost:5000/config
# 返回: { "uploads_dir": "...", "converted_dir": "...", "result_ttl_seconds": 600, "active_results": 2 }

# POST /cleanup - 手动清理过期的转换结果和文件
curl -X POST http://localhost:5000/cleanup
# 返回: { "status": "ok", "cleaned": 2, "active_results": 1 }

# GET /health - 健康检查
curl http://localhost:5000/health
```

#### Word 参数说明

| 参数 | 默认值 | 说明 |
|------|-------|------|
| `extract_images` | `true` | 是否提取嵌入图片 |
| `skip_cover` | `false` | 是否去除第一页封面 |
| `toc_mode` | `none` | 目录处理模式：`none` / `toc_only` / `before_toc` / `before_toc_keep_abstract` |

| toc_mode 值 | 说明 |
|------------|------|
| `none` | 保留所有内容 |
| `toc_only` | 只移除目录页 |
| `before_toc` | 移除目录及其之前的内容 |
| `before_toc_keep_abstract` | 移除目录及之前内容，但保留中英文摘要 |

## 文件存储与清理

Web 服务运行时，上传和转换的文件自动保存在项目目录下：

- **uploads/{session_id}/** — 用户上传的原始 .docx 文件
- **converted/{session_id}/{filename}/** — 转换生成的 Markdown 和提取的图片

### 自动清理

后台线程每 60 秒检查一次，自动清理 **超过 10 分钟** 的文件（可通过 `_RESULT_TTL` 配置）。

### 手动清理

```bash
curl -X POST http://localhost:5000/cleanup
```

### 自定义存储路径

通过环境变量可自定义文件存储位置，详见 [UPLOAD_STORAGE_GUIDE.md](UPLOAD_STORAGE_GUIDE.md)。

## 项目结构

```
doc2md/
├── pyproject.toml              # 项目配置 & 依赖
├── requirements.txt            # pip 依赖
├── README.md
├── README_zh.md                # 中文文档
├── UPLOAD_STORAGE_GUIDE.md     # 文件存储和清理策略说明
├── .gitignore
├── templates/
│   └── index.html              # Web 前端页面（预览 + 下载）
└── converter/
    ├── __init__.py
    ├── cli.py                  # CLI 入口 (argparse)
    ├── webapp.py               # Flask Web 服务
    ├── word2md.py              # Word 转换（直接解析 .docx XML）
    └── numbering.py            # Word 编号/样式/大纲级别解析
```

> 运行时会自动创建 `uploads/` 和 `converted/` 目录用于临时文件存储，已通过 `.gitignore` 排除。

## 技术方案说明

### Word (.docx) 转换流程

```
.docx (ZIP)  ──解压──▶  XML (document.xml, styles.xml, numbering.xml)
                              │
                              ├─ outlineLvl → 标题级别 (H1–H6)
                              ├─ numPr → 自动编号 ("第一章"、"1.1")
                              ├─ rPr → 加粗/斜体/删除线
                              ├─ hyperlink + rels → 超链接
                              ├─ drawing + media → 图片
                              └─ tbl → 表格
                              │
                              ▼
                          Markdown
```

- 直接解析 .docx 的 XML 结构，不依赖 mammoth 或 markdownify
- 标题级别来源于每个段落的 `outlineLvl` 属性（支持样式继承链 `basedOn`）
- 自动编号来源于 `numPr`（numId + ilvl），通过 numbering.xml 解析编号格式
- 支持中文编号（"第一章"、"一、"）、罗马数字、字母等格式
- 封面/目录/摘要检测基于 XML 属性（分页符、TOC 样式/SDT/域代码、标题关键词）

## License

MIT
