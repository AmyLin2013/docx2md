[English](README.md) | [中文](README_zh.md)

# doc2md — Pure Python Word-to-Markdown Converter, No Complex Dependencies ⚡

Zero third-party conversion dependencies. Directly parses `.docx` XML to produce well-structured Markdown. Offers both CLI and Web interfaces, ready to use out of the box.

## Features

| Feature | Word (.docx) |
|---------|:-----------:|
| Heading detection | ✅ Based on outlineLvl attribute |
| Auto numbering | ✅ Based on numPr attribute |
| Ordered/unordered lists | ✅ |
| Bold/Italic | ✅ |
| Hyperlinks | ✅ |
| Image extraction | ✅ Saved to images/ |
| Tables | ✅ Converted to Markdown tables |
| Smart paragraph merging | ✅ |
| Cover page removal | ✅ |
| TOC handling | ✅ 4 modes |
| Abstract preservation | ✅ For academic papers |
| Batch conversion | ✅ |
| Markdown preview | ✅ Rendered + source |
| Smart download | ✅ .md or .zip |

## Installation

```bash
# Recommended: use a virtual environment
python -m venv .venv
.venv\Scripts\activate      # Windows
# source .venv/bin/activate  # Linux/Mac

# Install dependencies
pip install -r requirements.txt

# Optional: install as a CLI tool
pip install -e .
```

## Dependencies

| Library | Purpose |
|---------|---------|
| [Flask](https://flask.palletsprojects.com/) | Web framework |

Word conversion uses only Python standard libraries (`xml.etree.ElementTree` + `zipfile`) to parse .docx XML directly — no extra dependencies needed.

## Usage

### Command Line

```bash
# Single file conversion
doc2md document.docx               # → document.md

# Specify output path
doc2md document.docx -o output.md

# Batch conversion
doc2md *.docx

# Print to stdout (no file saved)
doc2md document.docx --stdout

# Skip image extraction
doc2md document.docx --no-images

# Remove cover page
doc2md document.docx --skip-cover

# TOC handling options
doc2md document.docx --toc-mode toc_only              # Remove TOC only
doc2md document.docx --toc-mode before_toc             # Remove TOC and everything before it
doc2md paper.docx --toc-mode before_toc_keep_abstract   # Remove TOC and before, keep abstract

# Combined usage
doc2md paper.docx --skip-cover --toc-mode before_toc_keep_abstract --no-images
```

### Python API

```python
from converter.word2md import convert_word_to_markdown

# Word → Markdown
md = convert_word_to_markdown("input.docx", "output.md")

# Remove cover, remove TOC but keep abstract (academic paper scenario)
md = convert_word_to_markdown(
    "paper.docx", "paper.md",
    skip_cover=True,
    toc_mode="before_toc_keep_abstract",
)

# Remove TOC page only
md = convert_word_to_markdown("doc.docx", toc_mode="toc_only")

# Get string only, no file saved
md = convert_word_to_markdown("input.docx")
print(md)
```

## Web Service

### Starting the Server

```bash
# Development mode (uses default uploads/ and converted/ directories)
python -m converter.webapp

# Or use Flask CLI
flask --app converter.webapp run --port 5000

# Custom temporary file directories
set DOC2MD_UPLOAD_DIR=d:\my_uploads      # Custom upload directory
set DOC2MD_CONVERTED_DIR=d:\my_converted  # Custom converted file directory
python -m converter.webapp
```

Visit http://localhost:5000 after starting the server to open the Web interface.

### Web Interface Features

- **Drag & drop / click to upload** .docx files (supports multi-file batch upload)
- Conversion options displayed based on file format:
  - **Word**: Image extraction, cover page removal, TOC handling mode (4 modes)
- **Markdown preview**: Preview directly in the browser after conversion, with rendered view and source view toggle
- **One-click copy**: Copy Markdown source to clipboard
- **Smart download**:
  - Single file without images → direct `.md` download
  - With images or multiple files → bundled as `.zip` download

### API Endpoints

```bash
# POST /convert - Upload and convert files, returns JSON (preview content + download ID)
curl -X POST http://localhost:5000/convert \
  -F "files=@document.docx" \
  -F "extract_images=true" \
  -F "skip_cover=false" \
  -F "toc_mode=none"
# Returns: { "id": "xxx", "files": [{"name": "document.md", "content": "..."}], "needs_zip": false }

# GET /download/<id> - Download conversion result (auto-returns .md or .zip)
curl http://localhost:5000/download/xxx -o result.md

# POST /convert - Academic paper scenario
curl -X POST http://localhost:5000/convert \
  -F "files=@paper.docx" \
  -F "skip_cover=true" \
  -F "toc_mode=before_toc_keep_abstract"

# GET /config - View current configuration (file storage paths, active tasks)
curl http://localhost:5000/config
# Returns: { "uploads_dir": "...", "converted_dir": "...", "result_ttl_seconds": 600, "active_results": 2 }

# POST /cleanup - Manually clean up expired conversion results and files
curl -X POST http://localhost:5000/cleanup
# Returns: { "status": "ok", "cleaned": 2, "active_results": 1 }

# GET /health - Health check
curl http://localhost:5000/health
```

#### Word Parameters

| Parameter | Default | Description |
|-----------|---------|-------------|
| `extract_images` | `true` | Whether to extract embedded images |
| `skip_cover` | `false` | Whether to remove the first cover page |
| `toc_mode` | `none` | TOC handling mode: `none` / `toc_only` / `before_toc` / `before_toc_keep_abstract` |

| toc_mode Value | Description |
|----------------|-------------|
| `none` | Keep all content |
| `toc_only` | Remove TOC page only |
| `before_toc` | Remove TOC and everything before it |
| `before_toc_keep_abstract` | Remove TOC and everything before it, but keep abstract |

## File Storage & Cleanup

When the Web service is running, uploaded and converted files are automatically saved under the project directory:

- **uploads/{session_id}/** — Original .docx files uploaded by users
- **converted/{session_id}/{filename}/** — Converted Markdown files and extracted images

### Automatic Cleanup

A background thread checks every 60 seconds and automatically cleans up files **older than 10 minutes** (configurable via `_RESULT_TTL`).

### Manual Cleanup

```bash
curl -X POST http://localhost:5000/cleanup
```

### Custom Storage Paths

You can customize file storage locations via environment variables. See [UPLOAD_STORAGE_GUIDE.md](UPLOAD_STORAGE_GUIDE.md) for details.

## Project Structure

```
doc2md/
├── pyproject.toml              # Project config & dependencies
├── requirements.txt            # pip dependencies
├── README.md                   # English documentation
├── README_zh.md                # Chinese documentation
├── UPLOAD_STORAGE_GUIDE.md     # File storage and cleanup guide
├── .gitignore
├── templates/
│   └── index.html              # Web frontend (preview + download)
└── converter/
    ├── __init__.py
    ├── cli.py                  # CLI entry point (argparse)
    ├── webapp.py               # Flask Web service
    ├── word2md.py              # Word converter (direct .docx XML parsing)
    └── numbering.py            # Word numbering/style/outline level parser
```

> The `uploads/` and `converted/` directories are created automatically at runtime for temporary file storage and are excluded via `.gitignore`.

## Technical Details

### Word (.docx) Conversion Pipeline

```
.docx (ZIP)  ──unzip──▶  XML (document.xml, styles.xml, numbering.xml)
                              │
                              ├─ outlineLvl → Heading level (H1–H6)
                              ├─ numPr → Auto numbering ("Chapter 1", "1.1")
                              ├─ rPr → Bold/Italic/Strikethrough
                              ├─ hyperlink + rels → Hyperlinks
                              ├─ drawing + media → Images
                              └─ tbl → Tables
                              │
                              ▼
                          Markdown
```

- Directly parses .docx XML structure without relying on mammoth or markdownify
- Heading levels are derived from each paragraph's `outlineLvl` attribute (supports style inheritance via `basedOn`)
- Auto numbering is derived from `numPr` (numId + ilvl), parsed through numbering.xml
- Supports Chinese numbering ("第一章", "一、"), Roman numerals, letters, and more
- Cover/TOC/abstract detection based on XML attributes (page breaks, TOC styles/SDT/field codes, heading keywords)

## License

MIT
