# 文件存储结构说明

## 目录结构

自 v0.2.0 起，项目采用独立的上传文件和转换结果目录，方便测试和调试：

```
word2md/
├── uploads/              # 用户上传的原始文件目录
│   ├── {session_id_1}/
│   │   └── a1b2c3d4_document1.docx
│   ├── {session_id_2}/
│   │   └── i9j0k1l2_paper.docx
│   └── ...
│
├── converted/            # 转换生成的文件目录
│   ├── {session_id_1}/
│   │   └── document1/
│   │       ├── document1.md
│   │       └── images/
│   │           ├── img1.png
│   │           └── img2.png
│   ├── {session_id_2}/
│   │   └── paper/
│   │       ├── paper.md
│   │       └── images/
│   └── ...
│
└── .gitignore            # uploads/ 和 converted/ 已添加到 .gitignore
```

## 文件生命周期

| 阶段 | 位置 | 说明 |
|------|------|------|
| 1. 用户上传 | `uploads/{session_id}/` | 原始文件保存在此 |
| 2. 文件转换 | `converted/{session_id}/` | 生成 .md 和图片 |
| 3. 用户下载 | 内存中返回 | 不再保存到磁盘 |
| 4. 定时清理 | 删除 10 分钟前的 | 后台线程自动清理过期文件 |

### 清理策略

- **自动清理**：后台线程每 60 秒检查一次，删除超过 10 分钟（`_RESULT_TTL = 600s`）的文件
- **手动清理**：调用 `POST /cleanup` 立即清理过期文件
- **环境变量控制**：
  - `DOC2MD_UPLOAD_DIR` — 自定义上传文件目录
  - `DOC2MD_CONVERTED_DIR` — 自定义转换文件目录

## 使用示例

### 1. 启动服务，查看配置

```bash
python -m converter.webapp
```

或

```bash
$env:DOC2MD_UPLOAD_DIR = "e:\my\custom\uploads"
$env:DOC2MD_CONVERTED_DIR = "e:\my\custom\converted"
python -m converter.webapp
```

### 2. 检查配置

```bash
curl http://localhost:5000/config
```

返回示例：
```json
{
  "uploads_dir": "E:\\部门\\SDK AI Demo\\word2md\\uploads",
  "converted_dir": "E:\\部门\\SDK AI Demo\\word2md\\converted",
  "result_ttl_seconds": 600,
  "active_results": 2
}
```

### 3. 上传和转换文件

```bash
curl -X POST http://localhost:5000/convert \
  -F "files=@test_document.docx" \
  -F "extract_images=true"
```

上传即时出现在：
- `uploads/{session_id}/` — 原始 Word 文件
- `converted/{session_id}/` — 转换后的 .md 文件和图片

### 4. 手动清理过期文件

```bash
curl -X POST http://localhost:5000/cleanup
```

返回示例：
```json
{
  "status": "ok",
  "cleaned": 2,
  "active_results": 1
}
```

## 测试工作流

### 场景 1：保留文件用于调试

```bash
# 启动服务（使用默认目录）
python -m converter.webapp

# 上传文件
curl -X POST http://localhost:5000/convert \
  -F "files=@my_test_file.docx"

# 立即查看文件
Get-ChildItem e:\部门\SDK AI Demo\word2md\uploads\*
Get-ChildItem e:\部门\SDK AI Demo\word2md\converted\*

# 调试完毕后手动清理
curl -X POST http://localhost:5000/cleanup
```

### 场景 2：多文件批量测试

```bash
# 上传多个文件
curl -X POST http://localhost:5000/convert \
  -F "files=@test1.docx" \
  -F "files=@test2.docx" \
  -F "skip_cover=true" \
  -F "toc_mode=before_toc_keep_abstract"

# 查看转换结果
tree e:\部门\SDK AI Demo\word2md\converted
```

## 文件清理说明

### 自动清理规则

1. **后台线程** — 每 60 秒运行一次
2. **清理条件** — 创建时间 > 10 分钟（可通过 `_RESULT_TTL` 修改）
3. **清理内容**：
   - `uploads/{session_id}/` 目录整体删除
   - `converted/{session_id}/` 目录整体删除
   - 内存中的 `_results[{result_id}]` 条目删除

### 异常清理

如果转换失败或用户中途中止，会立即清理对应的上传目录和转换目录。

## 调试技巧

### 检查当前活跃的转换任务

```bash
curl http://localhost:5000/config | jq .active_results
```

### 查看文件生成在哪个会话ID下

```bash
# 上传后立即这行
Get-ChildItem e:\部门\SDK AI Demo\word2md\uploads
Get-ChildItem e:\部门\SDK AI Demo\word2md\converted
```

### 防止文件自动删除（调试用）

修改 converter/webapp.py 中的 `_RESULT_TTL` 值（单位：秒）：

```python
_RESULT_TTL = 600  # 改为 3600 (1小时) 或更大
```

### 禁用自动清理（仅限开发/调试）

在 `_start_cleanup_thread()` 调用处加注释：

```python
# Start cleanup thread on module load
# _start_cleanup_thread()  # 注释掉此行

# Ensure cleanup on exit
# atexit.register(_stop_cleanup_thread())  # 注释掉此行
```

然后需要手动清理：
```bash
curl -X POST http://localhost:5000/cleanup
```

## 常见问题

### Q: 为什么找不到上传的文件？
**A:** 检查 `/config` 端点查看实际目录位置，可能是：
- 使用了环境变量指定了自定义目录
- 文件已被后台清理线程删除（10 分钟后）
- 转换任务失败，文件被异常清理

### Q: 如何保留某个测试文件用于长期调试？
**A:** 
1. 找到对应的会话 ID 目录
2. 移动到其他位置（如 `test_cases/` 目录）
3. 修改 `_RESULT_TTL` 以防自动删除

### Q: 转换失败的文件会被清理吗？
**A:** 是的，无论成功失败，所有文件都会在 10 分钟后被自动清理。建议定期检查 `/config` 端点的 `active_results` 数量。

### Q: 如何同时处理多个用户的上传？
**A:** 每个转换请求都有唯一的 `session_id`，所以不同用户的文件会保存在不同的会话目录中，不会相互覆盖。

