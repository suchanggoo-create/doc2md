## doc2md (MVP)
 
将 `.docx` 转换为 Markdown，抽取图片与嵌入附件（OLE/embeddings）。

本仓库主要用途是**作为一个可调用工具**，供其他 agent/模型在本地自动执行“docx → md”转换，并拿到稳定的输出目录结构。
 
 ### 安装（本地）
 
 ```bash
python -m pip install -e .
 ```
 
### 推荐调用方式（给其他 agent/模型）

优先使用 **CLI**（最稳、最少集成成本）。如果你在写 Python 流水线，也可以直接调用库 API。

### CLI 调用

#### 最小示例

```bash
doc2md "input.docx" --out "out_dir" --force
```

#### Excel 单独转换（xlsx → md）

当你只需要把 Excel（`.xlsx`）转为 Markdown 时，用 `xlsx2md`：

```bash
xlsx2md "input.xlsx" --out "output.md"
```

说明：
- 默认不截断行数（输出全部行）
- 输出使用 HTML `<table>`，并用 `rowspan/colspan` 保留合并单元格

#### 参数说明
- `--out <dir>`：输出目录（必填）
- `--force`：若输出目录已存在则清空后重建（用于重复运行）
- `--render-embedded true|false`：是否尝试把嵌入的 `.docx/.xlsx` 附件读出来并内嵌到主 `index.md`（默认 `true`）
- `--excel-preview-rows <n>`：渲染 xlsx 的行数
  - `0` 表示**不截断（输出全部行）**（默认 `0`）
  - `n>0` 表示只输出前 `n` 行
- `--max-embed-depth <n>`：嵌入附件递归渲染深度（默认 `2`）
- `--max-embed-bytes <bytes>`：尝试渲染嵌入附件的大小上限（默认 `52428800`，即 50MB）
- `--keep-image-size`：用 HTML `<img>` 保留图片尺寸（默认关闭）
- `--toc`：输出一个简单 TOC（MVP 暂为占位）
- `--log-level error|warn|info|debug`：日志级别（默认 `info`）

#### 输出结构（稳定）

输出目录 `out_dir/` 包含：
- `out_dir/index.md`
- `out_dir/assets/images/`：从 docx 抽取的图片
- `out_dir/assets/attachments/`：从 docx 抽取的附件文件（同时 xlsx 会生成对应 `.md` 产物，便于追溯）

> 注：**嵌入 Excel/Word 附件**默认会尝试“读出内容并以内嵌文本/表格形式放入 `index.md` 原位置”。\n
> 对 Excel，使用 HTML `<table>`（支持合并单元格的 `rowspan/colspan`），以保证内容不丢失。

### Python API 调用（可选）

```python
from doc2md import convertDocxToMd

convertDocxToMd(
    inputPath="input.docx",
    outDir="out_dir",
    options={
        "force": True,
        "render_embedded": True,
        "excel_preview_rows": 0,  # 0 = all rows
        "max_embed_depth": 2,
        "max_embed_bytes": 50 * 1024 * 1024,
        "keep_image_size": False,
        "toc": False,
    },
)
```
 
### 验证与测试

```bash
python -m pip install -e ".[test]"
pytest -q
```
