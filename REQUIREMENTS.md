## 文档目的
定义“DOCX 转 Markdown（含图片与附件）工具”的需求规格、范围边界、输入输出、功能点、非功能约束与验收标准，作为开发与验收依据。

## 技术与交付形态
- **实现语言**：Python 3.11+
- **交付形态**：
  - CLI 可执行命令 `doc2md`
  - （可选）可被其他项目调用的库 API：`convertDocxToMd(inputPath, outDir, options)`

## 输入与输出
### 输入
- 单个 `.docx` 文件路径（第一版支持单文件；批量/目录作为后续增强）。

### 输出（固定结构）
输出为一个目录（`--out <out_dir>`），包含：
- `index.md`
- `assets/images/`：抽取的图片资源
- `assets/attachments/`：抽取的嵌入附件资源（OLE/嵌入文件）

Markdown 内引用一律使用**相对路径**（如 `assets/images/img_0001.png`）。

## 总体目标（核心保证）
- **内容顺序保持**：最大化保持 Word 中的**线性阅读顺序**，覆盖文字/图片/表格/附件等元素。
- **不丢元素**：遇到无法解析/无法渲染的元素，必须在 `index.md` 原位置插入**占位块**，并输出日志说明原因。
- **附件处理（尽力渲染 + 保底降级）**：
  - 必须抽取附件文件本体并在原位置生成链接
  - 对可识别的 Word/Excel 附件尽力解析为 Markdown；失败则降级为“链接 + 失败原因占位”

## 功能需求（FR）
### FR1 内容顺序与遍历粒度（必须）
- 以 DOCX 的 WordprocessingML 主文档流为准：`word/document.xml` 的 `w:body` 下块级节点顺序。
- 主要块级类型：
  - 段落：`w:p`
  - 表格：`w:tbl`
- 图片/附件通常位于段落内的 run（`w:r`）中，必须在其出现位置输出。
- 任一节点渲染失败不得导致整体失败：必须插入占位并继续。

### FR2 文本与结构（必须）
- **标题**：按样式（`Heading1..n`）映射为 `#..######`。
- **段落**：按段落输出，段落间空行分隔。
- **列表**：有序/无序列表，尽量保留嵌套层级；无法可靠还原时允许降级为平铺并保留缩进。
- **超链接**：若 DOCX 中存在超链接关系，需输出为 `[text](url)` 并保留其在行内的位置。

### FR3 行内样式（必须/尽力）
- **粗体/斜体**：尽力输出为 `**text**` / `*text*`。
- **下划线**：Markdown 无原生语法，第一版可忽略或降级为 HTML（如 `<u>text</u>`），但不得破坏文本内容。
- **删除线（必须）**：
  - **识别**：从 `w:rPr` 的 `w:strike` / `w:dstrike` 判断该 run 是否为删除线。
  - **渲染**：输出为 `~~text~~`。
  - **与其他样式叠加**：允许组合，需全局一致（例如统一使用 `~~**text**~~`）。
  - **链接内删除线**：优先输出 `~~[text](url)~~`；若实现复杂导致断裂，允许降级为 `~~text~~ ([link](url))`，但不得丢失“删除线语义”和“链接语义”。
  - **跨 run 合并**：相邻删除线 run 尽量合并为一个 `~~...~~`；无法安全合并时可逐 run 包裹。

### FR4 表格（必须）
- 默认渲染为 Markdown 表格。
- 若存在合并单元格（rowspan/colspan）等 Markdown 不支持的结构，必须降级为 HTML `<table>`（仍在原位置）。

### FR5 图片（必须）
- 抽取 `word/media/*` 到 `assets/images/`。
- 在 `index.md` 原位置插入引用：`![](assets/images/img_0001.png)`。
- 图片尺寸：
  - 默认不保留尺寸（避免污染 Markdown）。
  - 若启用 `--keep-image-size`，允许用 HTML `<img ... width height>` 表达。

### FR6 附件/嵌入对象（OLE）（必须 + 尽力渲染）
#### 抽取（必须）
- 从段落 run 中定位 OLE（常见：`w:r/w:object`、`o:OLEObject` 等），解析其关系 id（如 `r:id` / `r:embed`）。
- 通过 `word/_rels/document.xml.rels` 解析关系目标，抽取对应二进制（常见位于 `word/embeddings/*`）。
- 将附件写入 `assets/attachments/`，并在 `index.md` 原位置插入链接：
  - `[附件: <name>](assets/attachments/<file>)`
- 抽取失败时必须占位：
  - `> [ATTACHMENT_EXTRACT_FAILED: <reason>]`

#### 尽力渲染（你要求启用，失败可降级）
- 若附件可识别为 `.docx`：
  - 生成 `assets/attachments/<name>/index.md`（及其 `assets/`）
  - 主文档在原位置插入链接：`[查看附件内容: <name>](assets/attachments/<name>/index.md)`
- 若附件可识别为 `.xlsx`：
  - 生成 `assets/attachments/<name>.md`（按 sheet 分段输出 Markdown 表格）
  - 主文档在原位置插入链接，并可选插入前 N 行预览（由 `--excel-preview-rows` 控制）
- 若渲染失败或不可识别：
  - 保留抽取文件链接
  - 插入占位说明（包含失败原因摘要），但不得中断整体转换

### FR7 不支持元素的占位（必须）
无法渲染的元素在原位置输出占位块（示例格式）：
- `> [UNSUPPORTED: <type>]`
- `> [EMBED_RENDER_FAILED: <reason>]`

## CLI 需求（第一版建议规格）
- `doc2md <input.docx> --out <out_dir>`
- 关键参数与默认值建议：
  - `--out <dir>`：输出目录
  - `--force`：允许覆盖输出目录（覆盖策略见下）
  - `--render-embedded`：启用附件尽力渲染（默认 `true`）
  - `--max-embed-depth <n>`：递归渲染深度（默认 `2`）
  - `--max-embed-bytes <bytes>`：渲染尝试的附件大小上限（默认 `50MB`，超限只抽取不渲染）
  - `--excel-preview-rows <n>`：Excel 预览行数（默认 `20`，0 表示不预览仅链接）
  - `--keep-image-size`：保留图片尺寸（默认 `false`）
  - `--toc`：生成 TOC（默认 `false`）
  - `--log-level <level>`：`error|warn|info|debug`（默认 `info`）

### 覆盖策略（建议）
- 默认：若 `out_dir` 已存在则报错并退出（避免覆盖用户文件）。
- `--force`：允许覆盖，建议策略为“清空目录后再写入”（可预测、不留脏文件）。

## 资源命名规则（稳定、可预测）
### 图片命名
- 路径：`assets/images/img_<seq>.<ext>`
- `<seq>`：按出现顺序从 `0001` 开始
- `<ext>`：优先保留原始扩展名；无法判断用 `.bin` 并占位提示

### 附件命名
- 路径：`assets/attachments/att_<seq>_<safe_name>.<ext_or_bin>`
- `<safe_name>`：从显示名/标题/文件名提取并做安全净化（去除路径分隔符、控制字符，限制长度）
- 格式识别优先级：
  - 有扩展名则优先扩展名
  - 否则用 magic bytes（ZIP 识别 docx/xlsx；OLE2 识别传统 doc/xls；仍未知则 `.bin`）

## 非功能需求（NFR）
### NFR1 稳定性与容错
- 任一元素失败不影响整体产出：必须继续转换并占位。
- 对关系缺失、坏引用、损坏的嵌入对象等情况，必须输出可定位问题的日志。

### NFR2 安全
- 防 Zip Slip：写出路径必须限制在输出目录内。
- 对图片/附件抽取与渲染设置 size 上限，避免内存/磁盘被恶意文件拖垮。
- 限制递归深度，防附件套娃导致无限递归。
- 文件名与路径必须净化，禁止 `..`、绝对路径等。

### NFR3 性能（第一版原则）
- 以流式写入 `index.md` 为主，不要求一次性构建完整 DOM 再输出。
- 大文件优先“可用且不崩”，性能优化后置。

## 范围边界（MVP 明确）
### MVP 必做
- 顺序：段落/表格严格按 `document.xml` 顺序输出
- 图片：抽取 + 原位置引用
- 附件：抽取 + 原位置链接；`.docx/.xlsx` 尽力渲染，失败降级
- 删除线：必须体现为 `~~ ~~`
- 鲁棒性：失败占位 + 日志，不中断整体

### MVP 不做（但必须占位）
- 页眉页脚并入正文（可后续作为附录输出）
- 脚注/尾注完整回链
- 批注/修订
- 公式转 LaTeX
- 文本框/形状/SmartArt 等复杂对象的精准排版

## 验收标准（可测）
以样例文档 `视觉健康档案-数字护眼报告 技术方案文档.docx` 为主，验收至少包含：
- **结构**：产出 `index.md` 与 `assets/images/`、`assets/attachments/` 目录（即使为空也存在）。
- **顺序**：随机抽查 10 处“文字/图片/表格/附件（若有）”交错位置，`index.md` 顺序不乱；浮动对象无法定位时必须有占位且不丢。
- **图片**：`index.md` 的每个图片引用在磁盘上存在；打开路径有效。
- **附件**：若存在嵌入附件：
  - `assets/attachments/` 中存在被抽取文件
  - `index.md` 原位置存在链接
  - 若为 `.docx/.xlsx` 且可解析，存在对应 md 产物；否则存在失败占位并保留链接
- **删除线**：文档中存在删除线文本时，`index.md` 中能看到 `~~ ~~` 体现（且不丢内容）。
- **鲁棒性**：遇到坏关系/坏附件时仍能完成转换并在对应位置占位。

