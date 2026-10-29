# 来源绑定的局部版面修复

在 Word 和 WPS 的真实逐页结果中发现问题后，仍从不可变原稿重跑。不能直接修改候选文件再声称通过原稿完整性校验。所有决定文件都必须带原稿 `source_sha256`，下列位置 ID 来自同一原稿的 `inspect`。

## 段落、表格和图片段落

`decisions.layout_repairs` 是列表，每项包含 `kind`（`paragraph`、`table`、`figure`）、`id`、原稿证据和 `set`。`paragraph` 使用 `p000001`；`table` 使用 `t0001`；`figure` 使用包含原生 DrawingML 图片的段落 ID。`expected_text` 是原目标内全部 `w:t` 的连接。`expected_sha256` 是 `sha256(thesis_format.xml(原目标元素))`，使用库统一序列化，不能哈希界面复制的文本。

非空普通段落至少提供精确 `expected_text`；表格、图片和空段落必须提供 `expected_sha256`。两项都提供时两项都校验。示例：

```json
{"source_sha256":"原稿SHA256","layout_repairs":[{"kind":"paragraph","id":"p000123","expected_text":"原稿原段完整文字","reason":"Word和WPS均显示标题与内容分离","set":{"keep_with_next":true,"before_pt":0,"after_pt":6}}]}
```

允许的属性只有 `page_break_before`、`keep_with_next`、`keep_lines`（严格布尔值），以及 `before_pt`、`after_pt`、`line_min_pt`（磅值）。前后间距范围 0–144，最小行距范围 1–144，只接受 0.05 磅的整数倍。不得借此改变文字、字体、字号、缩进、数学对象或引用。表格决定作用于该表直接拥有的段落，段前分页只作用于第一段；不能保证某办公软件一定把表格分页到预期位置，仍须双引擎复核。

修复采用可见、可编辑、继承语义角色的原生段落样式，名称和 ID 由角色及受限属性确定。重跑会识别并保留它们，不累积样式。所有操作以 `layout_repair` 写入报告。

## 图片宽度

自动检查主文档中包含单个原生图片的 DrawingML `inline` / `anchor`。超出当前节版心或所在单元格可用宽度时，按同一比例缩小外框及图片变换尺寸。媒体字节、裁切区域和关系不变；不放大图片、不重绘、不解散组合。图表、组合图和旧 VML/OLE 不在此自动缩放范围，必须在视觉报告中逐项核对。缩放以 `picture_fit` 记录前后尺寸和对象哈希。

## 续前表

阅读 Word 和 WPS 的原生表格行跨页记录后，在原稿的表格决定中提供 `continuation_breaks`：在这些 **原表 0 起始行号之前**切分。它必须严格递增，每个位置大于 `header_rows` 且小于总行数；不能把只有表头的部分单独切出。必须同时提供原表 `expected_sha256`，可另加 `expected_text`。例如：

```json
{"source_sha256":"原稿SHA256","tables":{"t0003":{"kind":"three-line","header_rows":2,"expected_sha256":"原表元素SHA256","continuation_breaks":[18,34]}}}
```

每个后续片段前插入右对齐原生段落“续前表”，仅复制原文本表头；每一原数据行只移动一次。保留表格属性、网格和数据对象，重新建立各片段边线。表头存在数学、图形、域、书签、修订、注释、脚注、嵌套表格等对象时拒绝复制；跨断点纵向合并也拒绝，不猜测拆分。已有“续前表”后面的片段不可再次按原表断点分割。

`table_continuation` 报告源表 ID/哈希、断点、表头行数、增加的表头文字/XML、标签数和每个数据片段行范围。它是来源绑定的版面操作，不自动推测 Word/WPS 的最佳断点。切分后必须重新更新域、双引擎渲染、审核全部续页；长表未得到这些证据时不得交付 PASS。

## 附录代码

明确确认的 `appendix_code` 段落保留原有字体、缩进和空白，通过继承原样式的命名代码样式保留识别身份。普通附录段落也不会做段首/段尾空白清理。代码内容不得套用引文、公式引用或作者展示修复。

## 图表题位置

无歧义且紧邻的“表后表题”或“图前图题”会整段移动到表前、图后，不重写标题文字。不相邻或多目标位置用 `decisions.caption_targets` 列表绑定：`caption_id`、`expected_text`、`target_kind`（`table` 或 `figure`）、`target_id`、`target_sha256`，可加 `target_expected_text` 和 `reason`。目标哈希仍由原稿目标元素的 `thesis_format.xml` 计算；图目标使用只含图片的段落 ID。两者必须是主文档同级块。表题固定移到表前，图题固定移到图后，无任意 position 参数。移动以 `caption_position` 记录，须重渲染验证。

## 纯文本公式转写

先确认段落角色为 `equation`。只有源段为普通文本、带明确公式标签、含等号且没有分式/上下标等非线性语法时，才可原符号转成原生 OMML。已有 OMML 始终保留。复杂纯文本公式由宿主逐项核对后提供 `decisions.formula_transcriptions`：段落 ID 对应 `{expected_text, omml, reason}`。`omml` 必须是一个 `m:oMath` XML；解析器将分式、上下标、括号结构还原为线性表达并核验原字符和分组。失败作为未决问题保留，不猜数学含义。

旧 OLE、图片公式、带格式上下标或书签/域的源段不适用此入口，不自动重写。每项成功转写以 `formula_transcription` 记录原表达式、OMML 和哈希，独立完整性校验再次与原稿绑定；真实可编辑性仍须 Word/WPS 实测。
