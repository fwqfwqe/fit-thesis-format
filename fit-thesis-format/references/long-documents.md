# 长论文：分批读取，逐项决定，完整覆盖

适用于所有执行此 skill 的助手。不能仅阅读报告摘要、前若干段或第一批后宣称完成全文审核；也不能把审核草稿中的建议直接批量改为 `reviewed: true`。静态建议用于定位，最终角色必须结合每项完整原文、对象提示、前后边界和学校规范确认。

`review_plan.py` 不修改 DOCX。它沿用 `thesis_format.py inspect` 的每个 XML part 内段落 ID（`p000001`）和表格 ID（`t0001`），以 `part#id` 作为全文唯一键。正文、表内段落、空段落、图形内段落、已有页眉/页脚/脚注/尾注以及其他包含 Word 段落的 story（例如批注）均进入索引。表格作为独立审核项，表内段落也各有独立决定。上下文只是阅读提示，不重复计入决定。

## 1. 创建审核目录

在已经准备的工作目录中，使用 `original.docx` 作为绑定原稿：

```text
python scripts/review_plan.py create --source work/original.docx --work-dir work/review --draft work/style-map.draft.json
```

若 `prepare.py` 已创建 `work/review`，直接读取该目录，不重复创建。独立运行时 `--draft` 可省略，工具会创建空的未审草稿。

产物为轻量 `index.json`、保留原始字节的 `style-map.draft.json` 和 `chunks/chunk-0001.json` 等批次。原输入和草稿不被覆盖。默认每批最多 120 项、60,000 个原文字符，可用 `--chunk-size 100` 和 `--max-chars 40000` 调小；每批最多允许 150 项。单个超长段落或表格保持全文，因此可能单独超过字符上限，不允许截断后假装已经审完。

每个批次保存完整原文、原文与 XML 对象哈希、原样式、所属表格、图形/公式/域/文本框提示，以及上一项和下一项的边界预览。预览最多 500 字符并明确标注截断；完整内容仍在该项自己的批次中。ID 只对绑定的准确原稿有效；改过输入后必须重新建立计划。

## 2. 连续读取每批，逐项审核

切换助手或上下文后，先查看简短续审状态；只输出已审/未审计数、待审批次路径和下一批，不输出论文全文：

```text
python scripts/review_plan.py status --source work/original.docx --work-dir work/review
```

随后读取具体批次：

```text
python scripts/review_plan.py next --source work/original.docx --work-dir work/review
```

`next` 仅返回第一份仍有未审项的批次及路径。编辑其中每项的 `decision`，然后再次运行 `next` 继续后续批次。需要回看时按索引直接读取已有批次文件。不得修改 `text`、ID、哈希、对象提示、上下文或索引分批边界；发现原文不对时回到原稿重新检查。

段落确认应用现有规范角色：

```json
{"reviewed": true, "action": "format", "role": "body"}
```

正文表格确认类型，表内各段仍需分别审核：

```json
{"reviewed": true, "action": "format", "role": "three-line", "header_rows": 1}
```

表格角色只允许 `three-line` 或 `complex`，段落角色以 `references/style-spec.json` 为准。未开放格式修改的 story、story 中的表格、保护图形或对象内部段落必须明确保留，例如：

```json
{"reviewed": true, "action": "preserve", "reason": "位于原有图形文本框内，保留对象布局并在渲染页人工核查"}
```

其他无需修改的内容也可保留，但必须写具体原因。`preserve` 不能同时分配格式角色；原因至少 6 字符，不能仅填写“保持原样”等占位语句。此项检查无法替代实际理解，不得用同一空泛理由批量掩盖未读内容。`reviewed` 必须是 JSON 布尔值 `true`。

## 3. 组装正式映射和覆盖证据

只有 `next` 返回无待审批次后才运行：

```text
python scripts/review_plan.py assemble --source work/original.docx --work-dir work/review --style-map work/style-map.reviewed.json --coverage work/coverage.json
```

组装会从真实原稿重新生成全集，核对源 SHA、索引、每批来源、顺序、重复项、未知 ID、缺失项、完整原文、保护对象、审核状态及角色。任何一项不符都拒绝输出正式结果。两个输出都必须使用新路径，不能覆盖草稿或上一次结果。修订已经组装的决定后，需重新组装到新的正式文件名。

正式 style map 保留 `reviewed: true`、准确源 SHA 和草稿/索引来源，并仅把 `format` 决定转换为现有 `apply` 能接受的映射；`preserve` 决定全部留在 `coverage.json` 中供核验。页面设置默认 `page_setup: false`，需要单独审核布局阶段。此流程不自动继承草稿中的建议、页面设置或样式例外。

若全部项目都明确保留，仍可生成有效覆盖记录，返回 `apply_required: false`。这时不要把空 style map 交给 `apply`，也不能因此宣称格式合格。

## 4. 覆盖验证与交付边界

```text
python scripts/review_plan.py verify --source work/original.docx --coverage work/coverage.json
```

验证必须使用原稿，不能传入已排版的输出。它重新读取全部批次，检查草稿快照、索引、批次和正式 style map 的哈希，并重算全部决定与实际映射的一致性。应完整保留审核目录和正式文件；证据包含绝对路径，单独搬动文件会使验证失败。

覆盖完成只能证明每项都有可追溯的明确审核记录，不能证明审核者理解无误、角色判断正确、学校格式合格或已经逐页查看。交付还需要原文保护验证、最终静态格式检查、实际执行映射的证据及全文逐页视觉检查；保留理由不是内容变化的豁免名单。

页面检查也要分批续审。把页图以可辨认正文和公式的大小逐页打开，每批保存具体页码与观察；下次从第一个未审页继续。拥有看图工具时，不得因篇幅长而停在首中尾抽查，也不能以一份全为 pending 的记录结束本可完成的全文审核。问题页如实记 fail 或需要确认，随后继续检查剩余页；没有修改最终文件时可以沿用哈希一致的已审页，修改后必须重新渲染并复核受影响页面，完整记录仍须对应同一最终稿。

集成接口：

- `create_review_plan(source, work_dir, chunk_size=120, draft_path=None, max_chars=60000)` 创建新目录，返回索引位置和下一批路径。
- `assemble_review_plan(source, work_dir, style_map_path, coverage_path)` 返回计数、正式文件路径与 `apply_required`。
- `review_status(source, work_dir)` 返回无正文的简短续审状态；`next_chunk(source, work_dir)` 返回首个待审批次。
- `verify_coverage(source, coverage_path)` 验证失败抛异常；成功返回 `coverage_complete`、`source_sha256`、`style_map_path`、`style_map_sha256`、`mapped_item_count`、`preserve_count`、`counts`、`verifiable`。`counts` 包含 `total`、`paragraphs`、`tables`、`story_paragraphs`、`story_tables`、`reviewed`、`unreviewed`、`formatted`、`preserved`、`chunks`。`verifiable` 的布尔项分别说明源、全集、去重、原文/对象、审核状态、批次哈希、映射一致性与草稿保留已核验。

交付工具必须把 `style_map_sha256` 与实际执行链中的 styles 阶段映射对上，并核对该阶段从本记录绑定的原稿开始；仅提供一个未实际执行的完整映射，不能替代最终稿的证据。全部保留且没有 styles 阶段时，根据真实执行链处理，不应强迫套用空映射。
