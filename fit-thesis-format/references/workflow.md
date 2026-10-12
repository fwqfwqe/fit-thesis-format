# 排版执行与交付验证

这是在模型理解原稿后使用的基础排版助手，不是无需阅读就一键识别所有论文结构的转换器。全篇逐字不变与命名样式可手调是两项独立交付条件。

仅收到 GitHub 链接时先读[START_HERE.md](../START_HERE.md)，取得完整包。v7 先运行 `python scripts/prepare.py --source 原稿.docx --work-dir 新任务目录`，保留原稿、生成清单、待审映射及 `review/` 分批审核计划。完成全部批次后生成合并映射与覆盖报告，再修改完整 DOCX。最后运行 `scripts/delivery_check.py --coverage ... --chain ...`，绑定真实原稿、全文覆盖、完整阶段链及实际逐页检查记录；它返回 `delivery_ready=false` 时只可称为未验收稿或部分完成。

需要先审核时使用`format_audit.py`，见audit-mode.md。用户已经授权学校封面承诺书、自动标题编号、完整公式标签与编辑器转换时，使用school-standard-mode.md另做结构映射和核验；本基础助手不会实现这些功能，不能只运行它就声称已经完成学校全部格式。

## 运行环境与清单

使用当前环境可执行的Python 3.10+，基础助手仅需lxml/标准库，数学新建助手另需python-docx；完整依赖在包内requirements.txt。宿主有受管理的解释器时遵循其运行规则，否则可使用独立虚拟环境。命令中的PYTHON、SKILL、INPUT、OUTPUT、WORK均替换为当前环境路径，含空格路径作为独立参数传递；不硬编码机器用户名、模型平台或运行时版本。也可通过agent_cli.py提交JSON任务，见agent-protocol.md。

完整审核流程见[长文完整覆盖与断点续审](long-documents.md)。`prepare` 已生成 `WORK/review` 时直接用 `next` 接续；仅在尚无计划时调用 `create`，其工作目录必须是新目录：

```text
PYTHON SKILL/scripts/review_plan.py create --source INPUT --work-dir WORK/review --chunk-size 120
PYTHON SKILL/scripts/review_plan.py next --source INPUT --work-dir WORK/review
PYTHON SKILL/scripts/review_plan.py assemble --source INPUT --work-dir WORK/review --style-map WORK/role_map.json --coverage WORK/coverage.json
```

`next` 返回一批实际原文及审核项；阅读后保存逐项决定，再继续下一批。所有段落、表格及其他文档区域都必须分类或说明具体保留原因，包含空段、附录和末尾对象。不得用“论文过长”为由抽样、缩写、漏尾，也不得把未知角色自动猜成正文或标题。只能编辑审核决定，不能改分批文件中的原文、ID或哈希；中断后复用相同源文件与已保存计划继续。全部确认后才能 `assemble`。

需要样式修改时，下面阶段使用合并生成的 `role_map.json`，一次处理完整输入；每阶段保留完整 DOCX，不将正文拆批另存后拼接。全部项目均已明确保留时可以生成覆盖记录，但不要把空样式映射交给 apply：

```text
PYTHON SKILL/scripts/thesis_format.py inspect INPUT --out WORK/manifest.json
PYTHON SKILL/scripts/thesis_format.py apply INPUT OUTPUT --map WORK/role_map.json --report WORK/apply-report.json
PYTHON SKILL/scripts/thesis_format.py verify OUTPUT --source INPUT --map WORK/role_map.json
```

从inspect读取确切roles/key和段落/表格ID；不要猜ID或靠旧清单操作新文档。映射形状：

```json
{
  "source_sha256": "从inspect取得的完整SHA256",
  "reviewed": true,
  "paragraphs": {"p000001": "body"},
  "tables": {"t0001": {"kind": "three-line", "header_rows": 1}},
  "story_paragraphs": {
    "word/header1.xml": {"p000001": "header"},
    "word/footer1.xml": {"p000001": "page_number"}
  },
  "page_setup": true,
  "landscape_policy": "preserve"
}
```

示例仅说明schema，不能代替完整审核生成的映射；第一段通常不是正文，第一张表通常也不一定是数据表。`story_paragraphs`是可选项，只能列inspect.stories中已存在的页眉、页脚、脚注或尾注部件；每个部件的段落ID独立编号。基础助手只处理现有story；缺失页眉/页脚与编号由下文专用布局助手处理，不能假装基础助手已完成。完整覆盖中的 `preserve` 项不进入样式修改映射，但每项必须有具体保留理由；不能把“保留”全部当成“已排版”或据此豁免学校要求。正文公式/图片段单独识别，现有横向复杂节单独审核。

v7 拒绝没有任何段落、表格或 story 选择的空修改映射，也拒绝 `reviewed:false` 草稿；正式映射须经过实际审核并标记 `reviewed:true`。局部修改仍可用于后续明确的修正，但不能代替首次完整合并映射和全文覆盖记录。`verify` 检查真实样式数值，不能把文字未变当成字号、行距正确。确有版式例外时，可在该阶段映射中记录 `style_overrides`，例如 `{"body":{"reviewed":true,"reason":"经核对需要增加段后距离以容纳相邻对象","properties":{"after_pt":6}}}`；例外影响该命名样式的全部段落，apply/verify 必须使用同一映射。不要直接改动已经与覆盖报告绑定的合并映射；额外例外使用前一阶段完整输出作新输入，另建明确审核的映射并加入阶段链。例外不自动豁免学校要求，最终静态审核仍会报告偏差。

阅读清单后记录一份任务临时 `artifact.md`：输入/范本hash、页面/节/角色映射、特殊对象、允许的格式修改、未覆盖问题与已知冲突选择。格式依据来自技能范本；编辑对象是输入论文的副本。这是针对用户“内容一个字不改”的专用保留策略，不从学校示例重建正文。

## 引用字符样式阶段

先完成基础样式及其verify，再以该输出为新阶段source。`citation_style.py inspect`只列候选，不自动选用；人工审阅正文/综述范围，排除文末参考文献、公式、图表统计标记和脚注。

```text
PYTHON SKILL/scripts/citation_style.py inspect BASE --out WORK/citation-candidates.json
PYTHON SKILL/scripts/citation_style.py apply --source BASE --output CITED --map WORK/citation-map.json --report WORK/citation-report.json
PYTHON SKILL/scripts/citation_style.py verify --source BASE --output CITED --map WORK/citation-map.json
```

```json
{
  "schema_version": 1,
  "reviewed": true,
  "source_sha256": "本阶段BASE输入的SHA256",
  "spans": [{"pid":"p000001","start":13,"end":17,"expected_text":"[12]","role":"body_citation"}]
}
```

示例偏移仅对应合成字串`作者甲（2024）已有陈述[12]。`。start/end是Python Unicode码点的左闭右开范围，必须据真实清单重算。一个完整引用可能跨多个run。助手仅拆分纯文字run，保留每个字及未选部分格式，选中部分绑定可见字符样式“论文-文献引用上标”，只设上标并继承段落字体字号。引用的直接字号/字体/位置覆盖仅在该范围内清除，局部语义加粗斜体仍保留。

字段、书签、超链接、修订、公式、图形、脚注、换行/Tab、表格/文本框内段落以及与其他字符样式冲突的选区会被保守拒绝；整个apply失败，不会悄悄跳过。用保留原结构的定点方法处理并单独核验，或原样保留并报告。没有需调整的引用时可跳过这一阶段，不能为满足脚本而编造span。文献列表编号不属于正文上标，作者年份不由此脚本改成基线或改字。

## 分节、奇偶页眉和页码阶段

用户明确要求的页眉页脚属于允许生成的排版元素，文章逐字保护仍有效。引用阶段完成后（未运行则用基础输出），使用该阶段精确文件重新inspect：

```text
PYTHON SKILL/scripts/thesis_layout.py inspect CITED --out WORK/layout-manifest.json
PYTHON SKILL/scripts/thesis_layout.py apply CITED FINAL --map WORK/layout-map.json --report WORK/layout-report.json
PYTHON SKILL/scripts/thesis_layout.py verify FINAL --source CITED --map WORK/layout-map.json
```

```json
{
  "source_sha256": "本阶段CITED输入的SHA256",
  "title": {"part":"word/document.xml","paragraph_id":"p000001","expected_text":"本篇已确认的题目原文"},
  "regions": [
    {"start":"start","kind":"unpaged"},
    {"start":"p000020","expected_text":"摘要区域首段的完整原文","kind":"front"},
    {"start":"p000040","expected_text":"正文首段的完整原文","kind":"body"}
  ]
}
```

示例ID与文字仅展示schema，不可直接用于论文。`front`应从中文摘要页面真正的首段开始，可能是本篇中文题目，而不一定是“摘要”二字；`body`从第一章首段开始，后续章及参考文献、致谢、附录全归body。没有的前置区域可省略，首region的start固定为`start`，之后按unpaged→front→body顺序。保留所有原段落和文本换行，使用前一个顶层段落的节属性分界；边界紧随表格或复杂容器时会拒绝，需专门布局处理，不能移动/删除表格文字来适配脚本。

题目定位可用`word/document.xml`中纯文字段落，或本稿已存在页眉中的纯文字段落（part填写inspect发现的真实`word/headerN.xml`）。expected_text必须逐字一致。主副标题来源不一致或跨复杂域时需核实，不能自动缩写、拼凑或借用范文题目。

布局助手复用已安装的“论文-页眉”“论文-页码”段落样式，启用evenAndOddHeaders，每节显式绑定default/even/first以消除错误继承，取消首页不同，front大写罗马从I起且续号，body十进制从1起且所有后节续号，unpaged页眉页脚为空。页脚新建真实PAGE域，由分页显示实际数字，不设锁定、不更新整篇域。页眉题目从本稿逐字复用，偶数页固定校名；不修改文章题目。

现有页眉脚XML保留原字节，新建运行用部件，只改文档节属性、settings奇偶开关、所需关系和内容类型。含绘图、数学对象、表格、水印或复杂内容控件/修订的旧页眉页脚不会被粗暴替换，助手会拒绝；需辨认可见内容再定点修改，文本框兼容表示不可重复计数。新增分节如可能使按节重启的脚注/尾注或SECTION/SECTIONPAGES域改变显示，也会拒绝，不能以文字字节未变当作编号没变。

## 内容保留

- 保存输入SHA256、ZIP parts、原文本序列、段落/表格结构、math/drawing/media/embedding/chart、域/编号、脚注尾注/评论修订、超链接等基线。编号不能只比较numId：旧样式里无有效绑定的numPr在Word中可能不显示，迁移时不得将它激活；有编号与无编号的实际显示都要保持。
- 助手apply在临时包上检查所有原文本节点，并保护公式/图形/二进制对象；失败不写最终输出。仅主文档、styles及映射明确选择的现有story可发生纯格式变化，其他部件保持原字节。story内容骨架、文字、域及编号同样核验。verify后仍需检查变化仅是允许的格式属性。
- 不拆分/合并/删除原段落/表格，不迁移内容。对无法套用的原文仍保留并列出例外。
- 引用阶段仅对审核过的纯文字run做必要拆分，不拆段落；布局阶段仅新增/修改节格式及明确授权的运行页眉页脚。两者的verify从该阶段source+map重算精确允许的包内容，核对所有成员集合与逐部件字节，防止借“忽略页眉”漏掉正文变更。拒绝重复ZIP部件名。原页眉脚部件不删除，不能把误改原部件藏在新引用后。
- 最终保留 `原稿 → BASE → CITED（可选）→ FINAL` 的完整文件、来源SHA与各阶段报告，交付前用 `--coverage` 核对原稿全部项目、用 `--chain` 重新验证整条链；修改过的文件即使只改格式也不能省略链。中间文件只用于审计，不默认交付给用户。不能从最终稿创建新的baseline冒充原稿核验。基础助手不接受run拆分/新部件，其严格检查仅适用于BASE阶段。
- 若助手之外另改了内容部件、对象几何、分节或样式，运行verify并独立核验对应结构。不能通过跳过比较或重置baseline让修改“通过”。确需几何变更但助手严格对象检查拒绝时，保留原对象并说明，或另实现仅比较几何之外内容的明确格式审计；不可放宽文字/媒体检查。
- 有未决修订时保留所有插入/删除/作者时间和显示状态，不自动接受修订。强制保护、锁定控件应作为编辑限制报告。
- Word只读渲染不保存源稿/成稿。不要在渲染时调用Fields.Update、TOC.Update、Save等改写缓存；PDF仅内部检查。Word的打印/导出设置也可能自动重算目录：若使用专用Word实例，记录并临时关闭UpdateFieldsAtPrint，结束时恢复原值；不能将渲染引擎重新计算的目录内容写回成稿。最终内容校验以保存的DOCX及其原域缓存为准，动态域在用户自行刷新后可能重新计算。

## 手动调样式的验证

1. styles.xml中的论文样式为paragraph类型，有name/qFormat/uiPriority，非隐藏非锁定，不开启自动更新。
2. 选中段落实际pStyle指向其对应样式；字号/字体/行距/缩进等不被普通直接格式压住。正文强调用保留的局部格式/字符样式，检查不会使普通正文继承失效。
3. 未把正文或图表截图替代可编辑文字；数学对象保留，表格仍可编辑。
4. 在内部QA副本中只修改“论文-正文”的字号/行距，验证多个普通正文段落随之变化，再丢弃QA副本。最终成稿保留学院参数。也检查通过Word样式面板可选择标题/正文。
5. 内置TOC样式/原字符样式可能覆盖局部设置；保留域缓存时单独处理其显示格式，不能更新内容来掩盖不一致。原自动编号若继承于旧段落样式，需要保存原编号属性或创建保留同一numId/ilvl的适配样式，不重写文字或编号值。

## 上下空行及目录制表位

关键词、表格、图片与正文之间的空一行先检查原文是否已有空段。没有时调整对应命名样式的段前/段后（例如20磅），已有时避免重复加空白；不插入或删除原段落、空格、换行。根据上下文需要对最终DOCX的styles.xml做样式级微调，不改全局技能spec，不以逐段直接spacing代替样式。保持原角色styleId，便于map核验。

目录的现有文字、域和内容Tab保留。右对齐页码和点引导线属于制表位格式，可以在目录样式设置右制表位（位置按所在节正文宽度计算），清理相冲突的段落直接制表位；不能把目录文字重新拼写或重建域。正文run里的Tab是内容，pPr/tabs里的tab是格式，两者校验时分开。

## 版式QA

- 页面与边距正确，各标题层级、大纲及段间距正确，每章段前分页。
- 目录内容/引用/题注文字不变；目录陈旧页码属于需授权更新的内容限制，交付说明。
- 图表不超页宽、不裁切、不改数据，公式分式/上下标完整；不要把包含图片的行强设固定20磅造成裁切。
- 数据表三线或复杂框线正确，重复原表头、可读行高；跨页缺“续前表”报告但不添加。
- 页眉页脚核对封面与承诺页、罗马页码首末页、正文1/2/3/4、跨节边界及全文末页；显示页码奇数题目/偶数校名，每页均有对应PAGE且位置统一，后置部分不中断。检查Word分页/打印是否产生额外空白页，不复制范文空白页，不通过删文字解决。复杂结构尚未处理时报告实际缺口，不能声称全完成。
- 已知实测限制：原生Word在一个合成样例中，在罗马III之后重启正文1并启用奇偶页眉时，导出插入了一张空白物理页；仅将分节设为nextPage或continuous并不足以消除。助手不能保证不同Word/WPS版本的物理页数一致。遇此情况按目标软件检查分节及打印分页，不暗改原文分页符、整篇兼容设置或手动编号来掩盖；未解决就说明实际空白页，不能宣称版面全部核验通过。
- 引用核对作者、连接词、年份和括号基线、完整方括号上标、标点基线；文末编号仍在基线。作者—年份叙述另查漏号及人数：一位写该作者，两位中英文均用“和”，三位及以上中文用“等”、英文用“et al.”，见[引用一致性](citation-consistency.md)。纯排版保留原字串和位置；已授权引用规范化的有限字符变更必须逐项记录、绑定已有文献并重验。只改“论文-文献引用上标”样式应能调整选中引用，不能改变普通作者年份或公式符号；正文样式改字号也应带动引用的基础字号。
- 中文摘要超一页、文献缺失等不通过删字/缩字修复。
- 渲染最后一次输出并逐页检查。优先用户目标软件原生导出；已有Word、其他文档引擎或文档API均可提供渲染能力，具体调用和只读/更新域区别见portable-rendering.md。没有渲染能力时明确剩余检查，不假报全部通过。

## 交付措辞

交付新的DOCX，报告核验结果与真实限制。示例：“已按学院范本调整排版，原文字、公式和图表内容核验一致；各部分已绑定可编辑Word样式。统一手调：开始→样式→右键‘论文-正文’→修改。”

若部分章节/来源缺失或编号文本与规范不符，简短列明“按你的内容不变要求予以保留”。不要宣称全面合规。无原文件的首次调用只需说明可提交论文，不要生成空壳论文冒充成稿。
