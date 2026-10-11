# 只收到 GitHub 链接时：从这里开始（v6）

本文件写给接收仓库链接和用户论文的 AI / Agent。**先拿到完整资源并证明工具能执行，再修改原稿副本。** 不能只读网页说明，就根据提取文字新建一份 Word 并宣称完成格式修改。规则与脚本不绑定 AI 品牌，但网页浏览、文件执行和文档渲染是不同能力。

## 1. 取得完整包

官方仓库：https://github.com/fwqfwqe/fit-thesis-format

- v6 完整 ZIP：https://github.com/fwqfwqe/fit-thesis-format/releases/download/v6/fit-thesis-format-v6.zip
- 当前版本校验清单：https://raw.githubusercontent.com/fwqfwqe/fit-thesis-format/main/package-manifest.json
- 本启动指南 raw：https://raw.githubusercontent.com/fwqfwqe/fit-thesis-format/main/fit-thesis-format/START_HERE.md
- 标准库下载校验工具：https://raw.githubusercontent.com/fwqfwqe/fit-thesis-format/main/bootstrap.py
- 固定 v6 校验清单：https://raw.githubusercontent.com/fwqfwqe/fit-thesis-format/v6/package-manifest.json

有网络下载和 Python 工具时，把下载校验工具保存为 `bootstrap.py` 后运行：

```text
python bootstrap.py --dest package-new
```

它按仓库当前清单下载相应版本的包，核对 ZIP 和逐个文件的 SHA256，再解压到尚不存在的 `package-new/fit-thesis-format/`。记录实际版本；本指南对应 v6，不能混用不同版本的脚本和清单。

需要固定 v6 或环境没有网络时，用宿主文件工具取得**同版本**的 ZIP 与校验清单，保存为 `fit-thesis-format-v6.zip` 和 `manifest-v6.json`，然后运行：

```text
python bootstrap.py --manifest manifest-v6.json --archive fit-thesis-format-v6.zip --dest package-new
```

如果当前会话不能下载但能读附件，让用户上传完整 ZIP 和原始论文 DOCX。仅收到 `SKILL.md`、网页文本或几张截图时，模板、脚本和原文结构仍未取得；不要假装资源齐全。没有 Python 时可用宿主解压和 SHA256 工具做相同校验；无法验证的部分如实记录。

解压后至少应实际读到 `SKILL.md`、`references/format-requirements.md`、`references/style-spec.json`、`references/workflow.md`、`references/school-standard-mode.md`、`agent-tools.json`，并取得 `assets/fit-economics-management-template.docx` 和完整 `scripts/`。发现缺文件先修复获取步骤，不能靠记忆补规则。

## 2. 实测当前会话的能力

不要按“GPT 网页版”“WorkBuddy”或其他产品名称推测工具。先确认：能否取得附件原始字节、写新文件、执行 Python 或文档 API、导出 DOCX、渲染页面，以及实际查看页面图片。

| 只有这些能力 | 本次可以做什么 |
| --- | --- |
| 网页浏览 | 读取格式要求、说明缺少的工具；不能修改论文文件。 |
| 附件文字提取 | 对可见文字提出检查建议；不能确定真实样式、分节、编号、公式对象和分页。 |
| 文件读写与 Python | 执行清单、静态审核、已审查映射的格式修改与内容核验。 |
| 文档 API，无 Python | 用内部副本先实测公式、域、编号、图片、分节及导出能否保留；仅执行可核验的操作。 |
| 已有 DOCX 渲染引擎及图片查看 | 检查实际页面；记录真实引擎、字体问题和检查范围。 |

缺少 Word 不一定完全不能做：已有其他渲染引擎可辅助检查，但分页和公式兼容性不能视为与 Word 相同。不能导出或不能保留原对象时，交付可完成的审核，不要用重新手写正文来绕过。

有 Python 3.10+ 时使用当前环境；确实缺依赖且宿主允许安装时执行：

```text
python -m pip install -r package-new/fit-thesis-format/requirements.txt
```

将下面内容保存为 `doctor-request.json` 后执行环境发现：

```json
{"action":"doctor"}
```

```text
python package-new/fit-thesis-format/scripts/agent_cli.py --request doctor-request.json
```

doctor 只发现解释器、依赖和候选渲染工具。发现 Word 注册项不代表已经导出成功，发现渲染程序也不代表看过页面；后续必须用当前候选文件实际测试。

## 3. 留下原稿，运行审核

把论文附件落实为工具可读取的原始 DOCX，路径以当前工具环境为准，不能把用户电脑路径直接当成云端文件。再执行：

```text
python package-new/fit-thesis-format/scripts/prepare.py --source "论文.docx" --work-dir work-new
```

工作目录必须是新目录。该步骤产生 `original.docx`、`baseline.json`、`doctor.json`、样式/布局/引用三份清单、`audit.json` / `audit.md`、三个 `*.draft.json` 映射草稿、`START.md` 和 `prepare-result.json`。**prepare 不修改论文，也不代表已经完成排版。** 阅读生成的 `START.md`、清单和审核，确认每个标题、正文段、图表、公式、参考文献及附录的真实角色。

只要求审核时，到此结合实际可用的页面检查交付问题清单，不修改用户文件。需要修改时，阅读 `SKILL.md` 指向的规则和对应工具的映射格式；不能直接把草稿喂给 apply，也不能把所有段落标成正文。

## 4. 在副本定点修改，保护原对象

原稿有公式、自动标题编号、目录、交叉引用、书签、图片、图表、脚注尾注、批注或修订时，保留其原有结构。不能用 `paragraph.text`、全文抽取或空白文档重建来消掉这些对象。`academic_numbering.py` 用于示例和明确映射的结构辅助，不能当作任意已有论文的一键转换器。

最低格式要点是：学校真实封面和诚信承诺书；四级标题“一、”“（一）”“1.”“（1）”；两级目录；真实奇偶页眉，正文奇数页是论文题目、偶数页是学校固定文字；可编辑数学结构与完整“公式（章-序号）”；独立公式无表格、文本框和边框；引用编号上标，作者年份基线；致谢最后，已有附录完整置于致谢前。**具体字号、间距、分节和转换限制必须读完整规则，不能只靠这段摘要执行。**

对每一阶段保存输入、映射、输出与核验报告，绑定准确 SHA256。学校封面替换、手写编号转换、公式转写等需要额外逐项核验；普通格式脚本不能自动完成所有这类转换。任何未完成的转换须列出，不能通过重设内容基线或忽略差异来取得“通过”。

## 5. 对当前候选稿重新验收

先运行静态交付检查，路径换成实际文件：

```text
python package-new/fit-thesis-format/scripts/delivery_check.py --source "论文.docx" --final "work-new/candidate.docx" --report "work-new/delivery-check.json"
```

读取输出里的实际问题和未完成项。需要阶段链与页面证据时按 `references/agent-protocol.md` 提供 `--chain`、`--render-manifest`、`--page-manifest` 和 `--review`。不能用一个截图或旧版本 PDF 的检查结果代替当前最终稿。

渲染方法见 `references/portable-rendering.md`。实际逐页检查封面、目录、标题编号、显示页码、奇偶页眉、公式上下标、图表和最后的致谢。渲染报告绑定 DOCX 与 PDF 哈希，页面清单绑定每张图；检查记录写明实际看过的页。**机器核对记录与哈希，不能证明 AI 真的看过页面，不能编造检查记录。**

## 6. 诚实交付

报告本次达到的状态：只读到规则、完成审核、完成候选修改但未验收，或完成已声明范围的验收。没有实际产出文件就不能说“已生成 Word”；保护检查失败就不能把候选稿交付成完成品；没有逐页检查就标明“未完成逐页视觉核验”。

保留原文件，提供可编辑候选或通过验收的 Word 以及简短问题说明。不要把范文内容、内部日志或映射插入论文。用户可在 Word“开始→样式→修改”统一手调。
