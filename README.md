# 福州理工学院论文格式 · v7

**如果用户只发了这个 GitHub 链接，先读下面的启动指南，取得并校验完整包，再检查自己的执行能力。只读到首页，不能开始重新生成论文。**

[AI 单文件启动指南](https://raw.githubusercontent.com/fwqfwqe/fit-thesis-format/main/fit-thesis-format/START_HERE.md) · [下载 v7 完整包](https://github.com/fwqfwqe/fit-thesis-format/releases/download/v7/fit-thesis-format-v7.zip) · [当前包校验清单](https://raw.githubusercontent.com/fwqfwqe/fit-thesis-format/main/package-manifest.json) · [下载校验工具](https://raw.githubusercontent.com/fwqfwqe/fit-thesis-format/main/bootstrap.py)

本项目提供福州理工学院经济管理学院的格式规则、模板和执行工具，不绑定 AI 品牌。**能否完成排版，取决于当前会话实际拥有的文件、代码、文档和看图工具。** GPT 网页版、WorkBuddy 或其他 Agent 都按实际能力检查，不根据产品名字承诺成功。

## 用户怎么用

把本仓库链接和原始论文 DOCX 一起交给 AI，并说：

> 先按仓库的 START_HERE.md 取得并校验完整包，检查当前工具能做什么。用 prepare 建立原稿基线，分批审核全文，每一段都要确认角色或逐项说明保留原因，再合并映射对完整副本只改格式。不要因论文长而抽样、缩写或漏掉后半篇，不要从提取文字重新生成论文，不要丢失公式、目录、编号、图片、脚注或引用。最后核对全文覆盖、内容和每一页；做不到的环节要明确说明，未验收文件不能说成已完成。

如果 AI 只能浏览网页，它可以解释规则；如果它能读取论文、执行脚本或使用文档工具，才可以继续实际审核、修改。网络下载不可用但能读取 ZIP 附件时，可以把**完整 ZIP 和论文 DOCX**交给它，不能仅上传 `SKILL.md` 代替模板和脚本。

## 工作流程，用最简单的话说

1. **拿齐工具。** 下载完整包，检查有没有缺文件、下载错版本。
2. **留下原稿。** 保存副本和内容记录，先找出格式问题。
3. **分批读完全文。** 逐项确认标题、正文、公式、图表和参考文献；保存每批审核状态，可断点继续，最后合并映射修改完整副本。
4. **核对覆盖与内容。** 全部段落、表格和文末都有审核决定；检查文字、数字、公式、图片和引用有没有丢失或改变。
5. **逐页看。** 检查实际页眉、分页、目录、公式和图表。
6. **交付。** 给可编辑 Word 和真实的检查结果；未验收就明确标为未验收。

## 有 Python 的 Agent：从这里运行

需要 Python 3.10+。先用已有下载工具把上面的 `bootstrap.py` 保存到本地，再执行：

```text
python bootstrap.py --dest package-new
```

`package-new` 必须尚不存在。下载工具使用标准库取得仓库**当前**校验清单和对应 ZIP，逐项校验后解压到 `package-new/fit-thesis-format/`。记录实际取得的版本；不要把不同版本的 ZIP、清单或规则混用。离线方式及固定 v7 的用法见[启动指南](fit-thesis-format/START_HERE.md)。

在已经满足依赖的环境直接运行；缺依赖且宿主允许安装时，使用包内 `requirements.txt`：

```text
python -m pip install -r package-new/fit-thesis-format/requirements.txt
python package-new/fit-thesis-format/scripts/prepare.py --source "论文.docx" --work-dir work-new
```

`prepare` 只读原稿，在新的工作目录生成原稿副本、内容基线、三份结构清单、审核报告、待审映射草稿、`review/index.json` 与分批审核文件，以及 `START.md` 操作步骤。按批阅读原文并保存决定，全部完成后生成合并映射和 `coverage.json`。**它不会完成排版；未知角色与疑似标题须确认，不能自动把建议当决定。** 原有公式、自动编号、目录域和复杂对象需要按原结构保留，不能重建成普通文字。

长论文同样必须审核全部项目；每项明确选择格式角色，或说明具体保留原因。分批的是审核工作，修改时仍处理完整 DOCX，每阶段都保留完整文件。覆盖完整不代表格式合格，全部标成“保留”也不能免除格式与逐页检查。操作见[长文完整覆盖与断点续审](fit-thesis-format/references/long-documents.md)。

修改工具与映射规则见[通用执行协议](fit-thesis-format/references/agent-protocol.md)和[基础助手](fit-thesis-format/references/workflow.md)。工具参数 [agent-tools.json](fit-thesis-format/agent-tools.json) 是接入说明，打开它不会自动给聊天软件安装可调用工具。

候选文件生成后，重新检查，而不是复用修改前的审核结果：

```text
python package-new/fit-thesis-format/scripts/delivery_check.py --source "论文.docx" --final "work-new/candidate.docx" --coverage "work-new/coverage.json" --chain "work-new/chain.json" --report "work-new/delivery-check.json"
```

上例中的覆盖报告与阶段链必须来自本次真实操作；修改过的 DOCX 必须提供阶段链。没有渲染证据与逐页检查记录时，只能报告已完成的静态检查。完整交付证据的参数见[通用执行协议](fit-thesis-format/references/agent-protocol.md)；学校封面替换、公式转写等结构转换需要专门保留和核验，不能用一句“允许变化”跳过保护检查。

## 没有 Python 或 Word 时

| 当前实际能力 | 可以完成什么 |
| --- | --- |
| 只有网页浏览 | 读取规则、说明准备材料；不能声称已修改论文。 |
| 附件只可提取文字 | 给文字层面的提示；不能确认 Word 字号、公式对象和分页。 |
| 文件与 Python | 静态审核、经过角色映射的格式修改及内容核验。 |
| 文档 API，可读写并导出 DOCX | 先实测它是否保留公式、编号、域、分节和嵌入对象，再执行支持的修改。 |
| DOCX 渲染工具和看图能力 | 实际检查页面；记录使用的引擎和检查范围。 |

没有 Microsoft Word 时可以使用已有文档引擎；不同引擎的分页、字体和公式表现不保证相同。有工具调用能力也不意味着以上能力都存在。缺少保护原文所需的能力时，保留原稿并交付可完成的审核，不能通过重新手写一份 DOCX 来绕过。

## 格式与手调

保留学校封面和诚信承诺书、四级自动编号、两级目录、真实奇偶页眉、连续页码、可编辑公式和完整“公式（x-x）”。独立公式不用表格、文本框或边框；正文引用编号上标，作者年份保持基线；已有附录保留并置于最后的致谢之前。

Word 中通过“开始→样式→修改”统一手调。详细要求在[完整 skill](fit-thesis-format/SKILL.md)，渲染方法和证据记录在[便携渲染说明](fit-thesis-format/references/portable-rendering.md)。支持 skill 的宿主可加载完整文件夹；`agents/openai.yaml` 是可选界面配置。

仓库包含执行工具和学校格式资料，没有上传用户研究论文、数据集或运行日志。哈希可以确认使用的是同一份文件，不能证明 AI 真的看过每一页。
