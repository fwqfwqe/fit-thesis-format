# 通用渲染与逐页复核

这一步把当前候选 Word 文件变成实际页面，再由具备图像查看能力的 AI 或人工检查。它不依赖特定 AI 品牌、插件、作者标记、专有运行时目录或文件链接语法。若工具不能直接打开文件，只能给建议，不能声称已经执行排版或逐页验收。v6 将 DOCX、PDF、每张页面图片和实际检查记录用哈希连接起来；修改最终文件后旧证据不能继续算通过。

## 选择当前环境已有的工具

| 已有能力 | 可用方案 | 必须说明的限制 |
| --- | --- | --- |
| Windows 和 Microsoft Word | 本目录的 `scripts/word_render.ps1`，或等效的 Word 文档 API | Word 版本、字体和打印设置也会影响分页；需要实际查看输出。 |
| 已连接的文档应用 API | 用其读取副本、导出 PDF，并记录应用名称及更新域行为 | 先确认 API 能否只读导出、是否会保存原稿或自动更新字段。行为不明时不能用于严格保留任务。 |
| 已安装 LibreOffice | 用现有转换功能或 CLI 在独立临时目录导出 PDF | 与 Word 对 OMML、域、多级编号、奇偶页眉及分页的表现可能不同，不视为等价验收。 |
| 现成的 PDF 查看器或页面截图能力 | 打开导出的 PDF，逐页查看 | 必须能读清实际文字、公式和边界；只有抽取文字不能证明页面无重叠。 |
| Python 和可选的 pypdfium2、Pillow | 用 `scripts/render_pdf.py` 将全部页面转为 PNG | 这是可选工具，不具备这些依赖时可以用其他已有查看工具。 |

不要为满足本 skill 默默下载或安装办公软件、字体或 Python 包；先使用现成的能力。若没有渲染引擎或图像查看能力，仍可完成结构审核，但必须把“页面外观未验收”列为未完成项。不要把 XML 检查通过写成全部排版合格。

已有 LibreOffice 的示例命令如下，`soffice` 替换为本机真实可执行路径，输出目录和用户配置目录每次使用新的目录：

```text
soffice -env:UserInstallation=file:///ABSOLUTE/TEMP/unique-profile --headless --convert-to pdf --outdir ABSOLUTE/NEW/OUTPUT source-copy.docx
```

先确认输出不存在，转换后核对源文件哈希。即使传入原稿也不得使用会原位保存的参数；对于严格保持域缓存的要求，优先使用行为已知的只读方案。LibreOffice 可能在导入时重新计算字段，其导出结果必须按实际引擎记录。

## Word：默认只读导出

在 PowerShell 中从 skill 根目录执行，路径中有空格时保留引号。PDF 与报告都用新路径；保存本次真实标准输出，不手填引擎或哈希：

```powershell
& ./scripts/word_render.ps1 `
  -InputPath "C:\work\source.docx" `
  -OutputPdf "C:\work\qa-new\source.pdf" |
  Out-File -LiteralPath "C:\work\qa-new\word-export.json" -Encoding utf8 -NoClobber
```

运行需要 Windows、允许 COM 的 PowerShell 和已经安装的 Microsoft Word。脚本使用自己创建的隐藏 Word 实例，禁止宏和打开/打印时的链接更新，按只读方式打开 DOCX，不主动调用域或目录更新，不保存输入，不连接或关闭用户正在使用的 Word 实例，不按进程名强行结束 Word。它在 `finally` 中关闭自己创建的文档和实例、恢复修改过的应用设置，并核对输入 SHA256 没有变化。

输入必须为现存 DOCX；输出 PDF 必须是新的路径，已存在的文件或目录都会拒绝覆盖。脚本输出 JSON，记录 `engine`、`engine_version`、`pages`、`pdf`、`pdf_sha256`、源文件前后 SHA256 和是否显式更新域。只读模式没有另存的 DOCX，使用 `source_sha256_before` 绑定当前候选稿。成功输出 PDF 不代表已经看过每页；无有效报告或进程失败时不能把留下的 PDF 当作通过结果。

注意：只读和不主动更新域保证的是“不写回源文件、不调用字段更新操作”；Word 仍会在布局时计算页码等动态显示。不能据此承诺每一个显示值都与文件内旧缓存完全一致。如果任务要求缓存显示逐项不变，应额外比对域代码、缓存及渲染结果，并报告引擎不可避免的显示差异。

## 只有已授权副本才更新字段

新建示例通常需要更新目录、SEQ 和 REF；已有论文仅在用户授权该副本更新字段时使用下列模式。参数表示调用方确认已有授权，不能把设置开关本身当成用户授权。

```powershell
& ./scripts/word_render.ps1 `
  -InputPath "C:\work\authorized-copy.docx" `
  -OutputPdf "C:\work\qa-new\updated.pdf" `
  -UpdateFields -AuthorizedCopy `
  -SavedPath "C:\work\updated-new.docx" |
  Out-File -LiteralPath "C:\work\qa-new\word-export.json" -Encoding utf8 -NoClobber
```

三个更新参数必须一起给出；仍以只读方式打开输入，两轮更新正文与其他文档区域的域和目录，重新分页后另存为新的 DOCX，然后导出 PDF。不会覆盖输入或现存输出。自动更新也可能影响日期、链接或交叉引用显示，交付前仍需按授权范围核验内容。失败后可能留下未完成的输出，只能在确认成功及检查后交付；重试应选新输出路径。

此时报告额外给出 `saved_docx` 和 `saved_docx_sha256`，页面证据对应这个另存文件。不要把更新之前的文件哈希填作最终稿。当前 `delivery_check.py` 不用任意 allow-change 参数放行域缓存改写等自定义变化；更新导致受保护内容差异且没有支持的重放验证时，应保留差异和专门核验记录，标为未验收，不能通过替换原始基线取得通过。

## 可选：全部页面转为图片

```text
python scripts/render_pdf.py "qa-new/source.pdf" "qa-new/pages" --dpi 144
```

可选依赖为 `pypdfium2`、`Pillow`，实测版本见 skill 根目录的 [requirements-render.txt](../requirements-render.txt)；不在脚本中安装依赖。输出目录必须尚不存在。脚本逐页生成 `page-001.png` 等文件和 `manifest.json`，记录引擎版本、PDF 输入 SHA256、页数以及每张 PNG 的 SHA256。`pages[].page` 是 PDF 物理页序，不是正文显示页码。

使用自己的 PDF 截图工具也可以，但需覆盖全部页面并产出相同的可核验字段：`input_sha256` 是实际 PDF 哈希，`page_count` 是总页数，`pages` 中每项有 `page`、`file`、`sha256`。文件路径相对该清单自身位置或使用绝对路径。其他 DOCX 引擎的导出报告也须提供实际引擎及版本、页数、DOCX/PDF 哈希和源文件保留结果，不能直接填写“Microsoft Word”冒充。

## 记录真实的逐页观察

图片生成后，实际逐页打开，再创建 `review.json`。下面是**待填写格式**，不是通过记录，不可把 pending 批量替换为 pass：

```json
{
  "final_docx_sha256": "最终候选DOCX的实际SHA256",
  "pdf_sha256": "本次导出PDF的实际SHA256",
  "reviewer": "实际复核者或Agent标识",
  "pages": [
    {
      "page": 1,
      "image_sha256": "本页图片的实际SHA256",
      "status": "pending",
      "checks": {
        "layout": "pending",
        "headers_footers": "pending",
        "numbering_toc": "pending",
        "math_tables": "pending",
        "content_order": "pending"
      },
      "notes": ""
    }
  ]
}
```

每个物理页恰好一条记录。实际检查后填写具体观察，例如本页类型、显示页码、页眉和公式是否裁切；只有确实通过的项目才填 `pass`。本页没有目录或公式时也应确认不存在异常，不能省略对应检查项。发现问题就保持未通过并修复，不伪造通过观察。按显示页码判断奇偶页眉，同时在 notes 中记录相关显示页码。

`delivery_check.py` 重新计算实际 DOCX、PDF、PNG 哈希，并通过 `requirements-render.txt` 中的 pypdfium2 独立读取 PDF 实际页数，检查总页数和逐页记录是否完整一致。只审核或只做格式修改不必具备此渲染依赖；但运行完整页面证据校验时需要它，不能只凭三份报告写了相同页数就认为覆盖完整。检查器不会替你查看图片，也不能证明记录者真的看过或判断正确。执行方法见[通用执行协议](agent-protocol.md)。

## 页面验收清单

1. 校徽、封面横线、诚信承诺书与个人信息位置正确。
2. 摘要、目录、章标题的字体、间距和分页正确；目录页码与实际内容相符。
3. 按 Word 的显示页码核对奇偶页眉，不能仅按 PDF 的物理页序推断；空页需要解释原因。
4. 公式在普通段落中居中，右侧为完整“公式（章-序号）”，无表格网格、文本框或边框；分式、上下标无裁切。
5. 表格、图片、引用上下标及跨页位置清楚，页面无重叠、截断或意外大空白。
6. 附录若原稿已有，全文保留并置于致谢前；最后一部分是致谢。

在报告中写明“使用什么引擎、实际看了哪些页、哪些项目没做”。只有结构检查时用“结构检查完成”；只看部分页面时明确页号。最终交付文件应与已验收文件哈希一致，修改后重新检查受影响页面及其后续分页。
