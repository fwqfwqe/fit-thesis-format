# 福州理工学院经管论文格式 skill

`fit-thesis-format` 用于福州理工学院经济管理学院毕业论文的格式审核与排版。当前打包版本为 **v4**，研究内容保持原样，输出可在 Word 样式栏手调的 DOCX。

## 功能

- 只读格式审核，按位置列出现状、学校要求和改法，并区分需人工复核的项目。
- 学校封面与诚信承诺书、四级自动编号、可更新两级目录、真实奇偶页眉及连续页码。
- 可编辑数学公式、正确上下标及完整“公式（x-x）”编号；独立公式以普通段落排版，不使用表格、文本框或边框。
- 正文引用编号上标，作者与年份保持基线；字体、间距、三线表等按学校范本核查。
- 最新个人排版偏好：致谢置于全文最后；已有附录完整保留并移到致谢之前，没有附录时不增写。

正文论述、数据、公式含义、引用条目及附录代码不得改写。自动编号、目录和已授权学校组件的格式性转换单独记录核验。学校范本原批注与个人新增要求有差异时，以已明确的用户要求优先；源证据保留原样。

## 安装到 Codex

下载 [v4 ZIP](dist/fit-thesis-format-v4.zip)，将其中整个 `fit-thesis-format` 文件夹放到自己的 `~/.codex/skills/` 目录。

Windows 目录通常为 `%USERPROFILE%\.codex\skills\fit-thesis-format`，其下直接包含 `SKILL.md`。已安装旧版时先备份旧文件夹，再换入新版。新开 Codex 对话以加载 skill。

## 使用

将论文 DOCX 交给 Codex，并说明：

> 用 $fit-thesis-format 只审核这篇论文的格式，列出位置、现状、要求和改法，文章内容不要改。

或：

> 用 $fit-thesis-format 按学校格式排版，文章研究内容不要改，输出可手调样式的 Word。

脚本可独立运行静态审核：

```powershell
python -m pip install -r requirements.txt
python fit-thesis-format/scripts/format_audit.py "论文.docx" --json "格式审核.json" --markdown "格式审核.md"
```

高级映射、严格文字保护及数学编号调用见 [skill 入口](fit-thesis-format/SKILL.md) 和其链接参考。

## 运行与核验

本次验证环境：Python 3.12、lxml 6.1.1、python-docx 1.2.0；其余依赖为 Python 标准库。Codex 排版流程使用 documents skill 的文档依赖及渲染能力。Microsoft Word 用于原生动态域刷新、公式编辑和逐页复核。

静态审核不等于最终格式合格证；完整验收还须更新目录和域、检查公式含义，并实际查看每一页。Word 可通过“开始→样式→修改”统一手调；数学文字和公式编号也有各自字符样式。

## 仓库内容

- `fit-thesis-format/`：可安装的完整 skill、脚本、格式要求、范本和证据。
- `dist/fit-thesis-format-v4.zip`：同版本安装包，内容与 skill 源码逐项一致。
- `package-manifest.json`：文件大小与 SHA-256 校验值。
- `requirements.txt`：本次核验的 Python 依赖版本。

本仓库为私有备份，保留学校范本及格式证据原件；没有上传用户研究论文、数据集或本地运行日志。
