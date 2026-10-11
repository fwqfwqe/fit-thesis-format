# 福州理工学院论文格式 · 通用Agent版

**v5** 将学校论文格式规则和执行工具做成独立文件包。能读取文件并调用代码或文档工具的AI/Agent，都可以接入这套流程；无需绑定某个模型、账号或专有插件。排版仍按福州理工学院经济管理学院要求，研究内容保持原样。

[下载 v5 通用包](dist/fit-thesis-format-v5.zip) · [AI读取入口](fit-thesis-format/SKILL.md) · [通用执行协议](fit-thesis-format/references/agent-protocol.md) · [函数调用参数](fit-thesis-format/agent-tools.json)

## 工作流程，用最简单的话说

1. **收论文、留原稿。** 保存原文件，在副本上工作。
2. **检查格式。** 对照学校要求，找出哪里要改。
3. **只改格式。** 调整标题、字体、页眉页码、公式、图表和引用样式。
4. **核对内容。** 对照原稿，检查文字、数字、公式和参考文献有没有被改动。
5. **逐页看一遍。** 检查错页、挤压、截断和编号。
6. **交付。** 给你可编辑的Word和问题说明。

只要求审核时，执行检查并给问题清单。完整排版还需要相应的文件修改、DOCX导出和逐页查看工具；能力不足的部分会明确列出。

## 给任意AI使用

解压完整包，让AI能访问其中的文件，再把论文DOCX交给它，并发这段话：

> 先读取这个文件夹里的SKILL.md，再按里面的流程处理我的论文。只改格式，研究内容不改；先审核，再修改副本，核对内容，逐页检查，最后给我可编辑的Word和修改说明。如果我只要求审核，就只输出问题清单。

支持skill的Agent可以把`fit-thesis-format`文件夹交给自己的skill加载器；不支持skill机制的AI，直接读取Markdown说明并调用自己的文件/代码/文档工具即可。`agents/openai.yaml`只是可选界面配置。

原有Codex用户仍可把文件夹放入`~/.codex/skills/`，在新对话使用`$fit-thesis-format`。

## 通用工具调用

核心脚本适用于Python 3.10+。本次实测Python 3.12.14，依赖为lxml 6.1.1和python-docx 1.2.0：

```text
python -m pip install -r fit-thesis-format/requirements.txt
python fit-thesis-format/scripts/agent_cli.py --request task.json
```

`task.json`审核示例：

```json
{
  "action":"audit",
  "source":"论文.docx",
  "report":"work/格式审核.json",
  "markdown":"work/格式审核.md"
}
```

入口返回一个JSON结果，所有相对路径基于调用进程的工作目录。函数调用型AI将[工具schema](fit-thesis-format/agent-tools.json)按所在框架注册，再把参数转发给该入口，也可通过标准输入传入。接入示例和完整动作见[通用执行协议](fit-thesis-format/references/agent-protocol.md)。

工具包括环境探测、原稿清单、静态审核、样式修改、引用格式修改、页眉页脚布局、内容核验和编号/公式结构检查。Agent仍需阅读原文、确定段落角色和授权范围，不能把基础脚本当成所有论文的一键转换器。

## 渲染与手调

可以使用已有Word、文档API或其他DOCX渲染工具。包内提供独立的Windows Word只读导出脚本和可选PDF转图片脚本，调用与依赖见[便携渲染说明](fit-thesis-format/references/portable-rendering.md)。不同引擎的分页与公式兼容性须按实际结果检查。

Word中可通过“开始→样式→修改”统一手调。数学文字与公式编号也有各自的字符样式。静态脚本成功不代表已经逐页看过，未完成视觉核验会明确报告。

## 保留的格式要求

学校封面和诚信承诺书；四级自动编号及两级目录；真实奇偶页眉与连续页码；可编辑公式、正确上下标和完整“公式（x-x）”；独立公式不用表格、文本框或边框；正文引用编号上标，作者年份保持基线；已有附录完整保留并放在最后的致谢之前。

研究文字、数据、公式含义、引用条目和附录代码不改写。目录、编号表示及已授权学校组件的转换单独记录。只审核时不修改论文，缺少内容时不自动补写。

## 文件

- `fit-thesis-format/`：完整通用skill、工具schema、脚本、要求、范本与依赖文件。
- `dist/fit-thesis-format-v5.zip`：与当前源文件逐项校验的通用包。
- `package-manifest.json`：包内文件与ZIP的SHA-256。
- `tests/`：通用工具入口的行为验证。

仓库公开保存skill源码、安装包、学校范本和格式证据；没有上传用户研究论文、数据集或本地运行日志。
