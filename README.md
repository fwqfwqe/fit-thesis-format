# 福州理工学院论文格式 · v10

把[完整技能包](https://github.com/fwqfwqe/fit-thesis-format/releases/download/v10.0.0/fit-thesis-format-v10.zip)与原始论文 DOCX 交给具备文件、执行和看图工具的智能体，按 [START_HERE.md](fit-thesis-format/START_HERE.md) 启动。完整排版使用 `thesis_job.py`，会保存原稿、处理全文、返回需要判断的事项，并持续推进到验收。

v10 将规则、确定性修改和验收串成可恢复的任务。封面承诺书、中英文摘要、四级标题、自动目录、页眉页码、图表公式、引用和文后顺序都进入同一流程。完整委托包含内部修复与复验，智能体不能把生成清单或中间候选稿当作完成。

本工具适用于福州理工学院经济管理学院现有规范，依据附带的匿名化范本及 2026-10-29 使用者确认要求；属于非官方工具，不表示学校发布或认证。论文内容、数字、研究结论及原有对象受保护。规范来源与适用边界见 [规则目录](fit-thesis-format/references/rules-v10.json)。

## 直接交给智能体

> 使用本仓库的完整 fit-thesis-format 技能处理我的原始论文 DOCX。读 START_HERE.md 和 SKILL.md，用 thesis_job.py 启动完整任务。保留研究内容，在副本中完成全文格式修改。自行处理语义审核、实际页面复核和必要修复，直到最终同一 DOCX 通过 Word 与 WPS 验收，再交付可编辑 Word 和简短修改说明。缺工具或缺事实时准确说明，不把中间状态当作完成。

不同宿主使用同一 CLI/JSON 接口，不依赖特定 AI SDK。是否能完成取决于当前会话的实际工具：纯聊天环境无法修改文件；具备 Python 只能证明静态处理，完整验收还需实际 Word、WPS 导出及全页看图。未逐一测试所有智能体品牌。

## 获取和运行

可直接下载 Release 中的 ZIP；也可下载 [bootstrap.py](bootstrap.py) 后执行：

```text
python bootstrap.py --dest package-new
python -m pip install -r package-new/fit-thesis-format/requirements.txt -r package-new/fit-thesis-format/requirements-render.txt
python package-new/fit-thesis-format/scripts/thesis_job.py --request request.json
```

`package-new` 必须是新目录。`bootstrap.py` 取得当前 [校验清单](package-manifest.json) 和对应 ZIP，验证所有文件后解压，不覆盖已有目录。Python 需 3.10 或更高版本。`request.json` 使用真实绝对路径：

```json
{"action":"run","source":"原始论文.docx的绝对路径","work_dir":"尚不存在的工作目录绝对路径"}
```

根据返回状态执行 `next / submit / resume / verify`。参数以 [agent-tools.json](fit-thesis-format/agent-tools.json) 和 [完整执行合同](fit-thesis-format/references/v10-execution.md) 为准。读到 `semantic_review_required` 或 `visual_review_required` 后，宿主应继续处理待办；这些状态不是完成，也不是自动要求用户再次授权。

支持技能目录的宿主可安装完整 `fit-thesis-format/` 文件夹；只复制 `SKILL.md` 会缺少模板、规则和程序。旧 `prepare`、`apply` 工具保留给只读审核或局部格式任务，不代替 v10 完整入口。

## 如何确认完成

- 覆盖全部段落、表格和文档区域，未知语义交由宿主结合上下文确认。
- 每一步转换与源文件绑定，可重放；核对研究文字、公式、媒体、修订和注释等保护对象。
- 同一最终文件在 Word、WPS 分别导出，校对目录与实际页码。
- 宿主查看两套全部页面，记录具体观察；只允许复用图像哈希完全相同的真实检查记录。
- 修复编号延续、图表裁切、续表、空白页等问题后重新验证。只有 `delivery_ready: true` 才称完整验收。

任务目录和论文不上传到本仓库。发布包仅含通用工具、匿名化模板、规则和合成测试；范本保留的格式批注仅作证据，复用学校组件时使用无批注的净化模板。

## 开发和验收

```text
python -m unittest discover -s fit-thesis-format/tests -v
python -m unittest discover -s tests -v
```

第一组检查 v10 流程与保护边界，第二组保留旧接口兼容回归。CI 使用合成文档检查跨平台程序；它不能替代实际 Office 导出和逐页复核。原生验收范围见 [VALIDATION.md](VALIDATION.md)。
