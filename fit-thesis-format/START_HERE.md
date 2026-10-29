# 从完整技能包开始

完整包包含SKILL.md、agent-tools.json、scripts、references、assets及tests。只有入口Markdown无法执行排版。使用本次交付ZIP及对应校验清单，不把旧公开版本下载地址当成新版，也不要混用版本文件。

解压到新目录，确认原稿已成为真实可读文件；先读[SKILL.md](SKILL.md)。Python 3.10+以及requirements.txt和requirements-render.txt所列依赖可使用宿主管理的环境。不把用户名、私有运行时路径写死到调用中。

在技能根目录调用：

```text
python scripts/thesis_job.py --request request.json
```

request.json使用实际绝对路径：

```json
{"action":"run","source":"原稿.docx的绝对路径","work_dir":"尚不存在的工作目录绝对路径"}
```

该调用先保留原稿、实际检测Word和WPS，再执行确定性处理并返回下一步。智能体继续处理返回的语义判断、逐页观察和修复任务，不把返回中间状态当成交付。`next / submit / resume / verify`参数见[协议](references/agent-protocol.md)和agent-tools.json。

有文件读写与Python执行能力的不同智能体使用相同入口。只有文档API时，必须证明支持相同的对象保护、导出和证据核验；只有聊天能力则不能实际生成验收通过的Word。完整验收要求Word和WPS分别检查同一个最终DOCX；缺少工具准确报告缺项，不暗自上传文件。

旧版prepare和apply工具仍可用于只读审核或限定的基础格式任务，但不替代完整流程。原稿不覆盖，内部日志不插入论文；最终交付可编辑文件和简短真实说明。

开发验证：运行 `python -m unittest discover -s tests -v`。这些单元测试中模拟的Office调用不等于原生验收；原生能力必须由doctor的真实smoke测试及最终双引擎逐页检查证明。
