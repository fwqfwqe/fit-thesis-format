# 从完整技能包开始

完整包包含SKILL.md、agent-tools.json、scripts、references、assets及tests。只有入口Markdown无法执行排版。使用本次交付ZIP及对应校验清单，不把旧公开版本下载地址当成新版，也不要混用版本文件。

解压到新目录，确认原稿已成为真实可读文件；先读[SKILL.md](SKILL.md)。Python 3.10+以及requirements.txt和requirements-render.txt所列依赖可使用宿主管理的环境。不把用户名、私有运行时路径写死到调用中。

在技能根目录调用：

```text
python scripts/agent_cli.py --request request.json
```

request.json使用实际绝对路径：

```json
{"action":"run","source":"原稿.docx的绝对路径","work_dir":"尚不存在的工作目录绝对路径"}
```

该调用保留原稿，执行确定性处理、静态检查及环境探测，返回下一步。读取 `status`、`delivery_ready`、`next` 和 `repair_plan`；`ok=true` 仅表示调用成功。完整任务的唯一完成条件是 `status=complete` 且 `delivery_ready=true`。`correction_required` 必须继续修可修问题，不能直接交候选稿；缺身份等事实不抵消其余格式工作。`next / submit / resume / verify` 参数见[宿主调用与恢复表](references/host-adapters.md)。

DeepSeek Harness、WorkBuddy 等有文件和命令工具的宿主：让工具执行上述命令，把 stdout JSON 返回给智能体；用绝对路径打开 `next` 给出的文件和页面图片；在同一个 work_dir 中继续提交。不要把脚本重新翻译成提示词，不要每轮新建 run，不需要专用 API 密钥。`scripts/thesis_job.py` 是同一状态机的直接入口，可用于已有集成；旧 `prepare/apply_styles` 不是完整任务入口。

有文件读写与Python执行能力的不同智能体使用相同入口。只有文档API时，必须证明支持相同的对象保护、导出和证据核验；只有聊天能力则不能实际生成验收通过的Word。完整验收要求Word和WPS分别检查同一个最终DOCX；缺少工具准确报告缺项，不暗自上传文件。

旧版prepare和apply工具仍可用于只读审核或限定的基础格式任务，但不替代完整流程。原稿不覆盖，内部日志不插入论文；最终交付可编辑文件和简短真实说明。

开发验证：运行 `python -m unittest discover -s tests -v`。模拟 Office 调用、协议测试或 Python 跨平台 CI 不等于 DeepSeek Harness、WorkBuddy 客户端实测；原生能力由实际 smoke 测试及最终双引擎逐页检查证明，具体实测范围以该版本发布说明为准。
