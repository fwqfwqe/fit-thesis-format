# DeepSeek Harness / WorkBuddy / 其他宿主

不同宿主调用同一 Python 程序。这里约定工具交互，不假设某客户端的私有 API、工具名或安装路径，也不声称未经运行的客户端已通过。

## 最小接入

宿主需要本地文件读写、Python 3.10+ 执行、返回完整 stdout JSON、实际打开本地图片这四项能力。完整页面验收另需可实际导出的 Word 和 WPS。Python 依赖在 requirements.txt 和 requirements-render.txt；优先复用宿主已配置环境，缺依赖才安装。只支持聊天的环境不能执行本技能。

1. 解压完整包，工作目录设为技能根目录；读取 SKILL.md。
2. 用文件工具写 UTF-8 request.json，参数中使用实际绝对路径：
   `{"action":"run","source":"原稿.docx的绝对路径","work_dir":"新的工作目录绝对路径"}`。
3. 执行 `python scripts/agent_cli.py --request request.json`。工具若返回运行句柄，等待该进程完成；不要因单次工具超时另启动同一任务。
4. 保存响应 JSON 和 work_dir。按下表继续，最终交付经验证的 candidate 路径。不能凭退出码或 ok 判断论文完成。

stdin 可靠的宿主也可使用 `--request -` 传同一对象。不要把 JSON 内联拼接进 PowerShell/Bash；路径、引号和反斜杠由 JSON 库编码。跨宿主接续必须能访问同一工作目录和同一版本完整包；移动目录或换版本会使旧证据失效，不能改哈希绕过检查。

## 每轮返回后做什么

| status | 宿主动作 |
|---|---|
| semantic_review_required | 打开 next.inspection_file 指向的完整 JSON，确认全部待判对象。next.inspection 仅统计与最多30条完整样例，不能替代全文。读取 job.json 的 decisions 合并修改，用 next.submit_request 提交；decision_template 是空骨架。 |
| visual_review_required | 打开每个引擎 page_manifest 指向的实际图片，填 review_template。问题页记 fail 和具体定位；完成后 submit review。不能用提取文字或生成图片代替看图。 |
| correction_required | 读取 repair_plan.groups 和 next.evidence；批量修可修格式和能由原稿核实的问题，合并来源绑定决定后 submit。缺事实单列，不终止其余修复。 |
| capability_required | 已完成的修改和静态审计保留；修复真实工具能力后 resume。不能假冒另一办公软件，也不反复探测未变化的环境。 |
| execution_failed | 读 error 及对应阶段报告；修复具体执行原因后 resume，不新建 run 绕过失败。 |
| no_progress | 停止提交相同决定，诊断缺失修复能力或错误假设；有实质修正后再 submit。不能伪造 pass。 |
| complete | 还须 delivery_ready=true，才交付已绑定最终哈希的可编辑 DOCX。 |

通用请求（每次保存为 JSON 文件再执行同一命令）：

stdout 有样例数量和字节上限。inspection 中的 sample_omitted_count、repair_plan.sample_omitted_counts、blocking_items_omitted 明确标记未内联条目；完整问题分别在 complete_findings / blocking_items_file / issues_file 指向的 job.json 中。样例只用于定位下一步，不等于已审完整篇。按本地文件分段读取全量证据，避免反复把几百 KB 映射塞进模型上下文。

```json
{"action":"next","work_dir":"已有工作目录绝对路径"}
```

```json
{"action":"submit","work_dir":"已有工作目录绝对路径","kind":"decisions","decisions":"完整合并决定.json的绝对路径"}
```

```json
{"action":"submit","work_dir":"已有工作目录绝对路径","kind":"review","engine":"word","review":"实际逐页复核.json的绝对路径"}
```

WPS 复核将 engine 改为 wps。submit 会继续状态机，不需要无条件再跑一次 resume。中断后用 `{"action":"resume","work_dir":"已有工作目录绝对路径"}`；最终重验用同对象 action=verify。decisions 内 source_sha256 和定位必须来自原稿，不是刷新目录后改变的段落序号；具体修复字段见 [layout-repairs.md](layout-repairs.md)。

## 一次委托的交付边界

- 标题叠号、目录层级和旧缓存、公式编号和对齐、摘要分页、承诺书错位、可查证作者年份属于需要继续处理的问题，不能以“已记录未通过”替代修复。
- 缺英文题目仍要让 Abstract 独立起页；缺学生身份仍要排好封面；缺致谢不能耽误其他格式。未知内容不编写，不删除模拟数据等事实披露。
- 本稿书目已有唯一对应年份时先核对并来源绑定补齐，不能直接推给作者。确实缺依据的内容集中报告，不反复问。
- 一个 job、一份原稿基线、一次批量修改；有新失败才扩展检查。不能为节省时间跳过实际页面验收，也不能重复全流程来掩盖没有进展。
- 交付说明只写已修内容、实测范围与实际缺项。适配合同、单元测试、原生 Word/WPS 验证、具体客户端端到端实测是四种证据，不能互相替代。
