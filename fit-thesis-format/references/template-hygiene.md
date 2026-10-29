# 模板批注与修订检查

学校原范本包含编写时留下的教学批注。这些文字是格式证据，不是学生论文的一部分。删除范文正文只改变 `document.xml` 的节点，不能确保独立批注部件、页眉页脚等区域的锚点、关系和 `[Content_Types].xml` 登记一起消失。

## 默认使用净化模板

- [原范本](../assets/fit-economics-management-template.docx)及 [source-comments.json](source-comments.json)保留原样，仅供查证。
- [净化模板](../assets/fit-economics-management-template-clean.docx)用于复用学校组件，只去除历史批注。它仍含范文和示例信息，不能直接当作空白模板或合格成稿。
- 对已有论文从原稿副本定点修改，按需要迁入净化模板中的组件，不从模板重建整篇论文。原稿真实作者、导师批注和修订默认保留；最终稿不新增审计批注，审核意见写在外部说明中。

需要重新生成净化模板时，在完整技能目录内运行：

```text
python scripts/template_hygiene.py clean-template assets/fit-economics-management-template.docx work/template-clean.docx --report work/template-clean-report.json
```

输出和报告必须是新路径。命令仅接受 SHA-256 为 `83dac94cb138ab7b46fb100d9ac05c6e1ab8c7dec986ebefd2e5ca22bf359eac` 的学校原范本，保持输入不变；**不能将用户论文交给此命令清空批注**。报告记录真实删除数量和前后检查结果，不代表完成页面视觉核验。

## 检查原稿与当前候选稿

对每个需要核对的 DOCX 运行只读检查：

```text
python scripts/template_hygiene.py inspect work/candidate.docx
```

检查全部 XML 部件中的批注与锚点、批注相关部件和关系、内容类型登记、孤立引用及真实修订。`format_audit.py` 同时调用这项检查，交付检查据审核问题判断是否仍有阻断项。读报告的具体问题，不能把命令退出成功当成“零批注”或“格式合格”。

与已记录模板批注文字相同的项是需要核对来源的证据，不能单凭文字相同就删除用户批注。确认来自模板的历史批注应在新的副本中定点清理，并保留清理前文件与变化记录；用户真实批注须按其明确要求处理。来源不明时先保留并报告，不能用“全部删除”掩盖无法归因的问题。

只删除已确认的批注节点及其依赖。批注引用所在的 run 可能同时含域或图片，不能因它没有 `w:t` 就删除整个 run。修改关系和内容类型登记须按 XML 节点解析，不用容易被 `application/...` 等属性值打断的正则表达式。清理后再次检查批注、锚点、关系、部件和登记，并确认无关内容未变。

## 修订必须精确计数并保护

按 XML 命名空间和完整本地名识别 `w:ins`、`w:del`、移动及属性修订等节点。`w:insideH` 是表格内部横边线，`w:instrText` 是域代码；用字符串前缀 `<w:ins` 搜索会把它们误当插入修订。没有 `w:ins` 或 `w:del` 也不能代替其他修订类别的检查。

去除模板批注与处理真实修订是两项不同操作。不得顺带接受插入、删除删除内容或改动其他修订。报告应分别列出实测批注数量、锚点和登记残留、各修订类别数量；没有检测到什么，就只说明对应检查结果，不编造原先数量或把模板总数等同于某份成稿的继承数量。
