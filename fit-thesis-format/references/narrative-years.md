# 来源绑定的作者年份补齐

当正文“作者……[n]”缺作者后年份，先用本稿完整书目核对。扫描器提示 `missing_year` 不是自动批准；宿主确认唯一文献和出版年份后，在完整 decisions 中加入 `year_insertions`：

```json
{
  "source_sha256": "原稿SHA256",
  "year_insertions": [{
    "pid": "p000123",
    "expected_text": "Vial的综述讨论转型[5]。",
    "author_start": 0,
    "author_end": 4,
    "expected_author": "Vial",
    "reference_pid": "p000456",
    "reference_expected_text": "[5] VIAL G. Digital transformation[J]. Journal, 2019.",
    "year": "2019",
    "reviewed": true,
    "reason": "已核对本稿第5条书目，作者和出版年份唯一对应"
  }]
}
```

位置来自原稿，字符偏移按 Unicode 码点、左闭右开。整段、作者范围、书目整段必须逐字匹配。程序再次检查唯一 missing_year 候选、同句单个现有 `[n]`、书目编号不重复、书目只有一个明确年份；存在多个年份、多个文献标记、DOI/URL 数字或复杂对象时拒绝，不替宿主猜年份。

只在完整作者后插入基线“（年份）”，保留论述、作者写法和引文编号。每段记录 `narrative_year_insertion` 及原始证据，完整性验证从原稿独立核对。程序在正文空白清理、作者展示规范化之前执行；不能把已经改写后的段落定位当原稿定位。无依据的年份仍列待核实，不新编参考文献。
