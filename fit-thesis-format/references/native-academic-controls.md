# 新建学院实例的原生 Word 控件

本参考对应 `scripts/academic_numbering.py` 和 `scripts/validate_academic_controls.py`。这是**用户明确要求新建格式实例时的构建工具**，不能在“原论文只改格式、任何内容不变”模式下直接重输标题、公式或引用。原文转换须另外建立绑定输入 SHA-256 的人工审核映射，明确原编号、原表达式与目标结构，分别核验可见文本、编号和数学语义；没有该映射不得自动转换。尤其不得把公式图片或 MathType 对象猜写成 OMML。

## 学院四级标题

四级显示依次为“一、绪论”“（一）研究背景”“1.国内研究现状”“（1）……”。新建示例的标题文本只写标题词语，编号由同一个 Word 多级列表产生。样式绑定 `FIT_chapter`、`FIT_section`、`FIT_subsection`、`FIT_subsubsection`，大纲级别为 0、1、2、3；结论与展望及参考文献等后置标题使用各自不编章号的样式。

实际实现同时写明标题段的 `numPr/numId/ilvl`，并把每级绑定到对应命名样式；不能只给样式设置 `numPr` 就宣称 Word 已显示编号。新的 `abstractNum` 必须位于全部具体 `num` 之前。较低级别使用默认重启规则，不写显式 `lvlRestart`：在本机原生 Word 测试中，后者导致低级编号被读为空格式，XML 表面看似正确但页面不显示编号。

2026-10-11 原生 Word 验收得到：第一章“一、”，同章一级节“（一）”“（二）”，二级节“1.”“2.”，三级节“（1）”“（2）”；第二章为“二、”，其首个一级节重置“（一）”。验收应读取各标题实际 `Range.ListFormat.ListString`，不能仅检查 OOXML 标签。

## 原生公式与数学变量

`append_math` 输出可在 Word 公式编辑器中编辑的 OMML。`fraction`、`subscript`、`superscript`、`sub_superscript`、`summation`、`brackets`、`square_root`、`math_function` 可组合。上下标使用真正的数学结构，不用缩小普通文字或统一 `vertAlign` 代替。正文中的数学变量也调用 `append_math`。

新建实例采用可见字符样式“论文-数学变量”（`FIT_math`）管理数学 run 的字体、字号。Cambria Math 12 磅是实例选用值，**不是学院明确规定值**。默认不在每个 run 写死字体字号；右键修改该样式应同时作用于公式和行内变量。原生 Word 测试中，全部实际数学字符由 12 磅同步变成 16 磅；OMML 的内部控制字符仍继承段落字号，整体 `OMath.Range.Font.Size` 返回混合值属于正常现象，应检查实际数学字符，不能把混合值误判为样式不生效。

独立公式使用真正的普通段落，**不使用表格、文本框、框架或段落边框**。段落左对齐，首行和左右缩进为零；以页面正文宽度的一半设置居中制表位，以正文右边界设置右对齐制表位。内容顺序为“制表符、原生 OMML、制表符、完整动态标签”。以正文宽 9355 twips 为例，制表位分别为居中 4678、右对齐 9355。这样在 Word 中也不会出现布局表格的虚线网格。

段落套“论文-公式”，至少 20 磅并保持段内完整，不能用固定 20 磅裁切分式或求和上下限。完整“公式（x-x）”标签的全部 run 套可见字符样式“论文-公式编号文字”（`FIT_equation_label`），可单独手调；保留的旧“论文-公式编号”段落样式不用于新建这种单段公式。公式字体字号仍由“论文-数学变量”字符样式控制。宽公式必须渲染检查是否与右侧标签重叠或换行，不能为凑版删除数学项，也不能悄悄改回布局表格。

## 公式动态编号及限制

章标题末尾有两个不可见的原生域：`SEQ FITChapter \h` 计数阿拉伯章号，`SEQ FITEquation \r 0 \h` 重置本章公式序列。公式标签由可见词语与真实域构成：

```
公式（{ SEQ FITChapter \c }-{ SEQ FITEquation }）
```

域的缓存数值只是打开时的显示缓存，不能把它当作手打编号。完整标签设置书签，正文以 `REF 书签名 \h \* CHARFORMAT` 引用，因此显示“公式（2-1）”全标签，同时继承引用所在正文的字号。新建文档可在校验副本或授权输出中更新域，再按顺序检查序列与引用；保存后还要重开、渲染核对。

中文章号不能直接用 `STYLEREF` 加 Arabic 开关转换为阿拉伯数字。原生 Word 实测 `STYLEREF "论文-章标题" \n \* Arabic` 仍显示“一、”，`\s \* Arabic` 或 `\n \t \* Arabic` 仍显示“一”。因此本助手保留隐藏 SEQ 方法，而不宣称任意新套章标题样式就会联动公式章号。

**Word 中新增章的操作约束**：复制一个完整现有章标题段落（包括其隐藏域）再修改标题词语，或通过 `ThesisNumbering.heading` 创建，然后更新域。只新建普通段落再套“论文-章标题”样式，章标题列表会正确变为“三、”，但没有隐藏 SEQ 标记时公式章号仍可能留在“2”。每个章标题必须有这两个隐藏标记，结构验收脚本会检查。删除或移动完整章标题段落后也应重新更新域并验收公式。

## 真实奇偶页眉页脚

`configure_running_furniture` 必须在全部节创建完后调用，计划覆盖所有节。它设置真实 `evenAndOddHeaders` 和每节独立的 `default`/`even` 页眉及双 PAGE 页脚，关闭“首页不同”，断开前节链接。奇数正文页眉是本稿题目，偶数为“福州理工学院本科生毕业论文（设计）”；页码是真实 PAGE 域。封面、承诺书没有页码；摘要目录连续大写罗马，正文从阿拉伯 1 开始，后置节续号。

不得使用单一页眉中的条件 IF 冒充 Word 的奇偶页不同设置。对真实奇偶页眉，Word 在页码重置为奇数、下一物理页却为偶数时，可能生成无编号的单双面对齐空白页。本机 `nextPage` 和“连续节加下一段段前分页”测试都出现这种对齐行为。应以显示页码确认奇偶页眉，检查空白页是否正是该行为，不盲目删节、关掉奇偶页设置或改成 IF。可接受范围以用户/当前流程的明确要求为准，交付说明应诚实描述必要的对齐页，不能把多余空段产生的空白页解释成对齐页。

## 调用示例与验收

```python
from docx import Document
import academic_numbering as A

doc = Document()                     # 仅新建授权实例
A.install_styles(doc)
controls = A.ThesisNumbering(doc, authorized_new_document=True)
controls.heading("绪论", level=0)
controls.heading("研究背景", level=1)
expression = A.math_sequence(
    A.subscript("C", "t"), "=",
    A.fraction(A.subscript("D", "t"), A.subscript("Q", "t")),
)
paragraph, label = controls.equation(expression, bookmark="EqCost")
p = doc.add_paragraph("计算关系见", style="论文-正文")
A.add_equation_reference(p, "EqCost", label)
p.add_run("。")
A.configure_running_furniture(doc, [{"kind": "body", "start": 1}], "实例论文题目")
doc.save("new-example.docx")
```

`validate_academic_controls.py new-example.docx --report check.json` 检查真编号结构、数学对象/样式、完整公式标签域、隐藏章标记及真实奇偶部件，新增严格检查公式位于有两个制表符的普通段落，中心/右侧制表位正确，且公式不在表格或文本框内、不带边框/框架。它不替代 Word 验收：更新并检查四级真实 `ListString`、公式字段结果及 REF，测试普通正文和数学字符样式传播，核对物理页与显示页码，再逐页看 PDF 中标题、公式居中、完整编号、上下标、裁切和奇偶页眉。
