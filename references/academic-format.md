# 默认中文学术排版

本默认方案采用用户提供的《本科毕业设计（论文）管理办法》附件二中的通用排版要求（渲染后的第 30–41 页），保留用户明确指定的两个例外：公式单倍行距，代码采用等宽字体。它不自动要求报告内容变成毕业论文。

## 参数表

唯一参数文件：`../scripts/academic-profile.json`。

| 内容 | 中文 / 英文数字 | 字号 | 行距 | 段前 / 段后 | 缩进与对齐 |
|---|---|---|---|---|---|
| 正文与摘要 | 宋体 / Times New Roman | 12 pt | 固定 20 pt | 0 / 0 pt | 首行 2 汉字，两端对齐 |
| 章标题 | 黑体 / Times New Roman | 15 pt | 固定 20 pt | 24 / 18 pt | 无首行缩进，居中，新页 |
| 一级节标题 | 黑体 / Times New Roman | 14 pt | 固定 20 pt | 18 / 6 pt | 无首行缩进，左对齐 |
| 二级节标题 | 黑体 / Times New Roman | 14 pt | 固定 20 pt | 12 / 6 pt | 同上 |
| 三级节标题 | 黑体 / Times New Roman | 12 pt | 固定 20 pt | 12 / 6 pt | 同上 |
| 图题 | 宋体 / Times New Roman | 10.5 pt | 固定 20 pt | 6 / 12 pt | 居中，图下，不额外加粗 |
| 表题 | 宋体 / Times New Roman | 10.5 pt | 固定 20 pt | 12 / 6 pt | 居中，表上，不额外加粗 |
| 表中文字 | 宋体 / Times New Roman | 10.5 pt | 固定 20 pt | 0 / 0 pt | 无首行缩进 |
| 公式 | Cambria Math，编号 Times New Roman | 12 pt | 单倍，自动高度 | 6 / 6 pt | 公式居中，编号右对齐 |
| 目录 | 章黑体，其余宋体 / Times New Roman | 12 pt | 固定 20 pt | 0 / 0 pt | 层级缩进，页码右对齐 |
| 参考文献 | 宋体 / Times New Roman | 10.5 pt | 固定 20 pt | 0 / 0 pt | 悬挂缩进，单条保持完整 |
| 页眉 | 宋体 / Times New Roman | 10.5 pt | 固定 20 pt | 0 / 0 pt | 居中，0.5 pt 单横线 |
| 页码 | Times New Roman | 9 pt | 单倍 | 0 / 0 pt | 底端居中 |
| 脚注 | 宋体 / Times New Roman | 9 pt | 单倍 | 0 / 0 pt | 两端对齐，悬挂缩进 |
| 代码 | 中文宋体，英文数字 Consolas | 10.5 pt | 固定 14 pt | 0 / 0 pt | 左对齐，保留源文件缩进 |

相邻标题之间没有文字时，第二个标题的段前设为 0。不要插入空段落来实现上述间距。默认封面复用用户指定的课程设计报告模板，保留其字体、字号、间距和图像位置，见 [封面使用说明](default-cover.md)。

页面为 A4（11906×16838 DXA），四边 1701 DXA（3 cm），页眉页脚距边界 1134 DXA（2 cm），正文宽度为 8504 DXA。更换边距时，表格和公式宽度重新计算。表格默认 0.5 pt 网格，三线表可显式选择。

## 公共生成接口

```javascript
const {AcademicBuilder, createAcademicDocument, packAcademicDocument} = require('./academic_doc');
const b = new AcademicBuilder();
const doc = createAcademicDocument({
  title: '具体文档题目',
  coverFields: {college: '', major: '', studentId: '', author: '', course: ''},
  front: [
    {title: '摘要', children: [b.frontHeading('摘 要'), b.body('摘要内容')]},
    {title: '目录', children: [b.frontHeading('目 录'), b.toc()]}
  ],
  body: [
    b.heading('引言', 1), b.heading('研究背景', 2), b.body('正文内容'),
    b.formula(String.raw`F(\theta)=\frac{1}{n}\sum_{i=1}^{n}\ell_i(\theta)`),
    b.caption('实验参数', 'table'), b.table(['参数', '数值'], [['学习率', '0.01']]),
    b.unnumbered('参考文献'), b.ref('作者. 题名[M]. 出版地: 出版者, 年份.'),
    b.unnumbered('附录 A 源代码', {appendix: 'A'}), b.code('print("hello")')
  ]
});
packAcademicDocument(doc).then(buffer => require('fs').writeFileSync('output.docx', buffer));
```

`heading(text, level)` 不接收手工编号，级别为 1–4。`caption(description, 'figure'|'table')` 不接收图表编号。`formula()` 接收 LaTeX 或原生 docx 数学组件数组，默认分章自动编号。`image(file,width,height,description)` 的宽高为 96 dpi 显示像素，超过正文宽度时同比例缩放；调用者必须保证缩放后的图中文字清楚。`table()` 接收可选列宽，超出当前内容宽度时按比例归一化。

电子版为默认模式，使用普通下一页分节，封面、摘要、目录和正文之间不插入打印用空白背页。只有用户明确要求双面打印时才设置 `duplex: true`，用 odd-page 分节确保前置部分和正文第 1 页从正面开始；其他章仍使用普通新页。正文页眉在所有页面显示完整章号与章名，例如“第一章 引言”，参考文献和附录显示自身标题。创建文档后应更新域并渲染验证前置罗马页码、正文从 1 开始、页眉章号与章名，以及中文章名/阿拉伯节号。

每个章节边界仅由“下一页分节符”换页，`Heading1` 不带 `pageBreakBefore`，不要在章末再添加分页符。使用 `packAcademicDocument()` 输出：它将独立分节标记合并到章末正文段落，移除边界处冗余的空段/分页符；以表格结束的章节保留 1 磅行高的必要分节段，防止继承正文 20 磅行高后溢出成空白页。保留正文 20 磅、章标题段前段后间距及连续页码。

## 输出检查

- `audit_format.py` 检查默认样式、间距、字体和分节。存在明确用户覆盖时，可以用覆盖后的参数文件校验。
- Schema 校验、格式检查与 Word 渲染分别验证结构、样式和可见排版，不能互相替代。
- 逐页看最终大小的图片，而不是仅看缩略图或只抽查首页。
- 图内标签不得在大图源文件里很清楚、插入后却过小。科学图必须按最终物理宽度绘制。
- 更新目录、题注、页码后再次检查，尤其关注图文跨页、仅一两行的孤页、未要求的空白页和过大的下半页留白。
