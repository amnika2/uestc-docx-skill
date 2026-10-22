# 默认课程设计报告封面

来源为用户指定的《课程设计报告—Adam 最优化算法及其在逻辑回归分类中的应用》封面。`assets/uestc-course-cover.docx` 只保留封面及其图像，不包含报告正文；六个变量已替代原报告的信息。

固定部分为电子科技大学中英文校名、校徽、课程设计报告中英文标题及原有信息栏布局。保留原来的封面字体、字号、间距和下划线；封面不套用正文样式，不显示页眉页码，封面之后只用一个下一页分节符。

| 字段 | 内容 |
|---|---|
| `title` | 题目，默认使用文档标题 |
| `college` | 学院 |
| `major` | 专业 |
| `studentId` | 学号，按字符串传入 |
| `author` | 作者姓名 |
| `course` | 课程名称 |

未知信息留空。文档属性的作者不自动填入封面。`assets/cover-fields.json` 只保存原信息栏的宽度，用于替换内容时保留下划线长度和居中效果。

## 公共接口

```javascript
const doc = createAcademicDocument({
  title: '梯度下降算法及其在机器学习中的应用',
  coverFields: {
    college: '自动化工程学院',
    major: '控制科学与工程',
    studentId: '本次学号',
    author: '本次姓名',
    course: '最优化理论与应用'
  },
  front, body
});
const buffer = await packAcademicDocument(doc);
```

省略 `cover` 时自动使用模板；`cover: []` 或 `coverTemplate: false` 表示不生成封面；`cover: [自定义段落]` 使用用户另行指定的封面。必须通过 `packAcademicDocument()` 输出，以插入原始封面图像、布局和独立样式。

Markdown 的一级标题 `# 文档题目` 自动成为封面题目。其他信息可放入一个 JSON 文件，并作为第三个位置参数传入：

```text
node scripts/new_doc.js input.md output.docx cover-fields.json
```

JSON 顶层直接使用表中字段，不包含正文或样式。没有该文件时，题目自动填写，其余信息栏留空。较长题目或超出原信息栏宽度的字段，需要检查渲染结果并按实际内容调整，不能缩小正文字号。

## 使用范围

这是新建文档的默认封面。编辑已有文档时，不自动覆盖原封面。用户要求其他学校、其他报告类型、其他封面或无封面时，使用该次要求。
