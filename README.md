# docx-editor-cn

使用说明见 [SKILL.md](SKILL.md)。中文学术文档的排版参数统一定义在 `scripts/academic-profile.json`，生成入口为 `scripts/new_doc.js`，公共生成接口为 `scripts/academic_doc.js`。

```text
node scripts/new_doc.js input.md output.docx
python scripts/audit_format.py output.docx
```

默认生成电子版，不插入双面打印空白背页；只有明确要求双面打印时才启用 `duplex: true`。生成后需要更新 Word 域、渲染并逐页检查。学校模板或本次用户要求优先于默认参数。
