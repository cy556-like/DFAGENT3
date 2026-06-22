# CHANGELOG

## 2.0.0 (2026-06-21)

### 重大升级：从模板框架到可用的 8D 报告生成器

**template.json v2.0**
- 5Why: hints(候选列表) -> steps(完整连贯预填路径，含答案和证据)
- 6M: 纯文本 -> 数组格式 + judgment字段(根本原因/贡献因子/排除/待确认)
- 新增 root_cause_summary: 预填RC1/RC2
- 新增 verification: 预填验证结论(含具体数据)
- 5套模板全部重写

**generate_8d.py 视觉升级**
- 章节标题: 深蓝底 -> 黄色底+深蓝字, 视觉分层
- Sheet: 单 Sheet，标签颜色 003366 深蓝
- D4生成器: 兼容新旧template格式
- Word同步更新

**SKILL.md**
- 新增第5条价值主张: 真实行业参考资料
- 新增v2.0 template.json数据结构文档
- 新增web_fetch知识增强搜索指引
- 新增YAML frontmatter(name+description)

**references/**
- 基于Quality-One 8D方法论补充真实行业实践内容

---

## 1.0.0 (2026-06-20)

### 初始版本
- 完整目录结构: SKILL.md + templates/ + references/ + scripts/
- 5套行业模板
- scripts/generate_8d.py (xlsx+docx双输出)
- references/: 8d_guide.md / 5why_examples.md / fishbone_guide.md
