# DocMind CLI

**专治 Word 排版疑难杂症** — 本地离线 · 深度 OOXML · AI Agent 实时协作

---

## 安装

```bash
git clone https://github.com/baolongzhanshenNo1/docmind-cli.git
cd docmind-cli
pip install .
```

## 用法

```bash
# 对论文按学校规范模板排版（格式修正，不改文字）
docmind format my_thesis.docx --spec university_spec.docx

# 提取规范模板的格式规则（字体/边距/页眉/页码）→ YAML
docmind analyze university_spec.docx --output spec_rules.yaml

# 奇偶页强制（空白页、页眉重整）
docmind enforce my_thesis.docx --libreoffice /path/to/soffice.exe
```

## 命令参考

### `docmind format <论文.docx> --spec <规范.docx> [--output <out.docx>]`

自动分析规范模板的格式规则 → 发现文档章节结构 → 差异对比 → 应用修正。

**示例输出：**

```
┌──────────────────────────────────────────┐
│  DocMind — 专治 Word 排版疑难杂症        │
└──────────────────────────────────────────┘

  目标文档 → my_thesis.docx
  规范模板 → university_spec.docx

  正在排版: my_thesis.docx… ━━━━━━━━━━━━━

   📄 修正项  61
   🔄 残留    0
   ⚠️  诊断    0
   📦 输出    my_thesis_formatted.docx

✅ 排版完成
```

### `docmind analyze <规范.docx> [--output <template.yaml>]`

提取规范模板的字体/边距/页眉/页码格式规则，生成结构化 YAML。

### `docmind enforce <文档.docx> [--libreoffice ...] [--sections ...]`

对已有 docx 执行奇偶页强制：扉页空白页插入、各章节从奇数页开始、页眉重整。

> 需要 [LibreOffice](https://www.libreoffice.org/) 用于 PDF 渲染与页码检测。

---

## 特性

- 📐 **排版修正，非生成** — 规范是尺子，量哪里不对修哪里，一字不改
- 🏠 **本地离线** — 全部在本地运行，文档不用上传
- 🧩 **深度 OOXML** — 直接操作 OOXML ZIP 层级；页眉/页脚/分节符/页码/字体完全控制
- 📄 **支持** — 论文（thesis）、公文（gov_doc）、技术文档（tech_doc）、合同审查（contract）、文档翻译（translate）
- 🔧 **可扩展** — 基于可组合 Steps 架构，加新格式模板无需改引擎
- 🎨 **Rich 终端渲染** — 彩色状态、表格结果、进度提示

---

## 项目结构

```
docmind-cli/
├── cli_new.py               # CLI 入口（click + rich）
├── docmind/
│   ├── tools/               # 5 个工具（thesis / gov_doc / tech_doc / contract / translate）
│   ├── engine/              # 核心排版引擎（writer / fixer / enforce / spec_reader）
│   └── steps/               # 可组合排版步骤
├── pipeline/                # Agent v3 六步闭环 Pipeline
├── templates/               # YAML 模板（thesis.yaml 等）
└── tests/                   # 单元测试（pytest）
```

---

## 依赖

- Python ≥ 3.11
- LibreOffice（仅 `enforce` 命令需要）
- 核心 Python 包：`python-docx`, `PyYAML`, `lxml`, `click`, `rich`, `pymupdf`, `httpx`

---

## 开源协议

MIT License — 见 [LICENSE](LICENSE)

---

## 路线图

- [ ] 支持更多领域模板（law、gov、tech）
- [ ] 内置 AI Agent 对话式排版（需要 LLM API key）
- [ ] 企业版管理后台
- [ ] Word 插件

---

**DocMind** — *专治 Word 排版疑难杂症*
