<h1 align="center">DocMind CLI</h1>

<p align="center">
  <strong>专治 Word 排版疑难杂症</strong> — 本地离线 · 深度 OOXML · MIT 开源
</p>

<p align="center">
  <a href="LICENSE"><img src="https://img.shields.io/badge/license-MIT-blue.svg" alt="MIT"></a>
  <a href="#"><img src="https://img.shields.io/badge/python-≥3.11-blue" alt="Python 3.11+"></a>
  <a href="#"><img src="https://img.shields.io/badge/platform-Windows%20|%20macOS%20|%20Linux-lightgrey" alt="Platform"></a>
</p>

---

## 📸 Demo 画廊

<p align="center">
  <strong>目录页 — 修前（无页码） vs 修后（罗马数字"第Ⅰ页"）</strong><br/>
  <img src="assets/toc_before.png" width="45%" alt="Before - no page number"/>
  <img src="assets/toc_after.png" width="45%" alt="After - roman numeral Ⅰ"/>
</p>

<p align="center">
  <strong>封面 — 修后（页眉已清除）</strong><br/>
  <img src="assets/cover_after.png" width="45%" alt="Cover - header cleared"/>
</p>

---

## ⚡ 快速开始

```bash
git clone https://github.com/baolongzhanshenNo1/docmind-cli.git
cd docmind-cli
pip install .

# 排版论文（最常用）
docmind format my_thesis.docx --spec university_spec.docx

# 提取规范模板规则 → YAML
docmind analyze university_spec.docx

# 奇偶页强制（需要 LibreOffice）
docmind enforce my_thesis.docx
```

---

## 📋 它能修正什么？

| 能力 | 说明 |
|------|------|
| 📐 页眉 | 正文各章独立页眉（章标题）、封面/声明清空 |
| 📄 页码 | 罗马/阿拉伯分节编号、前置页不显示 |
| 🔢 奇偶页 | 每章奇数页起始、扉页空白页 |
| 🎨 字体 | 中文宋体/标题黑体、英文 TNR（Ascii/EastAsian 分离） |
| 📏 边距 | 严格对齐规范 |
| 📑 分节符 | 拆分合并节、补断链引用 |
| ✅ 诊断 | 排版后自动诊断，0 残留 |

---

## 📦 环境

| 依赖 | |
|------|------|
| Python ≥ 3.11 | `pip install .` 自动安装 |
| [LibreOffice](https://www.libreoffice.org/) | 仅 `enforce` 命令需要 |

---

## 🧪 测试

```bash
pip install -e ".[dev]"
pytest tests/ -q
```

---

## 📄 License

MIT — 详见 [LICENSE](LICENSE)

---

<p align="center">
  <strong>DocMind</strong> — <em>专治 Word 排版疑难杂症</em><br/>
  <sub>by <a href="https://github.com/baolongzhanshenNo1">baolongzhanshenNo1</a></sub>
</p>
