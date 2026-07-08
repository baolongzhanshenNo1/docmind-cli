<h1 align="center">DocMind CLI</h1>

<p align="center">
  <strong>AI 驱动的智能文档排版工具</strong> — 将"2 小时手工调整"压缩为"2 分钟对话指令"
</p>

<p align="center">
  <a href="LICENSE"><img src="https://img.shields.io/badge/license-MIT-blue.svg" alt="MIT"></a>
  <a href="#"><img src="https://img.shields.io/badge/python-≥3.11-blue" alt="Python 3.11+"></a>
  <a href="#"><img src="https://img.shields.io/badge/platform-Windows%20|%20macOS%20|%20Linux-lightgrey" alt="Platform"></a>
</p>

---

## 项目简介

**DocMind** 是一个基于深度 OOXML 的智能文档排版引擎。给它一篇格式混乱的论文，加上一份学校规范模板，它能自动完成字体、标题层级、页眉、页码、分节符、奇偶页等全套排版——**只修格式、不改一字**。

| 维度 | 说明 |
|------|------|
| 📂 定位 | 排版修正工具（非内容生成器、非 AI 代写） |
| 🏠 运行方式 | 本地离线，数据不出电脑 |
| 🔧 核心能力 | 字体分离（中文/西文）、页眉独立化、页码分节、奇偶页强制、分节符拆分、Fixer 诊断 + 验证 |
| 🎯 目标用户 | 高校学生（论文）、律师（合同）、政企（标书/公文）、开发团队（技术文档） |
| 📊 市场规模 | 年毕业 500 万本硕博 + 全国 4 万律所 + 年千亿政府采购 |
| 🧩 开源协议 | MIT — 完全免费，可商用 |
| 🔗 相关项目 | [DocMind Desktop](https://github.com/baolongzhanshenNo1/docmind-desktop) — 对话式 Agent 桌面应用 |

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
| Python ≥ 3.11 | |
| LibreOffice | 仅 `enforce` 命令需要 |

---

## 🧪 测试

```bash
pip install -e ".[dev]"
pytest tests/ -q
```

---

## 📄 License

MIT — 详见 [LICENSE](LICENSE)

<p align="center">
  <strong>DocMind</strong> — <em>专治 Word 排版疑难杂症</em><br/>
  <sub>by <a href="https://github.com/baolongzhanshenNo1">baolongzhanshenNo1</a></sub>
</p>
