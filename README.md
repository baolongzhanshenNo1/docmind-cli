<h1 align="center">DocMind CLI</h1>

<p align="center">
  <strong>AI 驱动的智能文档排版工具</strong> — 专治 Word 排版疑难杂症
</p>

<p align="center">
  <a href="LICENSE"><img src="https://img.shields.io/badge/license-MIT-blue.svg" alt="MIT"></a>
  <a href="#"><img src="https://img.shields.io/badge/python-≥3.11-blue" alt="Python 3.11+"></a>
  <a href="#"><img src="https://img.shields.io/badge/platform-Windows%20|%20macOS%20|%20Linux-lightgrey" alt="Platform"></a>
</p>

---

## 能解决什么问题

如果你写过长文档——毕业论文、合同、标书、技术手册——你一定经历过这些：

| 折磨 | DocMind 怎么做 |
|------|---------------|
| 封面莫名其妙出现页眉，删不掉 | 前置页（封面/声明）自动清空页眉页码 |
| 正文各章页眉全是"参考文献" | 每节创建独立页眉，各章自动显示对应标题 |
| 目录页码对不上实际内容 | 页码分节编号：前置罗马数字、正文阿拉伯数字 |
| 每章开头歪在偶数页上 | 奇偶页强制：空白页插入 + 奇数页起始 |
| 中英文混排字体乱套 | 中文宋体/标题黑体，英文 Times New Roman，分离设置不打架 |
| 改完一版学校规范又变了，重调一遍 | 规范模板驱动：换一份规范 docx，重新跑一下命令 |

**核心逻辑**：给它一份"格式乱的论文"和一份"学校规范模板"，把格式差异找出来修掉——**只修格式、不动文字内容**。

---

## ⚡ 快速开始

```bash
git clone https://github.com/baolongzhanshenNo1/docmind-cli.git
cd docmind-cli
pip install .

# 最常用：论文按规范排版
docmind format my_thesis.docx --spec university_spec.docx

# 提取规范模板的格式规则 → YAML
docmind analyze university_spec.docx
```

### `format` 示例输出

```
  目标文档 → my_thesis.docx
  规范模板 → university_spec.docx
  正在排版… ⣾

   📄 修正项  61
   🔄 残留    0
   ⚠️  诊断    0
   📦 输出    my_thesis_formatted.docx

✅ 排版完成
```

---

## 命令参考

| 命令 | 作用 |
|------|------|
| `docmind format <论文> --spec <规范>` | 按规范模板排版 |
| `docmind analyze <规范>` | 提取格式规则 → YAML |
| `docmind enforce <文档>` | 奇偶页强制（需 LibreOffice） |

---

## 能力清单

| 能力 | 说明 |
|------|------|
| 📐 页眉 | 全文各节独立页眉（封面/声明自动清空） |
| 📄 页码 | 前置罗马数字 / 正文阿拉伯数字 / 封面前置无页码 |
| 🔢 奇偶页 | 每章奇数页起始 / 扉页空白页插入 |
| 🎨 字体 | 中文宋体+标题黑体 / 英文 Times New Roman（Ascii/EastAsian 分离） |
| 📏 边距 | 严格对齐规范模板 |
| 📑 分节符 | 拆分合并节 / 补断链引用 |
| ✅ 诊断 | 排版后自动 Fixer 诊断，0 残留才通过 |

---

## 环境

- Python ≥ 3.11
- LibreOffice（仅 `enforce` 命令需要）

Python 依赖通过 `pip install .` 自动安装。

---

## 测试

```bash
pip install -e ".[dev]"
pytest tests/ -q
```

---

## 架构

```
docmind format xxx.docx --spec yyy.docx
      │
      ▼
  DocMindAgent (pipeline/agent.py)
      │
      ├── SpecReader   ── 从规范模板提取字体/边距/页眉格式
      ├── DocDiscover  ── 自动发现文档章节结构
      ├── Reconciler   ── 差异对比 → 生成修复计划
      ├── Writer       ── ZIP 级 OOXML 写入（7种操作）
      └── Fixer        ── 诊断执行结果（8种验证器）
```

所有操作基于 OOXML ZIP 层直接修改，不依赖 Word 或 LibreOffice 运行时。

---

## 相关项目

- [DocMind Desktop](https://github.com/baolongzhanshenNo1) — 对话式 Agent 桌面应用（Tauri + Vue 3），自然语言指令排版

---

## License

MIT — 详见 [LICENSE](LICENSE)

<p align="center">
  <sub>by <a href="https://github.com/baolongzhanshenNo1">baolongzhanshenNo1</a></sub>
</p>
