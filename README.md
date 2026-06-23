# 🧾 极简票流 (FlowInvoice)

**智能发票提取与排版引擎** — 一键合并打印单，自动价税分离，本地离线运行。

专为个人报销与小微财务场景设计。告别手动调整排版、计算税额、核对明细的重复劳动。

---

## ✨ 核心功能

- **一键智能排版**：扫描「初始发票箱」中的 PDF / JPG / PNG，自动居中缩放排版，生成可直接打印的合并报销单（横向 / 竖向可选）
- **AI 提取算税**：本地 OCR + 正则提取，自动分类 8 种常见票据（机票、高铁、打车、加油、通讯、餐饮、住宿、增值税），自动算出不含税金额与税额
- **置信度可视化**：每张票据标记 0-100 置信度分数，Excel 中低可信行自动红/黄高亮，网页端可实时编辑修正
- **封面汇总 + 重复票检测**：Excel 含独立封面页按类型分组汇总，自动标记同发票号的疑似重复票
- **本地隐私**：所有处理均在本地完成，不依赖任何云服务或外部 API

## 🧩 Excel 模板填报

当前版本会导出内置格式的 `发票报销明细汇总.xlsx`，尚未内置“写入公司固定报销单模板”的可视化配置。

如果公司报销单格式固定，推荐的实现路线是：

1. 准备一份 `.xlsx` 模板，保留原有样式、公式、合并单元格和打印设置。
2. 增加模板映射配置，声明识别字段写入哪些单元格，例如报销日期、总金额、明细起始行、金额列、税额列、备注列。
3. 使用 `openpyxl` 复制模板并只写入映射单元格，避免破坏原表格式。
4. 在网页端/桌面端增加“按模板导出”入口；没有模板时继续使用当前默认汇总表。

这也是 [Issue #1](https://github.com/VicLuoV5/FlowInvoice/issues/1) 的建议解决方向。它应作为独立功能实现，而不是把某一家公司的表格格式硬编码进 `core/processor.py`。

## 📖 使用指南

面向**零技术背景用户**的完整操作手册（含截图、3 大场景、8 个常见问题）：

👉 **[docs/使用手册.pdf](docs/使用手册.pdf)**（推荐下载阅读）
👉 [docs/使用手册.md](docs/使用手册.md)（线上预览）

## 📦 快速开始

### 方式 A：Windows 可执行文件（零门槛）

从 [Releases](https://github.com/VicLuoV5/FlowInvoice/releases) 下载 `FlowInvoice.exe`，双击运行。

### 方式 B：源码运行（开发者 / macOS / Linux）

```bash
git clone https://github.com/VicLuoV5/FlowInvoice.git
cd FlowInvoice
pip install -r requirements.txt

streamlit run web_app.py   # 推荐：网页端（跨平台体验最佳）
python app.py              # 桌面端（CustomTkinter，已适配 macOS / Linux）
```

> macOS / Linux 暂未提供签名过的 `.app` / `.dmg`，推荐走网页端，体验比桌面端更好。

### 自己打包 `.exe`

```bash
pip install pyinstaller
pyinstaller app.spec --clean
# 产物在 dist/FlowInvoice.exe
```

## 🛠️ 技术栈

| 层 | 技术 |
|---|---|
| OCR 引擎 | RapidOCR (ONNX 运行时，离线) |
| PDF 处理 | PyMuPDF (fitz) |
| 数据 / Excel | pandas + openpyxl |
| 桌面端 UI | CustomTkinter + pywinstyles (Windows Mica 效果) |
| 网页端 UI | Streamlit |

## 🗂️ 项目结构

```
app.py          桌面端入口 (CustomTkinter)
web_app.py      网页端入口 (Streamlit)
config.py       全局常量：税率、路径、品牌
core/processor.py  核心引擎：排版 + OCR 提取 + Excel 输出
docs/           使用手册（PDF / Markdown / 截图）
app.spec        PyInstaller 打包配置
```

## 🤝 贡献

欢迎 Issue 与 PR。新票据类型、OCR 准确率优化、UI 改进都是优先方向。

### 贡献者

- [VicLuoV5](https://github.com/VicLuoV5)：项目作者与维护者
- [OpenAI Codex](https://openai.com/codex)：v1.0.2 发票识别修复、回归测试补充与发布验证协作

## 🙏 致谢

本项目在代码编写、架构重构与文档整理过程中使用了 [Claude Code](https://claude.com/claude-code) 辅助开发。所有产品需求定义、设计决策、功能审美与最终验收均由人类作者主导，AI 作为 pair programmer 负责具体实现与迭代。

v1.0.2 的发票识别修复、样本回归验证、Windows 可执行文件打包与 release 更新过程中使用了 [OpenAI Codex](https://openai.com/codex) 辅助协作。

仓库中的 `CLAUDE.md` 是面向 AI 协作者的开发约定，公开透明地记录了项目规则与边界，方便未来贡献者（无论人类还是 AI）快速对齐。

## 📜 License

MIT
