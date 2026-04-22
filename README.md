# 🧾 极简票流 (FlowInvoice)

**智能发票提取与排版引擎** — 一键合并 A4 打印单，自动价税分离，本地离线运行。

专为个人报销与小微财务场景设计。告别手动调整排版、计算税额、核对明细的重复劳动。

---

## ✨ 核心功能

- **一键智能排版**：扫描「初始发票箱」中的 PDF / JPG / PNG，自动居中缩放到 A4 幅面，生成可直接打印的合并报销单（横向 / 竖向可选）
- **AI 提取算税**：本地 OCR 识别 + 正则提取，自动分类 8 种常见票据（机票、高铁、打车、加油、通讯、餐饮、住宿、增值税），自动算出不含税金额与税额
- **置信度可视化**：每张票据标记 0-100 置信度分数，Excel 中低可信行自动红/黄高亮，网页端可实时编辑修正
- **封面汇总**：生成的 Excel 含独立封面页，按票据类型分组汇总数量与金额
- **重复票检测**：自动标记同发票号的疑似重复票据
- **本地隐私**：所有票据处理均在本地完成，不依赖任何云服务或外部 API

## 🎯 适用场景

- 差旅报销：整理一趟出差的几十张票据，生成符合财务要求的合并打印单 + 明细表
- 小微企业记账：个体户 / 自由职业者按月汇总发票
- 财务助理：加速日常发票录入与核对

## 🛠️ 技术栈

| 层 | 技术 |
|---|---|
| OCR 引擎 | RapidOCR (ONNX 运行时，离线) |
| PDF 处理 | PyMuPDF (fitz) |
| 数据 / Excel | pandas + openpyxl |
| 桌面端 UI | CustomTkinter + pywinstyles (Windows Mica 效果) |
| 网页端 UI | Streamlit |

## 📦 快速开始

### 方式 A：源码运行（开发者 / 跨平台）

```bash
git clone https://github.com/VicLuoV5/FlowInvoice.git
cd FlowInvoice
pip install -r requirements.txt
```

启动桌面端：

```bash
python app.py
```

启动网页端：

```bash
streamlit run web_app.py
```

### 方式 B：Windows 可执行文件（非开发者）

从 [Releases](https://github.com/VicLuoV5/FlowInvoice/releases) 下载 `FlowInvoice.exe`，双击运行。

自己打包：

```bash
pip install pyinstaller
pyinstaller app.spec --clean
# 产物在 dist/FlowInvoice.exe
```

### 方式 C：macOS / Linux 用户

暂未提供 `.app` / `.dmg` 打包（未做代码签名，Gatekeeper 体验不佳），请走源码路线：

```bash
git clone https://github.com/VicLuoV5/FlowInvoice.git
cd FlowInvoice
pip install -r requirements.txt
streamlit run web_app.py     # 推荐：网页端体验最佳
python app.py                # 桌面端也可用，已适配跨平台
```

**推荐网页端**的理由：UI 响应更快、数据可编辑、macOS 原生字体渲染更自然。桌面端仅在需要离线双击启动时才建议。

## 📖 使用流程

### 桌面端

1. 启动 `app.py`，点击「打开」进入「初始发票箱」文件夹
2. 将所有发票（PDF / JPG / PNG）放入该文件夹
3. 回到主界面，选择排版方向（横向 / 竖向），点击「**1. 一键智能排版**」生成合并 PDF
4. 点击「**2. AI 提取算税**」生成 Excel 明细（含封面汇总页）
5. 点击「清空发票箱」开始下一轮

### 网页端

1. `streamlit run web_app.py` 打开浏览器
2. 拖拽发票文件到上传区
3. 点击「排版」下载合并 PDF，或点击「AI 提取」进入可视化编辑器
4. 在表格中核对 / 修正识别结果，点击「生成并下载 Excel」

## 💡 使用技巧

- **指定发票顺序**：在文件名前加数字编号，如 `01_高铁票.pdf` `02_打车票.pdf`，程序按文件名排序合并
- **低可信行**：Excel 中红色行（置信度 < 50）建议人工复核；黄色行（50-79）建议抽查
- **重复票**：同发票号的票据会在备注列标记「⚠️ 疑似重复」
- **图像模糊**：如 OCR 无法识别，程序会把该文件放入失败列表并给出原因，不会静默丢弃

## ❓ FAQ

**Q：所有数据都在本地处理吗？**
是。OCR 使用 RapidOCR + ONNX 本地模型，所有发票、图像、Excel 均不离开本机。

**Q：首次运行 .exe 很慢？**
PyInstaller 单文件包启动时需要解压到临时目录，首次冷启动 3-8 秒属正常。

**Q：餐饮发票为什么显示价税合计而没有分离税额？**
按国税规定，餐饮服务增值税不可抵扣进项，程序只记录价税合计并在备注中标注。

**Q：支持哪些票据类型？**
机票行程单、高铁/火车票、打车/出租、加油费、通讯费、餐饮、住宿、通用增值税发票。其余类型会进入失败列表，欢迎提 Issue 扩展。

**Q：可以修改税率吗？**
可以。编辑 `config.py` 中的 `HSR_TAX_RATE`、`FLIGHT_TAX_RATE`、`TAXI_TAX_RATE`。

## 🗂️ 项目结构

```
app.py              桌面端入口 (CustomTkinter)
web_app.py          网页端入口 (Streamlit)
config.py           全局常量：税率、路径、品牌
core/
  processor.py      核心引擎：排版 + OCR 提取 + Excel 输出
初始发票箱/          运行时输入目录（git 忽略）
app.spec            PyInstaller 打包配置
.streamlit/         Streamlit 主题配置
```

## 🤝 贡献

欢迎 Issue 与 PR。新票据类型、OCR 准确率优化、UI 改进都是优先方向。

## 🙏 致谢

本项目在代码编写、架构重构与文档整理过程中使用了 [Claude Code](https://claude.com/claude-code) 辅助开发。所有产品需求定义、设计决策、功能审美与最终验收均由人类作者主导，AI 作为 pair programmer 负责具体实现与迭代。

仓库中的 `CLAUDE.md` 是面向 AI 协作者的开发约定，公开透明地记录了项目规则与边界，方便未来贡献者（无论人类还是 AI）快速对齐。

## 📜 License

MIT
