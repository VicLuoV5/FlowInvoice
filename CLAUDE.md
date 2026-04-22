# CLAUDE.md — 极简票流 FlowInvoice

## 项目定位
纯本地运行的发票提取与排版引擎，面向个人报销和小微财务。
双端：桌面版 (`app.py` CustomTkinter) + 网页版 (`web_app.py` Streamlit)。

## 目录约定
```
app.py          桌面端入口
web_app.py      网页端入口
config.py       全局常量（税率、路径、品牌）
core/
  processor.py  唯一核心引擎，包含 merge_pdfs_logic + extract_data_logic
初始发票箱/     运行时输入目录（git 忽略，用户手动放入发票）
```

`core/extract_data.py` 和 `core/merge_pdf.py` 是已废弃模块，不要再编辑或导入。

## 改动原则
- 税率、金额阈值、缩放系数等数字常量全部放 `config.py`，不在 processor.py 硬编码
- 发票类型分类优先级：机票 > 高铁/火车 > 打车 > 餐饮 > 住宿 > 增值税（通用兜底）
- 桌面端耗时操作必须走 `threading.Thread`，不允许阻塞主线程
- UI 回调只在主线程执行（用 `self.after(0, callback)` 从子线程回主线程）

## 开发环境
Python 3.10+，依赖见 `requirements.txt`（已锁定最低兼容版本）。
不依赖任何云服务或外部 API。

## 不要做的事
- 不要在 `processor.py` 里直接 print/log，异常通过返回值 `(False, msg)` 传递
- 不要在提交里包含 `*.pdf`、`*.xlsx`、`*.png`（见 `.gitignore`）
- 不要修改 UI 的空格对齐文本（那是有意为之的视觉排版，不是 bug）
