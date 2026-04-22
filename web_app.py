import streamlit as st
import os
import tempfile
import pandas as pd
import config
from core.processor import merge_pdfs_logic, extract_invoices_data, write_excel_from_data

# ================= 网页基础配置 =================
st.set_page_config(page_title=config.PAGE_TITLE, page_icon="🧾", layout="centered")

# ================= 极简 UI 视觉 =================
st.markdown(
    f"<h1 style='text-align:center; color:#D97757;'>{config.APP_NAME}</h1>",
    unsafe_allow_html=True)
st.markdown(
    f"<p style='text-align:center; color:#8C8070; margin-top:-8px;'>{config.APP_SUBTITLE}</p>",
    unsafe_allow_html=True)
st.divider()

# ================= 上传区 =================
uploaded_files = st.file_uploader(
    f"本次报销发票请拖拽至此处 (等同于放入 [{config.INPUT_FOLDER_NAME}])：",
    type=['pdf', 'jpg', 'jpeg', 'png'],
    accept_multiple_files=True
)

# 文件列表变化时清空上一次的提取结果，避免张冠李戴
def _files_fingerprint(files):
    return tuple((f.name, f.size) for f in files) if files else None

current_fp = _files_fingerprint(uploaded_files)
if st.session_state.get("last_file_fp") != current_fp:
    st.session_state["last_file_fp"] = current_fp
    st.session_state.pop("invoices", None)
    st.session_state.pop("failures", None)

if uploaded_files:
    st.success(f"已就绪 {len(uploaded_files)} 份发票文件", icon="✅")

    layout_choice = st.radio("请选择排版方向：", ("横向排版", "竖向排版"), horizontal=True)
    mode = '横向' if '横' in layout_choice else '竖向'

    st.write("")

    # ================= 两大操作 =================
    c1, c2 = st.columns(2)

    with c1:
        if st.button("1. 一键智能排版", use_container_width=True, type="primary"):
            bar = st.progress(0, text="准备中...")
            with tempfile.TemporaryDirectory() as tmp:
                for f in uploaded_files:
                    with open(os.path.join(tmp, os.path.basename(f.name)), "wb") as b:
                        b.write(f.getbuffer())

                def cb(cur, tot, name):
                    bar.progress(cur / tot, text=f"排版中 {cur}/{tot}：{name}")

                out = os.path.join(tmp, f"合并后_报销单({mode}).pdf")
                success, msg = merge_pdfs_logic(tmp, out, layout_mode=mode, progress_callback=cb)
                bar.empty()

                if success:
                    with open(out, "rb") as rb:
                        pdf_bytes = rb.read()
                    st.success(msg)
                    st.download_button(
                        "📥 下载排版 PDF", pdf_bytes,
                        f"报销合并单_{mode}.pdf", "application/pdf",
                        use_container_width=True)
                else:
                    st.error(msg)

    with c2:
        if st.button("2. AI 提取算税", use_container_width=True):
            bar = st.progress(0, text="准备中...")
            with tempfile.TemporaryDirectory() as tmp:
                for f in uploaded_files:
                    with open(os.path.join(tmp, os.path.basename(f.name)), "wb") as b:
                        b.write(f.getbuffer())

                def cb(cur, tot, name):
                    bar.progress(cur / tot, text=f"AI 识别 {cur}/{tot}：{name}")

                invoices, failures = extract_invoices_data(tmp, progress_callback=cb)
                bar.empty()

            st.session_state["invoices"] = invoices
            st.session_state["failures"] = failures
            st.rerun()

# ================= 提取结果预览 + 编辑 + 下载 =================
if st.session_state.get("invoices"):
    st.divider()
    st.markdown("### 📊 识别结果（可编辑）")
    st.caption("置信度偏低的行请人工核对；修改后再点击下方按钮生成 Excel。")

    df = pd.DataFrame(st.session_state["invoices"])

    edited_df = st.data_editor(
        df,
        use_container_width=True,
        num_rows="fixed",
        column_config={
            "文件名": st.column_config.TextColumn(disabled=True, width="medium"),
            "业务分类": st.column_config.SelectboxColumn(
                options=["机票行程单", "高铁/火车票", "打车/交通票", "加油费",
                         "通讯费", "餐饮发票", "住宿发票", "增值税发票"],
                required=True),
            "置信度(%)": st.column_config.ProgressColumn(
                format="%d", min_value=0, max_value=100, width="small"),
        },
        key="invoice_editor",
    )

    failures = st.session_state.get("failures") or []
    if failures:
        with st.expander(f"⚠️ {len(failures)} 个文件未能识别（点击查看）", expanded=False):
            for f in failures:
                st.markdown(f"- **{f['file']}** — {f['reason']}")

    st.write("")

    if st.button("📥 生成并下载 Excel", use_container_width=True, type="primary"):
        with tempfile.TemporaryDirectory() as tmp:
            out = os.path.join(tmp, "发票报销明细汇总.xlsx")
            success, msg = write_excel_from_data(edited_df.to_dict("records"), out)
            if success:
                with open(out, "rb") as rb:
                    xlsx_bytes = rb.read()
                st.success(msg)
                st.download_button(
                    "💾 点击保存到本地", xlsx_bytes,
                    "发票报销明细汇总.xlsx",
                    use_container_width=True)
            else:
                st.error(msg)

# ================= 页脚：GitHub Star 入口 =================
st.markdown(
    """
    <div style='text-align:center; color:#8C8070; margin-top:48px;
                font-size:0.85em; padding-bottom:20px;'>
      ⭐ 觉得好用？
      <a href='https://github.com/VicLuoV5/FlowInvoice' target='_blank'
         style='color:#D97757; text-decoration:none;'>在 GitHub 点亮星标</a>
    </div>
    """,
    unsafe_allow_html=True)
