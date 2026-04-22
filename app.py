import os
import sys
import subprocess
import threading
import webbrowser
import customtkinter as ctk
from tkinter import messagebox

import config
from core.processor import merge_pdfs_logic, extract_data_logic

try:
    import pywinstyles
    HAS_WINDOW_STYLES = True
except ImportError:
    HAS_WINDOW_STYLES = False

GITHUB_URL = "https://github.com/VicLuoV5/FlowInvoice"

class FlowInvoiceApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title(config.PAGE_TITLE)
        self.geometry("450x560")
        self.eval('tk::PlaceWindow . center')

        if HAS_WINDOW_STYLES and os.name == 'nt':
            try:
                pywinstyles.apply_style(self, "mica")
            except Exception:
                pass

        if sys.platform == "darwin":
            self.MAIN_FONT = "PingFang SC"
        elif sys.platform == "win32":
            self.MAIN_FONT = "Microsoft YaHei"
        else:
            self.MAIN_FONT = "Noto Sans CJK SC"
        self.CORAL       = "#D97757"
        self.CORAL_HOVER = "#C4694A"
        self.DARK        = "#3D3929"
        self.DARK_HOVER  = "#2A271D"

        font_title    = (self.MAIN_FONT, 22, "bold")
        font_subtitle = (self.MAIN_FONT, 13)
        self.font_btn = (self.MAIN_FONT, 15, "bold")
        font_small    = (self.MAIN_FONT, 12)
        font_link     = (self.MAIN_FONT, 11)
        font_hint     = (self.MAIN_FONT, 10)

        self.check_env()

        # ── 标题区 ──────────────────────────────────────────────
        ctk.CTkLabel(self, text=config.APP_NAME,
                     font=font_title).pack(pady=(40, 5))
        ctk.CTkLabel(self, text=config.APP_SUBTITLE,
                     font=font_subtitle,
                     text_color=("gray45", "gray60")).pack(pady=(0, 26))

        # ── 文件入口卡片 ─────────────────────────────────────────
        # 左侧说明文字 + 右侧彩色「打开」按钮
        card_file = ctk.CTkFrame(self, fg_color=("gray91", "gray17"),
                                 corner_radius=8)
        card_file.pack(fill="x", padx=46, pady=(0, 20))

        ctk.CTkLabel(card_file,
                     text=f"📁   {config.INPUT_FOLDER_NAME}",
                     font=(self.MAIN_FONT, 13),
                     text_color=("gray20", "gray80")).pack(side="left",
                                                           padx=(16, 0), pady=12)
        ctk.CTkButton(card_file,
                      text="打  开",
                      font=font_small,
                      fg_color=self.CORAL,
                      hover_color=self.CORAL_HOVER,
                      text_color="white",
                      width=72, height=30,
                      corner_radius=6,
                      command=self.open_folder).pack(side="right", padx=12, pady=12)

        # ── 排版方向选择（紧贴排版按钮，视觉上属于同一操作组） ──
        self.radio_var = ctk.StringVar(value="横向")
        frame_radio = ctk.CTkFrame(self, fg_color="transparent")
        frame_radio.pack(pady=(0, 6))
        ctk.CTkRadioButton(frame_radio, text="横向排版",
                           variable=self.radio_var, value="横向",
                           font=(self.MAIN_FONT, 13),
                           fg_color=self.CORAL,
                           hover_color=self.CORAL_HOVER).pack(side="left", padx=20)
        ctk.CTkRadioButton(frame_radio, text="竖向排版",
                           variable=self.radio_var, value="竖向",
                           font=(self.MAIN_FONT, 13),
                           fg_color=self.CORAL,
                           hover_color=self.CORAL_HOVER).pack(side="left", padx=20)

        # ── 主操作按钮（两个按钮同父级、同 padx，保证绝对等宽对齐） ──
        self.btn_typeset = ctk.CTkButton(
            self,
            text="               1 .   一  键  智  能  排  版",
            anchor="w",
            font=self.font_btn,
            fg_color=self.CORAL,
            hover_color=self.CORAL_HOVER,
            height=50,
            corner_radius=6,
            command=self.click_merge
        )
        self.btn_typeset.pack(fill="x", padx=46, pady=(0, 14))

        # 两个主按钮之间的细分隔线，强化操作独立性
        ctk.CTkFrame(self, height=1,
                     fg_color=("gray82", "gray28")).pack(fill="x", padx=46, pady=(0, 14))

        self.btn_extract = ctk.CTkButton(
            self,
            text="               2 .   A   I   提  取  算  税",
            anchor="w",
            font=self.font_btn,
            fg_color=self.DARK,
            hover_color=self.DARK_HOVER,
            height=50,
            corner_radius=6,
            command=self.click_extract
        )
        self.btn_extract.pack(fill="x", padx=46, pady=(0, 22))

        # ── 清空（极度弱化的文字链接） ───────────────────────────
        self.btn_clear = ctk.CTkButton(
            self,
            text="清空发票箱",
            font=font_link,
            fg_color="transparent",
            text_color=("gray55", "gray50"),
            hover_color=("gray85", "gray22"),
            border_width=0,
            width=100,
            command=self.click_clear
        )
        self.btn_clear.pack(pady=(0, 14))

        # ── 底部提示文字 ─────────────────────────────────────────
        ctk.CTkLabel(self,
                     text="将报销发票放入「初始发票箱」后，点击上方按钮开始处理",
                     font=font_hint,
                     text_color=("gray60", "gray50")).pack(pady=(0, 6))

        # ── GitHub Star 入口（克制，低调） ───────────────────────
        ctk.CTkButton(self,
                      text="⭐ 觉得好用？在 GitHub 点亮星标",
                      font=font_link,
                      fg_color="transparent",
                      text_color=("gray60", "gray55"),
                      hover_color=("gray85", "gray22"),
                      border_width=0,
                      width=220,
                      command=lambda: webbrowser.open(GITHUB_URL)).pack(pady=(0, 18))

    def check_env(self):
        if not os.path.exists(config.INPUT_FOLDER):
            os.makedirs(config.INPUT_FOLDER)
        tip_file = os.path.join(config.INPUT_FOLDER, "💡说明：请将发票放入此文件夹.txt")
        if not os.path.exists(tip_file):
            with open(tip_file, "w", encoding="utf-8") as f:
                f.write("【排版顺序规则】\n程序默认按文件名的数字或字母顺序进行合并排版。\n"
                        "如果需要指定发票的先后顺序，请在文件名前加上数字编号，例如：\n"
                        "01_高铁票.pdf\n02_打车票.pdf\n03_酒店住宿.pdf\n\n"
                        "请将需要处理的 PDF 或图片发票放入此文件夹。")

    def open_folder(self):
        path = config.INPUT_FOLDER
        if sys.platform == "win32":
            os.startfile(path)
        elif sys.platform == "darwin":
            subprocess.Popen(["open", path])
        else:
            subprocess.Popen(["xdg-open", path])

    def _progress_text(self, prefix, cur, tot, name):
        display = name if len(name) <= 14 else name[:12] + ".."
        return f"               ⏳   {prefix} {cur}/{tot}   {display}"

    def click_merge(self):
        self.check_env()
        self.btn_typeset.configure(state="disabled",
                                   text="               ⏳    智  能  排  版  中 . . .")
        self.btn_extract.configure(state="disabled")
        mode = self.radio_var.get()

        def on_progress(cur, tot, name):
            txt = self._progress_text("排版", cur, tot, name)
            self.after(0, lambda t=txt: self.btn_typeset.configure(text=t))

        def worker():
            success, msg = merge_pdfs_logic(
                config.INPUT_FOLDER, f'合并后_报销单({mode}).pdf',
                layout_mode=mode, progress_callback=on_progress)
            self.after(0, lambda: self._on_merge_done(msg))

        threading.Thread(target=worker, daemon=True).start()

    def _on_merge_done(self, msg):
        self.btn_typeset.configure(state="normal",
                                   text="               1 .   一  键  智  能  排  版")
        self.btn_extract.configure(state="normal")
        messagebox.showinfo("处理结果", msg)

    def click_extract(self):
        self.check_env()
        self.btn_extract.configure(state="disabled",
                                   text="               ⏳    A   I   提  取  中 . . .")
        self.btn_typeset.configure(state="disabled")

        def on_progress(cur, tot, name):
            txt = self._progress_text("识别", cur, tot, name)
            self.after(0, lambda t=txt: self.btn_extract.configure(text=t))

        def worker():
            success, msg = extract_data_logic(
                config.INPUT_FOLDER, '发票报销明细汇总.xlsx',
                progress_callback=on_progress)
            self.after(0, lambda: self._on_extract_done(msg))

        threading.Thread(target=worker, daemon=True).start()

    def _on_extract_done(self, msg):
        self.btn_extract.configure(state="normal",
                                   text="               2 .   A   I   提  取  算  税")
        self.btn_typeset.configure(state="normal")
        messagebox.showinfo("处理结果", msg)

    def click_clear(self):
        self.check_env()
        if not messagebox.askyesno("确认", "确定要删除发票箱内所有发票文件吗？"):
            return
        count = 0
        for f in os.listdir(config.INPUT_FOLDER):
            if f.lower().endswith(('.pdf', '.jpg', '.jpeg', '.png')):
                try:
                    os.remove(os.path.join(config.INPUT_FOLDER, f))
                    count += 1
                except Exception:
                    pass
        messagebox.showinfo("清理完成", f"已清理 {count} 个文件。")

if __name__ == "__main__":
    app = FlowInvoiceApp()
    app.mainloop()
