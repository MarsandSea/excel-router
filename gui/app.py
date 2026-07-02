"""
Excel 通用拆分工具 - 图形界面

单屏自适应界面（先易后难，无需选"档"）：
  · 输入选「单个文件」→ 极简：选主拆分列 → 开始（每个取值一个文件）。
  · 输入选「文件夹」→ 自动展开批量区：每组打包 ZIP + 可选「同时拆到人」（仅指定的表）。
  · 「▸ 高级」折叠：表头识别策略、跳过值、取值归并、精确匹配、跨文件合并、保留格式。

网格汇总与到人是同一次运行的两个产出。拆分在子线程里跑，self.after(0, ...) 回主线程更新 UI。

Copyright (c) 2026 Abelin
MIT License
"""

import os
import sys
import json
import threading
import subprocess
import webbrowser

import customtkinter as ctk
from tkinter import filedialog, messagebox


# =====================================================
# 路径解析（兼容 PyInstaller --onefile 打包）
# =====================================================

def _is_frozen():
    return getattr(sys, 'frozen', False)


def _resource_dir():
    if _is_frozen():
        return sys._MEIPASS
    return os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def _app_dir():
    if _is_frozen():
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


DEFAULT_CONFIG_PATH = os.path.join(_resource_dir(), "config", "default_config.json")
USER_CONFIG_PATH    = os.path.join(_app_dir(), "user_config.json")

COL_PLACEHOLDER = "（请先识别列）"

APP_VERSION = "2.3"
# 匿名反馈问卷地址（问卷 URL 确定后替换此处即可，一行改动 + 打 tag 发版）
FEEDBACK_URL = "https://f.wps.cn/g/pBOAWUQc/"

# 兜底默认配置（通用、不绑定任何业务）
FALLBACK_CONFIG = {
    "input_path": "", "output_path": "",
    "header_mode": "auto", "header_row": 1,
    "grid_keys": [], "id_keys": [],
    "split_column": "", "selected_values": [],
    "person_column": "", "to_person": False, "person_file_filter": [],
    "value_alias_map": {},
    "skip_values": ["合计", "小计", "总计", "平均", ""],
    "merge_across_files": True,
    "make_zip": True,
    "exact_match": True,
    "preserve_format": True,
    "auto_open_output": True,
}

ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")


def load_config():
    for path in (USER_CONFIG_PATH, DEFAULT_CONFIG_PATH):
        if os.path.exists(path):
            try:
                with open(path, encoding='utf-8') as f:
                    cfg = json.load(f)
                merged = dict(FALLBACK_CONFIG)
                merged.update(cfg)
                return merged
            except Exception:
                continue
    return dict(FALLBACK_CONFIG)


def save_config(cfg):
    with open(USER_CONFIG_PATH, 'w', encoding='utf-8') as f:
        json.dump(cfg, f, ensure_ascii=False, indent=2)


class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title(f"ExcelRouter · Excel 业务数据自动分发工具 v{APP_VERSION}")
        self.geometry("780x860")
        self.minsize(700, 720)
        self._stop_flag = False
        self.cfg = load_config()
        self._columns = []
        self._adv_open = False
        self._build_ui()
        # 初始输入类型按已保存路径推断
        init_type = "文件夹" if os.path.isdir(self.cfg.get("input_path", "")) else "单个文件"
        self._input_type.set(init_type)
        self._on_input_type(init_type)

    # ── UI 构建 ──────────────────────────────────────────
    def _build_ui(self):
        self.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(self, text="Excel 通用拆分工具",
                     font=ctk.CTkFont(size=20, weight="bold")).grid(
            row=0, column=0, pady=(16, 0), padx=20, sticky="w")
        ctk.CTkLabel(self, text="选一列，把 Excel 拆成多个文件 · 保留原格式 · Copyright © 2026 Abelin · MIT",
                     font=ctk.CTkFont(size=11), text_color="gray").grid(
            row=1, column=0, padx=20, sticky="w")

        # 输入/输出
        pf = ctk.CTkFrame(self)
        pf.grid(row=2, column=0, padx=20, pady=(12, 0), sticky="ew")
        pf.grid_columnconfigure(1, weight=1)

        self._input_var  = ctk.StringVar(value=self.cfg.get("input_path", ""))
        self._output_var = ctk.StringVar(value=self.cfg.get("output_path", ""))
        self._input_type = ctk.StringVar(value="单个文件")

        ctk.CTkLabel(pf, text="输入").grid(row=0, column=0, padx=12, pady=8, sticky="w")
        seg = ctk.CTkSegmentedButton(pf, values=["单个文件", "文件夹"],
                                     variable=self._input_type, command=self._on_input_type)
        seg.grid(row=0, column=1, padx=8, pady=8, sticky="w")
        ctk.CTkEntry(pf, textvariable=self._input_var).grid(row=1, column=1, padx=8, pady=(0, 8), sticky="ew")
        ctk.CTkButton(pf, text="浏览", width=60, command=self._browse_input).grid(row=1, column=2, padx=8, pady=(0, 8))

        ctk.CTkLabel(pf, text="输出文件夹").grid(row=2, column=0, padx=12, pady=8, sticky="w")
        ctk.CTkEntry(pf, textvariable=self._output_var).grid(row=2, column=1, padx=8, pady=8, sticky="ew")
        ctk.CTkButton(pf, text="浏览", width=60,
                      command=lambda: self._browse_dir(self._output_var)).grid(row=2, column=2, padx=8)

        # 主拆分列
        mf = ctk.CTkFrame(self)
        mf.grid(row=3, column=0, padx=20, pady=(12, 0), sticky="ew")
        mf.grid_columnconfigure(1, weight=1)
        ctk.CTkLabel(mf, text="按这列拆", font=ctk.CTkFont(size=13, weight="bold")).grid(
            row=0, column=0, padx=12, pady=10, sticky="w")
        self._split_var = ctk.StringVar(value=self.cfg.get("split_column") or COL_PLACEHOLDER)
        self._split_menu = ctk.CTkOptionMenu(mf, variable=self._split_var,
                                             values=[self._split_var.get()], width=220)
        self._split_menu.grid(row=0, column=1, padx=8, pady=10, sticky="w")
        ctk.CTkButton(mf, text="🔄 识别列", width=90, command=self._detect_columns).grid(
            row=0, column=2, padx=8, pady=10)
        self._status = ctk.CTkLabel(mf, text="", font=ctk.CTkFont(size=11), text_color="gray")
        self._status.grid(row=1, column=0, columnspan=3, padx=12, pady=(0, 8), sticky="w")

        # 批量区（仅"文件夹"显示）
        self._batch_frame = self._build_batch(self)

        # 高级折叠
        self._adv_btn = ctk.CTkButton(self, text="▸ 高级设置", width=120, anchor="w",
                                      fg_color="transparent", text_color=("gray20", "gray80"),
                                      hover_color=("gray85", "gray25"), command=self._toggle_adv)
        self._adv_btn.grid(row=5, column=0, padx=20, pady=(10, 0), sticky="w")
        self._adv_frame = self._build_advanced(self)

        # 进度 + 按钮
        ctrl = ctk.CTkFrame(self, fg_color="transparent")
        ctrl.grid(row=7, column=0, padx=20, pady=12, sticky="ew")
        ctrl.grid_columnconfigure(3, weight=1)
        self._progress = ctk.CTkProgressBar(ctrl)
        self._progress.grid(row=0, column=0, columnspan=4, sticky="ew", pady=(0, 8))
        self._progress.set(0)
        self._start_btn = ctk.CTkButton(ctrl, text="▶ 开始处理", command=self._start, width=140)
        self._start_btn.grid(row=1, column=0, sticky="w")
        self._stop_btn = ctk.CTkButton(ctrl, text="⏹ 停止", command=self._stop, width=100,
                                       fg_color="gray40", hover_color="gray30", state="disabled")
        self._stop_btn.grid(row=1, column=1, padx=10, sticky="w")
        ctk.CTkButton(ctrl, text="💬 反馈建议", width=100, border_width=1,
                      fg_color="transparent", text_color=("gray20", "gray80"),
                      hover_color=("gray85", "gray25"),
                      command=self._open_feedback).grid(row=1, column=2, sticky="w")
        ctk.CTkButton(ctrl, text="保存配置", width=100, fg_color="green", hover_color="darkgreen",
                      command=self._save_cfg).grid(row=1, column=3, sticky="e")

        # 日志
        ctk.CTkLabel(self, text="处理日志", font=ctk.CTkFont(size=13)).grid(
            row=8, column=0, padx=20, pady=(0, 4), sticky="w")
        self._log_box = ctk.CTkTextbox(self, state="disabled", wrap="none")
        self._log_box.grid(row=9, column=0, padx=20, pady=(0, 16), sticky="nsew")
        self.grid_rowconfigure(9, weight=1)

    def _build_batch(self, parent):
        f = ctk.CTkFrame(parent)
        f.grid_columnconfigure(0, weight=1)

        self._zip_var = ctk.BooleanVar(value=self.cfg.get("make_zip", True))
        ctk.CTkCheckBox(f, text="每个分组打包成 ZIP（方便整体分发）",
                        variable=self._zip_var).grid(row=0, column=0, padx=12, pady=(10, 4), sticky="w")

        row = ctk.CTkFrame(f, fg_color="transparent")
        row.grid(row=1, column=0, padx=8, pady=(0, 4), sticky="ew")
        self._person_var = ctk.BooleanVar(value=self.cfg.get("to_person", False))
        ctk.CTkCheckBox(row, text="同时拆到人，按", variable=self._person_var).grid(row=0, column=0, padx=4)
        self._pcol_var = ctk.StringVar(value=self.cfg.get("person_column") or COL_PLACEHOLDER)
        self._pcol_menu = ctk.CTkOptionMenu(row, variable=self._pcol_var,
                                            values=[self._pcol_var.get()], width=140)
        self._pcol_menu.grid(row=0, column=1, padx=4)
        ctk.CTkLabel(row, text="拆；仅处理名字含").grid(row=0, column=2, padx=4)
        self._pfilter_var = ctk.StringVar(value=", ".join(self.cfg.get("person_file_filter", [])))
        ctk.CTkEntry(row, textvariable=self._pfilter_var, width=160,
                     placeholder_text="如 工资,费用（空=全部）").grid(row=0, column=3, padx=4)
        ctk.CTkLabel(row, text="的表").grid(row=0, column=4, padx=4)

        ctk.CTkLabel(f, text="到人只对指定的表生效，其余文件只进汇总。",
                     font=ctk.CTkFont(size=11), text_color="gray").grid(
            row=2, column=0, padx=12, pady=(0, 8), sticky="w")
        return f

    def _build_advanced(self, parent):
        f = ctk.CTkFrame(parent)
        f.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(f, text="表头识别").grid(row=0, column=0, padx=12, pady=8, sticky="w")
        mode_map = {"auto": "自动", "row": "指定行号", "keyword": "关键词"}
        self._header_seg = ctk.CTkSegmentedButton(f, values=["自动", "指定行号", "关键词"],
                                                  command=self._on_header_mode)
        self._header_seg.set(mode_map.get(self.cfg.get("header_mode", "auto"), "自动"))
        self._header_seg.grid(row=0, column=1, padx=8, pady=8, sticky="w")

        self._hrow_frame = ctk.CTkFrame(f, fg_color="transparent")
        self._hrow_frame.grid(row=1, column=0, columnspan=2, padx=12, pady=2, sticky="w")
        ctk.CTkLabel(self._hrow_frame, text="表头行号（1 基）：").grid(row=0, column=0)
        self._hrow_var = ctk.StringVar(value=str(self.cfg.get("header_row", 1)))
        ctk.CTkEntry(self._hrow_frame, textvariable=self._hrow_var, width=80).grid(row=0, column=1, padx=6)

        self._kw_frame = ctk.CTkFrame(f, fg_color="transparent")
        self._kw_frame.grid(row=2, column=0, columnspan=2, padx=12, pady=2, sticky="ew")
        self._kw_frame.grid_columnconfigure(1, weight=1)
        ctk.CTkLabel(self._kw_frame, text="必含关键词A（逗号分隔）").grid(row=0, column=0, padx=4, pady=4, sticky="w")
        self._grid_keys_var = ctk.StringVar(value=", ".join(self.cfg.get("grid_keys", [])))
        ctk.CTkEntry(self._kw_frame, textvariable=self._grid_keys_var).grid(row=0, column=1, padx=6, pady=4, sticky="ew")
        ctk.CTkLabel(self._kw_frame, text="必含关键词B（逗号分隔）").grid(row=1, column=0, padx=4, pady=4, sticky="w")
        self._id_keys_var = ctk.StringVar(value=", ".join(self.cfg.get("id_keys", [])))
        ctk.CTkEntry(self._kw_frame, textvariable=self._id_keys_var).grid(row=1, column=1, padx=6, pady=4, sticky="ew")

        ctk.CTkLabel(f, text="跳过值（逗号分隔）").grid(row=3, column=0, padx=12, pady=8, sticky="w")
        self._skip_var = ctk.StringVar(value=", ".join(self.cfg.get("skip_values", [])))
        ctk.CTkEntry(f, textvariable=self._skip_var).grid(row=3, column=1, padx=8, pady=8, sticky="ew")

        ctk.CTkLabel(f, text="取值归并映射（JSON，选填）").grid(row=4, column=0, padx=12, pady=(8, 2), sticky="nw")
        self._alias_box = ctk.CTkTextbox(f, height=64, wrap="none")
        self._alias_box.grid(row=4, column=1, padx=8, pady=(8, 2), sticky="ew")
        alias = self.cfg.get("value_alias_map", {})
        if alias:
            self._alias_box.insert("1.0", json.dumps(alias, ensure_ascii=False, indent=2))

        opts = ctk.CTkFrame(f, fg_color="transparent")
        opts.grid(row=5, column=0, columnspan=2, padx=8, pady=8, sticky="w")
        self._exact_var    = ctk.BooleanVar(value=self.cfg.get("exact_match", True))
        self._merge_var    = ctk.BooleanVar(value=self.cfg.get("merge_across_files", True))
        self._preserve_var = ctk.BooleanVar(value=self.cfg.get("preserve_format", True))
        self._auto_open_var = ctk.BooleanVar(value=self.cfg.get("auto_open_output", True))
        ctk.CTkCheckBox(opts, text="精确匹配", variable=self._exact_var).grid(row=0, column=0, padx=8, pady=4)
        ctk.CTkCheckBox(opts, text="跨文件合并汇总", variable=self._merge_var).grid(row=0, column=1, padx=8, pady=4)
        ctk.CTkCheckBox(opts, text="保留格式", variable=self._preserve_var).grid(row=0, column=2, padx=8, pady=4)
        ctk.CTkCheckBox(opts, text="完成后打开输出", variable=self._auto_open_var).grid(row=0, column=3, padx=8, pady=4)

        self._on_header_mode(self._header_seg.get())
        return f

    # ── 自适应 / 折叠 ───────────────────────────────────
    def _on_input_type(self, value):
        # 仅当现有路径与新类型不符时才清空（避免清掉启动时已保存的有效路径）
        p = self._input_var.get().strip()
        mismatch = ((value == "文件夹" and p and not os.path.isdir(p)) or
                    (value == "单个文件" and p and not os.path.isfile(p)))
        if mismatch:
            self._input_var.set("")
            self._columns = []
            self._split_menu.configure(values=[COL_PLACEHOLDER]); self._split_var.set(COL_PLACEHOLDER)
            self._set_status("")
        if value == "文件夹":
            self._batch_frame.grid(row=4, column=0, padx=20, pady=(12, 0), sticky="ew")
        else:
            self._batch_frame.grid_remove()

    def _toggle_adv(self):
        self._adv_open = not self._adv_open
        if self._adv_open:
            self._adv_frame.grid(row=6, column=0, padx=20, pady=(4, 0), sticky="ew")
            self._adv_btn.configure(text="▾ 高级设置")
        else:
            self._adv_frame.grid_remove()
            self._adv_btn.configure(text="▸ 高级设置")

    def _on_header_mode(self, label):
        if label == "指定行号":
            self._hrow_frame.grid(); self._kw_frame.grid_remove()
        elif label == "关键词":
            self._hrow_frame.grid_remove(); self._kw_frame.grid()
        else:
            self._hrow_frame.grid_remove(); self._kw_frame.grid_remove()

    # ── 浏览 / 识别 ─────────────────────────────────────
    def _browse_input(self):
        if self._input_type.get() == "单个文件":
            path = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx *.xls")])
        else:
            path = filedialog.askdirectory()
        if path:
            self._input_var.set(path)
            self._detect_columns()

    def _browse_dir(self, var):
        path = filedialog.askdirectory()
        if path:
            var.set(path)

    def _template_path(self):
        p = self._input_var.get().strip()
        if os.path.isfile(p):
            return p
        if os.path.isdir(p):
            for root, _, files in os.walk(p):
                for f in sorted(files):
                    if f.lower().endswith(('.xlsx', '.xls')) and not f.startswith('~$'):
                        return os.path.join(root, f)
        return None

    def _detect_columns(self):
        tpl = self._template_path()
        if not tpl:
            self._set_status("请先选择有效的输入")
            return
        self._set_status("正在识别列...")
        cfg = self._collect_config()

        def work():
            cols = []
            try:
                from core.splitter import list_columns
                cols = list_columns(tpl, cfg)
            except Exception:
                cols = []
            self.after(0, self._on_columns, cols)

        threading.Thread(target=work, daemon=True).start()

    def _on_columns(self, cols):
        self._columns = cols
        if cols:
            self._split_menu.configure(values=cols)
            if self._split_var.get() not in cols:
                self._split_var.set(cols[0])
            self._pcol_menu.configure(values=cols)
            if self._pcol_var.get() not in cols:
                self._pcol_var.set(cols[0])
            self._set_status(f"识别到 {len(cols)} 列")
        else:
            self._set_status("未识别到列（检查输入或高级里的表头设置）")

    def _set_status(self, text):
        self._status.configure(text=text)

    # ── 配置收集 ───────────────────────────────────────
    def _split_list(self, text):
        return [k.strip() for k in text.split(",") if k.strip()]

    def _parse_alias(self):
        raw = self._alias_box.get("1.0", "end").strip()
        if not raw:
            return {}, True
        try:
            obj = json.loads(raw)
            if isinstance(obj, dict):
                return obj, True
        except Exception:
            pass
        return {}, False

    def _collect_config(self):
        mode_map = {"自动": "auto", "指定行号": "row", "关键词": "keyword"}
        try:
            header_row = int(self._hrow_var.get())
        except (TypeError, ValueError):
            header_row = 1

        split_col = "" if self._split_var.get() == COL_PLACEHOLDER else self._split_var.get()
        pcol = "" if self._pcol_var.get() == COL_PLACEHOLDER else self._pcol_var.get()
        alias, _ = self._parse_alias()

        return {
            "input_path":  self._input_var.get().strip(),
            "output_path": self._output_var.get().strip(),
            "header_mode": mode_map.get(self._header_seg.get(), "auto"),
            "header_row":  header_row,
            "grid_keys":   self._split_list(self._grid_keys_var.get()),
            "id_keys":     self._split_list(self._id_keys_var.get()),
            "split_column": split_col,
            "selected_values": self.cfg.get("selected_values", []),
            "person_column": pcol,
            "to_person":   self._person_var.get(),
            "person_file_filter": self._split_list(self._pfilter_var.get()),
            "value_alias_map": alias,
            "skip_values": self._split_list(self._skip_var.get()),
            "merge_across_files": self._merge_var.get(),
            "make_zip":    self._zip_var.get(),
            "exact_match": self._exact_var.get(),
            "preserve_format": self._preserve_var.get(),
            "auto_open_output": self._auto_open_var.get(),
        }

    def _save_cfg(self):
        try:
            save_config(self._collect_config())
            messagebox.showinfo("已保存", f"配置已保存到：\n{USER_CONFIG_PATH}\n下次启动自动加载")
        except Exception as e:
            messagebox.showerror("保存失败", f"无法保存配置：{e}")

    # ── 运行 ───────────────────────────────────────────
    def _log(self, msg):
        self._log_box.configure(state="normal")
        self._log_box.insert("end", msg + "\n")
        self._log_box.see("end")
        self._log_box.configure(state="disabled")

    def _set_progress(self, val):
        self._progress.set(val)

    def _start(self):
        cfg = self._collect_config()
        if not cfg["input_path"] or not os.path.exists(cfg["input_path"]):
            messagebox.showerror("错误", "请选择有效的输入文件或文件夹")
            return
        if not cfg["output_path"]:
            messagebox.showerror("错误", "请选择输出文件夹")
            return
        if not cfg["split_column"]:
            messagebox.showerror("错误", "请先选择「按这列拆」（点「识别列」获取列名）")
            return
        _, alias_ok = self._parse_alias()
        if not alias_ok:
            messagebox.showerror("错误", "「取值归并映射」不是合法 JSON，请检查或清空")
            return

        self._stop_flag = False
        self._start_btn.configure(state="disabled")
        self._stop_btn.configure(state="normal")
        self._progress.set(0)
        self._log_box.configure(state="normal")
        self._log_box.delete("1.0", "end")
        self._log_box.configure(state="disabled")

        def run():
            try:
                from core.splitter import run_split
                output_path = run_split(
                    cfg,
                    log_fn=lambda m: self.after(0, self._log, m),
                    progress_fn=lambda v: self.after(0, self._set_progress, v),
                    stop_flag=lambda: self._stop_flag,
                )
            except Exception as e:
                self.after(0, self._log, f"\n❌ 运行出错：{e}")
                output_path = None
            self.after(0, self._on_done, cfg, output_path)

        threading.Thread(target=run, daemon=True).start()

    def _stop(self):
        self._stop_flag = True
        self._log("正在停止，请等待当前文件处理完成...")

    def _open_feedback(self):
        """打开匿名反馈页；顺手把版本号复制进剪贴板，方便用户粘贴到问卷。"""
        try:
            self.clipboard_clear()
            self.clipboard_append(f"ExcelRouter v{APP_VERSION}")
        except Exception:
            pass
        webbrowser.open(FEEDBACK_URL)
        self._log("💬 已打开反馈页，版本号已复制到剪贴板，粘贴到问卷即可。")

    def _on_done(self, cfg, output_path):
        self._start_btn.configure(state="normal")
        self._stop_btn.configure(state="disabled")
        self._progress.set(1)
        if output_path:
            self._log("💬 用得顺手或踩了坑？点「反馈建议」匿名告诉作者（1 分钟）。")
        if cfg.get("auto_open_output") and output_path and os.path.exists(output_path):
            try:
                subprocess.Popen(f'explorer "{os.path.normpath(output_path)}"')
            except Exception:
                pass


def run():
    app = App()
    app.mainloop()
