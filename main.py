import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import openpyxl
import os
import threading
import time
from tkinterdnd2 import DND_FILES, TkinterDnD


class BeisiExcelExtractor:
    def __init__(self, root):
        self.root = root
        self.root.title("贝思excel数据提取")
        self.root.geometry("780x600")
        self.root.resizable(True, True)
        self.root.configure(bg="#f0f4f8")

        self.files = []
        self.output_dir = tk.StringVar()
        self.col_letter = tk.StringVar(value="A")
        self.start_pos = tk.IntVar(value=1)
        self.end_pos = tk.IntVar(value=5)
        self.has_header = tk.BooleanVar(value=True)
        self.has_footer = tk.BooleanVar(value=False)
        self.footer_rows = tk.IntVar(value=1)
        self.processing = False

        self._build_ui()

    def _build_ui(self):
        title_frame = tk.Frame(self.root, bg="#2c5f8a", height=50)
        title_frame.pack(fill=tk.X)
        title_frame.pack_propagate(False)
        tk.Label(
            title_frame,
            text="贝思excel数据提取",
            font=("Microsoft YaHei", 16, "bold"),
            bg="#2c5f8a",
            fg="white",
        ).pack(expand=True)

        main_frame = tk.Frame(self.root, bg="#f0f4f8")
        main_frame.pack(fill=tk.BOTH, expand=True, padx=15, pady=10)

        # --- 文件区域 ---
        file_frame = tk.LabelFrame(
            main_frame,
            text=" 源文件 ",
            font=("Microsoft YaHei", 10, "bold"),
            bg="#f0f4f8",
            fg="#2c5f8a",
        )
        file_frame.pack(fill=tk.X, pady=(0, 8))

        drop_frame = tk.Frame(file_frame, bg="#dce8f5", relief=tk.RIDGE, bd=2, height=80)
        drop_frame.pack(fill=tk.X, padx=8, pady=6)
        drop_frame.pack_propagate(False)

        self.drop_label = tk.Label(
            drop_frame,
            text="📂  拖拽 Excel 文件到此处，或点击下方按钮选择文件（支持多选）",
            font=("Microsoft YaHei", 10),
            bg="#dce8f5",
            fg="#555",
        )
        self.drop_label.pack(expand=True)

        try:
            drop_frame.drop_target_register(DND_FILES)
            drop_frame.dnd_bind("<<Drop>>", self._on_drop)
            self.drop_label.drop_target_register(DND_FILES)
            self.drop_label.dnd_bind("<<Drop>>", self._on_drop)
        except Exception:
            pass

        btn_row = tk.Frame(file_frame, bg="#f0f4f8")
        btn_row.pack(fill=tk.X, padx=8, pady=(0, 6))
        tk.Button(
            btn_row,
            text="选择文件",
            command=self._select_files,
            bg="#2c5f8a",
            fg="white",
            font=("Microsoft YaHei", 9),
            relief=tk.FLAT,
            padx=12,
            pady=4,
        ).pack(side=tk.LEFT)
        tk.Button(
            btn_row,
            text="清空列表",
            command=self._clear_files,
            bg="#c0392b",
            fg="white",
            font=("Microsoft YaHei", 9),
            relief=tk.FLAT,
            padx=12,
            pady=4,
        ).pack(side=tk.LEFT, padx=(6, 0))
        self.file_count_label = tk.Label(
            btn_row,
            text="已选 0 个文件",
            font=("Microsoft YaHei", 9),
            bg="#f0f4f8",
            fg="#555",
        )
        self.file_count_label.pack(side=tk.RIGHT)

        list_frame = tk.Frame(file_frame, bg="#f0f4f8")
        list_frame.pack(fill=tk.X, padx=8, pady=(0, 6))
        self.file_listbox = tk.Listbox(
            list_frame,
            height=4,
            font=("Microsoft YaHei", 8),
            bg="white",
            selectmode=tk.EXTENDED,
        )
        self.file_listbox.pack(side=tk.LEFT, fill=tk.X, expand=True)
        sb = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=self.file_listbox.yview)
        sb.pack(side=tk.RIGHT, fill=tk.Y)
        self.file_listbox.configure(yscrollcommand=sb.set)

        # --- 提取参数 ---
        param_frame = tk.LabelFrame(
            main_frame,
            text=" 提取参数 ",
            font=("Microsoft YaHei", 10, "bold"),
            bg="#f0f4f8",
            fg="#2c5f8a",
        )
        param_frame.pack(fill=tk.X, pady=(0, 8))

        row1 = tk.Frame(param_frame, bg="#f0f4f8")
        row1.pack(fill=tk.X, padx=8, pady=6)

        tk.Label(row1, text="目标列（字母）:", font=("Microsoft YaHei", 9), bg="#f0f4f8").pack(side=tk.LEFT)
        tk.Entry(row1, textvariable=self.col_letter, width=6, font=("Microsoft YaHei", 9)).pack(side=tk.LEFT, padx=(4, 16))

        tk.Label(row1, text="起始位（第几位）:", font=("Microsoft YaHei", 9), bg="#f0f4f8").pack(side=tk.LEFT)
        tk.Spinbox(row1, textvariable=self.start_pos, from_=1, to=999, width=5, font=("Microsoft YaHei", 9)).pack(side=tk.LEFT, padx=(4, 16))

        tk.Label(row1, text="结束位（第几位）:", font=("Microsoft YaHei", 9), bg="#f0f4f8").pack(side=tk.LEFT)
        tk.Spinbox(row1, textvariable=self.end_pos, from_=1, to=999, width=5, font=("Microsoft YaHei", 9)).pack(side=tk.LEFT, padx=(4, 0))

        row2 = tk.Frame(param_frame, bg="#f0f4f8")
        row2.pack(fill=tk.X, padx=8, pady=(0, 6))
        tk.Checkbutton(
            row2,
            text="首行为表头（跳过）",
            variable=self.has_header,
            font=("Microsoft YaHei", 9),
            bg="#f0f4f8",
        ).pack(side=tk.LEFT)
        tk.Checkbutton(
            row2,
            text="有表尾（跳过末尾行数）:",
            variable=self.has_footer,
            font=("Microsoft YaHei", 9),
            bg="#f0f4f8",
        ).pack(side=tk.LEFT, padx=(16, 0))
        tk.Spinbox(row2, textvariable=self.footer_rows, from_=1, to=999, width=5, font=("Microsoft YaHei", 9)).pack(side=tk.LEFT, padx=(4, 0))

        # --- 输出目录 ---
        out_frame = tk.LabelFrame(
            main_frame,
            text=" 输出位置 ",
            font=("Microsoft YaHei", 10, "bold"),
            bg="#f0f4f8",
            fg="#2c5f8a",
        )
        out_frame.pack(fill=tk.X, pady=(0, 8))
        out_row = tk.Frame(out_frame, bg="#f0f4f8")
        out_row.pack(fill=tk.X, padx=8, pady=6)
        tk.Entry(out_row, textvariable=self.output_dir, font=("Microsoft YaHei", 9)).pack(side=tk.LEFT, fill=tk.X, expand=True)
        tk.Button(
            out_row,
            text="浏览",
            command=self._select_output,
            bg="#2c5f8a",
            fg="white",
            font=("Microsoft YaHei", 9),
            relief=tk.FLAT,
            padx=10,
            pady=3,
        ).pack(side=tk.LEFT, padx=(6, 0))

        # --- 进度 & 速度 ---
        prog_frame = tk.LabelFrame(
            main_frame,
            text=" 处理进度 ",
            font=("Microsoft YaHei", 10, "bold"),
            bg="#f0f4f8",
            fg="#2c5f8a",
        )
        prog_frame.pack(fill=tk.X, pady=(0, 8))

        self.progress_bar = ttk.Progressbar(prog_frame, orient=tk.HORIZONTAL, mode="determinate")
        self.progress_bar.pack(fill=tk.X, padx=8, pady=(6, 2))

        info_row = tk.Frame(prog_frame, bg="#f0f4f8")
        info_row.pack(fill=tk.X, padx=8, pady=(0, 6))
        self.progress_label = tk.Label(
            info_row, text="就绪", font=("Microsoft YaHei", 9), bg="#f0f4f8", fg="#333"
        )
        self.progress_label.pack(side=tk.LEFT)
        self.speed_label = tk.Label(
            info_row, text="", font=("Microsoft YaHei", 9), bg="#f0f4f8", fg="#2c5f8a"
        )
        self.speed_label.pack(side=tk.RIGHT)

        # --- 开始按钮 ---
        self.start_btn = tk.Button(
            main_frame,
            text="▶  开始提取",
            command=self._start_processing,
            bg="#27ae60",
            fg="white",
            font=("Microsoft YaHei", 12, "bold"),
            relief=tk.FLAT,
            pady=8,
        )
        self.start_btn.pack(fill=tk.X)

    # ---- 文件操作 ----
    def _select_files(self):
        paths = filedialog.askopenfilenames(
            title="选择 Excel 文件",
            filetypes=[("Excel 文件", "*.xlsx *.xls *.xlsm"), ("所有文件", "*.*")],
        )
        self._add_files(list(paths))

    def _on_drop(self, event):
        raw = event.data
        # tkinterdnd2 返回的路径可能被花括号包裹
        paths = self.root.tk.splitlist(raw)
        self._add_files(list(paths))

    def _add_files(self, paths):
        for p in paths:
            p = p.strip()
            if p and p not in self.files:
                self.files.append(p)
                self.file_listbox.insert(tk.END, os.path.basename(p))
        self.file_count_label.config(text=f"已选 {len(self.files)} 个文件")

    def _clear_files(self):
        self.files.clear()
        self.file_listbox.delete(0, tk.END)
        self.file_count_label.config(text="已选 0 个文件")

    def _select_output(self):
        d = filedialog.askdirectory(title="选择输出目录")
        if d:
            self.output_dir.set(d)

    # ---- 处理逻辑 ----
    def _start_processing(self):
        if self.processing:
            return
        if not self.files:
            messagebox.showwarning("提示", "请先选择要处理的 Excel 文件")
            return
        if not self.output_dir.get():
            messagebox.showwarning("提示", "请先选择输出目录")
            return
        col = self.col_letter.get().strip().upper()
        if not col.isalpha():
            messagebox.showwarning("提示", "目标列请输入字母，如 A、B、AB")
            return
        s = self.start_pos.get()
        e = self.end_pos.get()
        if s < 1 or e < s:
            messagebox.showwarning("提示", "起始位和结束位设置有误，请检查")
            return

        self.processing = True
        self.start_btn.config(state=tk.DISABLED)
        thread = threading.Thread(target=self._process_files, daemon=True)
        thread.start()

    def _col_letter_to_index(self, col: str) -> int:
        """列字母转0-based索引，如 A->0, B->1, AA->26"""
        idx = 0
        for ch in col.upper():
            idx = idx * 26 + (ord(ch) - ord("A") + 1)
        return idx - 1

    def _process_files(self):
        files = self.files[:]
        total = len(files)
        col = self.col_letter.get().strip().upper()
        col_idx = self._col_letter_to_index(col)
        s = self.start_pos.get() - 1  # 转为0-based
        e = self.end_pos.get()        # slice用法 [s:e]
        out_dir = self.output_dir.get()
        has_header = self.has_header.get()
        has_footer = self.has_footer.get()
        footer_rows = self.footer_rows.get() if has_footer else 0

        start_time = time.time()
        errors = []

        for i, fpath in enumerate(files):
            fname = os.path.basename(fpath)
            self._update_progress(i, total, f"正在处理: {fname}", start_time, i)

            try:
                ext = os.path.splitext(fpath)[1].lower()
                wb = openpyxl.load_workbook(fpath, data_only=True)
                for ws in wb.worksheets:
                    all_rows = list(ws.iter_rows())
                    data_rows = all_rows
                    if has_header and data_rows:
                        data_rows = data_rows[1:]
                    if footer_rows > 0 and len(data_rows) > footer_rows:
                        data_rows = data_rows[:-footer_rows]

                    for row in data_rows:
                        if col_idx < len(row):
                            cell = row[col_idx]
                            val = str(cell.value) if cell.value is not None else ""
                            extracted = val[s:e]
                            cell.value = extracted

                out_path = os.path.join(out_dir, fname)
                # 避免覆盖同名：自动加序号
                base, suffix = os.path.splitext(out_path)
                counter = 1
                while os.path.exists(out_path):
                    out_path = f"{base}_提取{counter}{suffix}"
                    counter += 1
                wb.save(out_path)

            except Exception as ex:
                errors.append(f"{fname}: {ex}")

        elapsed = time.time() - start_time
        self._update_progress(total, total, f"完成！共处理 {total} 个文件，耗时 {elapsed:.1f}s", start_time, total)
        self.root.after(0, self._finish, errors)

    def _update_progress(self, done, total, msg, start_time, processed):
        pct = int(done / total * 100) if total else 0
        elapsed = time.time() - start_time
        speed = processed / elapsed if elapsed > 0 and processed > 0 else 0
        speed_text = f"{speed:.1f} 文件/秒" if speed > 0 else ""
        self.root.after(
            0,
            lambda: (
                self.progress_bar.configure(value=pct),
                self.progress_label.config(text=msg),
                self.speed_label.config(text=speed_text),
            ),
        )

    def _finish(self, errors):
        self.processing = False
        self.start_btn.config(state=tk.NORMAL)
        if errors:
            err_msg = "\n".join(errors[:10])
            if len(errors) > 10:
                err_msg += f"\n...共 {len(errors)} 个错误"
            messagebox.showerror("部分文件处理失败", err_msg)
        else:
            messagebox.showinfo("完成", "所有文件已成功处理！\n输出文件已保存到指定目录。")


if __name__ == "__main__":
    try:
        root = TkinterDnD.Tk()
    except Exception:
        root = tk.Tk()
    app = BeisiExcelExtractor(root)
    root.mainloop()
