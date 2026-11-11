# -*- coding: utf-8 -*-
"""
华为基站MML任务结果智能解析与转换工具 - 桌面图形界面 (海量数据极速版)
Huawei MML Task Result Parser GUI

特性：
1. 界面友好简洁，支持直接选择单个/多个TXT文件，或选择整个文件夹批量扫描。
2. 支持选择导出 CSV 还是 Excel 文件格式，默认为 CSV 文件。
3. 支持设置文件编码，默认为 GBK 编码（与 Windows / 中文 Excel 完美契合）。
4. 深度优化海量数据解析速度：流式缓冲区、列表直写、毫秒级状态机。
5. 独立工作线程执行，界面流畅显示实时行数、处理速度（行/秒）与日志。
"""

import os
import sys
import threading
import subprocess
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from tkinter.scrolledtext import ScrolledText

from mml_parser import export_mml_files

class MMLParserGUI(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("华为MML任务结果解析与转换工具")
        self.geometry("860x680")
        self.minsize(900, 600)

        # 界面主题设置
        self.style = ttk.Style(self)
        if "vista" in self.style.theme_names():
            self.style.theme_use("vista")
        elif "clam" in self.style.theme_names():
            self.style.theme_use("clam")

        self.input_files = []
        self.output_dir = ""

        self.create_widgets()

    def create_widgets(self):
        # 顶部标题栏
        header_frame = ttk.Frame(self, padding=(18, 12))
        header_frame.pack(fill=tk.X)

        title_label = ttk.Label(
            header_frame,
            text="华为MML任务结果解析与转换工具",
            font=("Microsoft YaHei UI", 15, "bold")
        )
        title_label.pack(anchor=tk.W)

        sub_label = ttk.Label(
            header_frame,
            text="面向千万行海量数据深度优化 | 自动按中文查询命令分表输出 | 复合开关(分类集&子类)全展开 ",
            foreground="#555555",
            font=("Microsoft YaHei UI", 9)
        )
        sub_label.pack(anchor=tk.W, pady=(3, 0))

        # 主容器
        main_frame = ttk.Frame(self, padding=(18, 5))
        main_frame.pack(fill=tk.BOTH, expand=True)

        # ----------------------------------------------------
        # 1. 输入文件区域
        # ----------------------------------------------------
        file_group = ttk.LabelFrame(main_frame, text=" 1. 待解析TXT文件选择 ", padding=10)
        file_group.pack(fill=tk.X, pady=(0, 10))

        btn_bar = ttk.Frame(file_group)
        btn_bar.pack(fill=tk.X, pady=(0, 6))

        btn_add_files = ttk.Button(btn_bar, text="📁 添加TXT文件 (支持多选)", command=self.choose_files)
        btn_add_files.pack(side=tk.LEFT, padx=(0, 8))

        btn_add_folder = ttk.Button(btn_bar, text="📂 添加文件夹 (递归扫描)", command=self.choose_folder)
        btn_add_folder.pack(side=tk.LEFT, padx=(0, 8))

        btn_clear = ttk.Button(btn_bar, text="🗑️ 清空列表", command=self.clear_files)
        btn_clear.pack(side=tk.RIGHT)

        self.lbl_file_count = ttk.Label(btn_bar, text="已选文件：0 个", foreground="#0066cc", font=("Microsoft YaHei UI", 9, "bold"))
        self.lbl_file_count.pack(side=tk.LEFT, padx=10)

        # 选中的文件展示框
        list_frame = ttk.Frame(file_group)
        list_frame.pack(fill=tk.X)

        self.file_listbox = tk.Listbox(list_frame, height=4, selectmode=tk.EXTENDED, font=("Consolas", 9))
        self.file_listbox.pack(side=tk.LEFT, fill=tk.X, expand=True)

        list_scroll = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=self.file_listbox.yview)
        list_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        self.file_listbox.config(yscrollcommand=list_scroll.set)

        # ----------------------------------------------------
        # 2. 导出配置区域
        # ----------------------------------------------------
        cfg_group = ttk.LabelFrame(main_frame, text=" 2. 导出参数配置 ", padding=10)
        cfg_group.pack(fill=tk.X, pady=(0, 10))

        # 输出目录行
        out_path_frame = ttk.Frame(cfg_group)
        out_path_frame.pack(fill=tk.X, pady=(0, 8))

        ttk.Label(out_path_frame, text="输出目录：").pack(side=tk.LEFT)
        self.ent_output_dir = ttk.Entry(out_path_frame)
        self.ent_output_dir.insert(0, self.output_dir)
        self.ent_output_dir.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)

        btn_browse_out = ttk.Button(out_path_frame, text="浏览...", command=self.choose_output_dir)
        btn_browse_out.pack(side=tk.RIGHT)

        # 格式与编码配置行
        opt_frame = ttk.Frame(cfg_group)
        opt_frame.pack(fill=tk.X, pady=(2, 0))

        # 输出格式选择（默认为 csv）
        ttk.Label(opt_frame, text="导出格式：", font=("Microsoft YaHei UI", 9, "bold")).pack(side=tk.LEFT)
        self.var_format = tk.StringVar(value="csv")
        r_csv = ttk.Radiobutton(opt_frame, text="CSV 格式 (.csv，默认推荐)", variable=self.var_format, value="csv")
        r_csv.pack(side=tk.LEFT, padx=(3, 12))

        r_excel = ttk.Radiobutton(opt_frame, text="Excel 格式 (.xlsx)", variable=self.var_format, value="excel")
        r_excel.pack(side=tk.LEFT, padx=(0, 12))

        r_both = ttk.Radiobutton(opt_frame, text="CSV + Excel 同时导出", variable=self.var_format, value="both")
        r_both.pack(side=tk.LEFT, padx=(0, 25))

        # 文件编码选择（默认为 gbk）
        ttk.Label(opt_frame, text="文件编码：", font=("Microsoft YaHei UI", 9, "bold")).pack(side=tk.LEFT)
        self.var_encoding = tk.StringVar(value="gbk")
        r_gbk = ttk.Radiobutton(opt_frame, text="GBK (默认)", variable=self.var_encoding, value="gbk")
        r_gbk.pack(side=tk.LEFT, padx=(3, 10))

        r_utf8 = ttk.Radiobutton(opt_frame, text="UTF-8 (含BOM)", variable=self.var_encoding, value="utf-8-sig")
        r_utf8.pack(side=tk.LEFT)

        # ----------------------------------------------------
        # 3. 运行控制与进度区域
        # ----------------------------------------------------
        run_group = ttk.Frame(main_frame)
        run_group.pack(fill=tk.X, pady=(4, 8))

        self.btn_run = ttk.Button(
            run_group,
            text="🚀 开始极速解析并导出",
            command=self.start_parsing
        )
        self.btn_run.pack(side=tk.LEFT, ipadx=16, ipady=5)

        self.btn_open_dir = ttk.Button(
            run_group,
            text="📂 打开输出目录",
            command=self.open_output_dir,
            state=tk.NORMAL
        )
        self.btn_open_dir.pack(side=tk.LEFT, padx=12, ipady=5)

        self.lbl_speed_stats = ttk.Label(run_group, text="", foreground="#008000", font=("Microsoft YaHei UI", 9, "bold"))
        self.lbl_speed_stats.pack(side=tk.RIGHT, padx=5)

        # 进度条与状态文本
        progress_frame = ttk.Frame(main_frame)
        progress_frame.pack(fill=tk.X, pady=(0, 5))

        self.progress_bar = ttk.Progressbar(progress_frame, orient=tk.HORIZONTAL, mode='determinate')
        self.progress_bar.pack(fill=tk.X, pady=(0, 3))

        self.lbl_status = ttk.Label(progress_frame, text="就绪：请添加待解析的华为MML结果TXT文件后点击开始", foreground="#666666")
        self.lbl_status.pack(anchor=tk.W)

        # 执行日志区域
        log_group = ttk.LabelFrame(main_frame, text=" 解析执行日志 ", padding=6)
        log_group.pack(fill=tk.BOTH, expand=True)

        self.log_text = ScrolledText(log_group, height=10, font=("Consolas", 9))
        self.log_text.pack(fill=tk.BOTH, expand=True)

        # 记录默认浏览目录与输出路径是否被用户手动修改过
        self.last_dir = os.path.dirname(os.path.abspath(__file__))
        self.output_dir_manually_set = False

    def log(self, message: str):
        self.log_text.insert(tk.END, message + "\n")
        self.log_text.see(tk.END)

    def choose_files(self):
        initial_dir = self.last_dir
        if self.input_files:
            initial_dir = os.path.dirname(self.input_files[0])
        files = filedialog.askopenfilenames(
            title="选择华为MML结果TXT文件",
            initialdir=initial_dir,
            filetypes=[("TXT 文本文件", "*.txt"), ("所有文件", "*.*")]
        )
        if files:
            self.add_files(files)

    def choose_folder(self):
        initial_dir = self.last_dir
        if self.input_files:
            initial_dir = os.path.dirname(self.input_files[0])
        folder = filedialog.askdirectory(
            title="选择包含MML结果文件的文件夹",
            initialdir=initial_dir
        )
        if folder:
            txt_files = []
            for root, _, files in os.walk(folder):
                for f in files:
                    if f.lower().endswith(".txt"):
                        txt_files.append(os.path.join(root, f))
            if txt_files:
                self.add_files(txt_files)
            else:
                messagebox.showinfo("提示", "该文件夹下未找到任何 .txt 文件！")

    def add_files(self, file_paths):
        is_first_batch = (len(self.input_files) == 0)
        for f in file_paths:
            f = os.path.abspath(f)
            if f not in self.input_files:
                self.input_files.append(f)
                self.file_listbox.insert(tk.END, f)
        
        self.lbl_file_count.config(text=f"已选文件：{len(self.input_files)} 个")

        # 默认与第一个导入文件的路径保持一致
        if self.input_files:
            first_dir = os.path.dirname(self.input_files[0])
            self.last_dir = first_dir
            # 如果输出目录尚未被用户手动修改，则默认与第一个导入文件路径一致
            if not self.output_dir_manually_set:
                self.output_dir = first_dir
                self.ent_output_dir.delete(0, tk.END)
                self.ent_output_dir.insert(0, first_dir)

    def clear_files(self):
        self.input_files.clear()
        self.file_listbox.delete(0, tk.END)
        self.lbl_file_count.config(text="已选文件：0 个")
        self.lbl_speed_stats.config(text="")
        if not self.output_dir_manually_set:
            self.output_dir = ""
            self.ent_output_dir.delete(0, tk.END)

    def choose_output_dir(self):
        initial_dir = self.ent_output_dir.get().strip() or self.last_dir
        dir_path = filedialog.askdirectory(title="选择输出目录", initialdir=initial_dir)
        if dir_path:
            self.output_dir = os.path.abspath(dir_path)
            self.output_dir_manually_set = True
            self.ent_output_dir.delete(0, tk.END)
            self.ent_output_dir.insert(0, self.output_dir)

    def open_output_dir(self):
        out_dir = self.ent_output_dir.get().strip()
        if not out_dir and self.input_files:
            out_dir = os.path.dirname(self.input_files[0])
        if not out_dir:
            out_dir = self.last_dir
        if not os.path.exists(out_dir):
            os.makedirs(out_dir, exist_ok=True)
        if sys.platform == "win32":
            os.startfile(out_dir)
        else:
            subprocess.Popen(["xdg-open", out_dir])

    def start_parsing(self):
        if not self.input_files:
            messagebox.showwarning("提示", "请先选择需要解析的MML结果TXT文件！")
            return

        out_dir = self.ent_output_dir.get().strip()
        if not out_dir and self.input_files:
            out_dir = os.path.dirname(self.input_files[0])
            self.ent_output_dir.delete(0, tk.END)
            self.ent_output_dir.insert(0, out_dir)
        if not out_dir:
            messagebox.showwarning("提示", "请指定输出目录！")
            return

        self.btn_run.config(state=tk.DISABLED)
        self.progress_bar["value"] = 0
        self.log_text.delete(1.0, tk.END)
        self.lbl_speed_stats.config(text="正在准备...")

        export_fmt = self.var_format.get()
        export_enc = self.var_encoding.get()
        files = list(self.input_files)

        # 启动后台线程执行解析，界面保持 60fps 响应不卡顿
        worker = threading.Thread(
            target=self.run_worker,
            args=(files, out_dir, export_fmt, export_enc),
            daemon=True
        )
        worker.start()

    def run_worker(self, files, out_dir, export_fmt, export_enc):
        t_start = 0

        def on_progress(cur, total, rows, msg):
            def update_ui():
                if total > 0:
                    pct = int((cur / total) * 100)
                    self.progress_bar["value"] = pct
                self.lbl_status.config(text=f"[{cur}/{total}] {msg}")
                self.log(f"[{cur}/{total}] {msg}")
                if rows > 0:
                    self.lbl_speed_stats.config(text=f"已解析: {rows} 行")
            self.after(0, update_ui)

        try:
            self.log(f"开始执行MML解析任务...")
            self.log(f"待处理文件数: {len(files)}")
            self.log(f"导出格式: {export_fmt.upper()}, CSV编码: {export_enc.upper()}")
            self.log(f"输出目录: {os.path.abspath(out_dir)}")
            self.log("-" * 50)

            result = export_mml_files(
                input_files=files,
                output_dir=out_dir,
                output_format=export_fmt,
                csv_encoding=export_enc,
                progress_callback=on_progress
            )

            def on_success():
                self.btn_run.config(state=tk.NORMAL)
                self.progress_bar["value"] = 100
                spd = int(result['total_rows'] / (result['elapsed_seconds'] + 0.001))
                self.lbl_status.config(text=f"解析完成！共处理 {result['total_rows']} 行，耗时 {result['elapsed_seconds']} 秒")
                self.lbl_speed_stats.config(text=f"速度: {spd} 行/秒 | 耗时: {result['elapsed_seconds']}s")

                self.log("\n" + "=" * 55)
                self.log("解析完成统计：")
                self.log(f"处理文件数: {result['total_files']}")
                self.log(f"解析总记录数: {result['total_rows']}")
                self.log(f"总耗时: {result['elapsed_seconds']} 秒 (平均吞吐率: {spd} 行/秒)")
                self.log("生成的报表列表（以中文查询命令命名）：")
                for cmd, cnt in result['records_per_command'].items():
                    title = result['cmd_display_names'].get(cmd, cmd)
                    self.log(f"  - 【{title}】({cmd}): {cnt} 条记录")
                self.log(f"输出目录: {os.path.abspath(out_dir)}")
                self.log("=" * 55)
                messagebox.showinfo("完成", f"MML文件解析完成！\n共解析 {result['total_rows']} 行数据\n耗时 {result['elapsed_seconds']} 秒\n输出目录：\n{out_dir}")

            self.after(0, on_success)
        except Exception as e:
            def on_error(err_str):
                self.btn_run.config(state=tk.NORMAL)
                self.lbl_status.config(text=f"解析出错: {err_str}", foreground="red")
                self.lbl_speed_stats.config(text="错误中断", foreground="red")
                self.log(f"\n[错误] 解析中断: {err_str}")
                messagebox.showerror("解析错误", f"处理过程中发生异常：\n{err_str}")
            self.after(0, lambda: on_error(str(e)))

def main():
    app = MMLParserGUI()
    app.mainloop()

if __name__ == "__main__":
    main()
