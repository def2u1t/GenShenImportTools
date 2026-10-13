# -*- coding: utf-8 -*-
"""
圣遗物分析.exe - 圣遗物截图批量OCR解析工具(图形界面)
================================================================
流程: 选择圣遗物截图文件夹(run_xxx, 由「圣遗物采集.exe」生成)
      -> 点击「开始分析」-> 逐张 OCR 识别(约2秒/张) -> 生成 artifacts.json
结果: 截图文件夹下的 artifacts.json, 之后用「圣遗物导出Excel.exe」生成评分表格
打包: python -m PyInstaller --onefile --windowed --collect-all rapidocr_onnxruntime ...
"""
import os
import queue
import sys
import threading

import tkinter as tk
from tkinter import filedialog, scrolledtext

TITLE = "圣遗物分析"
TIP = (
    "使用说明:\n"
    "  1. 选择圣遗物截图文件夹(即「圣遗物采集」生成的 run_时间戳 文件夹)\n"
    "  2. 点击「开始分析」, 程序将逐张 OCR 识别截图中的圣遗物详情(约 2 秒/张, 请耐心等待)\n"
    "  3. 解析结果自动保存为该文件夹下的 artifacts.json\n"
    "  4. 之后打开「圣遗物导出Excel.exe」, 选择该 JSON 即可生成直观的评分表格\n"
    "  5. 采集过程中生成的每张截图都会被解析, 中途关闭不会损坏已保存结果(每10张自动保存)"
)


def app_dir():
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


def fix_stdio():
    """windowed exe 中 stdout/stderr 为 None, print 会崩溃 -> 重定向到日志文件"""
    if sys.stdout is None:
        sys.stdout = open(os.path.join(app_dir(), "圣遗物分析_日志.txt"), "a", encoding="utf-8")
    if sys.stderr is None:
        sys.stderr = sys.stdout


class LogQ(queue.Queue):
    def write(self, s):
        self.put(s)

    def flush(self):
        pass


class App:
    def __init__(self, root):
        self.root = root
        root.title(TITLE)
        root.geometry("800x560")
        self.q = LogQ()
        frm = tk.Frame(root)
        frm.pack(fill="both", expand=True, padx=12, pady=10)

        tk.Label(frm, text=TIP, justify="left", anchor="w",
                 font=("Microsoft YaHei UI", 10)).pack(fill="x")

        row = tk.Frame(frm)
        row.pack(fill="x", pady=(10, 4))
        tk.Label(row, text="截图文件夹:", font=("Microsoft YaHei UI", 10)).pack(side="left")
        self.dir_var = tk.StringVar(value="")
        tk.Entry(row, textvariable=self.dir_var).pack(side="left", fill="x", expand=True, padx=6)
        tk.Button(row, text="选择文件夹...", command=self.pick_dir).pack(side="left")

        self.btn = tk.Button(frm, text="开 始 分 析", font=("Microsoft YaHei UI", 11, "bold"),
                             bg="#4472C4", fg="white", height=2, command=self.on_start)
        self.btn.pack(fill="x", pady=8)

        tk.Label(frm, text="运行日志:", anchor="w").pack(fill="x")
        self.txt = scrolledtext.ScrolledText(frm, font=("Microsoft YaHei UI", 9))
        self.txt.pack(fill="both", expand=True)
        self.root.after(100, self.poll)

    def pick_dir(self):
        d = filedialog.askdirectory(initialdir=app_dir())
        if d:
            self.dir_var.set(d)

    def poll(self):
        try:
            while True:
                s = self.q.get_nowait()
                if s == "__DONE__":
                    self.btn.config(state="normal")
                    continue
                self.txt.insert("end", s)
                self.txt.see("end")
        except queue.Empty:
            pass
        self.root.after(100, self.poll)

    def on_start(self):
        run_dir = self.dir_var.get().strip()
        if not run_dir or not os.path.isdir(run_dir):
            from tkinter import messagebox
            messagebox.showerror(TITLE, "请先选择有效的截图文件夹!")
            return
        pngs = [f for f in os.listdir(run_dir) if f.endswith(".png")]
        if not pngs:
            from tkinter import messagebox
            messagebox.showerror(TITLE, f"该文件夹下没有截图(.png):\n{run_dir}")
            return
        self.btn.config(state="disabled")
        self.q.put(f"[开始] 待解析截图 {len(pngs)} 张, 目录: {run_dir}\n")
        threading.Thread(target=self.worker, args=(run_dir,), daemon=True).start()

    @staticmethod
    def run_analyze(run_dir):
        """分析主流程(GUI与自动测试共用)"""
        print("[准备] 正在加载OCR引擎(首次约几秒)...")
        from genshin_analyzer import batch_analyze
        result = batch_analyze(run_dir=run_dir)
        print(f"\n[完成] 成功解析 {len(result)} 张")
        print(f"[完成] 结果文件: {os.path.join(run_dir, 'artifacts.json')}")

    def worker(self, run_dir):
        old_out, old_err = sys.stdout, sys.stderr
        sys.stdout = self.q
        sys.stderr = self.q
        try:
            self.run_analyze(run_dir)
        except Exception as e:
            import traceback
            print(f"\n[错误] {e}")
            print(traceback.format_exc())
        finally:
            sys.stdout, sys.stderr = old_out, old_err
            self.q.put("__DONE__")


def autorun(run_dir):
    """无界面直跑(自动化测试/命令行使用): 圣遗物分析.exe --autorun <截图文件夹>"""
    fix_stdio()
    App.run_analyze(run_dir)


def main():
    fix_stdio()
    if "--autorun" in sys.argv:
        i = sys.argv.index("--autorun")
        if i + 1 >= len(sys.argv) or not os.path.isdir(sys.argv[i + 1]):
            print("用法: 圣遗物分析.exe --autorun <截图文件夹>")
            sys.exit(2)
        autorun(sys.argv[i + 1])
        return
    root = tk.Tk()
    App(root)
    if "--selftest" in sys.argv:
        root.after(3000, root.destroy)
    root.mainloop()


if __name__ == "__main__":
    main()
