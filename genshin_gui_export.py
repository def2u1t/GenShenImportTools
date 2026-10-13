# -*- coding: utf-8 -*-
"""
圣遗物导出Excel.exe - 圣遗物评分直观展示工具(图形界面)
================================================================
流程: 选择评分数据 artifacts.json(由「圣遗物分析.exe」生成)
      -> 选择 Excel 输出文件夹 -> 点击「开始导出」-> 生成 圣遗物评分明细.xlsx
结果: 评分明细(按分数降序, 满级给当前评分/未满级给预期, 良好绿底/及格黄底)
      + 汇总(统计口径/等级档/建议分布/Top5) 两个工作表
打包: python -m PyInstaller --onefile --windowed ...
"""
import os
import queue
import sys
import threading

import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext

TITLE = "圣遗物导出Excel"
EXCEL_NAME = "圣遗物评分明细.xlsx"
TIP = (
    "使用说明(评分结果直观展示):\n"
    "  1. 选择评分数据文件 artifacts.json(由「圣遗物分析.exe」解析生成)\n"
    "  2. 选择 Excel 输出文件夹\n"
    "  3. 点击「开始导出」, 生成「圣遗物评分明细.xlsx」\n"
    "  4. 评分口径: 满级件(+20)给当前评分; 未满级件给预期分数(期望潜力)\n"
    "     明细按分数降序排列, 良好/优秀绿底、及格黄底, 一眼挑出值得保留的圣遗物"
)


def app_dir():
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


def fix_stdio():
    if sys.stdout is None:
        sys.stdout = open(os.path.join(app_dir(), "圣遗物导出日志.txt"), "a", encoding="utf-8")
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
        root.geometry("800x600")
        self.q = LogQ()
        frm = tk.Frame(root)
        frm.pack(fill="both", expand=True, padx=12, pady=10)

        tk.Label(frm, text=TIP, justify="left", anchor="w",
                 font=("Microsoft YaHei UI", 10)).pack(fill="x")

        row1 = tk.Frame(frm)
        row1.pack(fill="x", pady=(10, 4))
        tk.Label(row1, text="评分数据(JSON):", font=("Microsoft YaHei UI", 10)).pack(side="left")
        self.json_var = tk.StringVar(value="")
        tk.Entry(row1, textvariable=self.json_var).pack(side="left", fill="x", expand=True, padx=6)
        tk.Button(row1, text="选择JSON...", command=self.pick_json).pack(side="left")

        row2 = tk.Frame(frm)
        row2.pack(fill="x", pady=(0, 4))
        tk.Label(row2, text="Excel输出文件夹:", font=("Microsoft YaHei UI", 10)).pack(side="left")
        self.out_var = tk.StringVar(value=app_dir())
        tk.Entry(row2, textvariable=self.out_var).pack(side="left", fill="x", expand=True, padx=6)
        tk.Button(row2, text="选择文件夹...", command=self.pick_out).pack(side="left")

        self.btn = tk.Button(frm, text="开 始 导 出", font=("Microsoft YaHei UI", 11, "bold"),
                             bg="#4472C4", fg="white", height=2, command=self.on_start)
        self.btn.pack(fill="x", pady=8)

        tk.Label(frm, text="运行日志:", anchor="w").pack(fill="x")
        self.txt = scrolledtext.ScrolledText(frm, font=("Microsoft YaHei UI", 9))
        self.txt.pack(fill="both", expand=True)
        self.root.after(100, self.poll)

    def pick_json(self):
        f = filedialog.askopenfilename(
            title="选择评分数据 artifacts.json",
            filetypes=[("JSON数据", "*.json"), ("所有文件", "*.*")],
            initialdir=app_dir())
        if f:
            self.json_var.set(f)

    def pick_out(self):
        d = filedialog.askdirectory(
            initialdir=self.out_var.get() if os.path.isdir(self.out_var.get()) else app_dir())
        if d:
            self.out_var.set(d)

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
        json_path = self.json_var.get().strip()
        out_dir = self.out_var.get().strip()
        if not json_path or not os.path.isfile(json_path):
            messagebox.showerror(TITLE, "请先选择有效的 artifacts.json 文件!")
            return
        if not os.path.isdir(out_dir):
            if not messagebox.askyesno("目录不存在", f"目录不存在:\n{out_dir}\n\n是否自动创建并继续?"):
                return
            os.makedirs(out_dir, exist_ok=True)
        self.btn.config(state="disabled")
        self.q.put(f"[开始] 评分数据: {json_path}\n")
        threading.Thread(target=self.worker, args=(json_path, out_dir), daemon=True).start()

    @staticmethod
    def run_export(json_path, out_dir):
        """导出主流程(GUI与自动测试共用)"""
        import genshin_scoring as gs
        excel_path = os.path.join(out_dir, EXCEL_NAME)
        rows, stats, saved = gs.score_file(json_path, excel_path=excel_path)
        print(f"\n[完成] Excel已保存: {saved or excel_path}")
        print(f"[完成] 共评分 {stats['总数']} 件, 最高 {stats['主最高']:.1f} 分")

    def worker(self, json_path, out_dir):
        old_out, old_err = sys.stdout, sys.stderr
        sys.stdout = self.q
        sys.stderr = self.q
        try:
            self.run_export(json_path, out_dir)
        except Exception as e:
            import traceback
            print(f"\n[错误] {e}")
            print(traceback.format_exc())
        finally:
            sys.stdout, sys.stderr = old_out, old_err
            self.q.put("__DONE__")


def autorun(json_path, out_dir):
    """无界面直跑(自动化测试/命令行使用): 圣遗物导出Excel.exe --autorun <json> <输出文件夹>"""
    fix_stdio()
    App.run_export(json_path, out_dir)


def main():
    fix_stdio()
    if "--autorun" in sys.argv:
        i = sys.argv.index("--autorun")
        if i + 2 >= len(sys.argv) or not os.path.isfile(sys.argv[i + 1]):
            print("用法: 圣遗物导出Excel.exe --autorun <artifacts.json> <输出文件夹>")
            sys.exit(2)
        out_dir = sys.argv[i + 2] if i + 2 < len(sys.argv) else "."
        if not os.path.isdir(out_dir):
            os.makedirs(out_dir, exist_ok=True)
        autorun(sys.argv[i + 1], out_dir)
        return
    root = tk.Tk()
    App(root)
    if "--selftest" in sys.argv:
        root.after(3000, root.destroy)
    root.mainloop()


if __name__ == "__main__":
    main()
