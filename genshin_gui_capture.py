# -*- coding: utf-8 -*-
"""
圣遗物采集.exe - 原神圣遗物截图采集工具(图形界面)
================================================================
流程: 打开原神背包(列表最顶部) -> 选保存目录 -> 点「开始采集」-> 确认弹窗
      -> 3秒倒计时(切回游戏) -> 自动逐格点击并截图详情面板
保存: 所选目录下 run_时间戳 子文件夹(可用「圣遗物分析.exe」解析)
中止: 鼠标快速甩到屏幕左上角
打包: python -m PyInstaller --onefile --windowed --collect-all rapidocr_onnxruntime ...
"""
import os
import queue
import sys
import threading
import time
from datetime import datetime

import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext

TITLE = "圣遗物采集"
TIP = (
    "使用说明:\n"
    "  1. 先打开原神圣遗物背包界面, 确保列表从最顶部开始显示(第一行可见, 不要提前滚动)\n"
    "  2. 选择截图保存目录(默认: 本程序所在目录下的 screenshots 文件夹)\n"
    "  3. 点击「开始采集」并确认后, 有 3 秒倒计时, 请立即切回原神背包界面\n"
    "  4. 采集过程中请勿移动鼠标、不要改变窗口位置; 紧急中止: 鼠标快速甩到屏幕左上角\n"
    "  5. 截图保存在所选目录下 run_时间戳 子文件夹, 之后用「圣遗物分析.exe」解析成 JSON"
)


def app_dir():
    """exe所在目录(打包后) 或 脚本所在目录"""
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


def fix_stdio():
    """windowed exe 中 stdout/stderr 为 None, print 会崩溃 -> 重定向到日志文件"""
    if sys.stdout is None:
        sys.stdout = open(os.path.join(app_dir(), "圣遗物采集_日志.txt"), "a", encoding="utf-8")
    if sys.stderr is None:
        sys.stderr = sys.stdout


class LogQ(queue.Queue):
    """worker 线程写日志, GUI 主线程轮询取出显示(避免 tkinter 跨线程)"""

    def write(self, s):
        self.put(s)

    def flush(self):
        pass


class App:
    def __init__(self, root):
        self.root = root
        root.title(TITLE)
        root.geometry("780x560")
        self.q = LogQ()
        frm = tk.Frame(root)
        frm.pack(fill="both", expand=True, padx=12, pady=10)

        tk.Label(frm, text=TIP, justify="left", anchor="w",
                 font=("Microsoft YaHei UI", 10)).pack(fill="x")

        row = tk.Frame(frm)
        row.pack(fill="x", pady=(10, 4))
        tk.Label(row, text="截图保存目录:", font=("Microsoft YaHei UI", 10)).pack(side="left")
        self.dir_var = tk.StringVar(value=os.path.join(app_dir(), "screenshots"))
        tk.Entry(row, textvariable=self.dir_var).pack(side="left", fill="x", expand=True, padx=6)
        tk.Button(row, text="浏览...", command=self.pick_dir).pack(side="left")

        self.btn = tk.Button(frm, text="开 始 采 集", font=("Microsoft YaHei UI", 11, "bold"),
                             bg="#4472C4", fg="white", height=2, command=self.on_start)
        self.btn.pack(fill="x", pady=8)

        tk.Label(frm, text="运行日志:", anchor="w").pack(fill="x")
        self.txt = scrolledtext.ScrolledText(frm, font=("Microsoft YaHei UI", 9))
        self.txt.pack(fill="both", expand=True)
        self.root.after(100, self.poll)

    def pick_dir(self):
        d = filedialog.askdirectory(
            initialdir=self.dir_var.get() if os.path.isdir(self.dir_var.get()) else app_dir())
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
        save_dir = self.dir_var.get().strip()
        if not os.path.isdir(save_dir):
            if not messagebox.askyesno("目录不存在", f"目录不存在:\n{save_dir}\n\n是否自动创建并继续?"):
                return
            os.makedirs(save_dir, exist_ok=True)
        if not messagebox.askokcancel(
                "开始采集", "点击「确定」后开始 3 秒倒计时。\n请立即切到原神圣遗物背包界面(列表在最顶部)!"):
            return
        self.btn.config(state="disabled")
        self.q.put(f"[开始] 截图保存目录: {save_dir}\n")
        threading.Thread(target=self.worker, args=(save_dir,), daemon=True).start()

    @staticmethod
    def run_capture(save_dir):
        """采集主流程(GUI与自动测试共用)"""
        print("[准备] 正在加载OCR引擎(首次约几秒)...")
        from genshin_capture import get_artifact_count, traverse_artifacts, PER_PAGE
        for i in range(3, 0, -1):
            print(f"[倒计时] {i} 秒后开始, 请切到原神背包界面!")
            time.sleep(1)
        count, total = get_artifact_count()
        print(f"[识别] 圣遗物数量: {count} / {total}")
        run_dir = os.path.join(save_dir, datetime.now().strftime("run_%Y%m%d_%H%M%S"))
        pages = count // PER_PAGE + 3          # 页数上限(保险用, +3页富余)
        clicked = traverse_artifacts(artifact_count=count, pages=pages, run_dir=run_dir)
        print(f"\n[完成] 共采集 {clicked} 张截图")
        print(f"[完成] 保存目录: {run_dir}")

    def worker(self, save_dir):
        old_out, old_err = sys.stdout, sys.stderr
        sys.stdout = self.q
        sys.stderr = self.q
        try:
            self.run_capture(save_dir)
        except Exception as e:
            import traceback
            print(f"\n[错误] {e}")
            print(traceback.format_exc())
        finally:
            sys.stdout, sys.stderr = old_out, old_err
            self.q.put("__DONE__")


def main():
    fix_stdio()
    root = tk.Tk()
    App(root)
    if "--selftest" in sys.argv:        # 打包后自检: 界面3秒自动关闭
        root.after(3000, root.destroy)
    root.mainloop()


if __name__ == "__main__":
    main()
