# -*- coding: utf-8 -*-
"""
原神圣遗物背包遍历+采集+解析脚本(主入口, 只做命令行调度)

依赖: pip install pyautogui pillow rapidocr-onnxruntime

模块拆分(按功能):
  genshin_config.py    共享配置: 坐标/步长/路径 + OCR引擎单例
  genshin_capture.py   采集功能: 函数1数量识别 + 函数2遍历点击截图
  genshin_analyzer.py  分析功能: 函数3单张解析 + 函数4批量解析
  genshin_pipeline.py  并行模式: 采集截图的同时后台线程实时OCR解析

运行方式(与拆分前完全一致):
  采集:         python genshin_artifact.py
  单张解析:     python genshin_artifact.py parse 001   (序号或完整图片路径)
  批量解析:     python genshin_artifact.py analyze [run目录]
  采集并行分析: python genshin_artifact.py pipeline   (采集完成分析基本同步完成)
紧急中止: 鼠标快速甩到屏幕左上角触发 pyautogui 保险机制

说明: 各模式按需延迟导入对应模块 ——
  解析模式不会加载 pyautogui(不碰鼠标键盘), 采集模式才会加载采集模块
"""
import glob
import os
import sys

import genshin_config  # noqa: F401  导入即完成DPI设置与OCR引擎初始化
from genshin_config import SHOT_ROOT


def _resolve_parse_path(arg):
    """parse 参数 -> 截图路径。支持完整图片路径或序号(序号默认取最近一次run目录)"""
    if os.path.isfile(arg):
        return arg
    runs = sorted(glob.glob(os.path.join(SHOT_ROOT, "run_*")))
    if not runs:
        print(f"在 {SHOT_ROOT} 下找不到截图目录")
        return None
    return os.path.join(runs[-1], f"artifact_{int(arg):03d}.png")


def main():
    # --- 解析模式: python genshin_artifact.py parse 001 ---
    if len(sys.argv) >= 2 and sys.argv[1] == "parse":
        from genshin_analyzer import analyze_artifact
        arg = sys.argv[2] if len(sys.argv) > 2 else "001"
        path = _resolve_parse_path(arg)
        if not path:
            sys.exit(1)
        data = analyze_artifact(path)
        if data:
            print(f"\n===== {os.path.basename(path)} 解析结果 =====")
            for k, v in data.items():
                print(f"{k}：{v}")
        return

    # --- 批量解析模式: python genshin_artifact.py analyze [run目录] ---
    if len(sys.argv) >= 2 and sys.argv[1] == "analyze":
        from genshin_analyzer import batch_analyze
        batch_analyze(sys.argv[2] if len(sys.argv) > 2 else None)
        return

    # --- 采集+并行分析模式: python genshin_artifact.py pipeline ---
    if len(sys.argv) >= 2 and sys.argv[1] == "pipeline":
        from genshin_pipeline import run_pipeline
        run_pipeline()
        return

    # --- 采集模式(默认): python genshin_artifact.py ---
    from genshin_capture import run_capture
    run_capture()


if __name__ == "__main__":
    main()
