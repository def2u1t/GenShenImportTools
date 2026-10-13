# -*- coding: utf-8 -*-
"""
采集并行分析模块: 采集点击截图的同时, 后台线程立即OCR解析,
采集结束时分析基本同步完成, 无需再等十几分钟批量解析。

  run_pipeline() - 倒计时 -> 识别数量 -> 启动分析线程 -> 遍历采集(截图即入队)
                   -> 队列排空 -> 汇总 artifacts.json

运行: python genshin_pipeline.py
  或: python genshin_artifact.py pipeline

架构:
  主线程[采集]  点击格子 -> 截图落盘 -> 入队 (复用 genshin_capture.traverse_artifacts)
  分析线程[分析] 出队 -> analyze_artifact OCR解析 -> 结果缓存 -> 每10张增量保存JSON
  两线程通过 queue 通信; OCR单线程顺序调用(线程安全), 采集期间主线程不碰OCR引擎。

落盘模式(默认, in_memory=False):
  截图保存到 run_dir, 行为与原并行模式一致;
  截图可事后复核/重新解析(排障时能直观发现抓图区域偏移等问题)。

内存模式(可选, in_memory=True):
  截图不落盘, (文件名, PIL Image) 元组直接入队, 分析完即释放引用;
  结束仅保留 artifacts.json。
  注意: 曾在实战中出现过抓图内容偏移无法复核的情况, 建议优先用默认落盘模式。

依赖: pyautogui pillow rapidocr-onnxruntime
紧急中止: 鼠标快速甩到屏幕左上角触发 pyautogui 保险机制
"""
import json
import os
import queue
import re
import threading
import time
from datetime import datetime

import genshin_config  # noqa: F401  导入即完成DPI设置与OCR引擎初始化
from genshin_analyzer import analyze_artifact
from genshin_capture import get_artifact_count, traverse_artifacts, PER_PAGE
from genshin_config import SHOT_ROOT

_print_lock = threading.Lock()


def log(tag, msg):
    """线程安全日志, [采集]/[分析]/[系统] 前缀区分来源"""
    with _print_lock:
        print(f"[{tag}] {msg}", flush=True)


def _save(results, out_path):
    """按序号排序保存JSON(先写临时文件再原子替换, 防中途写坏)"""
    ordered = [results[k] for k in sorted(results)]
    tmp = out_path + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(ordered, f, ensure_ascii=False, indent=1)
    os.replace(tmp, out_path)


def _fmt_sec(sec):
    """秒 -> "分:秒"格式, 如 123 -> 2:03 (与genshin_analyzer一致)"""
    m, s = divmod(int(max(0, sec)), 60)
    return f"{m}:{s:02d}"


def _analyze_worker(q, results, errors, out_path, total=0, save_every=10):
    """
    分析线程主体: 逐张OCR解析队列中的截图, 增量保存JSON。

    队列消息协议:
      None            -> 采集结束, 存盘后退出
      (name, img)元组 -> 内存模式截图, OCR后引用即释放
      path字符串      -> 落盘模式截图, 读取文件OCR

    逐张日志(沿用genshin_analyzer批量解析的信息要素):
      当前进度N/总数 套装名 耗时X.Xs 预计还有M:SS完成

    OCR 约2秒/张慢于采集(约0.1~0.2秒/张), 采集结束后本线程继续消化积压,
    总耗时 ≈ OCR总时长(省掉了原先"先采集再批量解析"的串行等待)。
    """
    n_done = 0
    t_all = time.time()
    while True:
        item = q.get()
        if item is None:
            break
        if isinstance(item, tuple):
            name, img = item      # 内存模式
            src = img
        else:
            name, img = os.path.basename(item), None  # 落盘模式
            src = item
        t0 = time.time()
        try:
            data = analyze_artifact(src)
        except Exception as e:  # 单张异常不中断整个分析线程
            data = None
            log("分析", f"{name} 解析异常: {e}")
        dt = time.time() - t0
        if img is not None:
            img.close()  # 内存模式: 立即释放图像内存
        if data:
            num = int(re.search(r"(\d+)", name).group(1))
            results[num] = {"序号": num, **data}
            set_name = data.get("套装类型") or "<套装缺失>"
        else:
            errors.append(name)
            set_name = "<解析失败>"
        n_done += 1
        if total:
            avg = (time.time() - t_all) / n_done
            remain = avg * max(0, total - n_done)
            log("分析", f"当前进度{n_done}/{total} {set_name} 耗时{dt:.1f}s "
                       f"预计还有{_fmt_sec(remain)}完成")
        if n_done % save_every == 0:
            _save(results, out_path)  # 增量保存, 中断不丢
    _save(results, out_path)
    log("分析", f"分析线程结束: 共解析 {len(results)} 张, 失败 {len(errors)} 张")


def run_pipeline(countdown=3, in_memory=False):
    """
    采集并行分析主流程:

    1. 倒计时(留时间切到游戏背包界面)
    2. OCR识别圣遗物数量(停止条件)
    3. 启动分析线程
    4. 遍历点击+截图, 每张截图实时入队给分析线程
    5. 采集完成后等队列排空, 汇总保存 artifacts.json

    参数:
        countdown:  开始前倒计时秒数
        in_memory: False=落盘模式(默认): 截图保留在 run_dir, 可复核/重解析
                   True=内存模式: 截图不落盘, 结束仅保留 artifacts.json

    返回:
        分析结果字典 {序号: 数据}
    """
    print("=" * 50)
    print("原神圣遗物 采集+并行分析 脚本")
    mode = "落盘(截图保留可复核)" if not in_memory else "内存(截图不落盘,结束仅保留JSON)"
    print(f"模式: {mode}")
    print("运行后请立即切到原神圣遗物背包界面!")
    print("紧急中止: 鼠标快速甩到屏幕左上角")
    print("=" * 50)
    for i in range(countdown, 0, -1):
        print(f"{i}秒后开始...")
        time.sleep(1)

    # ---- 步骤1: 识别圣遗物数量 ----
    artifact_count, artifact_total = get_artifact_count()
    log("采集", f"圣遗物数量: {artifact_count} / {artifact_total}")

    # ---- 步骤2: 准备本次运行目录与分析线程 ----
    run_dir = os.path.join(SHOT_ROOT, datetime.now().strftime("run_%Y%m%d_%H%M%S"))
    os.makedirs(run_dir, exist_ok=True)
    out_path = os.path.join(run_dir, "artifacts.json")
    log("系统", f"截图与结果目录: {run_dir}")

    q = queue.Queue()
    results, errors = {}, []
    worker = threading.Thread(
        target=_analyze_worker, args=(q, results, errors, out_path, artifact_count),
        daemon=True, name="AnalyzerThread")
    worker.start()

    pages = artifact_count // PER_PAGE + 3       # 页数上限(保险用,+3页富余; PER_PAGE=40个/页)

    # ---- 步骤3: 遍历采集, 每张截图实时入队 ----
    t0 = time.time()
    try:
        clicked = traverse_artifacts(
            artifact_count=artifact_count, pages=pages, run_dir=run_dir,
            on_captured=lambda item: q.put(item), in_memory=in_memory)
    finally:
        # ---- 步骤4: 通知分析线程收尾, 等积压队列消化完 ----
        q.put(None)
        worker.join()

    dt = time.time() - t0
    log("系统", f"全部完成: 采集 {clicked} 张, 分析 {len(results)} 张, "
              f"失败 {len(errors)} 张, 总耗时 {dt:.0f} 秒")
    log("系统", f"结果已保存: {out_path}")
    if errors:
        log("系统", f"失败列表: {errors}")
    return results


# ==================== 直接运行入口 ====================
if __name__ == "__main__":
    run_pipeline()
