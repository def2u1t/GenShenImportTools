# -*- coding: utf-8 -*-
"""
采集功能模块(需要操作鼠标/截屏, 只在采集模式加载)

  get_artifact_count()  - 函数1: 截图OCR识别圣遗物数量
  traverse_artifacts()  - 函数2: 遍历点击圣遗物并截图详情面板(点满数量自动停止;
                          尾页按剩余行数翻页, 翻页后从倒数第r行开始点击)
  run_capture()        - 采集主流程(倒计时 -> 识别数量 -> 遍历截图)

依赖: pyautogui pillow rapidocr-onnxruntime
紧急中止: 鼠标快速甩到屏幕左上角触发 pyautogui 保险机制
"""
import os
import re
import time
from datetime import datetime

import pyautogui
from PIL import ImageGrab

from genshin_config import (COUNT_REGION, START_POS, DETAIL_REGION,
                            STEP_X, H_CLICKS, STEP_Y, Y_OFFSET_LIMIT,
                            SCROLL_TICK, SCROLL_TICKS, SCROLL_TICK_DELAY,
                            CLICK_DELAY, SCROLL_DELAY, SHOT_ROOT,
                            ocr)


# ==================== 函数1: 识别圣遗物数量 ====================
def get_artifact_count(region=COUNT_REGION, retries=3, interval=1.0):
    """
    截图数量显示区域并OCR识别（失败自动重试）。

    参数:
        region:   截图区域 (左上x, 左上y, 右下x, 右下y)
        retries:  最大尝试次数
        interval: 每次重试间隔秒数

    返回:
        (当前数量, 背包上限) 元组，如 (547, 2700)
    """
    last_text = ""
    for attempt in range(1, retries + 1):
        img = ImageGrab.grab(bbox=region)
        result, _ = ocr(img)

        # 合并所有识别文本并去掉空格，匹配 "当前/上限" 格式
        last_text = "".join(line[1] for line in result).replace(" ", "") if result else ""
        m = re.search(r"(\d+)/(\d+)", last_text)
        if m:
            return int(m.group(1)), int(m.group(2))

        print(f"[第{attempt}/{retries}次] 未识别到数量格式，识别文本: {last_text!r}，{interval}秒后重试...")
        time.sleep(interval)

    raise RuntimeError(f"OCR 连续{retries}次未解析到数量格式，最后识别文本: {last_text!r}")


# ==================== 函数2: 遍历点击并截图 ====================
COLS = H_CLICKS + 1                 # 每行格子数 = 行首1个 + 累加7次 = 8
PER_PAGE = Y_OFFSET_LIMIT * COLS    # 每页格子数 = 5行 x 8列 = 40


def _rows_needed(n):
    """n 个物品在屏幕上占用的行数(整数向上取整, 上限 Y_OFFSET_LIMIT 行); n<=0 返回 0"""
    if n <= 0:
        return 0
    rows = -(-n // COLS)            # 等价于 ceil(n / COLS) 的整数写法
    return rows if rows < Y_OFFSET_LIMIT else Y_OFFSET_LIMIT


def traverse_artifacts(artifact_count=None, pages=1,
                       click_delay=CLICK_DELAY, scroll_delay=SCROLL_DELAY,
                       run_dir=None, on_captured=None, in_memory=False):
    """
    遍历点击背包圣遗物格子，每次点击后截图详情面板:

    每页 5行 x 8列 = 40个, 第k页点击 min(40, 剩余数量) 个:
    1. 本页数据占 page_rows = ceil(本页个数/8) 行
    2. 第一页: 列表从顶部开始显示, 从第一行(175,200)开始点击(未满页也从第一行);
       翻页后的尾页: 新数据只占屏幕底部 page_rows 行 -> 从"倒数第 page_rows 行"
       开始点击, 即起始 y 偏移 (5-page_rows) 次 (y=380/560/740/920 对应剩4/3/2/1行)
    3. 每行从行首(175,y)开始, x+145 累加点击本行剩余的格子
    4. 页点完后按"下一页数据行数"换页: 滚动格数 = round(49 x 行数/5),
       即满页49格(5行)/39格(4行)/29格(3行)/20格(2行)/10格(1行)
       (游戏对单次大滚动只按一格处理, 所以逐次下滚一格)
    5. 点击计数达到 artifact_count 时, 完成最后一次截图后立即停止
       (不再点击、不再截图、不再翻页)

    参数:
        artifact_count: 圣遗物总数。点满该数量即停止遍历(None=点满 pages 页)
        pages:          最大页数上限(保险用)
        click_delay:    每次点击后等待秒数
        scroll_delay:   换页滚动全部完成后的等待秒数
        run_dir:        截图保存目录(None=自动创建时间戳目录; 传入则复用)
        on_captured:    截图回调(见下)
        in_memory:      截图不落盘, 直接把 (文件名, PIL Image) 元组交给回调(内存模式)

    回调约定:
        in_memory=False: on_captured(path)          # 截图已落盘
        in_memory=True:  on_captured((name, img))   # 图像仅在内存

    返回:
        总点击/截图数
    """
    # 每次运行使用独立的时间戳子目录，避免覆盖
    if run_dir is None:
        run_dir = os.path.join(SHOT_ROOT, datetime.now().strftime("run_%Y%m%d_%H%M%S"))
    os.makedirs(run_dir, exist_ok=True)  # 内存模式下仅存放 artifacts.json 等结果文件
    total = f"/{artifact_count}" if artifact_count else ""
    print(f"截图保存目录: {run_dir}")

    clicked = 0  # 点击计数

    def click_cell(x, y):
        """点击格子 -> 等待刷新 -> 截图详情面板(落盘保存或仅入内存)"""
        nonlocal clicked
        pyautogui.click(x, y)
        clicked += 1
        time.sleep(click_delay)  # 等详情面板刷新
        name = f"artifact_{clicked:03d}"
        if in_memory:
            img = ImageGrab.grab(bbox=DETAIL_REGION)   # 不落盘
            if on_captured:
                on_captured((name, img))                # 内存模式: 传(文件名, 图像)
            print(f"  [{clicked:03d}{total}] 点击({x},{y}) -> {name}(内存)")
        else:
            path = os.path.join(run_dir, f"{name}.png")
            ImageGrab.grab(bbox=DETAIL_REGION).save(path)
            if on_captured:
                on_captured(path)                      # 落盘模式: 传路径
            print(f"  [{clicked:03d}{total}] 点击({x},{y}) -> {os.path.basename(path)}")

    def finished():
        """是否已点满圣遗物数量(最后一次截图已完成)"""
        return artifact_count is not None and clicked >= artifact_count

    def click_row(y, cells):
        """从行首(x=START_POS[0], y)开始点击一行中的 cells 个格子; 返回是否已点满"""
        x = START_POS[0]
        for i in range(cells):
            if i > 0:
                x += STEP_X
            click_cell(x, y)
            if finished():
                return True
        return False

    page = 0
    while page < pages:
        # --- 计算本页要点的数量 / 占用行数 / 起始行 ---
        remaining = artifact_count - clicked if artifact_count is not None else PER_PAGE
        page_items = min(PER_PAGE, remaining)
        if page_items <= 0:
            break
        page_rows = _rows_needed(page_items)           # 本页数据占用的行数
        if page == 0:
            start_offset = 0   # 第一页: 列表从顶部开始显示(未满页也从第一行开始点)
        else:
            start_offset = Y_OFFSET_LIMIT - page_rows   # 翻页后: 新数据贴屏幕底部, 从倒数第page_rows行开始
        start_y = START_POS[1] + start_offset * STEP_Y
        print(f"===== 第{page + 1}页 (本页{page_items}个/{page_rows}行, 起始y偏移{start_offset}次) =====")

        # --- 逐行点击: 每行从行首开始, 点击本行剩余的格子 ---
        for r in range(page_rows):
            y = start_y + r * STEP_Y
            cells = min(COLS, page_items - r * COLS)    # 本行格子数
            print(f"  -- 第{r + 1}/{page_rows}行: 行首({START_POS[0]},{y}) 点击{cells}个 --")
            if click_row(y, cells):
                return _done(clicked, artifact_count)

        # --- y方向偏移第5次: 不偏移，按下一页数据行数分多次连续下滚换页 ---
        if finished():
            return _done(clicked, artifact_count)
        next_items = min(PER_PAGE, artifact_count - clicked) if artifact_count is not None else PER_PAGE
        next_rows = _rows_needed(next_items)
        scroll_ticks = round(SCROLL_TICKS * next_rows / Y_OFFSET_LIMIT)
        print(f"  y第{Y_OFFSET_LIMIT}次偏移 -> 不偏移, 连续下滚{scroll_ticks}格(每次{abs(SCROLL_TICK)})换页"
              f" -> 下一页{next_rows}行, 从倒数第{next_rows}行开始点击")
        # 暂时关闭pyautogui内置的每次调用0.1秒停顿(卡顿主因), 滚动节奏交给 SCROLL_TICK_DELAY 控制
        pyautogui.PAUSE = 0
        try:
            for _ in range(scroll_ticks):
                pyautogui.scroll(SCROLL_TICK)
                time.sleep(SCROLL_TICK_DELAY)
        finally:
            pyautogui.PAUSE = 0.1  # 恢复默认, 保证点击节奏与已校准值一致
        time.sleep(scroll_delay)

        page += 1

    return _done(clicked, artifact_count)


def _done(clicked, artifact_count):
    """遍历结束提示，返回总点击数"""
    target = f"(目标 {artifact_count})" if artifact_count is not None else ""
    print(f"遍历完成: 共点击/截图 {clicked} 个 {target}")
    return clicked


# ==================== 采集主流程 ====================
def run_capture(countdown=3):
    """
    采集主流程: 倒计时 -> OCR识别数量 -> 遍历点击+截图(点满数量自动停止)

    参数:
        countdown: 开始前倒计时秒数(留时间切到原神背包界面)

    返回:
        总点击/截图数
    """
    print("=" * 46)
    print("原神圣遗物背包遍历+采集脚本")
    print("运行后请立即切到原神圣遗物背包界面!")
    print("紧急中止: 鼠标快速甩到屏幕左上角")
    print("=" * 46)
    for i in range(countdown, 0, -1):
        print(f"{i}秒后开始...")
        time.sleep(1)

    # ---- 函数1: OCR识别圣遗物数量 ----
    artifact_count, artifact_total = get_artifact_count()
    print(f"圣遗物数量: {artifact_count} / {artifact_total}")

    # ---- 函数2: 遍历点击+截图, 点满数量自动停止 ----
    pages = artifact_count // PER_PAGE + 3       # 页数上限(保险用,+3页富余; PER_PAGE=40个/页)
    return traverse_artifacts(artifact_count=artifact_count, pages=pages)
