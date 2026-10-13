# -*- coding: utf-8 -*-
"""
屏幕自动化任务程序
- 直接运行：python auto_click.py
- 打包后：双击 auto_click.exe

新增步骤：在 main() 的「任务步骤区」继续往下添加 click / double_click / right_click 即可
"""

import time
import ctypes

import numpy as np
import mss

# ============ 底层：Win32 鼠标控制（零第三方依赖） ============
user32 = ctypes.windll.user32

MOUSE_LEFT_DOWN = 0x0002
MOUSE_LEFT_UP = 0x0004
MOUSE_RIGHT_DOWN = 0x0008
MOUSE_RIGHT_UP = 0x0010


def move(x, y):
    """移动鼠标到屏幕坐标 (x, y)"""
    user32.SetCursorPos(int(x), int(y))


def click(x, y, delay=0.05):
    """移动到 (x, y) 并单击鼠标左键"""
    move(x, y)
    time.sleep(delay)
    user32.mouse_event(MOUSE_LEFT_DOWN, 0, 0, 0, 0)
    time.sleep(delay)
    user32.mouse_event(MOUSE_LEFT_UP, 0, 0, 0, 0)


def double_click(x, y, delay=0.05):
    """移动到 (x, y) 并双击左键"""
    click(x, y, delay)
    time.sleep(0.1)
    click(x, y, delay)


def right_click(x, y, delay=0.05):
    """移动到 (x, y) 并单击右键"""
    move(x, y)
    time.sleep(delay)
    user32.mouse_event(MOUSE_RIGHT_DOWN, 0, 0, 0, 0)
    time.sleep(delay)
    user32.mouse_event(MOUSE_RIGHT_UP, 0, 0, 0, 0)


def step(msg):
    """打印带时间戳的步骤日志"""
    print("[%s] %s" % (time.strftime("%H:%M:%S"), msg), flush=True)


# ============ 底层：区域截图 + OCR（纯内存管线，引擎只加载一次） ============
_sct = None        # 全局复用的 mss 截图实例
_ocr_engine = None  # 全局复用的 OCR 引擎


def _get_sct():
    """mss 实例全局复用，避免每次截图都创建/销毁 GDI 句柄"""
    global _sct
    if _sct is None:
        # mss>=10 用 MSS 类；老版本用 mss() 工厂函数
        _sct = mss.MSS() if hasattr(mss, "MSS") else mss.mss()
    return _sct


def _get_ocr():
    """OCR 引擎懒加载、全局复用：模型只加载一次（性能关键）
    默认配置：PP-OCRv6 small 检测+识别模型，ONNX Runtime 推理，线程数自动
    """
    global _ocr_engine
    if _ocr_engine is None:
        from rapidocr import RapidOCR
        _ocr_engine = RapidOCR(params={
            # 屏幕文字均为正向，跳过方向分类器（省掉一个模型前向，提速 15%~25%）
            "Global.use_cls": False,
            # 只保留警告以上日志，输出干净
            "Global.log_level": "warning",
        })
    return _ocr_engine


def capture_region(x1, y1, x2, y2):
    """截图屏幕矩形区域 (x1,y1)~(x2,y2)，返回 BGR numpy 数组（纯内存，不落盘）"""
    left, top = int(x1), int(y1)
    w, h = int(x2) - left, int(y2) - top
    img = _get_sct().grab({"left": left, "top": top, "width": w, "height": h})
    return np.asarray(img)[:, :, :3].copy()  # BGRA -> BGR


def ocr_region(x1, y1, x2, y2, min_score=0.5):
    """截图指定区域并 OCR 识别。
    返回 [(文本, 置信度), ...]，按阅读顺序（自上而下、从左到右）排列。
    min_score: 置信度过滤阈值，低于该值的结果丢弃。
    """
    arr = capture_region(x1, y1, x2, y2)
    out = _get_ocr()(arr)
    if out is None or out.txts is None:
        return []
    items = []
    for box, txt, score in zip(out.boxes, out.txts, out.scores):
        if float(score) < min_score:
            continue
        cx = float(np.mean([p[0] for p in box]))
        cy = float(np.mean([p[1] for p in box]))
        items.append((cy, cx, str(txt), float(score)))
    # 按中心 y 聚行（20px 一行），行内按 x 排序，得到阅读顺序
    items.sort(key=lambda t: (int(t[0] // 20), t[1]))
    return [(t[2], t[3]) for t in items]


# ============ 任务步骤区（以后加步骤就在这里往下写） ============
def main():
    step("步骤一：点击屏幕坐标 (180, 250)")
    click(180, 250)

    step("步骤二：截图区域 (1308,119)~(1800,961) 并 OCR 识别")
    t0 = time.perf_counter()
    lines = ocr_region(1308, 119, 1800, 961)
    cost = (time.perf_counter() - t0) * 1000
    if lines:
        for txt, score in lines:
            print("    %-40s 置信度 %.2f" % (txt, score))
    else:
        print("    （该区域未识别到文字）")
    step("步骤二完成：识别出 %d 行，耗时 %.0f ms" % (len(lines), cost))

    # 步骤三示例（需要时去掉注释，或继续往下添加）：
    # step("步骤三：双击屏幕坐标 (500, 300)")
    # double_click(500, 300)

    step("全部步骤执行完毕")


if __name__ == "__main__":
    main()
