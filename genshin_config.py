# -*- coding: utf-8 -*-
"""
共享配置模块: 坐标/步长/滚动参数、截图路径、OCR引擎单例

采集(genshin_capture)与分析(genshin_analyzer)两个模块共用的部分统一放这里:
- 坐标参数改动只改这一个文件
- OCR引擎全局只初始化一次, 两个模块共用同一实例(避免重复加载模型)
- DPI感知设置在导入本模块时即生效(必须先于任何点击/截屏)

依赖: rapidocr-onnxruntime
"""
import ctypes

from rapidocr_onnxruntime import RapidOCR

# ==================== 环境初始化 ====================
# DPI感知: 保证点击/截图使用物理像素坐标(100%缩放下无影响)
try:
    ctypes.windll.user32.SetProcessDPIAware()
except Exception:
    pass

# ==================== 数量识别配置 ====================
COUNT_REGION = (1600, 0, 1800, 80)   # 圣遗物数量显示区域(左上x,左上y,右下x,右下y)

# ==================== 遍历采集配置 ====================
START_POS = (175, 200)    # 每页起始点击位置(第一行行首)
DETAIL_REGION = (1306, 118, 1803, 964)  # 详情面板截图区域(左上x,左上y,右下x,右下y)
STEP_X = 145              # 横向点击步长(px)
H_CLICKS = 7              # 每行横向累加点击次数(含行首共8个格子)
STEP_Y = 180              # 纵向偏移步长(px)
Y_OFFSET_LIMIT = 5        # y方向偏移第5次时不偏移，改为滚动换页
SCROLL_TICK = -85        # 单次滚动量(负数=向下; 单次发大数值无效,游戏只按一格处理)
SCROLL_TICKS = 49         # 换页需要滚动的格数(每次一格,逐次发送)
SCROLL_TICK_DELAY = 0.02  # 每次滚动之间的间隔(秒),约1个游戏帧;过短会被游戏合并吞掉
CLICK_DELAY = 0.001         # 每次点击后等待界面刷新(秒)
SCROLL_DELAY = 0.001        # 换页全部滚动完成后的等待(秒)

# ==================== 截图保存 ====================
SHOT_ROOT = "screenshots"  # 截图保存根目录(每次运行建带时间戳的子目录)

# ==================== OCR引擎(全局单例) ====================
# 导入即初始化; 采集/分析共用, 不要在业务模块里另建 RapidOCR 实例
print("[init] 正在加载OCR引擎(首次约1~2秒)...", flush=True)
ocr = RapidOCR()
print("[init] OCR引擎加载完成", flush=True)
