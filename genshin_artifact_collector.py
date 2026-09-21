"""
原神圣遗物自动采集工具
功能：自动打开原神 → 识别圣遗物总数 → 逐页点击+截图保存
"""
import ctypes
import threading
import time
import os
import re
import cv2
import numpy as np
import pyautogui
from PIL import ImageGrab, ImageEnhance, Image
import pytesseract

# ========== 通用配置 ==========
try:
    ctypes.windll.shcore.SetProcessDpiAwareness(2)
except Exception:
    ctypes.windll.user32.SetProcessDPIAware()

# ========== 打开原神配置 ==========
ICON_PATH = r"./icon.png"
THRESHOLD = 0.80
SCALES = [1.0, 0.75, 0.5, 1.25, 1.5, 2.0]
CLICK_INTERVAL = 0.3
GAME_WAIT = 15

# ========== 圣遗物采集配置 ==========
START_X = 150
START_Y = 200
X_STEP  = 150
Y_STEP  = 175
X_REPEATS = 7
Y_REPEATS = 4
CLICK_DELAY = 0.15
SHOT_REGION = (1310, 120, 1800, 960)
NUM_REGION  = (1520, 0, 1820, 90)
SAVE_DIR = "screenshots"
GAME_LAUNCH_DELAY = 3

ITEMS_PER_ROW  = X_REPEATS + 1
ROWS_PER_PAGE  = Y_REPEATS + 1
ITEMS_PER_PAGE = ITEMS_PER_ROW * ROWS_PER_PAGE

# ========== Windows API ==========
user32 = ctypes.windll.user32
MOUSEEVENTF_LEFTDOWN = 0x0002
MOUSEEVENTF_LEFTUP   = 0x0004
MOUSEEVENTF_WHEEL    = 0x0800
WHEEL_DELTA          = 120

os.makedirs(SAVE_DIR, exist_ok=True)


# ========================================================
#  第一阶段：打开原神
# ========================================================

def load_template(path):
    tpl = cv2.imread(path, cv2.IMREAD_COLOR)
    if tpl is None:
        raise FileNotFoundError(f"模板读取失败: {path}")
    return tpl

def grab_frame(region=None):
    shot = ImageGrab.grab(bbox=region)
    return cv2.cvtColor(np.asarray(shot), cv2.COLOR_RGB2BGR)

def locate_all(frame, template, threshold=THRESHOLD, scales=SCALES):
    hits = []
    for s in scales:
        t = cv2.resize(template, None, fx=s, fy=s, interpolation=cv2.INTER_AREA)
        th, tw = t.shape[:2]
        if th >= frame.shape[0] or tw >= frame.shape[1]:
            continue
        res = cv2.matchTemplate(frame, t, cv2.TM_CCOEFF_NORMED)
        ys, xs = np.where(res >= threshold)
        for x, y in zip(xs, ys):
            cx, cy = int(x + tw / 2), int(y + th / 2)
            if all(abs(cx - px) > tw // 2 or abs(cy - py) > th // 2 for px, py in hits):
                hits.append((cx, cy))
    return hits

def open_genshin():
    """用模板匹配找到原神图标并点击打开"""
    tpl = load_template(ICON_PATH)
    print(f"[打开原神] 模板: {ICON_PATH}  尺寸: {tpl.shape[1]}x{tpl.shape[0]}")
    frame = grab_frame()
    pts = locate_all(frame, tpl)
    if pts:
        cx, cy = pts[0]
        pyautogui.moveTo(cx, cy)
        pyautogui.click()
        print(f"[打开原神] 点击图标 ({cx}, {cy})")
    else:
        print("[打开原神] 未找到原神图标，假设原神已在前台")
    print(f"[打开原神] 等待 {GAME_WAIT} 秒让游戏加载...")
    time.sleep(GAME_WAIT)
    print("[打开原神] 游戏应已就绪，开始采集")


# ========================================================
#  第二阶段：圣遗物采集（点击+截图保存）
# ========================================================

shot_idx = 0
lock = threading.Lock()

def take_screenshot():
    global shot_idx
    with lock:
        shot_idx += 1
        idx = shot_idx
    img = ImageGrab.grab(bbox=SHOT_REGION)
    path = os.path.join(SAVE_DIR, f"shot_{idx:03d}.png")
    img.save(path)

def take_num_screenshot():
    img = ImageGrab.grab(bbox=NUM_REGION)
    img.save("num.png")
    print("已保存截图 num.png")

def read_text():
    if os.path.exists("num.png"):
        img = Image.open("num.png")
    else:
        img = ImageGrab.grab(bbox=NUM_REGION)
    img = img.resize((img.width * 3, img.height * 3))
    img = img.convert("L")
    img = ImageEnhance.Contrast(img).enhance(2.0)
    text = pytesseract.image_to_string(img, lang="chi_sim+eng", config="--psm 7")
    text = text.strip()
    print(f"识别到的文本: {text}")
    match = re.search(r"(\d+)\s*/", text)
    if match:
        value = int(match.group(1))
        print(f"/ 前面的数值: {value}")
        return value
    print("未找到 / 前面的数值")
    return None

def scroll_up(times=10):
    for _ in range(times):
        user32.mouse_event(MOUSEEVENTF_WHEEL, 0, 0, -WHEEL_DELTA, 0)
        time.sleep(0.05)

def do_click(x, y):
    user32.SetCursorPos(x, y)
    user32.mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
    user32.mouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)

def click_and_shoot(x, y):
    t = threading.Thread(target=take_screenshot)
    t.start()
    do_click(x, y)
    t.join()

def process_page(items_count):
    y = START_Y
    full_rows = items_count // ITEMS_PER_ROW
    extra = items_count % ITEMS_PER_ROW
    for row in range(full_rows):
        if row > 0:
            y += Y_STEP
        click_and_shoot(START_X, y)
        for i in range(1, ITEMS_PER_ROW):
            click_and_shoot(START_X + X_STEP * i, y)
            time.sleep(CLICK_DELAY)
    if extra > 0:
        if full_rows > 0:
            y += Y_STEP
        click_and_shoot(START_X, y)
        for i in range(1, extra):
            click_and_shoot(START_X + X_STEP * i, y)
            time.sleep(CLICK_DELAY)


# ========================================================
#  主流程
# ========================================================

def main():
    print(f"{GAME_LAUNCH_DELAY} 秒后开始...")
    time.sleep(GAME_LAUNCH_DELAY)

    # 第一阶段：打开原神
    open_genshin()

    # 第二阶段：采集圣遗物
    print("\n===== 开始圣遗物采集 =====")
    take_num_screenshot()
    total = read_text()
    if total is None:
        print("无法识别圣遗物总数，退出")
        return

    full_pages = total // ITEMS_PER_PAGE
    remaining = total % ITEMS_PER_PAGE
    total_pages = full_pages + (1 if remaining > 0 else 0)
    print(f"总数: {total}，每页{ITEMS_PER_PAGE}，完整页: {full_pages}，剩余: {remaining}，总页数: {total_pages}")

    for page in range(total_pages):
        if page < full_pages:
            items_this_page = ITEMS_PER_PAGE
        else:
            items_this_page = remaining
        print(f"--- 第 {page + 1}/{total_pages} 页，本页 {items_this_page} 个 ---")
        process_page(items_this_page)
        if page < total_pages - 1:
            scroll_up(49)
            time.sleep(0.5)

    print(f"采集完成，共截图 {shot_idx} 张，保存在 {SAVE_DIR}/ 目录")


if __name__ == "__main__":
    main()
