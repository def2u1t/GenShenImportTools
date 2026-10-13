# -*- coding: utf-8 -*-
"""
分析功能模块(纯离线: 只读截图文件, 不操作鼠标/键盘/屏幕)

  analyze_artifact() - 函数3: OCR解析单张详情截图, 提取结构化数据
  batch_analyze()    - 函数4: 批量解析一次运行目录下的全部截图, 汇总JSON

依赖: rapidocr-onnxruntime (本模块不加载 pyautogui)
直接运行: python genshin_analyzer.py 即批量解析最近一次run目录
"""
import glob
import json
import os
import re
import time

from genshin_config import SHOT_ROOT, ocr

# ==================== 解析规则常量 ====================
SLOT_NAMES = ("生之花", "死之羽", "时之沙", "空之杯", "理之冠")


# ==================== 函数3: OCR解析圣遗物详情截图 ====================
def analyze_artifact(source):
    """
    OCR解析单张圣遗物详情截图，提取结构化数据。

    参数:
        source: 截图文件路径(str) 或 PIL Image 对象(内存模式, RapidOCR直接支持)

    识别规则(基于详情面板自上而下的文本行):
    - 名称:   最顶部一行
    - 部位:   匹配五大部位名
    - 主属性: 部位行与星级行之间的 名称行+数值行，合并为 "名称 数值"
    - 等级:   +数字 格式(兼容OCR反序如'6+')
    - 词条:   位于等级行与套装说明之间、含"词条名+数值"的行(按顺序即词条1~4)
              带"（待激活）"的词条保留原样标记
    - 套装类型: 首条"件套"说明文字上方最近的文本行(去尾部冒号)
    - 装备角色: 含"已装备"的行

    返回: 字典，如
        {"名称": "熔铸者的观测", "部位": "死之羽", "主属性": "攻击力 311",
         "等级": "+20", "词条1": "防御力+20.4%", ...,
         "套装类型": "炉火融炼之心", "装备角色": "Default已装备"}
    """
    result, _ = ocr(source)
    if not result:
        label = os.path.basename(source) if isinstance(source, str) else "<内存图像>"
        print(f"[警告] {label} OCR未识别到内容")
        return None

    # 按 y 坐标整理为自上而下的行 (y中心, 文本)
    rows = []
    for box, text, score in result:
        y = sum(p[1] for p in box) / 4
        rows.append((y, text.strip()))
    rows.sort(key=lambda r: r[0])
    texts = [t for _, t in rows]

    # 关键行定位
    LEVEL_RE = r"\+\d{1,2}|\d{1,2}\+"   # 等级: '+6'或OCR偶发反序的'6+'
    star_y = next((y for y, t in rows if "★" in t or "☆" in t), None)   # 星级行
    slot_y = next((y for y, t in rows if t in SLOT_NAMES), None)          # 部位行
    level_y = next((y for y, t in rows if re.fullmatch(LEVEL_RE, t)), None)  # 等级行
    setdesc_y = next((y for y, t in rows if "件套" in t), None)           # 首条套装效果行

    # 等级文本归一化: '6+' -> '+6'
    lv_text = next((t for t in texts if re.fullmatch(LEVEL_RE, t)), "")
    lv_num = re.search(r"\d{1,2}", lv_text)

    data = {
        "名称": texts[0] if texts else "",
        "部位": next((t for t in texts if t in SLOT_NAMES), ""),
        "主属性": "",
        "等级": f"+{lv_num.group()}" if lv_num else "",
    }

    # --- 主属性: 部位行与星级行之间的 名称行+数值行 ---
    main_name, main_val = "", ""
    for y, t in rows:
        if slot_y is not None and y <= slot_y:
            continue
        if star_y is not None and y >= star_y:
            break
        if re.fullmatch(r"[\d,]+(\.\d+)?%?", t):   # 数值行(兼容千位逗号/百分比)
            main_val = t.replace(",", "")
        elif not main_val and t not in (data["名称"], data["部位"]):
            main_name = t                            # 第一段文字行即主属性名
    data["主属性"] = f"{main_name} {main_val}".strip()

    # --- 副词条: 等级行与首条套装效果之间、含"+数值"的行 ---
    subs = []
    for y, t in rows:
        if level_y is not None and y <= level_y:
            continue
        if setdesc_y is not None and y >= setdesc_y:
            break
        t2 = re.sub(r"^[^\u4e00-\u9fa5]+", "", t).strip()  # 去掉词条前的符号前缀(·/?/乱码等)
        if "+" in t2 and re.search(r"\+\d+(\.\d+)?%?", t2):
            subs.append(t2)
    for i, s in enumerate(subs, 1):
        data[f"词条{i}"] = s

    # --- 套装类型: 首条"件套"说明上方最近的文本行(去尾部冒号) ---
    used = {data["名称"], data["部位"], data["等级"], main_name, main_val} | set(subs)
    set_name = ""
    if setdesc_y is not None:
        for y, t in rows:
            if y >= setdesc_y:
                break
            if "★" in t or "☆" in t or "·" in t or "+" in t or t in used:
                continue
            set_name = t                               # 持续覆盖，取最靠近套装效果的一行
    data["套装类型"] = set_name.rstrip("：:").strip()

    # --- 装备角色 ---
    data["装备角色"] = next((t for t in texts if "已装备" in t), "")

    return data


# ==================== 函数4: 批量解析全部截图 ====================
def _fmt_sec(sec):
    """秒数 -> "分:秒"格式, 如 123 -> 2:03"""
    m, s = divmod(int(sec), 60)
    return f"{m}:{s:02d}"


def batch_analyze(run_dir=None, out_path=None, verbose=True, log_file=True):
    """
    批量解析一次运行目录下的全部圣遗物截图，汇总保存为 JSON。

    详细日志(逐张实时输出, 避免长时间运行看不出是否卡住):
      每张:   [序号/总数] 文件名 单张耗时 名称|部位|等级|词条数|套装
      失败:   [序号/总数] 文件名 耗时 [失败] 原因
      每10张: 平均速度/已用时间/预计剩余时间
    日志同步落盘 run_dir/analyze_log.txt, 跑完可回查每张耗时与失败原因

    参数:
        run_dir:  截图目录(默认取最近一次运行 run_*)
        out_path: 输出JSON路径(默认 run_dir/artifacts.json)
        verbose:  逐张详细日志开关(失败行始终输出)
        log_file: 日志落盘开关

    返回:
        解析结果列表
    """
    if run_dir is None:
        runs = sorted(glob.glob(os.path.join(SHOT_ROOT, "run_*")))
        if not runs:
            print(f"在 {SHOT_ROOT} 下找不到截图目录")
            return []
        run_dir = runs[-1]
    if out_path is None:
        out_path = os.path.join(run_dir, "artifacts.json")

    images = sorted(glob.glob(os.path.join(run_dir, "artifact_*.png")))
    if not images:
        print(f"{run_dir} 下没有截图")
        return []

    total = len(images)
    print(f"开始批量解析: {total} 张截图 ({run_dir})")

    all_data, errors = [], []
    log_lines = []            # 与控制台同步的日志缓冲(结束后落盘)
    start = time.time()

    def log(line):
        """详细日志: 控制台实时输出(flush防管道缓冲) + 暂存待写文件"""
        print(line, flush=True)
        log_lines.append(line)

    for i, img in enumerate(images, 1):
        name = os.path.basename(img)
        t0 = time.time()
        data = analyze_artifact(img)
        dt = time.time() - t0

        if data:
            # 文件名序号即背包内顺序
            num = int(re.search(r"(\d+)", name).group(1))
            all_data.append({"序号": num, **data})
            if verbose:
                # 解析摘要: 一眼看懂这张图识别出了什么
                sub_count = sum(1 for k in data if k.startswith("词条"))
                pending = sum(1 for k in data if k.startswith("词条") and "待激活" in data[k])
                subs = f"{sub_count}词条" + (f"({pending}待激活)" if pending else "")
                parts = [data["名称"] or "<名称缺失>", data["部位"] or "<部位缺失>",
                         data["等级"] or "<等级缺失>", subs, data["套装类型"] or "<套装缺失>"]
                log(f"[{i:3d}/{total}] {name} {dt:4.1f}s {'|'.join(parts)}")
        else:
            errors.append(name)
            log(f"[{i:3d}/{total}] {name} {dt:4.1f}s [失败] OCR未识别到内容")

        # 速度统计: 每10张及最后一张输出一次, 让长时间运行有节奏感
        if verbose and (i % 10 == 0 or i == total):
            elapsed = time.time() - start
            avg = elapsed / i
            remain = avg * (total - i)
            log(f"  ---- 平均{avg:.2f}s/张, 已用{_fmt_sec(elapsed)}, 预计剩余{_fmt_sec(remain)} ----")

        # 增量保存: 每25张落盘一次, 中断不丢已解析进度
        if i % 25 == 0 or i == total:
            with open(out_path, "w", encoding="utf-8") as f:
                json.dump(all_data, f, ensure_ascii=False, indent=1)

    with open(out_path, "w", encoding="utf-8") as f:
        json.dump(all_data, f, ensure_ascii=False, indent=1)

    print(f"\n解析完成: 成功 {len(all_data)} / 共 {total} 张")
    if errors:
        print(f"失败列表({len(errors)}): {errors}")
    print(f"已保存: {out_path}")

    # 日志落盘(与控制台内容一致, 回查耗时/失败原因用)
    if log_file and log_lines:
        log_path = os.path.join(run_dir, "analyze_log.txt")
        with open(log_path, "w", encoding="utf-8") as f:
            f.write("\n".join(log_lines) + "\n")
        print(f"日志已保存: {log_path}")

    return all_data


# 直接运行本文件 = 批量解析最近一次run目录
# (main保护: 避免被主入口parse模式导入时意外触发全量解析)
if __name__ == "__main__":
    batch_analyze()