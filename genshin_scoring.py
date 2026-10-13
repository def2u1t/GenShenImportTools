# -*- coding: utf-8 -*-
"""
原神圣遗物评分脚本 (滚数当量法) — 依据 genshin-artifact-analysis 技能
====================================================================
- 满级件(+20): 给出当前评分(通用权值口径, 与三表Excel表一口径完全一致)
- 未满级件: 给出"预期分数"=期望潜力 —— 剩余强化等概率滚到已有副词条之一,
  期望增益 = 现有有效词条滚数当量均值(区分度真实, 非乐观上界)
- 评分等级: 不及格(<50) / 及格(50~70) / 良好(70~90) / 优秀(≥90)
- 建议(单一结论): 满级看当前评分, 未满级看期望潜力, 按装备角色定位(主C/副C/辅助)差异化

用法:
  python genshin_scoring.py [artifacts.json路径] [--top N] [--excel 输出.xlsx] [--out 明细.json]
  不传路径 -> 自动使用 screenshots 目录下最新 run_* 的 artifacts.json

输出:
  Excel(默认开启): "圣遗物评分明细.xlsx" = sheet[评分明细](主分数降序, 满级当前评分/
      未满级预期分列, 良好绿色/及格黄色着色) + sheet[汇总]
      文件被Excel占用时自动带时间戳另存; --no-excel 关闭Excel输出
  --top N: 仅控制控制台逐件打印前N件(0=不逐件打印, 默认只看汇总)
  --out 路径: 评分明细另存JSON

依赖: 标准库; Excel输出用 openpyxl(未装时提示并跳过, 仍可用 --out 输出JSON)
输入: artifacts.json 中文格式(截图OCR采集), 字段:
  序号/名称/部位/主属性/等级/词条1..4/套装类型/装备角色
"""
import argparse
import glob
import json
import os
import re
import sys

if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8')

# ============================================================
# 1. 常量(与 genshin-artifact-analysis 技能参考文件一致)
# ============================================================
# 五星副词条上限(单次最大滚动值)
SUB_TAG_MAX = {
    "critical": 0.0389, "criticalDamage": 0.0777,
    "attackPercentage": 0.0583, "lifePercentage": 0.0583,
    "defendPercentage": 0.0729, "attackStatic": 19.45,
    "lifeStatic": 298.75, "defendStatic": 23.15,
    "elementalMastery": 23.3, "recharge": 0.0648,
}

# 通用权值(未装备/跨件可比口径, 即三表Excel表一口径)
GENERIC_WEIGHTS = {
    "critical": 2.0, "criticalDamage": 1.0,
    "attackPercentage": 1.5, "elementalMastery": 0.5,
    "recharge": 0.8, "lifePercentage": 0.5, "defendPercentage": 0.3,
    "attackStatic": 0.3, "lifeStatic": 0.2, "defendStatic": 0.2,
}

# 属性英文名 -> 中文
MAIN_TAG_MAP = {
    "lifeStatic": "生命值", "attackStatic": "攻击力", "defendStatic": "防御力",
    "attackPercentage": "攻击力%", "defendPercentage": "防御力%", "lifePercentage": "生命值%",
    "elementalMastery": "元素精通", "recharge": "元素充能效率",
    "critical": "暴击率", "criticalDamage": "暴击伤害", "cureEffect": "治疗加成",
    "physicalBonus": "物理伤害加成", "fireBonus": "火元素伤害加成",
    "waterBonus": "水元素伤害加成", "iceBonus": "冰元素伤害加成",
    "thunderBonus": "雷元素伤害加成", "windBonus": "风元素伤害加成",
    "rockBonus": "岩元素伤害加成", "dendroBonus": "草元素伤害加成",
}
POSITION_MAP = {"flower": "生之花", "feather": "死之羽", "sand": "时之沙", "cup": "空之杯", "head": "理之冠"}
POS_CN2EN = {v: k for k, v in POSITION_MAP.items()}

# 属性中文名 -> 英文键
TAG_FIXED = {"生命值": "lifeStatic", "攻击力": "attackStatic", "防御力": "defendStatic", "元素精通": "elementalMastery"}
TAG_PCT = {
    "生命值": "lifePercentage", "攻击力": "attackPercentage", "防御力": "defendPercentage",
    "暴击率": "critical", "暴击伤害": "criticalDamage", "元素充能效率": "recharge",
    "治疗加成": "cureEffect", "火元素伤害加成": "fireBonus", "冰元素伤害加成": "iceBonus",
    "雷元素伤害加成": "thunderBonus", "水元素伤害加成": "waterBonus", "风元素伤害加成": "windBonus",
    "岩元素伤害加成": "rockBonus", "草元素伤害加成": "dendroBonus", "物理伤害加成": "physicalBonus",
}

# OCR 装备角色错别字归一化(圣遗物面板小字识别偏差)
EQUIP_ALIAS = {"神里华": "神里绫华", "纳西姐": "纳西妲"}

# 角色 profile(权重/主词条需求/定位) — 取自技能 character_profiles 实战配置
# 结构: name -> {role, weights, sand, cup, head[]}
PROFILE_DATA = {
    "神里绫华": {"role": "主C", "weights": {"critical": 2.0, "criticalDamage": 1.0, "attackPercentage": 1.5, "recharge": 1.0, "elementalMastery": 0.3}, "sand": "攻击力%", "cup": "冰元素伤害加成", "head": ["暴击率", "暴击伤害"]},
    "甘雨": {"role": "主C", "weights": {"critical": 2.0, "criticalDamage": 1.0, "attackPercentage": 1.5, "recharge": 0.5, "elementalMastery": 0.5}, "sand": "攻击力%", "cup": "冰元素伤害加成", "head": ["暴击率", "暴击伤害"]},
    "迪卢克": {"role": "主C", "weights": {"critical": 2.0, "criticalDamage": 1.0, "attackPercentage": 1.5, "elementalMastery": 1.0}, "sand": "攻击力%", "cup": "火元素伤害加成", "head": ["暴击率", "暴击伤害"]},
    "胡桃": {"role": "主C", "weights": {"critical": 2.0, "criticalDamage": 1.0, "lifePercentage": 1.5, "elementalMastery": 1.0, "recharge": 0.3}, "sand": "生命值%", "cup": "火元素伤害加成", "head": ["暴击率", "暴击伤害"]},
    "行秋": {"role": "副C", "weights": {"critical": 2.0, "criticalDamage": 1.0, "attackPercentage": 1.5, "recharge": 1.5}, "sand": "攻击力%", "cup": "水元素伤害加成", "head": ["暴击率", "暴击伤害"]},
    "香菱": {"role": "副C", "weights": {"critical": 2.0, "criticalDamage": 1.0, "attackPercentage": 1.5, "recharge": 1.5, "elementalMastery": 0.5}, "sand": "攻击力%", "cup": "火元素伤害加成", "head": ["暴击率", "暴击伤害"]},
    "班尼特": {"role": "辅助", "weights": {"attackPercentage": 1.5, "recharge": 2.0, "lifePercentage": 1.0}, "sand": "攻击力%", "cup": "攻击力%", "head": ["治疗加成"]},
    "钟离": {"role": "辅助", "weights": {"critical": 2.0, "criticalDamage": 1.0, "lifePercentage": 1.5, "recharge": 0.5}, "sand": "生命值%", "cup": "岩元素伤害加成", "head": ["暴击率", "暴击伤害", "生命值%"]},
    "枫原万叶": {"role": "辅助", "weights": {"elementalMastery": 2.0, "recharge": 1.5, "attackPercentage": 0.5}, "sand": "元素精通", "cup": "元素精通", "head": ["元素精通"]},
    "温迪": {"role": "辅助", "weights": {"elementalMastery": 1.5, "recharge": 1.5, "critical": 1.5, "criticalDamage": 1.0}, "sand": "元素精通", "cup": "元素精通", "head": ["暴击率", "暴击伤害", "元素精通"]},
    "雷电将军": {"role": "副C", "weights": {"critical": 2.0, "criticalDamage": 1.0, "recharge": 1.5, "attackPercentage": 1.5}, "sand": "元素充能效率", "cup": "雷元素伤害加成", "head": ["暴击率", "暴击伤害"]},
    "夜兰": {"role": "副C", "weights": {"critical": 2.0, "criticalDamage": 1.0, "lifePercentage": 1.5, "recharge": 1.5}, "sand": "生命值%", "cup": "水元素伤害加成", "head": ["暴击率", "暴击伤害"]},
    "纳西妲": {"role": "副C", "weights": {"critical": 2.0, "criticalDamage": 1.0, "elementalMastery": 1.5, "attackPercentage": 0.5}, "sand": "元素精通", "cup": "草元素伤害加成", "head": ["暴击率", "暴击伤害"]},
    "珊瑚宫心海": {"role": "辅助", "weights": {"lifePercentage": 1.5, "recharge": 1.5, "elementalMastery": 0.5}, "sand": "生命值%", "cup": "水元素伤害加成", "head": ["治疗加成"]},
    "刻晴": {"role": "主C", "weights": {"critical": 2.0, "criticalDamage": 1.0, "attackPercentage": 1.5, "elementalMastery": 1.0, "recharge": 0.5}, "sand": "攻击力%", "cup": "雷元素伤害加成", "head": ["暴击率", "暴击伤害"]},
    "菲谢尔": {"role": "副C", "weights": {"critical": 2.0, "criticalDamage": 1.0, "attackPercentage": 1.5, "recharge": 0.5}, "sand": "攻击力%", "cup": "雷元素伤害加成", "head": ["暴击率", "暴击伤害"]},
    "北斗": {"role": "副C", "weights": {"critical": 2.0, "criticalDamage": 1.0, "attackPercentage": 1.5, "recharge": 1.0}, "sand": "攻击力%", "cup": "雷元素伤害加成", "head": ["暴击率", "暴击伤害"]},
    "芭芭拉": {"role": "辅助", "weights": {"lifePercentage": 1.5, "recharge": 1.0, "elementalMastery": 0.5}, "sand": "生命值%", "cup": "水元素伤害加成", "head": ["治疗加成"]},
    "罗莎莉亚": {"role": "副C", "weights": {"critical": 2.0, "criticalDamage": 1.0, "attackPercentage": 1.5, "recharge": 0.5}, "sand": "攻击力%", "cup": "冰元素伤害加成", "head": ["暴击率", "暴击伤害"]},
    "五郎": {"role": "辅助", "weights": {"defendPercentage": 1.5, "critical": 2.0, "criticalDamage": 1.0, "recharge": 1.0, "defendStatic": 0.5}, "sand": "防御力%", "cup": "岩元素伤害加成", "head": ["暴击率", "暴击伤害", "防御力%"]},
    "凝光": {"role": "主C", "weights": {"critical": 2.0, "criticalDamage": 1.0, "attackPercentage": 1.5, "recharge": 0.5}, "sand": "攻击力%", "cup": "岩元素伤害加成", "head": ["暴击率", "暴击伤害"]},
    "云堇": {"role": "辅助", "weights": {"defendPercentage": 1.5, "defendStatic": 1.0, "recharge": 1.5}, "sand": "防御力%", "cup": "岩元素伤害加成", "head": ["防御力%"]},
    "八重神子": {"role": "副C", "weights": {"critical": 2.0, "criticalDamage": 1.0, "attackPercentage": 1.5, "recharge": 1.0, "elementalMastery": 0.5}, "sand": "攻击力%", "cup": "雷元素伤害加成", "head": ["暴击率", "暴击伤害"]},
    "流浪者": {"role": "主C", "weights": {"critical": 2.0, "criticalDamage": 1.0, "attackPercentage": 1.5, "recharge": 0.5}, "sand": "攻击力%", "cup": "风元素伤害加成", "head": ["暴击率", "暴击伤害"]},
    "闲云": {"role": "副C", "weights": {"critical": 2.0, "criticalDamage": 1.0, "attackPercentage": 1.5, "recharge": 1.0}, "sand": "攻击力%", "cup": "风元素伤害加成", "head": ["暴击率", "暴击伤害", "治疗加成"]},
    "宵宫": {"role": "主C", "weights": {"critical": 2.0, "criticalDamage": 1.0, "attackPercentage": 1.5, "recharge": 0.5}, "sand": "攻击力%", "cup": "火元素伤害加成", "head": ["暴击率", "暴击伤害"]},
    "诺艾尔": {"role": "主C", "weights": {"defendPercentage": 1.5, "critical": 2.0, "criticalDamage": 1.0, "recharge": 0.5, "defendStatic": 0.5}, "sand": "防御力%", "cup": "岩元素伤害加成", "head": ["暴击率", "暴击伤害", "防御力%"]},
    "早柚": {"role": "辅助", "weights": {"elementalMastery": 1.5, "recharge": 1.5, "lifePercentage": 0.5}, "sand": "元素精通", "cup": "元素精通", "head": ["治疗加成", "元素精通"]},
    "迪奥娜": {"role": "辅助", "weights": {"lifePercentage": 1.5, "recharge": 1.5}, "sand": "生命值%", "cup": "冰元素伤害加成", "head": ["治疗加成", "生命值%"]},
    "莫娜": {"role": "副C", "weights": {"critical": 2.0, "criticalDamage": 1.0, "recharge": 1.5, "elementalMastery": 0.5}, "sand": "元素充能效率", "cup": "水元素伤害加成", "head": ["暴击率", "暴击伤害"]},
    "雷泽": {"role": "主C", "weights": {"critical": 2.0, "criticalDamage": 1.0, "attackPercentage": 1.5}, "sand": "攻击力%", "cup": "物理伤害加成", "head": ["暴击率", "暴击伤害"]},
    "哥伦比娅": {"role": "副C", "weights": {"lifePercentage": 1.8, "critical": 1.8, "criticalDamage": 1.0, "recharge": 1.2, "elementalMastery": 0.6, "lifeStatic": 0.3}, "sand": "生命值%", "cup": "生命值%", "head": ["暴击伤害", "暴击率", "生命值%"]},
}

SHOT_ROOT = "screenshots"   # 默认在该目录下找最新 run_*/artifacts.json


# ============================================================
# 2. 中文条目(artifacts.json) -> 内部结构
# ============================================================
def parse_sub_tag(text):
    """'防御力+20.4%（待激活）' -> ('defendPercentage', 0.204)；'攻击力+16' -> ('attackStatic', 16)"""
    t = re.sub(r"[（(]待激活[)）]", "", text).strip()
    m = re.match(r"(.+?)\s*\+?\s*([0-9.]+)(%?)$", t)
    if not m:
        return None, 0.0
    name, val_str, pct = m.group(1).strip(), float(m.group(2)), bool(m.group(3))
    key = TAG_PCT.get(name) if pct else TAG_FIXED.get(name)
    if key is None:
        return None, 0.0
    return key, (val_str / 100.0 if pct else val_str)


def parse_main_tag(text):
    """'攻击力 46.6%' -> ('attackPercentage', 0.466)。花位OCR漏属性名只剩数值时兜底为生命值"""
    t = text.strip()
    m = re.match(r"(.+?)\s+([0-9.]+)(%?)$", t)
    if m:
        name, val_str, pct = m.group(1).strip(), float(m.group(2)), bool(m.group(3))
        key = TAG_PCT.get(name) if pct else TAG_FIXED.get(name)
        if key:
            return key, (val_str / 100.0 if pct else val_str)
    # 纯数值行(花位常见): "4,780" -> 生命值 4780
    m2 = re.fullmatch(r"[\d,]+", t)
    if m2:
        return "lifeStatic", float(t.replace(",", ""))
    return None, 0.0


def convert_item(item):
    """OCR中文条目 -> 内部结构 {position, mainTag, normalTags, level, equip, sub_raw}"""
    pos_cn = item.get("部位", "")
    pos = POS_CN2EN.get(pos_cn, pos_cn)
    # 花位主属性固定生命值: OCR缺属性名时兜底
    main_txt = item.get("主属性", "")
    if pos == "flower":
        m = re.match(r"^(?:生命值\s+)?([\d,.]+)$", main_txt.strip())
        if m and "生命值" not in main_txt:
            main_txt = "生命值 " + main_txt.strip().replace(",", "")
    main_key, main_val = parse_main_tag(main_txt)

    normal = []
    pending = []       # 待激活标记(词条序号)
    for i in range(1, 6):
        txt = item.get(f"词条{i}")
        if not txt:
            break
        is_pending = "待激活" in txt
        k, v = parse_sub_tag(txt)
        if k is not None:
            normal.append({"name": k, "value": v})
            if is_pending:
                pending.append(len(normal))

    lv = item.get("等级", "+0")
    m = re.search(r"\d+", lv)
    lv_num = int(m.group()) if m else 0

    equip = (item.get("装备角色") or "").strip()
    if equip.endswith("已装备"):
        equip = equip[:-3]
    equip = EQUIP_ALIAS.get(equip, equip)

    return {
        "position": pos,
        "mainTag": {"name": main_key or "", "value": main_val or 0},
        "normalTags": normal,
        "pending": pending,
        "level": min(lv_num, 20),
        "equip": equip or None,
        "set_cn": item.get("套装类型", ""),
        "main_txt": main_txt.replace(",", ""),
        "sub_raw": [item.get(f"词条{i}", "") for i in range(1, 5)],
    }


def quality_check(items):
    """分析前数据质量检查(技能陷阱#8: 防带病分析)"""
    problems = []
    n = len(items)
    miss = {"部位": 0, "主属性": 0, "等级": 0, "词条": 0, "套装": 0}
    for it in items:
        if not it.get("部位"):
            miss["部位"] += 1
        if not it.get("主属性"):
            miss["主属性"] += 1
        if not it.get("等级"):
            miss["等级"] += 1
        if not any(it.get(f"词条{i}") for i in range(1, 5)):
            miss["词条"] += 1
        if not it.get("套装类型"):
            miss["套装"] += 1
    for k, v in miss.items():
        if v:
            problems.append(f"{k}缺失{v}/{n}")
    # 序号连续性
    seqs = [it.get("序号") for it in items if isinstance(it.get("序号"), int)]
    if seqs and sorted(seqs) != list(range(1, n + 1)):
        problems.append(f"序号不连续(1~{n})")
    return problems


# ============================================================
# 3. 评分核心(滚数当量法)
# ============================================================
def get_tag_cn(name):
    return MAIN_TAG_MAP.get(name, name)


def roll_equiv(tname, tval, weights):
    """单条词条的滚数当量: (值/上限)×权值; 充能达标型(recharge高权值)3条封顶"""
    max_val = SUB_TAG_MAX[tname]
    weight = weights[tname]
    if tname == "recharge" and weight >= 1.5:
        return min((tval / max_val) * weight, weight * 3)
    return (tval / max_val) * weight


def current_score(tags, weights, level):
    """
    当前评分%: 已有有效词条滚数当量之和 ÷ 满分基准
    满分基准 = 前4大权值之和 + 5×最大权值 (+20共9次滚动: 4初始+5强化)
    """
    eff = [t for t in tags if t["name"] in SUB_TAG_MAX and weights.get(t["name"], 0) > 0]
    if not eff:
        return 0.0, []
    roll_sum = sum(roll_equiv(t["name"], t["value"], weights) for t in eff)
    sorted_w = sorted(weights.values(), reverse=True)
    full_score = sum(sorted_w[:4]) + 5 * sorted_w[0]
    if full_score <= 0:
        return 0.0, eff
    return round(min(roll_sum / full_score * 100, 100.0), 1), eff


def expected_potential(eff_tags, weights, level):
    """
    期望潜力%(未满级件预期分数):
    剩余强化等概率滚到已有副词条之一, 期望增益 = 现有有效词条滚数当量均值。
    剩余次数 = 5 - level÷4 (+20共5次强化)。无有效词条时期望增益为0。
    区别于乐观上界(假设全滚最大权值词条, +0件也有65%+, 无区分度)。
    """
    level = min(level, 20)
    eqs = []
    for t in eff_tags:
        if t["name"] in SUB_TAG_MAX and weights.get(t["name"], 0) > 0:
            eqs.append(roll_equiv(t["name"], t["value"], weights))
    sorted_w = sorted(weights.values(), reverse=True)
    full_score = sum(sorted_w[:4]) + 5 * sorted_w[0]
    if full_score <= 0:
        return 0.0
    remaining = max(0, 5 - level // 4)
    roll_sum = sum(eqs)
    avg_eq = sum(eqs) / len(eqs) if eqs else 0.0
    return round(min((roll_sum + remaining * avg_eq) / full_score * 100, 100.0), 1)


def score_grade(score):
    """评分等级: 不及格(<50) / 及格(50~70) / 良好(70~90) / 优秀(≥90)"""
    if score < 50:
        return "不及格"
    if score < 70:
        return "及格"
    if score < 90:
        return "良好"
    return "优秀"


# ============================================================
# 4. 主词条匹配 & 统一建议(单一结论)
# ============================================================
def check_main_tag_match(art, profile):
    """装备该件的角色需要的主词条是否匹配(sand/cup/head有约束, 花/羽固定)"""
    pos = art.get("position", "")
    main_cn = get_tag_cn(art.get("mainTag", {}).get("name", ""))
    if pos in ("flower", "feather"):
        return True
    if pos == "sand":
        return main_cn == profile.get("sand", "")
    if pos == "cup":
        cup_need = profile.get("cup", "")
        if cup_need == "元素精通":
            return main_cn == "元素精通" or main_cn in ("攻击力%", "生命值%", "防御力%")
        fallback = ["攻击力%"]
        sand = profile.get("sand", "")
        if sand in ("生命值%", "防御力%"):
            fallback.append(sand)
        return main_cn == cup_need or main_cn in fallback
    if pos == "head":
        return main_cn in profile.get("head", [])
    return False


def build_advice(cur, level, role_type, main_ok, eff_tags, weights):
    """
    单一结论建议:
    - 满级件: 只看当前评分(90优秀/70良好/50及格按定位分化/否则分解)
    - 未满级件: 只看期望潜力(≥70建议强化/50~70可强化赌词条或作狗粮/<50分解)
    - 依据注解: 主词条不符 / 核心词条构成
    """
    core_keys = [k for k, v in weights.items() if v >= 1.5]
    core_have = [get_tag_cn(t["name"]) for t in eff_tags if t["name"] in core_keys]

    if level >= 20:
        if cur >= 90:
            base = "优秀，必留"
        elif cur >= 70:
            base = "良好，保留"
        elif cur >= 50:
            if role_type == "主C":
                base = "及格但未达主C标准，替换后分解"
            elif role_type == "副C":
                base = "及格，副C过渡可用，有更优再替换"
            elif role_type == "辅助":
                base = "及格，辅助位可用，按需保留"
            else:
                base = "及格，过渡件"
        else:
            base = "不及格，分解"
    else:
        exp_pot = expected_potential(eff_tags, weights, level)
        if exp_pot >= 70:
            base = f"期望潜力良好({exp_pot:.0f})，建议强化"
        elif exp_pot >= 50:
            if role_type == "主C":
                base = f"期望潜力及格({exp_pot:.0f})，不建议主C投入，可作狗粮"
            else:
                base = f"期望潜力及格({exp_pot:.0f})，可强化赌词条或作狗粮"
        else:
            base = f"期望潜力不及格({exp_pot:.0f})，分解"

    notes = []
    if role_type and not main_ok:
        notes.append("主词条与角色需求不符")
    if role_type in ("主C", "副C") and not core_have:
        notes.append("副词条缺暴击/暴伤等核心词条")
    if core_have and len(core_have) >= 2:
        notes.append("核心词条" + "、".join(core_have[:3]))
    return base + ("；" + "，".join(notes) if notes else "")


# ============================================================
# 5. 主流程
# ============================================================
def default_input():
    """screenshots 下最新 run_* 的 artifacts.json"""
    cands = sorted(glob.glob(os.path.join(SHOT_ROOT, "run_*", "artifacts.json")))
    if not cands:
        return None
    return cands[-1]


def format_lvl(art):
    """等级显示: 未满级附带剩余强化次数"""
    lv = art["level"]
    if lv >= 20:
        return "+20"
    remaining = max(0, 5 - lv // 4)
    return f"+{lv}(剩{remaining}强)"


def compute_stats(rows, path):
    """汇总统计(控制台与Excel汇总sheet共用)"""
    from collections import Counter

    def _avg(xs):
        return sum(xs) / len(xs) if xs else 0.0

    def _med(xs):
        if not xs:
            return 0.0
        s = sorted(xs)
        return s[len(s) // 2]

    n = len(rows)
    n_max = sum(1 for r in rows if r["未满级预期"] is None)
    n_sub = n - n_max
    max_cur = [r["当前评分"] for r in rows if r["未满级预期"] is None]
    exps = [r["未满级预期"] for r in rows if r["未满级预期"] is not None]

    adv_c = Counter()
    for r in rows:
        a = r["建议"]
        if "分解" in a:
            adv_c["分解"] += 1
        elif "建议强化" in a:
            adv_c["建议强化"] += 1
        elif "赌词条" in a or "狗粮" in a:
            adv_c["可强化赌词条/狗粮"] += 1
        elif "保留" in a or "可用" in a or "过渡" in a or "必留" in a:
            adv_c["保留/可用"] += 1
        else:
            adv_c["其他"] += 1

    return {
        "路径": path, "总数": n, "满级数": n_max, "未满级数": n_sub,
        "主均": _avg([r["主分数"] for r in rows]),
        "主中位": _med([r["主分数"] for r in rows]),
        "主最高": max((r["主分数"] for r in rows), default=0),
        "满级均": _avg(max_cur), "满级最高": max(max_cur, default=0),
        "未满级均": _avg(exps), "未满级中位": _med(exps), "未满级最高": max(exps, default=0),
        "强化件数": sum(1 for x in exps if x >= 70),
        "等级档": dict(Counter(r["等级档"] for r in rows)),
        "建议": dict(adv_c),
        "Top5": [f"#{r['序号']}{r['部位']}{r['等级']}={r['主分数']:.1f}" for r in rows[:5]],
    }


def write_excel(rows, stats, path):
    """评分明细输出Excel: sheet[评分明细](主分数降序+等级着色) + sheet[汇总]"""
    try:
        from openpyxl import Workbook
        from openpyxl.styles import Alignment, Font, PatternFill
        from openpyxl.utils import get_column_letter
    except ImportError:
        print("[提示] 未安装openpyxl, 跳过Excel输出(可用 --out 输出JSON)")
        return None

    wb = Workbook()
    ws = wb.active
    ws.title = "评分明细"
    ws.freeze_panes = "A2"

    headers = ["序号", "套装", "部位", "等级", "主词条", "词条1", "词条2", "词条3", "词条4",
               "装备角色", "当前评分", "未满级预期", "等级档", "建议"]
    widths = [6, 20, 9, 13, 16, 15, 15, 15, 15, 16, 10, 11, 9, 60]
    head_font = Font(bold=True, color="FFFFFF")
    head_fill = PatternFill("solid", fgColor="4472C4")
    for c, h in enumerate(headers, 1):
        cell = ws.cell(row=1, column=c, value=h)
        cell.font = head_font
        cell.fill = head_fill
        cell.alignment = Alignment(horizontal="center")
        ws.column_dimensions[get_column_letter(c)].width = widths[c - 1]

    # 等级档整行着色(不及格不上色, 避免大片区视觉噪音)
    grade_style = {
        "优秀": ("C6EFCE", "006100"),
        "良好": ("C6EFCE", "006100"),
        "及格": ("FFEB9C", "9C6500"),
    }
    r_i = 2
    for r in rows:
        vals = [r["序号"], r["套装"], r["部位"], r["等级"], r["主词条"],
                r["词条1"], r["词条2"], r["词条3"], r["词条4"],
                r["装备角色"], r["当前评分"],
                (r["未满级预期"] if r["未满级预期"] is not None else "—"),
                r["等级档"], r["建议"]]
        for c, v in enumerate(vals, 1):
            cell = ws.cell(row=r_i, column=c, value=v)
            if c in (11, 12) and isinstance(v, (int, float)):
                cell.number_format = "0.0"
            if c == 13:
                cell.alignment = Alignment(horizontal="center")
        style = grade_style.get(r["等级档"])
        if style:
            fill = PatternFill("solid", fgColor=style[0])
            font = Font(color=style[1])
            for c in range(1, len(vals) + 1):
                ws.cell(row=r_i, column=c).fill = fill
                ws.cell(row=r_i, column=c).font = font
        r_i += 1

    # ---- 汇总sheet ----
    ws2 = wb.create_sheet("汇总")
    ws2.column_dimensions["A"].width = 30
    ws2.column_dimensions["B"].width = 92
    kv = [
        ("数据源", stats["路径"]),
        ("件数", f"{stats['总数']} 件 (满级 {stats['满级数']} / 未满级 {stats['未满级数']})"),
        ("口径", "满级件分数=当前评分(通用权值) / 未满级件分数=期望潜力(装备角色权值); 明细按主分数降序"),
        ("主分数 平均/中位/最高", f"{stats['主均']:.1f} / {stats['主中位']:.1f} / {stats['主最高']:.1f}"),
        ("满级件 当前评分", f"平均 {stats['满级均']:.1f}, 最高 {stats['满级最高']:.1f}"),
        ("未满级件 预期", f"平均 {stats['未满级均']:.1f}, 中位 {stats['未满级中位']:.1f}, 最高 {stats['未满级最高']:.1f}; 预期≥70(建议强化) {stats['强化件数']} 件"),
        ("等级档分布(按主分数)", "  ".join(f"{k}:{v}" for k, v in stats["等级档"].items())),
        ("建议分布", "  ".join(f"{k}:{v}" for k, v in stats["建议"].items())),
        ("Top5", "; ".join(stats["Top5"])),
    ]
    for i, (k, v) in enumerate(kv, 1):
        ws2.cell(row=i, column=1, value=k).font = Font(bold=True)
        ws2.cell(row=i, column=2, value=v)

    # ---- 保存(被占用时自动改名重试一次) ----
    try:
        wb.save(path)
        print(f"Excel已保存: {path}")
        return path
    except PermissionError:
        import datetime
        stem, _, ext = path.rpartition(".")
        if not stem:
            stem, ext = path, "xlsx"
        alt = f"{stem}_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.{ext}"
        try:
            wb.save(alt)
            print(f"[提示] {path} 被占用(可能正被Excel打开), 已另存: {alt}")
            return alt
        except PermissionError:
            print(f"[提示] {path} 与 {alt} 均无法写入, 本次未保存Excel")
            return None


def load_and_score(path):
    """
    读入 artifacts.json → 数据质量检查 → 转换 → 逐件评分 → 排序。
    返回 (rows, stats, problems)。
    """
    with open(path, "r", encoding="utf-8") as f:
        raw = json.load(f)
    if not isinstance(raw, list):
        raise ValueError("输入应为 artifacts.json 中文格式(数组)。mona.json 请用三表分析脚本。")

    problems = quality_check(raw)

    items = [convert_item(it) for it in raw]

    # ---- 逐件评分 ----
    rows = []
    for i, art in enumerate(items, 1):
        eq = art["equip"] or ""
        prof = PROFILE_DATA.get(eq)
        adv_weights = prof["weights"] if prof else GENERIC_WEIGHTS
        role_type = prof["role"] if prof else ""

        # 当前评分: 通用权值口径(与三表Excel表一口径一致, 跨件可比)
        g_cur, g_eff = current_score(art["normalTags"], GENERIC_WEIGHTS, art["level"])
        # 建议口径: 装备角色权值(无角色时为通用); 主词条匹配用角色需求
        a_cur, a_eff = current_score(art["normalTags"], adv_weights, art["level"])
        main_ok = check_main_tag_match(art, prof) if prof else True
        advice = build_advice(g_cur, art["level"], role_type, main_ok, a_eff, adv_weights)

        if art["level"] >= 20:
            main_score, score_kind = g_cur, "满级评分"
            exp_txt = "—"
        else:
            main_score = expected_potential(a_eff, adv_weights, art["level"])
            score_kind = "预期(期望潜力)"
            exp_txt = f"{main_score}"

        rows.append({
            "序号": i,
            "套装": art["set_cn"], "部位": POSITION_MAP.get(art["position"], art["position"]),
            "等级": format_lvl(art),
            "主词条": art["main_txt"],
            "词条1": art["sub_raw"][0],
            "词条2": art["sub_raw"][1],
            "词条3": art["sub_raw"][2],
            "词条4": art["sub_raw"][3],
            "副词条": "; ".join(t for t in art["sub_raw"] if t),
            "装备角色": (f"{eq}({role_type})" if role_type else (eq or "未装备")),
            "当前评分": g_cur,
            "未满级预期": (main_score if art["level"] < 20 else None),
            "主分数": main_score,
            "类型": score_kind,
            "等级档": score_grade(main_score),
            "建议": advice,
        })

    rows.sort(key=lambda r: r["主分数"], reverse=True)
    stats = compute_stats(rows, path)
    return rows, stats, problems


def score_file(json_path, excel_path=None, top=0, out_json=""):
    """
    完整评分流程(命令行与GUI共用):
    评分 → 数据质量警告 → 控制台报告 → Excel导出 → 可选JSON明细。
    返回 (rows, stats, excel_saved_path)。
    """
    rows, stats, problems = load_and_score(json_path)
    if problems:
        print(f"[数据质量警告] {'; '.join(problems)} — 请先核对截图采集(陷阱#8)")

    # ---- 控制台: top>0 时先打印前N件明细 ----
    if top > 0:
        print("=" * 76)
        print(f"原神圣遗物评分 (滚数当量法)  数据: {json_path}")
        print(f"共 {stats['总数']} 件 (满级 {stats['满级数']} / 未满级 {stats['未满级数']})  排序: 主分数降序")
        print(f"口径: 当前评分=通用权值(与三表Excel表一一致); 预期=期望潜力(装备角色权值)")
        print("=" * 76)
        for r in rows[:top]:
            print(f"[{r['序号']:03d}] {r['套装']} | {r['部位']} | {r['等级']:>9} | {r['主词条']} | {r['装备角色']}")
            print(f"      副词条: {r['副词条'] or '(无)'}")
            if r["未满级预期"] is not None:
                print(f"      预期 {r['未满级预期']:.1f} 【{r['等级档']}】 (当前 {r['当前评分']:.1f}) —— {r['建议']}")
            else:
                print(f"      评分 {r['当前评分']:.1f} 【{r['等级档']}】 —— {r['建议']}")

    # ---- 控制台汇总 ----
    print("=" * 76)
    print(f"原神圣遗物评分汇总  数据: {json_path}")
    print(f"共 {stats['总数']} 件 (满级 {stats['满级数']} / 未满级 {stats['未满级数']})")
    print(f"口径: 满级件分数=当前评分; 未满级件分数=期望潜力")
    print(f"汇总: 主分数平均 {stats['主均']:.1f} | 中位 {stats['主中位']:.1f} | 最高 {stats['主最高']:.1f}")
    print(f"  满级 {stats['满级数']} 件: 当前评分平均 {stats['满级均']:.1f}, 最高 {stats['满级最高']:.1f}")
    if stats["未满级数"]:
        print(f"  未满级 {stats['未满级数']} 件: 预期平均 {stats['未满级均']:.1f}, 中位 {stats['未满级中位']:.1f}, "
              f"最高 {stats['未满级最高']:.1f}; 预期≥70(建议强化) {stats['强化件数']} 件")
    print(f"  等级档(按主分数): {stats['等级档']}")
    print(f"  建议: {stats['建议']}")
    print("Top5: " + "; ".join(stats["Top5"]))

    # ---- Excel输出 ----
    excel_saved = None
    if excel_path:
        excel_saved = write_excel(rows, stats, excel_path)

    # ---- JSON明细(可选) ----
    if out_json:
        slim = [{k: r[k] for k in ("序号", "套装", "部位", "等级", "主词条", "副词条", "装备角色",
                                    "当前评分", "未满级预期", "主分数", "等级档", "建议")} for r in rows]
        with open(out_json, "w", encoding="utf-8") as f:
            json.dump(slim, f, ensure_ascii=False, indent=1)
        print(f"明细已保存: {out_json}")

    return rows, stats, excel_saved


def main():
    ap = argparse.ArgumentParser(description="原神圣遗物评分(滚数当量法): 满级给当前评分, 未满级给期望潜力")
    ap.add_argument("input", nargs="?", default=None, help="artifacts.json路径(默认取screenshots下最新run)")
    ap.add_argument("--top", type=int, default=0, help="控制台逐件打印前N件(0=不逐件打印, 默认)")
    ap.add_argument("--excel", default="圣遗物评分明细.xlsx", help="Excel输出路径(默认: 圣遗物评分明细.xlsx)")
    ap.add_argument("--no-excel", action="store_true", help="关闭Excel输出")
    ap.add_argument("--out", default="", help="评分明细另存JSON路径(可选)")
    args = ap.parse_args()

    path = args.input or default_input()
    if not path or not os.path.exists(path):
        print(f"找不到输入文件: {path}")
        sys.exit(1)

    score_file(path,
               excel_path=None if args.no_excel else args.excel,
               top=args.top, out_json=args.out)


if __name__ == "__main__":
    main()
