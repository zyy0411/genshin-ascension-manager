# -*- coding: utf-8 -*-
"""生成《原神角色养成管理表 v1.0》
特点：
- 记录等级/武器/天赋/圣遗物逐件词条
- 圣遗物逐件评分 + 有效词条判断 + 调整建议
- 面板数值 vs WIKI 毕业参考面板自动判定是否练好
"""
import json
import os
import re
import sys
from datetime import date

from openpyxl import Workbook
from openpyxl.styles import Protection, Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import CellIsRule, FormulaRule, ColorScaleRule, DataBarRule
from openpyxl.chart import BarChart, PieChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.chart.marker import DataPoint
from openpyxl.chart.shapes import GraphicalProperties
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.comments import Comment

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SRC = os.path.join(BASE, "_source")
sys.path.insert(0, SRC)
import genshin_data as D  # noqa: E402

OUT = os.path.join(BASE, "原神角色养成管理表_v1.0.xlsx")

FONT = "微软雅黑"
INK = "1F2937"
MUTED = "6B7280"
LINE = "D9E2EC"
NAVY = "1F3B57"
GOLD = "C9A227"
INPUT = "FFF9E6"
SOFT = "F5F8FB"

GROUPS = {
    "角色": ("1F3B57", "E8F0F8"),
    "等级": ("117A65", "E6F4EA"),
    "武器": ("1F5F7A", "E3F1F7"),
    "圣遗物": ("6C3F8F", "F0EAF7"),
    "天赋": ("9A5B00", "FDF1E0"),
    "面板": ("1F4E79", "E9F0F7"),
    "评估": ("1E6B47", "E8F5EE"),
    "其他": ("5D6D7E", "EEF1F5"),
    "状态": ("5D6D7E", "EEF1F5"),
}
ELEMENTS = {
    "火": ("FDEDEC", "C0392B"), "水": ("EAF2F8", "2471A3"),
    "风": ("E9F7EF", "1E8449"), "雷": ("F4ECF7", "7D3C98"),
    "草": ("F0F7E6", "4F6228"), "冰": ("E8F8F5", "117A65"),
    "岩": ("FEF9E7", "9A7D0A"),
}
STATUS_COLORS = {
    "已毕业": ("D5F5E3", "1E6B47"),
    "需调整": ("FDF3D0", "A66B00"),
    "未毕业": ("FADBD8", "A93226"),
    "未开始": ("EEF1F5", "6B7280"),
}
READY_COLORS = {
    "✔ 已练好": ("C6F6D5", "1E6B47"),
    "未练好": ("FDEBD0", "A66B00"),
    "待补圣遗物": ("EAF2F8", "21618C"),
    "未开始": ("EEF1F5", "6B7280"),
}
OWNED_COLORS = {
    "已有": ("E8F5EE", "1E6B47"),
    "未拥有": ("EEF1F5", "8A94A0"),
}
PANEL_COLORS = {
    "达标": ("D5F5E3", "1E6B47"),
    "接近": ("FDF3D0", "A66B00"),
    "未达标": ("FADBD8", "A93226"),
    "无参考": ("EEF1F5", "6B7280"),
}
WEAPON_COLORS = {
    "✔ 推荐武器": ("D5F5E3", "1E6B47"),
    "△ 可用武器": ("FDF3D0", "A66B00"),
    "✘ 非推荐": ("FADBD8", "A93226"),
}

THIN = Side(style="thin", color=LINE)
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
MED = Side(style="medium", color="B8C4D0")

DATA_FIRST, DATA_LAST = 3, 135
LIB_FIRST = 3
DETAIL_FIRST = 3
DETAIL_LAST = DETAIL_FIRST + 133 * 5 - 1   # 667
LIB_LAST = LIB_FIRST + 133 + 12 - 1       # 147

MAIN = "养成总表"
DETAIL = "圣遗物明细"
LIB = "配装数据库"
CARD = "角色参考卡"
TODAY = "今日练谁"
DASH = "仪表盘"
SRC_SHEET = "数据源"
UPDATE = "数据与更新"
HELP = "使用说明"

Q_MAIN = f"'{MAIN}'"
Q_DETAIL = f"'{DETAIL}'"
Q_LIB = f"'{LIB}'"


def LRA(col):
    return f"${col}$2:${col}$200"


# ---------------- 基础工具 ----------------
def F(size=10, bold=False, color=INK, italic=False, underline=None):
    kw = dict(name=FONT, size=size, bold=bold, color=color, italic=italic)
    if underline:
        kw["underline"] = underline
    return Font(**kw)


def fill(color):
    return PatternFill("solid", fgColor=color)


def A(h="center", v="center", wrap=False, indent=0):
    return Alignment(horizontal=h, vertical=v, wrap_text=wrap, indent=indent)


def cells_of(ws, rng):
    obj = ws[rng]
    if not isinstance(obj, tuple):
        return [obj]
    return [c for row in obj for c in row]


def style_range(ws, rng, font=None, fill_color=None, align=None, border=BORDER, number_format=None):
    for c in cells_of(ws, rng):
        if font is not None:
            c.font = font
        if fill_color is not None:
            c.fill = fill(fill_color)
        if align is not None:
            c.alignment = align
        if border is not None:
            c.border = border
        if number_format is not None:
            c.number_format = number_format


def merge(ws, rng, value=None, font=None, fill_color=None, align=None, border=BORDER):
    style_range(ws, rng, font=font, fill_color=fill_color, align=align, border=border)
    ws.merge_cells(rng)
    if value is not None:
        ws[rng.split(":")[0]] = value


def section(ws, row, title, end_col="AO", color=NAVY, bg="E8EEF5"):
    merge(ws, f"A{row}:{end_col}{row}", f"  {title}", font=F(11, True, color),
          fill_color=bg, align=A("left", "center", False, 1), border=None)
    ws.row_dimensions[row].height = 26


def align_hdr(h):
    return A("center", "center", True, 0)


# ---------------- 静态清单 ----------------
SUBSTATS = ["生命值", "攻击力", "防御力", "生命值%", "攻击力%", "防御力%",
            "元素精通", "元素充能效率", "暴击率", "暴击伤害"]
SUBSTAT_ROLL = {"生命值": 299, "攻击力": 19, "防御力": 23, "生命值%": 5.8,
                "攻击力%": 5.8, "防御力%": 7.3, "元素精通": 23,
                "元素充能效率": 6.5, "暴击率": 3.9, "暴击伤害": 7.8}
MAINSTATS = ["生命值", "攻击力", "生命值%", "攻击力%", "防御力%", "元素精通",
             "元素充能效率", "治疗加成", "暴击率", "暴击伤害",
             "火元素伤害加成", "水元素伤害加成", "雷元素伤害加成", "冰元素伤害加成",
             "风元素伤害加成", "岩元素伤害加成", "草元素伤害加成", "物理伤害加成"]
PIECES = ["生之花", "死之羽", "时之沙", "空之杯", "理之冠"]
MARKERS = ["暂缓", "不练了"]

_STAT_MAP = {
    "生命值百分比": "生命值%", "攻击力百分比": "攻击力%", "防御力百分比": "防御力%",
    "数值生命值": "生命值", "数值攻击力": "攻击力", "数值防御力": "防御力",
    "元素充能": "元素充能效率", "充能效率": "元素充能效率",
}


def norm_stat(s):
    s = (s or "").strip()
    return _STAT_MAP.get(s, s)


def norm_list(text, allowed=None):
    """把 wiki 文本里的词条列表规范成「A、B」形式。"""
    text = re.sub(r"（[^）]*）|\([^)]*\)", "", text or "")
    out = []
    for part in re.split(r"[、，,＞>]", text):
        n = norm_stat(part)
        if not n:
            continue
        if allowed is not None and n not in allowed:
            continue
        if n not in out:
            out.append(n)
    return "、".join(out)


def substat_tiers(text):
    """副词条优先级 → 3 档有效词条（每档「A、B」）。"""
    text = re.sub(r"（[^）]*）|\([^)]*\)", "", text or "")
    tiers = []
    for tier in re.split(r"[＞>]", text):
        names = []
        for part in re.split(r"[、，,]", tier):
            n = norm_stat(part)
            if n in SUBSTATS and n not in names:
                names.append(n)
        if names:
            tiers.append("、".join(names))
    while len(tiers) < 3:
        tiers.append("")
    return tiers[:3]


def split_sets(sets):
    out = []
    for s in sets or []:
        name = s.get("name", "") if isinstance(s, dict) else str(s)
        desc = s.get("desc", "") if isinstance(s, dict) else ""
        out.append((name, desc))
    return out


def talent_targets(b):
    """从 WIKI「技能升级推荐」文本推到 3 个目标天赋等级（参考值，可自行调整）。"""
    text = re.sub(r"\s+", "", b.get("talent_text") or "")
    order = b.get("talent_order") or ["普攻", "战技", "爆发"]
    alts = {"普攻": ("普通攻击", "普攻"), "战技": ("元素战技", "战技"), "爆发": ("元素爆发", "爆发")}
    up = {}
    for k, names in alts.items():
        up[k] = not any(f"不升级{n}" in text or f"不升{n}" in text for n in names)
    values, used = {}, [10, 9, 8]
    idx = 0
    prev = None
    for k in order:
        if not up.get(k, True):
            values[k] = 1
            prev = None
            continue
        tie = False
        if prev:
            for a in alts.get(prev, (prev,)):
                for b in alts.get(k, (k,)):
                    if f"{a}＝{b}" in text or f"{b}＝{a}" in text or f"{a}={b}" in text:
                        tie = True
        if tie:
            values[k] = values[prev]
        else:
            values[k] = used[min(idx, 2)]
        idx += 1
        prev = k
    for k in ("普攻", "战技", "爆发"):
        values.setdefault(k, 9)
    return values["普攻"], values["战技"], values["爆发"]


def load_all():
    roster = json.load(open(os.path.join(SRC, "wiki_roster.json"), encoding="utf-8"))
    builds = {}
    bp = os.path.join(SRC, "wiki_builds.json")
    if os.path.exists(bp):
        raw = json.load(open(bp, encoding="utf-8"))
        builds = raw.get("characters", raw)
    panels = {}
    pp = os.path.join(SRC, "wiki_panels.json")
    if os.path.exists(pp):
        panels = json.load(open(pp, encoding="utf-8"))

    rows = []
    for name, info in roster.items():
        b = builds.get(name) or {}
        p = panels.get(name) or {}
        sets = split_sets(b.get("artifacts"))
        t1, t2, t3 = substat_tiers(b.get("substats", ""))
        tna, tsk, tbu = talent_targets(b) if b else (9, 9, 9)
        ttotal = tna + tsk + tbu
        order = b.get("talent_order") or []
        order = " > ".join(order) if order else ""
        exp = 75
        if ttotal >= 27:
            exp = 80
        status = "WIKI" if b and sets else "待补"
        rows.append({
            "name": name,
            "element": info.get("element", ""),
            "version": info.get("version", ""),
            "sets": sets,
            "sand": norm_list(b.get("sand", ""), MAINSTATS),
            "goblet": norm_list(b.get("goblet", ""), MAINSTATS),
            "circlet": norm_list(b.get("circlet", ""), MAINSTATS),
            "substat": b.get("substats", ""),
            "eff1": t1, "eff2": t2, "eff3": t3,
            "best": b.get("best_weapon", ""),
            "best_desc": b.get("best_weapon_desc", ""),
            "opts": b.get("optional_weapons") or [],
            "opt_descs": b.get("optional_weapon_descs") or [],
            "order": order,
            "tna": tna, "tskill": tsk, "tburst": tbu, "ttotal": ttotal,
            "expect": exp,
            "cr": p.get("cr"), "cd": p.get("cd"), "atk": p.get("atk"),
            "hp": p.get("hp"), "def": p.get("def"), "em": p.get("em"),
            "er": p.get("er"), "panel_note": p.get("note", ""),
            "source": b.get("source_url", ""),
            "updated": (b.get("rev_timestamp") or "")[:10] or date.today().isoformat(),
            "variant": b.get("variant", ""),
            "status": status,
        })
    rows.sort(key=lambda x: (x["status"] != "待补", [int(v) if v.isdigit() else 0 for v in x["version"].split(".")], x["name"]), reverse=True)
    return rows, roster, builds, panels


def set_list(rows):
    """下拉用的圣遗物套装名：把 2+2 散搭的「A、B」拆成单项。"""
    names = set()
    for r in rows:
        for n, _ in r["sets"]:
            for part in str(n or "").replace("，", "、").split("、"):
                part = part.strip()
                if part:
                    names.add(part)
    for n in D.ARTIFACT_SETS:
        names.add(n)
    return sorted(names)


def unlock(cell):
    """标记单元格可编辑（配合工作表保护）。"""
    cell.protection = Protection(locked=False)


def protect(ws, allow_filter=True):
    """保护工作表：默认锁公式，输入格需预先 unlock。"""
    ws.protection.sheet = True
    ws.protection.autoFilter = allow_filter
    ws.protection.sort = allow_filter
    ws.protection.formatCells = False
    ws.protection.formatColumns = False
    ws.protection.formatRows = False
    ws.protection.insertRows = False
    ws.protection.deleteRows = False
