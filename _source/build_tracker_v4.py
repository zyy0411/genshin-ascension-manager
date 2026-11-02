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


# -*- coding: utf-8 -*-
# ============ 数据源 ============
def build_data(wb, rows):
    ws = wb.create_sheet(SRC_SHEET)
    ws.sheet_properties.tabColor = "7F8C8D"
    merge(ws, "A1:O1", "数据源 · 下拉选项与评分参数（可增补；不要删除整列）",
          font=F(11, True, "FFFFFF"), fill_color="5D6D7E", align=A("left", "center", False, 1), border=None)
    ws.row_dimensions[1].height = 28

    headers = ["元素", "圣遗物套装", "沙·主词条", "杯·主词条", "冠·主词条",
               "天赋等级", "目标等级", "精炼", "练度标记", "圣遗物部位",
               "主词条（18种）", "副词条（10种）", "单词条上限", "上限数值", "拥有状态",
               "命座", "强化等级"]
    for i, h in enumerate(headers, 1):
        c = ws.cell(row=2, column=i, value=h)
        c.font = F(10, True, "FFFFFF")
        c.fill = fill("5D6D7E")
        c.alignment = align_hdr(h)
        c.border = BORDER
    ws.row_dimensions[2].height = 40

    set_names = set_list(rows)
    cols = [
        D.ELEMENTS,
        set_names,
        D.SANDS,
        D.GOBLETS,
        D.CIRCLETS,
        D.TALENT_LEVELS,
        D.TARGET_LEVELS,
        D.REFINES,
        MARKERS,
        PIECES,
        MAINSTATS,
        SUBSTATS,
        SUBSTATS,
        [SUBSTAT_ROLL[s] for s in SUBSTATS],
        ["已有", "未拥有"],
        list(range(0, 7)),
        list(range(0, 21)),
    ]
    for ci, values in enumerate(cols, 1):
        for ri, v in enumerate(values, 3):
            c = ws.cell(row=ri, column=ci, value=v)
            c.font = F(10)
            c.alignment = A("center", "center")
            c.border = BORDER
            c.fill = fill(SOFT if ri % 2 else "FFFFFF")
    for ci, w in enumerate([9, 24, 20, 26, 18, 10, 10, 8, 10, 12, 20, 18, 14, 10, 12, 7, 9], 1):
        ws.column_dimensions[get_column_letter(ci)].width = w

    names = {
        "元素列表": ("A", 3, 9), "套装列表": ("B", 3, 2 + len(set_names)), "沙列表": ("C", 3, 7),
        "杯列表": ("D", 3, 14), "冠列表": ("E", 3, 9), "天赋等级列表": ("F", 3, 15),
        "目标等级列表": ("G", 3, 9), "精炼列表": ("H", 3, 7), "标记列表": ("I", 3, 6),
        "部位列表": ("J", 3, 7), "主词条列表": ("K", 3, 20), "词条列表": ("L", 3, 12),
        "词条上限表": ("M", 3, 12), "拥有列表": ("O", 3, 4),
        "命座列表": ("P", 3, 9), "强化列表": ("Q", 3, 23),
    }
    return names


# ============ 配装数据库 ============
LIB_HEADERS = [
    ("角色名", 15), ("元素", 7), ("实装", 8),
    ("套装1", 22), ("说明1", 32), ("套装2", 22), ("说明2", 32), ("套装3", 22), ("说明3", 32),
    ("沙·主词条", 20), ("杯·主词条", 24), ("冠·主词条", 20),
    ("副词条优先级", 46), ("有效词条·核心", 28), ("有效词条·次要", 28), ("有效词条·可选", 28),
    ("毕业武器", 16), ("武器说明", 34), ("可选武器1", 14), ("可选武器2", 14), ("可选武器3", 14),
    ("可选说明1", 30), ("可选说明2", 30), ("可选说明3", 30),
    ("加点顺序", 22), ("目标普攻", 8), ("目标战技", 8), ("目标爆发", 8), ("目标总级", 8),
    ("建议期望", 8), ("暴击率%", 9), ("暴击伤害%", 10), ("攻击力", 10), ("生命值", 10),
    ("防御力", 9), ("元素精通", 9), ("元素充能%", 10), ("面板说明", 40),
    ("WIKI攻略页", 46), ("攻略更新", 11), ("数据来源", 9), ("流派备注", 16),
]


def build_library(wb, rows):
    ws = wb.create_sheet(LIB)
    ws.sheet_properties.tabColor = "8E44AD"
    merge(ws, "A1:AP1", "配装数据库 · biligame 原神 WIKI 攻略页结构化数据 · 共 133 名角色（「待补」= WIKI 暂无攻略，可手工补）",
          font=F(11, True, "FFFFFF"), fill_color="6C3F8F", align=A("left", "center", False, 1), border=None)
    ws.row_dimensions[1].height = 30

    purple = ["F0EAF7", "E7DBF2"]
    for i, (h, w) in enumerate(LIB_HEADERS, 1):
        c = ws.cell(row=2, column=i, value=h)
        c.font = F(9, True, "FFFFFF")
        c.fill = fill("6C3F8F")
        c.alignment = align_hdr(h)
        c.border = BORDER
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.row_dimensions[2].height = 42

    for ri, r in enumerate(rows, LIB_FIRST):
        sets = (r["sets"] + [("", "")] * 3)[:3]
        opts = (r["opts"] + [""] * 3)[:3]
        odesc = (r["opt_descs"] + [""] * 3)[:3]
        panel = r
        vals = [
            r["name"], r["element"], r["version"],
            sets[0][0], sets[0][1], sets[1][0], sets[1][1], sets[2][0], sets[2][1],
            r["sand"], r["goblet"], r["circlet"],
            r["substat"], r["eff1"], r["eff2"], r["eff3"],
            r["best"], r["best_desc"], opts[0], opts[1], opts[2],
            odesc[0], odesc[1], odesc[2],
            r["order"], r["tna"], r["tskill"], r["tburst"], r["ttotal"],
            r["expect"], panel["cr"], panel["cd"], panel["atk"], panel["hp"],
            panel["def"], panel["em"], panel["er"], r["panel_note"],
            r["source"], r["updated"], r["status"], r["variant"],
        ]
        for ci, v in enumerate(vals, 1):
            c = ws.cell(row=ri, column=ci, value=v)
            c.font = F(9)
            c.alignment = A("center" if ci in (2, 3, 10, 11, 12, 13, 14, 15, 16, 19, 20, 21, 26, 27, 28, 29, 30, 31, 32, 33, 34, 35, 36, 37, 40, 41) else "left", "center", True)
            c.border = BORDER
            c.fill = fill(purple[ri % 2])
        ws.row_dimensions[ri].height = 44

    # 空白补充行
    for ri in range(LIB_FIRST + len(rows), LIB_LAST + 1):
        for ci in range(1, len(LIB_HEADERS) + 1):
            c = ws.cell(row=ri, column=ci)
            c.fill = fill(INPUT)
            c.border = BORDER
    merge(ws, f"B{LIB_LAST + 2}:AP{LIB_LAST + 2}",
          "提示：在上面黄色空白行可手工补充角色；「数据来源」写 人工。修改后总表下拉会自动出现。",
          font=F(9, False, MUTED), fill_color=None, align=A("left", "center", False, 1), border=None)

    ws.freeze_panes = "C3"
    ws.auto_filter.ref = f"A2:AP{LIB_LAST}"
    ws.sheet_view.zoomScale = 85
    return ws


# -*- coding: utf-8 -*-
LIBR = lambda c: f"${c}$3:${c}$135"   # noqa: E731
DR = lambda c: f"${c}$3:${c}$667"     # noqa: E731
Q_LIB = f"'{LIB}'"
Q_DETAIL = f"'{DETAIL}'"


def lib_idx(col, r, match_col="A"):
    return f"INDEX({Q_LIB}!{LIBR(col)},MATCH($B{r},{Q_LIB}!{LIBR(match_col)},0))"


# ============ 圣遗物明细 ============
DETAIL_HEADERS = [
    ("角色名", 13), ("部位", 10), ("套装", 20), ("主词条", 16), ("强化", 7),
    ("词条1", 14), ("词条2", 14), ("词条3", 14), ("词条4", 14),
    ("数值1", 8), ("数值2", 8), ("数值3", 8), ("数值4", 8),
    ("有效词条", 24), ("单品评分", 9), ("调整提示", 34), ("推荐主词条", 16), ("推荐套装", 18),
    ("权重1", 7), ("权重2", 7), ("权重3", 7), ("权重4", 7),
    ("roll1", 8), ("roll2", 8), ("roll3", 8), ("roll4", 8), ("主词条匹配", 8),
    ("键", 20), ("排序键", 9), ("核心词条", 22), ("次要词条", 22), ("可选词条", 22),
]


def build_detail(wb, rows):
    ws = wb.create_sheet(DETAIL)
    ws.sheet_properties.tabColor = "6C3F8F"
    merge(ws, "A1:AF1", "圣遗物明细 · 每件填套装/主词条/强化等级与 4 条副词条数值，自动计算有效词条与单品评分",
          font=F(11, True, "FFFFFF"), fill_color="6C3F8F", align=A("left", "center", False, 1), border=None)
    ws.row_dimensions[1].height = 30
    purple = ["F7F2FB", "FFFFFF"]
    for i, (h, w) in enumerate(DETAIL_HEADERS, 1):
        c = ws.cell(row=2, column=i, value=h)
        c.font = F(9, True, "FFFFFF")
        c.fill = fill("6C3F8F")
        c.alignment = align_hdr(h)
        c.border = BORDER
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.row_dimensions[2].height = 40

    by_name = {r["name"]: r for r in rows}
    # (套装, 主词条, 强化等级, 词条1,值1, 词条2,值2, 词条3,值3, 词条4,值4)
    sample = {
        "胡桃": {
            "生之花": ("炽烈的炎之魔女", "生命值", 20, "暴击率", 11.7, "暴击伤害", 23.4, "生命值%", 11.6, "元素精通", 23.0),
            "死之羽": ("炽烈的炎之魔女", "攻击力", 20, "暴击率", 11.7, "暴击伤害", 23.4, "生命值%", 11.6, "元素精通", 23.0),
            "时之沙": ("炽烈的炎之魔女", "生命值%", 20, "暴击率", 11.7, "暴击伤害", 23.4, "生命值%", 11.6, "元素精通", 23.0),
            "空之杯": ("炽烈的炎之魔女", "火元素伤害加成", 20, "暴击率", 11.7, "暴击伤害", 23.4, "生命值%", 11.6, "元素精通", 23.0),
            "理之冠": ("炽烈的炎之魔女", "暴击伤害", 20, "暴击率", 11.7, "暴击伤害", 23.4, "生命值%", 11.6, "元素精通", 23.0),
        },
        "芙宁娜": {
            "生之花": ("黄金剧团", "生命值", 20, "暴击率", 3.1, "暴击伤害", 6.2, "元素充能效率", 5.2, "生命值%", 4.7),
            "死之羽": ("黄金剧团", "攻击力", 20, "暴击率", 2.7, "暴击伤害", 14.0, "生命值%", 5.8, "元素充能效率", 6.5),
            "时之沙": ("黄金剧团", "元素充能效率", 16, "暴击率", 3.1, "暴击伤害", 7.0, "生命值%", 5.2, "元素精通", 21.0),
        },
    }
    row = DETAIL_FIRST
    for r in rows:
        for piece in PIECES:
            ws.cell(row=row, column=1, value=r["name"])
            ws.cell(row=row, column=2, value=piece)
            ws.cell(row=row, column=30, value=r["eff1"])
            ws.cell(row=row, column=31, value=r["eff2"])
            ws.cell(row=row, column=32, value=r["eff3"])
            sp = sample.get(r["name"], {}).get(piece)
            if sp:
                sset, mainstat, level, n1, v1, n2, v2, n3, v3, n4, v4 = sp
                ws.cell(row=row, column=3, value=sset)
                ws.cell(row=row, column=4, value=mainstat)
                ws.cell(row=row, column=5, value=level)
                ws.cell(row=row, column=6, value=n1)
                ws.cell(row=row, column=7, value=n2)
                ws.cell(row=row, column=8, value=n3)
                ws.cell(row=row, column=9, value=n4)
                ws.cell(row=row, column=10, value=v1)
                ws.cell(row=row, column=11, value=v2)
                ws.cell(row=row, column=12, value=v3)
                ws.cell(row=row, column=13, value=v4)
            row += 1
    last = row - 1

    for r in range(DETAIL_FIRST, last + 1):
        # 推荐主词条
        q = (f'=IFERROR(IF($B{r}="生之花","生命值",IF($B{r}="死之羽","攻击力",'
             f'IF($B{r}="时之沙",INDEX({Q_LIB}!$J$3:$J$135,MATCH($A{r},{Q_LIB}!$A$3:$A$135,0)),'
             f'IF($B{r}="空之杯",INDEX({Q_LIB}!$K$3:$K$135,MATCH($A{r},{Q_LIB}!$A$3:$A$135,0)),'
             f'INDEX({Q_LIB}!$L$3:$L$135,MATCH($A{r},{Q_LIB}!$A$3:$A$135,0)))))),"")')
        ws.cell(row=r, column=17, value=q)
        ws.cell(row=r, column=18, value=f'=IFERROR({lib_idx("D", r)},"")')
        # 权重
        for off, name_col in enumerate(["F", "G", "H", "I"]):
            weight_col = get_column_letter(19 + off)
            ws.cell(row=r, column=19 + off, value=(
                f'=IF(${name_col}{r}="",0,IF(ISNUMBER(SEARCH("、"&${name_col}{r}&"、","、"&$AD{r}&"、")),1,'
                f'IF(ISNUMBER(SEARCH("、"&${name_col}{r}&"、","、"&$AE{r}&"、")),0.75,'
                f'IF(ISNUMBER(SEARCH("、"&${name_col}{r}&"、","、"&$AF{r}&"、")),0.5,0))))'
            ))
        # roll
        for off, (name_col, val_col) in enumerate([("F", "J"), ("G", "K"), ("H", "L"), ("I", "M")]):
            ws.cell(row=r, column=23 + off, value=(
                f'=IF(OR(${name_col}{r}="",${val_col}{r}=""),0,'
                f"${val_col}{r}/IFERROR(INDEX('{SRC_SHEET}'!$N$3:$N$12,MATCH(${name_col}{r},'{SRC_SHEET}'!$M$3:$M$12,0)),1))"
            ))
        # 主词条匹配
        ws.cell(row=r, column=27, value=(
            f'=IF($D{r}="",0,IF(ISNUMBER(SEARCH("、"&$D{r}&"、","、"&$Q{r}&"、")),1,0))'
        ))
        # 有效词条
        p1 = f'IF($S{r}>0,$F{r},"")'
        p2 = f'IF(AND($G{r}<>"",$T{r}>0),IF({p1}="","","、")&$G{r},"")'
        p3 = f'IF(AND($H{r}<>"",$U{r}>0),IF({p1}&{p2}="","","、")&$H{r},"")'
        p4 = f'IF(AND($I{r}<>"",$V{r}>0),IF({p1}&{p2}&{p3}="","","、")&$I{r},"")'
        ws.cell(row=r, column=14, value=f'=IF($F{r}="","",{p1}&{p2}&{p3}&{p4})')
        # 单品评分
        ws.cell(row=r, column=15, value=(
            f'=IF(OR($E{r}="",$D{r}=""),"",ROUND(25*MIN(1,$E{r}/20)+20*$AA{r}'
            f'+55*MIN(1,($W{r}+$X{r}+$Y{r}+$Z{r})/8),1))'
        ))
        # 调整提示
        ws.cell(row=r, column=16, value=(
            f'=IF(AND($E{r}="",$D{r}=""),"",'
            f'IF($E{r}="","填强化等级",IF($E{r}<20,"强化 +"&$E{r}&" → +20",""))'
            f'&IF($D{r}=""," · 填主词条",IF($AA{r}<1," · 主词条不符(推荐"&$Q{r}&")",""))'
            f'&IF(AND($F{r}<>"",$S{r}=0,$G{r}<>"",$T{r}=0,$H{r}<>"",$U{r}=0,$I{r}<>"",$V{r}=0),'
            f'" · 副词条全部无效",""))'
        ))
        ws.cell(row=r, column=28, value=f'=$A{r}&"|"&$B{r}')
        ws.cell(row=r, column=29, value=f'=IF($O{r}="",1000000,$O{r}-ROW()/100000)')

        # 样式
        for ci in range(1, 33):
            c = ws.cell(row=r, column=ci)
            c.font = F(9)
            c.border = BORDER
            c.alignment = A("center", "center", True)
            if ci in (3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13):
                c.fill = fill(INPUT)
            elif ci in (30, 31, 32):
                c.fill = fill("F7F2FB")
            else:
                c.fill = fill(purple[r % 2])
        ws.row_dimensions[r].height = 22

    # 数据验证
    def add_dv(formula, rng, allow_blank=True):
        dv = DataValidation(type="list", formula1=formula, allow_blank=allow_blank, showDropDown=False)
        dv.error = "请从下拉列表中选择"
        dv.prompt = "从下拉列表选择"
        ws.add_data_validation(dv)
        dv.add(rng)

    add_dv("=套装列表", f"C{DETAIL_FIRST}:C{last}")
    add_dv("=主词条列表", f"D{DETAIL_FIRST}:D{last}")
    add_dv("=词条列表", f"F{DETAIL_FIRST}:I{last}")
    dv_num = DataValidation(type="list", formula1="=强化列表", allow_blank=True, showDropDown=False)
    ws.add_data_validation(dv_num)
    dv_num.add(f"E{DETAIL_FIRST}:E{last}")
    dv_val = DataValidation(type="decimal", operator="between", formula1="0", formula2="9999", allow_blank=True)
    ws.add_data_validation(dv_val)
    dv_val.add(f"J{DETAIL_FIRST}:M{last}")

    # 条件格式
    ws.conditional_formatting.add(
        f"O{DETAIL_FIRST}:O{last}",
        ColorScaleRule(start_type="num", start_value=40, start_color="F5B7B1",
                       mid_type="num", mid_value=75, mid_color="FCF3CF",
                       end_type="num", end_value=100, end_color="ABEBC6"))
    ws.conditional_formatting.add(
        f"N{DETAIL_FIRST}:N{last}",
        FormulaRule(formula=[f'LEN($N{DETAIL_FIRST})>0'], fill=fill("D5F5E3"), font=F(9, True, "1E6B47")))

    ws.column_dimensions.group("S", "AF", hidden=True)
    ws.freeze_panes = "C3"
    ws.auto_filter.ref = f"A2:AC{last}"
    ws.sheet_view.zoomScale = 90
    for _r in range(DETAIL_FIRST, last + 1):
        for _c in range(3, 14):
            unlock(ws.cell(row=_r, column=_c))
    protect(ws)
    return ws


# ============ 养成总表 ============
MAIN_HEADERS = [
    ("序号", 5, "auto"), ("角色名", 12, "in"), ("元素", 6, "auto"),
    ("当前\n等级", 7.5, "in"), ("目标\n等级", 7.5, "in"), ("等级\n进度", 7, "auto"), ("下一\n突破", 7.5, "auto"),
    ("当前武器", 16, "in"), ("精炼", 5.5, "in"), ("武器\n等级", 7, "in"), ("武器\n匹配", 12, "auto"),
    ("圣遗物套装", 18, "auto"), ("主词条\n匹配", 8, "auto"), ("总级", 6.5, "auto"), ("完成度", 7, "auto"),
    ("圣遗物\n评分", 8, "auto"), ("圣遗物\n状态", 8, "auto"), ("调整建议", 34, "auto"),
    ("普攻", 5.5, "in"), ("战技", 5.5, "in"), ("爆发", 5.5, "in"), ("总级", 6.5, "auto"),
    ("完成度", 7, "auto"), ("参考天赋", 12, "auto"), ("加点顺序", 15, "auto"),
    ("暴击率\n%", 7, "in"), ("暴击\n伤害%", 7, "in"), ("攻击力", 8, "in"), ("生命值", 8, "in"),
    ("防御力", 7, "in"), ("元素\n精通", 7, "in"), ("元素\n充能%", 7, "in"),
    ("面板\n达标", 7, "auto"), ("面板\n判定", 8, "auto"), ("期望值", 7, "in"),
    ("当前评分", 7.5, "auto"), ("评分\n差距", 7, "auto"), ("练好判定", 10, "auto"),
    ("更新日期", 11, "in"), ("备注", 24, "in"), ("特殊标记", 10, "in"),
    ("命座", 6, "in"), ("拥有状态", 10, "in"),
]

MAIN_GROUPS = [
    ("角色", 1, 3), ("等级", 4, 7), ("武器", 8, 11), ("圣遗物", 12, 18),
    ("天赋", 19, 25), ("面板", 26, 32), ("评估", 33, 38), ("其他", 39, 41),
    ("状态", 42, 43),
]

MAIN_INPUT_COLS = {"B", "D", "E", "H", "I", "J", "S", "T", "U", "Z", "AA", "AB",
                   "AC", "AD", "AE", "AF", "AI", "AM", "AN", "AO", "AP", "AQ"}

HELPERS = [
    ("AR", "目标暴击率", 8), ("AS", "目标暴伤", 8), ("AT", "目标攻击力", 8),
    ("AU", "目标生命值", 8), ("AV", "目标防御力", 8), ("AW", "目标精通", 8),
    ("AX", "目标充能", 8), ("AY", "面板目标数", 7), ("AZ", "面板命中数", 7),
    ("BA", "待练键", 9), ("BB", "下一步", 10), ("BC", "目标总级", 7), ("BD", "主套装", 18), ("BE", "最低分键", 9),
]


def main_formulas(r):
    f = {}
    f["A"] = f'=IF($B{r}="","",COUNTA($B$3:$B{r}))'
    f["C"] = f'=IF($B{r}="","",IFERROR({lib_idx("B", r)},""))'
    f["F"] = f'=IF(OR($B{r}="",$D{r}="",$E{r}=""),"",MIN(1,$D{r}/$E{r}))'
    f["G"] = (f'=IF(OR($B{r}="",$D{r}=""),"",IF($D{r}>=90,"已满级",'
              f'IF($D{r}>=80,90,IF($D{r}>=70,80,IF($D{r}>=60,70,IF($D{r}>=50,60,IF($D{r}>=40,50,40)))))))')
    f["K"] = (f'=IF(OR($B{r}="",$H{r}=""),"",IF($H{r}={lib_idx("Q", r)},"✔ 推荐武器",'
              f'IF(ISNUMBER(SEARCH($H{r},{lib_idx("S", r)}&" / "&{lib_idx("T", r)}&" / "&{lib_idx("U", r)})),"△ 可用武器","✘ 非推荐")))')
    f["BB"] = (f'=IF($B{r}="","",IFERROR(IF(INDEX({Q_DETAIL}!{DR("C")},MATCH($B{r}&"|生之花",{Q_DETAIL}!{DR("AB")},0))="","",'
               f'INDEX({Q_DETAIL}!{DR("C")},MATCH($B{r}&"|生之花",{Q_DETAIL}!{DR("AB")},0))),""))')
    f["L"] = (f'=IF($B{r}="","",IF($BB{r}="",IFERROR(IF({lib_idx("D", r)}="","WIKI 待补","推荐："&{lib_idx("D", r)}),"WIKI 待补"),'
              f'$BB{r}&IF(COUNTIFS({Q_DETAIL}!{DR("A")},$B{r},{Q_DETAIL}!{DR("C")},$BB{r})>=4," 4件套",'
              f'IF(COUNTIFS({Q_DETAIL}!{DR("A")},$B{r},{Q_DETAIL}!{DR("C")},$BB{r})=2," 2件套"," 散搭"))))')
    f["M"] = (f'=IF($B{r}="","",IF(COUNTIFS({Q_DETAIL}!{DR("A")},$B{r},{Q_DETAIL}!{DR("E")},">0")=0,"",'
              f'MAX(0,SUMIFS({Q_DETAIL}!{DR("AA")},{Q_DETAIL}!{DR("A")},$B{r})-2)&"/3"))')
    f["N"] = (f'=IF($B{r}="","",IF(COUNTIFS({Q_DETAIL}!{DR("A")},$B{r},{Q_DETAIL}!{DR("E")},">0")=0,"",'
              f'SUMIFS({Q_DETAIL}!{DR("E")},{Q_DETAIL}!{DR("A")},$B{r})))')
    f["O"] = (f'=IF($B{r}="","",IF(COUNTIFS({Q_DETAIL}!{DR("A")},$B{r},{Q_DETAIL}!{DR("E")},">0")=0,"",'
              f'COUNTIFS({Q_DETAIL}!{DR("A")},$B{r},{Q_DETAIL}!{DR("E")},">0")/5))')
    f["P"] = (f'=IF($B{r}="","",IF(COUNTIFS({Q_DETAIL}!{DR("A")},$B{r},{Q_DETAIL}!{DR("O")},">0")=0,"",'
              f'ROUND(SUMIFS({Q_DETAIL}!{DR("O")},{Q_DETAIL}!{DR("A")},$B{r})/5,1)))')
    f["Q"] = (f'=IF($B{r}="","",IF(COUNTIFS({Q_DETAIL}!{DR("A")},$B{r},{Q_DETAIL}!{DR("O")},">0")=0,"未开始",'
              f'IF($P{r}>=75,"已毕业",IF($P{r}>=55,"需调整","未毕业"))))')
    f["BA"] = f'=IFERROR({lib_idx("AC", r)},27)'
    f["R"] = (f'=IF(OR($B{r}="",$Q{r}="已毕业",$Q{r}="未开始"),"",'
               f'IF(COUNTIFS({Q_DETAIL}!{DR("A")},$B{r},{Q_DETAIL}!{DR("E")},">0")<5,'
               f'"先补齐 5 件圣遗物（当前 "&COUNTIFS({Q_DETAIL}!{DR("A")},$B{r},{Q_DETAIL}!{DR("E")},">0")&"/5）",'
               f'IFERROR(IF(INDEX({Q_DETAIL}!{DR("P")},MATCH($BC{r},{Q_DETAIL}!{DR("AC")},0))="","副词条质量一般，可继续优化",'
               f'"优先调整 "&INDEX({Q_DETAIL}!{DR("B")},MATCH($BC{r},{Q_DETAIL}!{DR("AC")},0))&"："'
               f'&INDEX({Q_DETAIL}!{DR("P")},MATCH($BC{r},{Q_DETAIL}!{DR("AC")},0))),"词条质量一般，可继续优化")))')
    f["V"] = f'=IF($B{r}="","",SUM($S{r}:$U{r}))'
    f["W"] = f'=IF(OR($B{r}="",$BA{r}=""),"",MIN(1.2,$V{r}/$BA{r}))'
    f["X"] = f'=IF($B{r}="","",IFERROR("目标 "&{lib_idx("Z", r)}&" / "&{lib_idx("AA", r)}&" / "&{lib_idx("AB", r)},""))'
    f["Y"] = f'=IF($B{r}="","",IFERROR({lib_idx("Y", r)},""))'
    f["AP"] = f'=IFERROR({lib_idx("AE", r)},"")'
    f["AQ"] = f'=IFERROR({lib_idx("AF", r)},"")'
    f["AR"] = f'=IFERROR({lib_idx("AG", r)},"")'
    f["AS"] = f'=IFERROR({lib_idx("AH", r)},"")'
    f["AT"] = f'=IFERROR({lib_idx("AI", r)},"")'
    f["AU"] = f'=IFERROR({lib_idx("AJ", r)},"")'
    f["AV"] = f'=IFERROR({lib_idx("AK", r)},"")'
    f["AW"] = f'=IF($B{r}="","",SUMPRODUCT(--($AP{r}:$AV{r}>0)))'
    f["AX"] = f'=IF($B{r}="","",SUMPRODUCT(($AP{r}:$AV{r}>0)*($Z{r}:$AF{r}>=$AP{r}:$AV{r})))'
    f["AG"] = f'=$AW{r}'
    f["AH"] = (f'=IF($B{r}="","",IF($AW{r}=0,"无参考",IF($AX{r}>=$AW{r},"达标",'
               f'IF($AX{r}/$AW{r}>=0.7,"接近","未达标"))))')
    f["AI"] = f'=IF($B{r}="","",IFERROR({lib_idx("AD", r)},80))'
    f["AJ"] = (f'=IF($B{r}="","",ROUND(100*(0.15*N($F{r})'
               f'+0.10*IF($K{r}="✔ 推荐武器",1,IF($K{r}="△ 可用武器",0.75,0.5))'
               f'+0.10*MIN(1,N($J{r})/90)'
               f'+0.35*MIN(1,N($O{r})*0.7+N($M{r})/3*0.3)'
               f'+0.15*MIN(1.2,N($W{r}))'
               f'+0.15*IF($AH{r}="无参考",1,IF($AH{r}="达标",1,IF($AH{r}="接近",0.75,0.5)))),1))')
    f["AK"] = f'=IF($B{r}="","",ROUND(N($AJ{r})-N($AI{r}),1))'
    f["AL"] = (f'=IF($B{r}="","",IF(COUNTIFS({Q_DETAIL}!{DR("A")},$B{r},{Q_DETAIL}!{DR("E")},">0")<5,"待补圣遗物",'
               f'IF($AW{r}=0,"缺参考面板",IF(AND($J{r}>=90,$K{r}="✔ 推荐武器",$Q{r}="已毕业",$V{r}>=$BA{r},$AX{r}>=$AW{r}),"✔ 已练好","未练好"))))')
    f["AY"] = (f'=IF(OR($B{r}="",$AL{r}="✔ 已练好",$AO{r}="暂缓",$AO{r}="不练了"),"",'
               f'(N($AI{r})-N($AJ{r}))*10+MAX(0,N($E{r})-N($D{r}))+MAX(0,90-N($J{r}))'
               f'+IF($AL{r}="待补圣遗物",15,0)+IF($AL{r}="缺参考面板",8,0)+IF($AH{r}="未达标",10,IF($AH{r}="接近",5,0))-ROW()/100000)')
    f["AZ"] = (f'=IF($B{r}="","",IF($AL{r}="✔ 已练好","已练好",IF($AL{r}="待补圣遗物","补圣遗物",'
               f'IF($AL{r}="缺参考面板","补面板参考",IF($D{r}<$E{r},"升等级",IF($J{r}<90,"升武器",IF($V{r}<$BA{r},"点天赋",'
               f'IF($Q{r}<>"已毕业","刷圣遗物",IF($AH{r}<>"达标","调面板","查漏补缺")))))))))')
    f["BC"] = (f'=IF($B{r}="","",IF(COUNTIFS({Q_DETAIL}!{DR("A")},$B{r},{Q_DETAIL}!{DR("E")},">0")=0,"",'
               f'SUMPRODUCT(MIN(({Q_DETAIL}!{DR("A")}=$B{r})*{Q_DETAIL}!{DR("AC")}+({Q_DETAIL}!{DR("A")}<>$B{r})*1000000))))')
    # v0.5：助手列右移两列（AP→AR … BC→BE），AP=命座、AQ=拥有状态
    shift = {"AP": "AR", "AQ": "AS", "AR": "AT", "AS": "AU", "AT": "AV", "AU": "AW",
             "AV": "AX", "AW": "AY", "AX": "AZ", "AY": "BA", "AZ": "BB",
             "BA": "BC", "BB": "BD", "BC": "BE"}
    pat = re.compile(r"\$([A-Z]{1,2})(?=\d)")
    for _k in list(f):
        f[_k] = pat.sub(lambda m: "$" + shift.get(m.group(1), m.group(1)), f[_k])
    # 未拥有角色：自动列全部留空（拥有状态在 AQ）
    for _k in list(f):
        f[_k] = f[_k].replace(f'$B{r}=""', f'OR($B{r}="",$AQ{r}="未拥有")')
    # 键名同步右移，否则公式会写回旧助手列/覆盖数据列
    f = {shift.get(_k, _k): _v for _k, _v in f.items()}
    f["AG"] = f'=IF(OR($AQ{r}="未拥有",$AY{r}=0),"",$AZ{r}&"/"&$AY{r})'  # 面板命中/目标（无参考留空）
    return f


def build_main(wb, rows):
    ws = wb.create_sheet(MAIN)
    ws.sheet_properties.tabColor = "1F3B57"
    # 分组
    for title, c1, c2 in MAIN_GROUPS:
        dark, light = GROUPS[title]
        merge(ws, f"{get_column_letter(c1)}1:{get_column_letter(c2)}1", f"  {title}",
              font=F(10, True, "FFFFFF"), fill_color=dark, align=A("center", "center", False, 0), border=None)
    ws.row_dimensions[1].height = 26

    for i, (h, w, kind) in enumerate(MAIN_HEADERS, 1):
        c = ws.cell(row=2, column=i, value=h)
        c.font = F(9, True, "FFFFFF")
        c.fill = fill("2C3E50")
        c.alignment = align_hdr(h)
        c.border = BORDER
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.row_dimensions[2].height = 46

    for hcol, title, width in HELPERS:
        ws.column_dimensions[hcol].width = width
        c = ws.cell(row=2, column=ws[f"{hcol}2"].column, value=title)
        c.font = F(8, True, MUTED)
        c.alignment = A("center", "center", True)
        c.border = BORDER
    ws.column_dimensions.group("AR", "BE", hidden=True)
    for _c in ("AR", "AS", "AT", "AU", "AV", "AW", "AX", "AY", "AZ", "BA", "BB", "BC", "BD", "BE"):
        ws.column_dimensions[_c].hidden = True

    # v0.5：预置全部角色名（未拥有），真实数据由导入脚本写入
    for idx, _row in enumerate(rows):
        _ri = DATA_FIRST + idx
        if _ri > DATA_LAST:
            break
        ws[f"B{_ri}"] = _row["name"]
        ws[f"AQ{_ri}"] = "未拥有"
    # 数字格式
    for r in range(DATA_FIRST, DATA_LAST + 1):
        for col in ["F", "O", "W"]:
            ws[f"{col}{r}"].number_format = "0%"
        for col in ["P", "AJ", "AK"]:
            ws[f"{col}{r}"].number_format = "0.0"
        ws[f"AM{r}"].number_format = "yyyy-mm-dd"

    # 数据行：样式 + 公式
    for r in range(DATA_FIRST, DATA_LAST + 1):
        ws.row_dimensions[r].height = 24
        for ci, (_h, _w, kind) in enumerate(MAIN_HEADERS, 1):
            cell = ws.cell(row=r, column=ci)
            cell.font = F(9, ci == 2)
            cell.border = BORDER
            cell.alignment = A("left", "center", False)
            cell.fill = fill(INPUT if kind == "in" else "FFFFFF")
            if kind == "in":
                unlock(cell)
        for _col, _fml in main_formulas(r).items():
            ws["{}{}".format(_col, r)] = _fml

    # 数据验证
    def add_dv(formula, rng):
        dv = DataValidation(type="list", formula1=formula, allow_blank=True, showDropDown=False)
        dv.error = "请从下拉列表中选择"
        dv.prompt = "从下拉列表选择"
        ws.add_data_validation(dv)
        dv.add(rng)

    add_dv("=角色列表", f"B{DATA_FIRST}:B{DATA_LAST}")
    add_dv("=目标等级列表", f"E{DATA_FIRST}:E{DATA_LAST}")
    add_dv("=精炼列表", f"I{DATA_FIRST}:I{DATA_LAST}")
    add_dv("=天赋等级列表", f"S{DATA_FIRST}:U{DATA_LAST}")
    add_dv("=标记列表", f"AO{DATA_FIRST}:AO{DATA_LAST}")
    add_dv("=拥有列表", f"AQ{DATA_FIRST}:AQ{DATA_LAST}")
    dv_seat = DataValidation(type="list", formula1="=命座列表", allow_blank=True, showDropDown=False)
    ws.add_data_validation(dv_seat)
    dv_seat.add(f"AP{DATA_FIRST}:AP{DATA_LAST}")
    dv_expect = DataValidation(type="decimal", operator="between", formula1="0", formula2="100", allow_blank=True)
    ws.add_data_validation(dv_expect)
    dv_expect.add(f"AI{DATA_FIRST}:AI{DATA_LAST}")
    dv_num = DataValidation(type="decimal", operator="between", formula1="0", formula2="99999", allow_blank=True)
    ws.add_data_validation(dv_num)
    dv_num.add(f"D{DATA_FIRST}:D{DATA_LAST}")
    dv_num.add(f"J{DATA_FIRST}:J{DATA_LAST}")
    dv_num.add(f"Z{DATA_FIRST}:AF{DATA_LAST}")

    # 条件格式
    for el, (bg, fg) in ELEMENTS.items():
        ws.conditional_formatting.add(f"C{DATA_FIRST}:C{DATA_LAST}",
            FormulaRule(formula=[f'$C{DATA_FIRST}="{el}"'], fill=fill(bg), font=F(9, True, fg)))
    for text, (bg, fg) in WEAPON_COLORS.items():
        ws.conditional_formatting.add(f"K{DATA_FIRST}:K{DATA_LAST}",
            CellIsRule(operator="equal", formula=[f'"{text}"'], fill=fill(bg), font=F(9, True, fg)))
    for text, (bg, fg) in STATUS_COLORS.items():
        ws.conditional_formatting.add(f"Q{DATA_FIRST}:Q{DATA_LAST}",
            CellIsRule(operator="equal", formula=[f'"{text}"'], fill=fill(bg), font=F(9, True, fg)))
    for text, (bg, fg) in PANEL_COLORS.items():
        ws.conditional_formatting.add(f"AH{DATA_FIRST}:AH{DATA_LAST}",
            CellIsRule(operator="equal", formula=[f'"{text}"'], fill=fill(bg), font=F(9, True, fg)))
    for text, (bg, fg) in READY_COLORS.items():
        ws.conditional_formatting.add(f"AL{DATA_FIRST}:AL{DATA_LAST}",
            CellIsRule(operator="equal", formula=[f'"{text}"'], fill=fill(bg), font=F(9, True, fg)))
    ws.conditional_formatting.add(f"M{DATA_FIRST}:M{DATA_LAST}",
        CellIsRule(operator="equal", formula=['"3/3"'], fill=fill("D5F5E3"), font=F(9, True, "1E6B47")))
    ws.conditional_formatting.add(f"M{DATA_FIRST}:M{DATA_LAST}",
        CellIsRule(operator="equal", formula=['"2/3"'], fill=fill("FCF3CF"), font=F(9, True, "A66B00")))
    ws.conditional_formatting.add(f"F{DATA_FIRST}:F{DATA_LAST}",
        DataBarRule(start_type="num", start_value=0, end_type="num", end_value=1, color="8EA9DB"))
    ws.conditional_formatting.add(f"O{DATA_FIRST}:O{DATA_LAST}",
        DataBarRule(start_type="num", start_value=0, end_type="num", end_value=1, color="B39DDB"))
    ws.conditional_formatting.add(f"W{DATA_FIRST}:W{DATA_LAST}",
        DataBarRule(start_type="num", start_value=0, end_type="num", end_value=1, color="F0B27A"))
    ws.conditional_formatting.add(f"AJ{DATA_FIRST}:AJ{DATA_LAST}",
        ColorScaleRule(start_type="num", start_value=40, start_color="F5B7B1",
                       mid_type="num", mid_value=70, mid_color="FCF3CF",
                       end_type="num", end_value=95, end_color="ABEBC6"))
    ws.conditional_formatting.add(f"AK{DATA_FIRST}:AK{DATA_LAST}",
        CellIsRule(operator="greaterThanOrEqual", formula=["0"], font=F(9, True, "1E6B47")))
    ws.conditional_formatting.add(f"AK{DATA_FIRST}:AK{DATA_LAST}",
        CellIsRule(operator="lessThan", formula=["0"], font=F(9, True, "A93226")))
    for _text, (_bg, _fg) in OWNED_COLORS.items():
        ws.conditional_formatting.add(f"AQ{DATA_FIRST}:AQ{DATA_LAST}",
            CellIsRule(operator="equal", formula=[f'"{_text}"'], fill=fill(_bg), font=F(9, True, _fg)))

    # 批注
    notes = {
        "B2": "从下拉选角色；没有的角色可到「配装数据库」补一行。",
        "H2": "填你实际装备的武器名（与 WIKI 推荐比对）。",
        "I2": "精炼 1-5。",
        "Z2": "以下 7 列填当前面板数值：暴击率/暴击伤害填百分比数字（如 85、180），其他填数值。",
        "AK2": "评分差距 = 当前评分 - 期望值。",
        "AL2": "自动判定：武器满级+推荐、天赋达标、圣遗物五件齐全且评分≥期望、面板达标 → 已练好。",
        "AI2": "默认取配装库建议期望值，可直接改成你自己的目标分。",
        "AO2": "标记 暂缓 / 不练了 后，不再出现在「今日练谁」；留空=在练。",
        "R2": "自动根据圣遗物明细里评分最低的一件，提示先调整哪个部位。",
        "AP2": "命座数量 0-6，导入账号数据时自动写入。",
        "AQ2": "已有/未拥有；未拥有角色只保留名字，其他列自动留空。",
    }
    for addr, text in notes.items():
        ws[addr].comment = Comment(text, "养成表")
        ws[addr].comment.width = 260
        ws[addr].comment.height = 80

    ws.freeze_panes = "C3"
    ws.auto_filter.ref = f"A2:AQ{DATA_LAST}"
    ws.sheet_view.zoomScale = 85
    ws.print_title_rows = "1:2"
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    protect(ws)
    return ws


# -*- coding: utf-8 -*-
# ============ 今日练谁 ============
def build_today(wb):
    ws = wb.create_sheet(TODAY)
    ws.sheet_properties.tabColor = "C0392B"
    merge(ws, "A1:J1", "今日练谁 · 我的 5 人清单", font=F(16, True, "FFFFFF"), fill_color=NAVY,
          align=A("left", "center", False, 1), border=None)
    ws.row_dimensions[1].height = 46
    merge(ws, "A2:J2",
          "从下拉选择最多 5 个想练的角色；当该角色在总表 AO 列标记「已练好」或自动判定为已练好，本格会自动移除；每格下方给出毕业面板与推荐方案。",
          font=F(9, False, MUTED), fill_color="F8FAFC", align=A("left", "center", False, 1), border=None)
    ws.row_dimensions[2].height = 24

    for col, w in {"A": 10, "B": 12, "C": 14, "D": 14, "E": 14, "F": 14,
                   "G": 14, "H": 14, "I": 14, "J": 22}.items():
        ws.column_dimensions[col].width = w

    merge(ws, "A4:J4",
          f'="已拥有 "&COUNTIF({Q_MAIN}!$AQ$3:$AQ$135,"已有")&" 名 ｜ 未拥有 "&COUNTIF({Q_MAIN}!$AQ$3:$AQ$135,"未拥有")'
          f'&" 名 ｜ 已练好 "&COUNTIF({Q_MAIN}!$AL$3:$AL$135,"✔ 已练好")&" 名 ｜ 待补圣遗物 "&COUNTIF({Q_MAIN}!$AL$3:$AL$135,"待补圣遗物")&" 名"',
          font=F(10, True, INK), fill_color="FFFFFF", align=A("left", "center", False, 1), border=None)
    ws.row_dimensions[4].height = 28

    def mr(r): return f'MATCH($C{r},{Q_MAIN}!$B$3:$B$135,0)'
    def lr(r): return f'MATCH($C{r},{Q_LIB}!$A$3:$A$135,0)'
    def mainidx(r, col): return f'INDEX({Q_MAIN}!${col}$3:${col}$135,{mr(r)})'
    def libidx(r, col): return f'INDEX({Q_LIB}!${col}$3:${col}$135,{lr(r)})'
    def pct(r, col):
        idx = mainidx(r, col)
        return f'IF(IFERROR({idx},0)=0,"—",TEXT({idx},"0.0")&"%")'
    def num(r, col):
        idx = mainidx(r, col)
        return f'IF(IFERROR({idx},0)=0,"—",{idx})'
    def lpct(r, col):
        idx = libidx(r, col)
        return f'IF(IFERROR({idx},0)=0,"—",TEXT({idx},"0.0")&"%")'
    def lnum(r, col):
        idx = libidx(r, col)
        return f'IF(IFERROR({idx},0)=0,"—",{idx})'

    titles = ["① 第 1 位", "② 第 2 位", "③ 第 3 位", "④ 第 4 位", "⑤ 第 5 位"]
    for i in range(5):
        r0 = 7 + i * 6
        section(ws, r0 - 1, titles[i], end_col="J")
        merge(ws, f"A{r0}:B{r0}", "角色", font=F(10, True, INK), fill_color="E8EEF5",
              align=A("center", "center"))
        c = ws.cell(row=r0, column=3, value=None)
        c.fill = fill(INPUT)
        c.font = F(11, True, "C0392B")
        c.alignment = A("center", "center")
        c.border = BORDER
        status = (
            f'=IF($C{r0}="","请从下拉选择角色（最多 5 位）",IFERROR(IF({mainidx(r0,"AQ")}="未拥有","⚠ 未拥有，无法规划",'
            f'IF(OR({mainidx(r0,"AO")}="已练好",{mainidx(r0,"AL")}="✔ 已练好"),"✔ 已练好，已从清单移除",'
            f'"在练 ｜ Lv "&{mainidx(r0,"D")}&" / "&{mainidx(r0,"E")}&" ｜ 评分 "&TEXT({mainidx(r0,"AJ")},"0.0")'
            f'&" / 期望 "&TEXT({mainidx(r0,"AI")},"0.0")&" ｜ "&{mainidx(r0,"AL")})),'
            f'"角色名不在总表，请到『配装数据库』补录"))')
        merge(ws, f"D{r0}:J{r0}", status, font=F(10, True, INK), fill_color="FFFFFF",
              align=A("left", "center", False, 1))
        ws.row_dimensions[r0].height = 30

        def _gap_expr():
            stats = [("暴击率", "Z", "AE"), ("暴击伤害", "AA", "AF"), ("攻击力", "AB", "AG"),
                     ("生命值", "AC", "AH"), ("防御力", "AD", "AI"), ("元素精通", "AE", "AJ"),
                     ("元素充能", "AF", "AK")]
            parts = []
            for label, cur_col, tgt_col in stats:
                cur = f'IFERROR({mainidx(r0,cur_col)},0)'
                tgt = f'IFERROR({libidx(r0,tgt_col)},0)'
                parts.append(f'IF(AND({tgt}>0,{cur}<{tgt}),"；{label} "&TEXT({cur},"0.0")&"/"&TEXT({tgt},"0.0"),"")')
            expr = "&".join(parts)
            return f'IF({mainidx(r0,"AH")}="未达标"," ｜ 面板未达标："&SUBSTITUTE({expr},"；","",1)&" ｜ 期望评分 "&TEXT({mainidx(r0,"AI")},"0.0"),"")'

        rows = [
            ("毕业面板", f'=IF(LEFT($D{r0},2)<>"在练","","暴击率 "&{lpct(r0,"AE")}&" ｜ 暴击伤害 "&{lpct(r0,"AF")}'
                         f'&" ｜ 攻击力 "&{lnum(r0,"AG")}&" ｜ 生命值 "&{lnum(r0,"AH")}&" ｜ 防御力 "&{lnum(r0,"AI")}'
                         f'&" ｜ 元素精通 "&{lnum(r0,"AJ")}&" ｜ 元素充能 "&{lpct(r0,"AK")})', "E9F0F7"),
            ("当前面板", f'=IF(LEFT($D{r0},2)<>"在练","","暴击率 "&{pct(r0,"Z")}&" ｜ 暴击伤害 "&{pct(r0,"AA")}'
                         f'&" ｜ 攻击力 "&{num(r0,"AB")}&" ｜ 生命值 "&{num(r0,"AC")}&" ｜ 防御力 "&{num(r0,"AD")}'
                         f'&" ｜ 元素精通 "&{num(r0,"AE")}&" ｜ 元素充能 "&{pct(r0,"AF")})', "FFFFFF"),
            ("推荐方案", f'=IF(LEFT($D{r0},2)<>"在练","","套装 "&{libidx(r0,"D")}&" ｜ 武器 "&{libidx(r0,"Q")}'
                         f'&" ｜ 主词条 "&{libidx(r0,"J")}&" / "&{libidx(r0,"K")}&" / "&{libidx(r0,"L")}'
                         f'&" ｜ 有效词条 "&{libidx(r0,"N")}&" ｜ 加点 "&{libidx(r0,"Y")})', "FDF1E0"),
            ("备注", f'=IF(LEFT($D{r0},2)<>"在练","",HYPERLINK("https://www.miyoushe.com/ys/search?keyword="&$C{r0},"打开米游社攻略搜索："&$C{r0})'
                     f'&" ｜ WIKI 攻略更新："&IFERROR({libidx(r0,"AN")},"")&{_gap_expr()})', "F8FAFC"),
        ]
        for j, (label, formula, bg) in enumerate(rows, 1):
            rr = r0 + j
            merge(ws, f"A{rr}:B{rr}", label, font=F(9, True, MUTED), fill_color="F8FAFC",
                  align=A("center", "center"))
            merge(ws, f"C{rr}:J{rr}", formula, font=F(9), fill_color=bg,
                  align=A("left", "center", True, 1))
            ws.row_dimensions[rr].height = 42

    dv = DataValidation(type="list", formula1="=角色列表", allow_blank=True, showDropDown=False)
    ws.add_data_validation(dv)
    for i in range(5):
        dv.add(f"C{7 + i * 6}")
        unlock(ws[f"C{7 + i * 6}"])

    for i in range(5):
        r0 = 7 + i * 6
        for el, (bg, fg) in ELEMENTS.items():
            ws.conditional_formatting.add(f"C{r0}", FormulaRule(formula=[f'$C{r0}="{el}"'], fill=fill(bg)))
        ws.conditional_formatting.add(f"D{r0}", FormulaRule(formula=[f'ISNUMBER(SEARCH("已练好",$D{r0}))'],
                                                            fill=fill("C6F6D5"), font=F(10, True, "1E6B47")))
        ws.conditional_formatting.add(f"D{r0}", FormulaRule(formula=[f'ISNUMBER(SEARCH("未拥有",$D{r0}))'],
                                                            fill=fill("EEF1F5"), font=F(10, True, "8A94A0")))
        ws.conditional_formatting.add(f"D{r0}", FormulaRule(formula=[f'ISNUMBER(SEARCH("在练",$D{r0}))'],
                                                            fill=fill("EAF2F8"), font=F(10, True, "21618C")))

    merge(ws, "A40:J40",
          "说明：这里不再自动排名，由你自己选 5 个想练的角色；角色一旦标记「已练好」（或总表判定为已练好），该格自动清空并从清单移除。",
          font=F(9, False, MUTED), fill_color="F8FAFC", align=A("left", "center", False, 1), border=None)
    merge(ws, "A41:J41",
          "未拥有角色会提示无法规划；毕业面板取自 WIKI 攻略页，推荐方案取自配装数据库；备注可直接跳转米游社攻略搜索。",
          font=F(9, False, MUTED), fill_color="F8FAFC", align=A("left", "center", False, 1), border=None)
    ws.sheet_view.showGridLines = False
    ws.sheet_view.zoomScale = 100
    protect(ws)


# ============ 角色参考卡 ============
def build_card(wb, rows):
    ws = wb.create_sheet(CARD)
    ws.sheet_properties.tabColor = "2E86C1"
    merge(ws, "A1:H1", "角色参考卡", font=F(18, True, "FFFFFF"), fill_color=NAVY,
          align=A("left", "center", False, 1), border=None)
    ws.row_dimensions[1].height = 52
    merge(ws, "A2:H2", "选择角色，查看该角色的推荐配装、有效词条与毕业面板参考（数据来自 biligame 原神 WIKI 攻略页）",
          font=F(9, False, MUTED), fill_color="F8FAFC", align=A("left", "center", False, 1), border=None)
    widths = {"A": 13, "B": 18, "C": 12, "D": 12, "E": 12, "F": 12, "G": 12, "H": 26}
    for col, w in widths.items():
        ws.column_dimensions[col].width = w

    merge(ws, "A4:A4", "选择角色", font=F(10, True, INK), fill_color="E8EEF5", align=A("center", "center"))
    merge(ws, "B4:C4", "胡桃", font=F(12, True, "C0392B"), fill_color=INPUT, align=A("center", "center"))
    merge(ws, "D4:D4", "数据来源", font=F(9, False, MUTED), fill_color="F8FAFC", align=A("right", "center"))
    merge(ws, "E4:H4", "biligame WIKI · 配装数据库", font=F(9, False, MUTED), fill_color="F8FAFC",
          align=A("left", "center", False, 1))
    ws.row_dimensions[4].height = 30
    _ci = lambda col: f'INDEX({Q_LIB}!${col}$3:${col}$135,MATCH($B$4,{Q_LIB}!$A$3:$A$135,0))'  # noqa: E731
    merge(ws, "A5:H5",
          f'=IFERROR("元素 "&{_ci("B")}&" ｜ 实装版本 "&{_ci("C")}&" ｜ "&IF({_ci("AO")}="待补","WIKI 暂无攻略，数据待补","数据：WIKI 攻略页"),"未收录该角色，请到「配装数据库」补充")',
          font=F(10, True, "FFFFFF"), fill_color="1F5F7A", align=A("left", "center", False, 1), border=None)
    ws.row_dimensions[5].height = 26

    card_idx = lambda col: f'INDEX({Q_LIB}!${col}$3:${col}$135,MATCH($B$4,{Q_LIB}!$A$3:$A$135,0))'  # noqa: E731

    def sec(row, title, color="6C3F8F", bg="F0EAF7"):
        merge(ws, f"A{row}:H{row}", f"  {title}", font=F(10, True, color), fill_color=bg,
              align=A("left", "center", False, 1), border=None)
        ws.row_dimensions[row].height = 26

    def kv(row, label, formula, num_fmt=None, bold=False):
        merge(ws, f"A{row}:A{row}", label, font=F(9, True, MUTED), fill_color="FFFFFF", align=A("center", "center"))
        merge(ws, f"B{row}:H{row}", formula, font=F(10, bold, INK), fill_color="FFFFFF",
              align=A("left", "center", True, 1))
        if num_fmt:
            ws[f"B{row}"].number_format = num_fmt
        ws.row_dimensions[row].height = 26

    def kv2(row, label, name_formula, desc_formula, bold=False):
        merge(ws, f"A{row}:A{row}", label, font=F(9, True, MUTED), fill_color="FFFFFF", align=A("center", "center"))
        merge(ws, f"B{row}:B{row}", name_formula, font=F(10, bold, INK), fill_color="FFFFFF",
              align=A("left", "center", True, 1))
        merge(ws, f"C{row}:H{row}", desc_formula, font=F(9, False, MUTED), fill_color="FFFFFF",
              align=A("left", "center", True, 1))
        ws.row_dimensions[row].height = 28

    sec(7, "① 推荐圣遗物")
    _or = lambda c: f'=IFERROR(IF({card_idx(c)}="","—",{card_idx(c)}),"")'  # noqa: E731
    kv2(8, "首选", f'=IFERROR(IF({card_idx("D")}="","WIKI 暂无该角色配装（待补）",{card_idx("D")}),"")',
        f'=IFERROR(IF({card_idx("D")}="","可在「配装数据库」手工补充，或等待 WIKI 更新",{card_idx("E")}),"")', bold=True)
    kv2(9, "备选", _or("F"), _or("G"))
    kv2(10, "备选", _or("H"), _or("I"))

    _ti = lambda c: f'IF(IFERROR({card_idx(c)},0)=0,"—",{card_idx(c)})'  # noqa: E731
    sec(12, "② 主词条与有效词条", "1F5F7A", "E3F1F7")
    kv(13, "时之沙", f'=IFERROR({_ti("J")},"—")')
    kv(14, "空之杯", f'=IFERROR({_ti("K")},"—")')
    kv(15, "理之冠", f'=IFERROR({_ti("L")},"—")')
    kv(16, "副词条优先级", f'=IFERROR({_ti("M")},"—")')
    merge(ws, "A17:A17", "有效词条", font=F(9, True, MUTED), align=A("center", "center"))
    merge(ws, "B17:H17",
          f'=IFERROR("核心："&{_ti("N")}&"  ｜  次要："&{_ti("O")}&"  ｜  可选："&{_ti("P")},"—")',
          font=F(9, False, INK), align=A("left", "center", True, 1))
    ws.row_dimensions[17].height = 44

    sec(19, "③ 武器推荐", "117A65", "E6F4EA")
    kv2(20, "毕业", f'=IFERROR({_ti("Q")},"—")', f'=IFERROR(IF({card_idx("Q")}="","—",{card_idx("R")}),"—")', bold=True)
    kv2(21, "可选1", f'=IFERROR({_ti("S")},"—")', f'=IFERROR(IF({card_idx("S")}="","—",{card_idx("V")}),"—")')
    kv2(22, "可选2", f'=IFERROR({_ti("T")},"—")', f'=IFERROR(IF({card_idx("T")}="","—",{card_idx("W")}),"—")')
    kv2(23, "可选3", f'=IFERROR({_ti("U")},"—")', f'=IFERROR(IF({card_idx("U")}="","—",{card_idx("X")}),"—")')

    sec(25, "④ 天赋加点", "9A5B00", "FDF1E0")
    kv(26, "加点顺序", f'=IFERROR({_ti("Y")},"—")', bold=True)
    kv(27, "目标天赋",
       f'=IFERROR(IF({_ti("Y")}="—","—","普攻 "&{_ti("Z")}&" / 战技 "&{_ti("AA")}&" / 爆发 "&{_ti("AB")}),"—")')

    sec(29, "⑤ 毕业面板参考（WIKI 攻略）", "1F4E79", "E9F0F7")
    labels = ["暴击率", "暴击伤害", "攻击力", "生命值", "防御力", "元素精通", "元素充能"]
    for i, lb in enumerate(labels):
        c = ws.cell(row=30, column=2 + i, value=lb)
        c.font = F(9, True, "FFFFFF")
        c.fill = fill("1F4E79")
        c.alignment = A("center", "center")
        c.border = BORDER
        v = ws.cell(row=31, column=2 + i, value=f'=IF({get_column_letter(2+i)}32<>"",{get_column_letter(2+i)}32,IFERROR(IF({card_idx(["AE","AF","AG","AH","AI","AJ","AK"][i])}=0,"—",{card_idx(["AE","AF","AG","AH","AI","AJ","AK"][i])}),"—"))')
        v.font = F(10, True, INK)
        v.alignment = A("center", "center")
        v.border = BORDER
        v.number_format = '0"%"' if i in (0, 1) else "#,##0"
    # 手动覆盖行：留空=用 WIKI；填写后优先使用
    merge(ws, "A32:A32", "手动覆盖", font=F(9, True, "9A5B00"), fill_color="FDF1E0", align=A("center", "center"))
    for i, lb in enumerate(labels):
        mc = ws.cell(row=32, column=2 + i, value=None)
        mc.fill = fill("FFF9E6")
        mc.font = F(10, False, INK)
        mc.alignment = A("center", "center")
        mc.border = BORDER
        mc.number_format = '0"%"' if i in (0, 1) else "#,##0"
        unlock(mc)
    ws.row_dimensions[32].height = 22
    merge(ws, "A33:H33", f'=IFERROR("说明："&SUBSTITUTE(SUBSTITUTE(SUBSTITUTE({card_idx("AL")},"<br>","；"),"<br/>","；"),"<br />","；"),"")', font=F(9, False, MUTED),
          align=A("left", "center", True, 1))
    ws.row_dimensions[33].height = 40

    sec(35, "⑥ 数据来源与更新", "1E6B47", "E8F5EE")
    merge(ws, "A36:B36", "WIKI 攻略页", font=F(9, True, MUTED), align=A("center", "center"))
    merge(ws, "C36:H36", f'=IFERROR(HYPERLINK({card_idx("AM")},"打开该角色的 WIKI 攻略页 ↗"),"")',
          font=Font(name=FONT, size=10, color="2E86C1", underline="single"), align=A("left", "center", False, 1))
    kv(37, "攻略更新时间", f'=IFERROR({card_idx("AN")},"")')
    merge(ws, "A38:C38", "米游社参考", font=F(9, True, MUTED), align=A("center", "center"))
    merge(ws, "D38:H38", "米游社-原神攻略区（一图流）：https://www.miyoushe.com/ys/ 搜索角色名",
          font=F(9, False, MUTED), align=A("left", "center", False, 1))
    merge(ws, "A40:H40", "提示：面板行可手工覆盖（优先于 WIKI）；武器/天赋推荐可在『配装数据库』对应角色行直接改；数据随版本更新。",
          font=F(9, False, MUTED), align=A("left", "center", False, 1), border=None)

    dv = DataValidation(type="list", formula1="=角色列表", allow_blank=False, showDropDown=False)
    ws.add_data_validation(dv)
    dv.add("B4")
    ws.sheet_view.showGridLines = False

    # 默认选胡桃
    if any(r["name"] == "胡桃" for r in rows):
        ws["B4"] = "胡桃"
    unlock(ws["B4"])
    protect(ws)
    return ws


# -*- coding: utf-8 -*-
# ============ 仪表盘 ============
def build_dash(wb):
    ws = wb.create_sheet(DASH)
    ws.sheet_properties.tabColor = "1F6B4A"
    merge(ws, "A1:H1", "养成总览", font=F(18, True, "FFFFFF"), fill_color=NAVY,
          align=A("left", "center", False, 1), border=None)
    ws.row_dimensions[1].height = 52
    merge(ws, "A2:H2", "数据随『养成总表』自动更新 · 无需手动维护 · 当前模板 v1.0",
          font=F(9, False, MUTED), fill_color="F8FAFC", align=A("left", "center", False, 1), border=None)

    def card(col, label, formula, color, r=4):
        c1, c2 = get_column_letter(col), get_column_letter(col + 1)
        merge(ws, f"{c1}{r}:{c2}{r}", label, font=F(9, False, MUTED), fill_color="F8FAFC",
              align=A("center", "center"))
        merge(ws, f"{c1}{r + 1}:{c2}{r + 1}", formula, font=F(18, True, color), fill_color="FFFFFF",
              align=A("center", "center"))
        ws.row_dimensions[r].height = 24
        ws.row_dimensions[r + 1].height = 44

    card(1, "已拥有角色", f'=COUNTIF({Q_MAIN}!$AQ$3:$AQ$135,"已有")', "1F3B57")
    card(3, "已练好", f'=COUNTIF({Q_MAIN}!$AL$3:$AL$135,"✔ 已练好")', "1E6B47")
    card(5, "平均评分", f'=IFERROR(AVERAGEIF({Q_MAIN}!$AQ$3:$AQ$135,"已有",{Q_MAIN}!$AJ$3:$AJ$135),0)', "A66B00", r=4)
    ws["E5"].number_format = "0.0"
    card(7, "圣遗物毕业", f'=COUNTIF({Q_MAIN}!$Q$3:$Q$135,"已毕业")', "6C3F8F")
    card(1, "面板达标", f'=COUNTIF({Q_MAIN}!$AH$3:$AH$135,"达标")', "1F4E79", r=7)
    card(3, "满级角色", f'=COUNTIF({Q_MAIN}!$D$3:$D$135,90)', "1F5F7A", r=7)
    card(5, "天赋毕业", f'=COUNTIF({Q_MAIN}!$W$3:$W$135,">=1")', "9A5B00", r=7)
    card(7, "未拥有", f'=COUNTIF({Q_MAIN}!$AQ$3:$AQ$135,"未拥有")', "8A94A0", r=7)

    # ① 元素分布
    section(ws, 11, "① 元素分布", end_col="D")
    headers = ["元素", "人数", "已练好", "平均评分"]
    for i, h in enumerate(headers, 1):
        c = ws.cell(row=12, column=i, value=h)
        c.font = F(9, True, "FFFFFF")
        c.fill = fill(NAVY)
        c.alignment = A("center", "center")
        c.border = BORDER
    for i, (el, (bg, fg)) in enumerate(ELEMENTS.items()):
        r = 13 + i
        ws.cell(row=r, column=1, value=el).font = F(10, True, fg)
        ws.cell(row=r, column=1).fill = fill(bg)
        ws.cell(row=r, column=2, value=f'=COUNTIFS({Q_MAIN}!$C$3:$C$135,$A{r},{Q_MAIN}!$AQ$3:$AQ$135,"已有")')
        ws.cell(row=r, column=3, value=f'=COUNTIFS({Q_MAIN}!$C$3:$C$135,$A{r},{Q_MAIN}!$AL$3:$AL$135,"✔ 已练好")')
        ws.cell(row=r, column=4, value=f'=IFERROR(AVERAGEIFS({Q_MAIN}!$AJ$3:$AJ$135,{Q_MAIN}!$C$3:$C$135,$A{r},{Q_MAIN}!$AQ$3:$AQ$135,"已有"),"")')
        ws.cell(row=r, column=4).number_format = "0.0"
        for ci in range(1, 5):
            c = ws.cell(row=r, column=ci)
            c.border = BORDER
            c.alignment = A("center", "center")
            if ci > 1:
                c.font = F(10)
                c.fill = fill(SOFT if i % 2 else "FFFFFF")

    # ② 练好判定
    section(ws, 20, "② 练好判定分布", end_col="D", color="1E6B47", bg="E8F5EE")
    for i, h in enumerate(["判定", "人数", "占比"], 1):
        c = ws.cell(row=21, column=i, value=h)
        c.font = F(9, True, "FFFFFF")
        c.fill = fill("1E6B47")
        c.alignment = A("center", "center")
        c.border = BORDER
    ready_list = ["✔ 已练好", "未练好", "待补圣遗物", "缺参考面板"]
    for i, st in enumerate(ready_list):
        r = 22 + i
        ws.cell(row=r, column=1, value=st)
        ws.cell(row=r, column=2, value=f'=COUNTIF({Q_MAIN}!$AL$3:$AL$135,$A{r})')
        ws.cell(row=r, column=3, value=f'=IF(SUM($B$22:$B$25)=0,0,$B{r}/SUM($B$22:$B$25))')
        ws.cell(row=r, column=3).number_format = "0%"
        for ci in range(1, 4):
            c = ws.cell(row=r, column=ci)
            c.border = BORDER
            c.alignment = A("center", "center")
            c.font = F(10, True if ci == 1 else False, READY_COLORS.get(st, ("", INK))[1])
            c.fill = fill(READY_COLORS.get(st, ("FFFFFF", ""))[0] if ci == 1 else (SOFT if i % 2 else "FFFFFF"))
    merge(ws, "A26:D26", f'="已拥有 "&COUNTIF({Q_MAIN}!$AQ$3:$AQ$135,"已有")&" 名"&"（未拥有 "&COUNTIF({Q_MAIN}!$AQ$3:$AQ$135,"未拥有")&" 名不计入）"', font=F(9, True, MUTED),
          fill_color="F8FAFC", align=A("right", "center", False, 0))

    # ③ 圣遗物状态
    section(ws, 28, "③ 圣遗物状态分布", end_col="D", color="6C3F8F", bg="F0EAF7")
    for i, h in enumerate(["状态", "人数", "占比"], 1):
        c = ws.cell(row=29, column=i, value=h)
        c.font = F(9, True, "FFFFFF")
        c.fill = fill("6C3F8F")
        c.alignment = A("center", "center")
        c.border = BORDER
    for i, st in enumerate(["已毕业", "需调整", "未毕业", "未开始"]):
        r = 30 + i
        ws.cell(row=r, column=1, value=st)
        ws.cell(row=r, column=2, value=f'=COUNTIF({Q_MAIN}!$Q$3:$Q$135,$A{r})')
        ws.cell(row=r, column=3, value=f'=IF(SUM($B$30:$B$33)=0,0,$B{r}/SUM($B$30:$B$33))')
        ws.cell(row=r, column=3).number_format = "0%"
        for ci in range(1, 4):
            c = ws.cell(row=r, column=ci)
            c.border = BORDER
            c.alignment = A("center", "center")
            c.font = F(10, True if ci == 1 else False, STATUS_COLORS.get(st, ("", INK))[1])
            c.fill = fill(STATUS_COLORS.get(st, ("FFFFFF", ""))[0] if ci == 1 else (SOFT if i % 2 else "FFFFFF"))
    merge(ws, "A34:D34", f'="合计 "&SUM(B30:B33)&" 名"', font=F(9, True, MUTED),
          fill_color="F8FAFC", align=A("right", "center", False, 0))

    # ④ 面板判定
    section(ws, 36, "④ 面板判定分布", end_col="D", color="1F4E79", bg="E9F0F7")
    for i, h in enumerate(["判定", "人数", "占比"], 1):
        c = ws.cell(row=37, column=i, value=h)
        c.font = F(9, True, "FFFFFF")
        c.fill = fill("1F4E79")
        c.alignment = A("center", "center")
        c.border = BORDER
    for i, st in enumerate(["达标", "接近", "未达标", "无参考"]):
        r = 38 + i
        ws.cell(row=r, column=1, value=st)
        ws.cell(row=r, column=2, value=f'=COUNTIF({Q_MAIN}!$AH$3:$AH$135,$A{r})')
        ws.cell(row=r, column=3, value=f'=IF(SUM($B$38:$B$41)=0,0,$B{r}/SUM($B$38:$B$41))')
        ws.cell(row=r, column=3).number_format = "0%"
        for ci in range(1, 4):
            c = ws.cell(row=r, column=ci)
            c.border = BORDER
            c.alignment = A("center", "center")
            c.font = F(10, True if ci == 1 else False, PANEL_COLORS.get(st, ("", INK))[1])
            c.fill = fill(PANEL_COLORS.get(st, ("FFFFFF", ""))[0] if ci == 1 else (SOFT if i % 2 else "FFFFFF"))
    merge(ws, "A42:D42", f'="合计 "&SUM(B38:B41)&" 名"', font=F(9, True, MUTED),
          fill_color="F8FAFC", align=A("right", "center", False, 0))

    for col, w in {"A": 10, "B": 10, "C": 10, "D": 10, "E": 10, "F": 10, "G": 10, "H": 10,
                   "I": 2, "J": 16, "K": 16, "L": 16}.items():
        ws.column_dimensions[col].width = w

    bar = BarChart()
    bar.type = "col"
    bar.title = "各元素角色数量"
    bar.y_axis.title = "人数"
    bar.y_axis.majorUnit = 1
    bar.height = 8.5
    bar.width = 14
    data = Reference(ws, min_col=2, min_row=12, max_row=19)
    cats = Reference(ws, min_col=1, min_row=13, max_row=19)
    bar.add_data(data, titles_from_data=True)
    bar.set_categories(cats)
    bar.legend = None
    ws.add_chart(bar, "J5")

    pie = PieChart()
    pie.title = "练好判定占比"
    pie.height = 8.5
    pie.width = 14
    data = Reference(ws, min_col=2, min_row=21, max_row=25)
    cats = Reference(ws, min_col=1, min_row=22, max_row=25)
    pie.add_data(data, titles_from_data=True)
    pie.set_categories(cats)
    pie.legend.position = "r"
    for i, color in enumerate(["1E6B47", "E67E22", "2E86C1", "BDC3C7"]):
        pt = DataPoint(idx=i)
        pt.graphicalProperties = GraphicalProperties(solidFill=color)
        pie.series[0].data_points.append(pt)
    pie.dataLabels = DataLabelList()
    pie.dataLabels.showPercent = True
    pie.dataLabels.showVal = False
    pie.dataLabels.showCatName = False
    pie.dataLabels.showSerName = False
    pie.dataLabels.showLegendKey = False
    ws.add_chart(pie, "J23")

    merge(ws, "A45:H45", "提示：在『养成总表』用筛选（Ctrl+Shift+L）按元素/练好判定/圣遗物状态查看子集，本页自动刷新。",
          font=F(9, False, MUTED), fill_color="F8FAFC", align=A("left", "center", False, 1), border=None)
    ws.sheet_view.showGridLines = False
    return ws


# ============ 数据与更新 ============
def build_update(wb, rows, panels):
    ws = wb.create_sheet(UPDATE)
    ws.sheet_properties.tabColor = "7F8C8D"
    merge(ws, "A1:E1", "数据来源与更新", font=F(18, True, "FFFFFF"), fill_color=NAVY,
          align=A("left", "center", False, 1), border=None)
    ws.row_dimensions[1].height = 52
    merge(ws, "A2:E2", "本表数据来自 biligame 原神 WIKI；账号数据可用米游社 Cookie 本地导入，不上传任何服务器",
          font=F(9, False, MUTED), fill_color="F8FAFC", align=A("left", "center", False, 1), border=None)
    for col, w in {"A": 22, "B": 20, "C": 44, "D": 56, "E": 14}.items():
        ws.column_dimensions[col].width = w

    r = 4
    section(ws, r, "数据概况", end_col="D", color="1F3B57", bg="E8EEF5")
    info = [
        ("收录角色", len(rows), "全部实装角色（含旅行者变体）"),
        ("WIKI 攻略覆盖", sum(1 for r0 in rows if r0.get("status") == "WIKI"), "有攻略页的角色数"),
        ("毕业面板覆盖", len(panels), "有可用参考面板的角色数（无面板判「无参考」，练好判定不会误判）"),
        ("数据版本", max((r0["version"] for r0 in rows if r0["version"]), default="7.2"), "最新实装版本"),
        ("数据抓取时间", date.today().isoformat(), "抓取 WIKI 数据的日期"),
        ("当前模板", "v1.0", "工作表版本"),
    ]
    for i, (k, v, note) in enumerate(info):
        row = r + 1 + i
        merge(ws, f"A{row}:A{row}", k, font=F(10, True, INK), fill_color="F8FAFC", align=A("right", "center"))
        ws[f"B{row}"] = v
        ws[f"B{row}"].font = F(11, True, "1F3B57")
        ws[f"B{row}"].alignment = A("center", "center")
        merge(ws, f"C{row}:D{row}", note, font=F(9, False, MUTED), align=A("left", "center", False, 1))
        for ci in range(1, 5):
            ws.cell(row=row, column=ci).border = BORDER
    r += len(info) + 2

    section(ws, r, "数据来源", end_col="D", color="1F3B57", bg="E8EEF5")
    sources = [
        ("原神 WIKI 首页", "总入口", "https://wiki.biligame.com/ys/首页"),
        ("角色一览", "角色名单 / 基础信息", "https://wiki.biligame.com/ys/角色"),
        ("角色攻略筛选", "全部角色攻略入口", "https://wiki.biligame.com/ys/角色攻略筛选"),
        ("伤害期望计算器", "期望伤害参考", "https://wiki.biligame.com/ys/伤害期望计算器"),
        ("米游社·原神", "攻略区（角色培养一图流）", "https://www.miyoushe.com/ys/"),
        ("米游社·观测枢", "角色/圣遗物图鉴（可补充）", "https://baike.mihoyo.com/ys/obc/"),
        ("玉衡杯数据库", "角色/武器基础数值", "https://ambr.top/"),
        ("KQM 攻略", "英文毕业面板参考", "https://keqingmains.com/"),
    ]
    for i, (name, note, url) in enumerate(sources):
        row = r + 1 + i
        ws[f"A{row}"] = name
        ws[f"A{row}"].font = F(10, True, INK)
        ws[f"A{row}"].alignment = A("left", "center")
        merge(ws, f"B{row}:C{row}", note, font=F(9, False, MUTED), align=A("left", "center"))
        ws[f"D{row}"] = url
        ws[f"D{row}"].font = F(9, False, "2E86C1")
        ws[f"D{row}"].hyperlink = url
        for ci in range(1, 5):
            ws.cell(row=row, column=ci).border = BORDER
    r += len(sources) + 2

    section(ws, r, "如何随版本更新", end_col="D", color="1F3B57", bg="E8EEF5")
    steps = [
        "1. 双击文件夹里的「更新数据.bat」（需本机可用 Python），自动重新抓取角色名单与攻略配装，并重建本表。",
        "2. 也可以直接把 WIKI 链接发给 AI 助手，让它按最新版本重跑一次更新。",
        "3. 只新增少量角色：到「配装数据库」底部黄色空白行填写角色与配装，总表下拉会自动出现。",
        "4. 账号装备有变化：先关掉 Excel，双击『更新我的数据.bat』重新拉取并覆盖主文件（自动备份，保留手填的期望/备注/标记）。",
        "5. 游戏大版本更新后，旧角色推荐可能变化；运行『更新数据.bat』重建模板（会覆盖示例行与手填内容，请先备份）。",
    ]
    for i, t in enumerate(steps):
        row = r + 1 + i
        merge(ws, f"A{row}:D{row}", t, font=F(10), align=A("left", "center", True, 1), border=BORDER)
        ws.row_dimensions[row].height = 30
    r += len(steps) + 2

    section(ws, r, "导入米游社账号数据（可选）", end_col="D", color="1F6B3A", bg="E8F5EE")
    imp = [
        "1. 浏览器登录米游社（https://www.miyoushe.com/ys/），按 F12 → Application → Cookies，复制 ltuid/ltoken 所在的一整段 Cookie。",
        "2. 把 Cookie 粘贴到本文件夹的「米游社cookie.txt」，保存；再运行「导入米游社数据.bat」。",
        "3. 脚本会读取账号角色，把等级/武器/天赋/圣遗物填进本表，生成带「我的数据」后缀的新文件；原模板保留。",
        "4. Cookie 只保存在本机文件里，脚本只访问米游社官方接口，不上传任何地方；不用时删除该文件即可。",
    ]
    for i, t in enumerate(imp):
        row = r + 1 + i
        merge(ws, f"A{row}:D{row}", t, font=F(10), align=A("left", "center", True, 1), border=BORDER)
        ws.row_dimensions[row].height = 30
    r += len(imp) + 2

    section(ws, r, "更新日志", end_col="D", color="1F3B57", bg="E8EEF5")
    logs = [
        ("v1.0", "2026-11-03", "开源发布版，默认空白模板；Windows 启动、文档与离线重建完善。公式锁定保护（只可填浅黄格）；命座/强化等级下拉；主交付为空白表，可一键重新拉取并覆盖更新（保留手填期望/备注/标记）；今日练谁备注显示面板未达标项与期望评分；参考卡⑤可手动覆盖毕业面板；武器/天赋等推荐可在配装数据库直接改。"),
        ("v0.5", "2026-11-02", "全角色预置（未拥有只显示名字）；新增命座列；今日练谁改为自选 5 人清单，标记已练好自动移除，并显示毕业面板/推荐方案/米游社攻略链接；导入账号数据时自动写入命座与拥有状态。"),
        ("v0.4.2", "2026-11-02", "修复：未填圣遗物时总表 L 列改为显示推荐套装；WIKI 缺面板时不再误判「已练好」，改显示「缺参考面板」；参考卡无配装时显示待补提示；新增米游社账号数据本地导入脚本与更多数据来源。"),
        ("v0.4.1", "2026-11-02", "修复圣遗物套装下拉出现拼接项的问题（拆分「A、B」散搭套装名）。"),
        ("v0.4", "2026-11-02", "删除武器类型/定位/星级/优先级；新增『圣遗物明细』逐件词条评分与调整建议；新增面板数值对比 WIKI 毕业面板，自动判定是否练好。"),
        ("v0.3", "2026-11-02", "新增『今日练谁』与练度标记，自动回答今天练谁、谁还没练。"),
        ("v0.2", "2026-11-02", "数据改为 WIKI 攻略页结构化（133 名角色），新增『角色参考卡』，总表精简。"),
    ]
    for i, (v, d, t) in enumerate(logs):
        row = r + 1 + i
        ws[f"A{row}"] = v
        ws[f"A{row}"].font = F(10, True, "9A5B00")
        ws[f"A{row}"].alignment = A("center", "center")
        ws[f"B{row}"] = d
        ws[f"B{row}"].font = F(9, False, MUTED)
        ws[f"B{row}"].alignment = A("center", "center")
        merge(ws, f"C{row}:D{row}", t, font=F(9), align=A("left", "center", True, 1))
        ws.row_dimensions[row].height = 40
        for ci in range(1, 5):
            ws.cell(row=row, column=ci).border = BORDER
    r += len(logs) + 2

    merge(ws, f"A{r}:E{r}",
          "说明：WIKI 攻略内容带一定主观性，且随版本调整；『配装数据库』中的套装/主词条/期望值/面板参考可自行修改，改完后总表与参考卡自动生效。"
          "米游社一图流为图片，无法被公式直接解析；表格使用同源结构化数据，并新增米游社账号导入做补充。",
          font=F(9, False, MUTED), fill_color="FEF9E7", align=A("left", "center", True, 1), border=None)
    ws.row_dimensions[r].height = 48
    ws.sheet_view.showGridLines = False
    return ws


# ============ 使用说明 ============
def build_help(wb):
    ws = wb.create_sheet(HELP)
    ws.sheet_properties.tabColor = "C9A227"
    merge(ws, "A1:H1", "使用说明 · 原神角色养成管理表", font=F(18, True, "FFFFFF"), fill_color=NAVY,
          align=A("left", "center", False, 1), border=None)
    ws.row_dimensions[1].height = 52
    merge(ws, "A2:H2", "v1.0（2026-11-02）｜ 数据：biligame 原神 WIKI（收录至 7.2）｜ 圣遗物逐件评分 + 面板毕业判定",
          font=F(9, False, MUTED), fill_color="F8FAFC", align=A("left", "center", False, 1), border=None)
    ws.column_dimensions["A"].width = 7
    for col in "BCDEFGH":
        ws.column_dimensions[col].width = 16.5
    r = 4
    section(ws, r, "① 五步上手", end_col="H"); r += 1
    lines = [
        "1. 在『养成总表』找到角色（133 个角色已全部预置）：浅黄格可填；白格是公式已锁定，双击也改不了，防止误触。未拥有的行只显示「未拥有」。",
        "2. 填当前/目标等级、武器与精炼、武器等级、天赋等级、期望值（黄色格子）。",
        "3. 到『圣遗物明细』给这个角色填 5 件圣遗物：套装、主词条、强化等级、4 条副词条及数值。",
        "4. 回到『养成总表』填当前面板：暴击率、暴击伤害、攻击力、生命值、防御力、元素精通、元素充能。",
        "5. 到『今日练谁』从下拉选最多 5 个想练的角色，自动显示毕业面板/推荐方案/米游社攻略链接；面板未达标时备注会列出「哪项差多少、期望多少」；标记「已练好」后该格自动移除。",
    ]
    for t in lines:
        merge(ws, f"A{r}:H{r}", t, font=F(10), align=A("left", "center", True, 1), border=None)
        ws.row_dimensions[r].height = 30
        r += 1
    r += 1
    section(ws, r, "② 颜色与符号", end_col="H"); r += 1
    legend = [
        ("黄底", "需要你填写", "FFF9E6"),
        ("白底", "自动计算或自动带出", "FFFFFF"),
        ("✔", "达标 / 推荐匹配 / 已练好", "C6F6D5"),
        ("△", "接近 / 可用 / 部分匹配", "FDF3D0"),
        ("✘", "未匹配 / 非推荐 / 未达标", "FADBD8"),
    ]
    for key, desc, color in legend:
        ws.cell(row=r, column=1, value=key).fill = fill(color)
        ws.cell(row=r, column=1).font = F(10, True, INK)
        ws.cell(row=r, column=1).alignment = A("center", "center")
        ws.cell(row=r, column=1).border = BORDER
        merge(ws, f"B{r}:H{r}", desc, font=F(10), align=A("left", "center", False, 1))
        for ci in range(1, 9):
            ws.cell(row=r, column=ci).border = BORDER
        ws.row_dimensions[r].height = 28
        r += 1
    r += 1
    section(ws, r, "③ 评分与判定规则", end_col="H"); r += 1
    rules = [
        "圣遗物单品评分（0-100）= 强化 25% + 主词条 20% + 有效副词条 55%；有效词条按 WIKI 副词条优先级分核心/次要/可选三档。",
        "圣遗物总评分 = 五件单品评分平均（缺一件会自动降低总分）；≥75 已毕业、55-75 需调整、<55 未毕业。",
        "调整建议：自动挑评分最低的一件，在总表 R 列提示先调整哪个部位。",
        "综合评分（0-100）= 等级 15% + 武器匹配 10% + 武器等级 10% + 圣遗物 35% + 天赋 15% + 面板 15%。",
        "练好判定：武器 90 且匹配、圣遗物五件齐全且已毕业、天赋达标、面板达标 → ✔ 已练好；若该角色没有 WIKI 参考面板，判定为「缺参考面板」，不会误显示已练好。",
        "面板判定：与 WIKI 毕业面板逐项比较，全部达标=达标，≥70% 项达标=接近，否则未达标；无参考面板=无参考。",
        "面板缺口：在『今日练谁』备注列自动列出未达标项目（当前值/期望值）与期望评分；想更宽松可手工把角色的『特殊标记』设为「已练好」。",
    ]
    for t in rules:
        merge(ws, f"A{r}:H{r}", t, font=F(10), align=A("left", "center", True, 1), border=None)
        ws.row_dimensions[r].height = 30
        r += 1
    r += 1
    section(ws, r, "④ 数据来源与更新", end_col="H"); r += 1
    ups = [
        "角色基础信息、推荐配装、有效词条、毕业面板均来自 biligame 原神 WIKI 攻略页；版本更新后运行『更新数据.bat』自动重抓。",
        "『配装数据库』里的套装/主词条/期望值/面板参考可自己改，改完全表与参考卡自动生效。",
        "米游社一图流是图片，公式无法直接读取；表格采用同源结构化数据，可对照参考。",
        "账号数据可用米游社 Cookie 本地导入：把 Cookie 粘到『米游社cookie.txt』，运行『导入米游社数据.bat』，自动填写等级/武器/天赋/圣遗物（仅本机，不上传）。",
        "后续更新：运行『更新我的数据.bat』会重新拉取账号数据并覆盖主文件里的游戏数据；手填的期望值、备注、特殊标记会保留（原文件自动备份到「旧版本」）。",
        "面板参考覆盖约 70 名角色；无参考面板时显示「缺参考面板」，请先补充参考，不会自动判为已练好。",
    ]
    for t in ups:
        merge(ws, f"A{r}:H{r}", t, font=F(10), align=A("left", "center", True, 1), border=None)
        ws.row_dimensions[r].height = 30
        r += 1
    r += 1
    section(ws, r, "⑤ 常见问题", end_col="H"); r += 1
    faqs = [
        ("角色显示「未收录」？", "到『配装数据库』补该角色，或在空白行填写。"),
        ("圣遗物评分很低怎么办？", "看总表 R 列调整建议；先补满 5 件，再把有效副词条数值填准确。"),
        ("面板怎么填？", "填当前实战面板；暴击率/暴击伤害填数字（如 85、180），其他填数值。"),
        ("为什么练好判定还是「未练好」？", "逐项检查：武器 90 且推荐、圣遗物已毕业、天赋达标、面板达标。"),
        ("某角色怎么没出现在『今日练谁』？", "今日练谁由你自己选（最多 5 位）；已练好/暂缓/不练了会自动移除。"),
        ("为什么有些格子不能改？", "白底是公式（自动计算），已锁定；只能填浅黄格。确实要改公式，先「审阅→撤销工作表保护」即可。"),
        ("推荐武器/天赋想自己改？", "到『配装数据库』对应角色行改；总表与参考卡会自动带出。参考卡⑤面板也可手填覆盖。"),
    ]
    for q, a in faqs:
        merge(ws, f"A{r}:H{r}", f"Q：{q}  A：{a}", font=F(9), align=A("left", "center", True, 1), border=None)
        ws.row_dimensions[r].height = 26
        r += 1
    r += 1
    section(ws, r, "⑥ 版本记录", end_col="H"); r += 1
    vers = [
        "v1.0（2026-11-03）：开源发布版。全角色预置、公式锁定、全下拉、今日练谁自选 5 人、面板缺口提示、参考卡手动覆盖、米游社账号导入与一键更新。",
        "v1.0（2026-11-02）：公式锁定与下拉完善；今日练谁备注显示面板缺口；参考卡可手动覆盖面板；支持一键重新拉取更新。v0.5：全角色预置+自选 5 人清单。v0.4.1：修复套装下拉拼接。",
        "v0.3（2026-11-02）：新增『今日练谁』与练度标记。",
        "v0.2（2026-11-02）：接入 WIKI 数据（133 角色），新增角色参考卡。",
        "v0.1（2026-11-02）：初版。",
    ]
    for t in vers:
        merge(ws, f"A{r}:H{r}", t, font=F(9, False, MUTED), align=A("left", "center", True, 1), border=None)
        ws.row_dimensions[r].height = 26
        r += 1
    merge(ws, f"A{r + 1}:H{r + 1}", "需要加配装对比/命座/专武标记/材料清单或队伍筛选？告诉我，可以直接加。",
          font=F(9, True, "9A5B00"), fill_color="FEF9E7", align=A("left", "center", False, 1), border=None)
    ws.sheet_view.showGridLines = False
    return ws


# ============ 主流程 ============
def main():
    rows, roster, builds, panels = load_all()
    wb = Workbook()
    wb.remove(wb.active)
    names = build_data(wb, rows)
    build_library(wb, rows)
    build_detail(wb, rows)
    build_main(wb, rows)
    build_today(wb)
    build_card(wb, rows)
    build_dash(wb)
    build_update(wb, rows, panels)
    build_help(wb)

    order = [MAIN, TODAY, DETAIL, CARD, LIB, DASH, SRC_SHEET, UPDATE, HELP]
    wb._sheets = [wb[name] for name in order]
    wb.active = 0

    for name, (col, r1, r2) in names.items():
        wb.defined_names[name] = DefinedName(name, attr_text=f"'{SRC_SHEET}'!${col}${r1}:${col}${r2}")
    wb.defined_names["角色列表"] = DefinedName(
        "角色列表", attr_text=f"OFFSET('{LIB}'!$A${LIB_FIRST},0,0,COUNTA('{LIB}'!$A${LIB_FIRST}:$A$1000),1)")

    wb.calculation.fullCalcOnLoad = True
    for ws in wb.worksheets:
        if not ws.protection.sheet:
            protect(ws)
    wb.save(OUT)
    print("saved:", OUT, "size:", os.path.getsize(OUT))


if __name__ == "__main__":
    main()
