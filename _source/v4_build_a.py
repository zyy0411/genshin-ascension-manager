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
