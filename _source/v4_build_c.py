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
