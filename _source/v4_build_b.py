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
