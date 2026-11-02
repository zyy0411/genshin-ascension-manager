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
        "5. 到『今日练谁』从下拉选最多 5 个想练的角色，自动显示毕业面板/推荐方案/米游社攻略链接；面板未达标时备注会列出「哪项差多少、期望多少」；标记「已练好」后隐藏详细规划，保留角色选择以便更换。",
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
        ("某角色怎么没出现在『今日练谁』？", "今日练谁由你自己选（最多 5 位）；已练好会隐藏详细规划，选择单元格保留角色名。"),
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
