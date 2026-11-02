# -*- coding: utf-8 -*-
"""从缓存的 WIKI「角色攻略」wikitext 中提取毕业参考面板数值。

输出 _source/wiki_panels.json:
  { name: {cr, cd, atk, hp, def, em, er, note, source_url} }
数值单位为百分比（暴击率/暴击伤害/充能）或整数（攻击/生命/防御/精通）。
"""
import json
import glob
import re
import os

HERE = os.path.dirname(os.path.abspath(__file__))
TMP = os.path.join(os.path.dirname(HERE), "_wiki_tmp")
OUT = os.path.join(HERE, "wiki_panels.json")

STAT_KEYS = ["cr", "cd", "atk", "hp", "def", "em", "er"]
LABEL = {"cr": "暴击率", "cd": "暴击伤害", "atk": "攻击力", "hp": "生命值",
         "def": "防御力", "em": "元素精通", "er": "元素充能效率"}


def load_pages():
    pages = {}
    for f in sorted(glob.glob(os.path.join(TMP, "pages", "batch_*.json"))):
        data = json.load(open(f, encoding="utf-8"))
        for title, p in data.get("results", {}).items():
            if p.get("missing"):
                continue
            pages[p.get("title") or title] = p.get("content", "")
    return pages


def candidate_cells(text, name):
    """返回角色本人在「参考面板」表格行中、含数值阈值的单元格文本。"""
    cells = []
    pats = [
        r"\n\|\s*\{\{图标/攻略\|角色\|" + re.escape(name) + r"\}\}",
        r"\n\|\s*\{\{图标/攻略\|" + re.escape(name) + r"\}\}",
    ]
    for pat in pats:
        for m in re.finditer(pat, text):
            seg = text[m.end():]
            lines = seg.split("\n")
            content = []
            for line in lines[1:]:
                if content and line.startswith("|"):
                    break
                content.append(line.lstrip().lstrip("|").strip())
                if len(content) > 8:
                    break
            cell = " ".join(x for x in content if x)
            if re.search(r"\d", cell) and ("以上" in cell or "面板" in cell):
                cells.append((m.start(), cell))
    cells.sort()
    return [c for _, c in cells]


def num(s):
    try:
        return int(round(float(s)))
    except Exception:
        return None


def parse_cell(cell):
    got = {}
    txt = cell.replace("，", ",").replace("％", "%")

    # 暴击率/暴击伤害：85/180%  或 75%/150%  或 55~64%/200%
    m = re.search(r"暴击率\s*/\s*暴击伤害[^\d\n]{0,6}(\d+)\s*%?\s*(?:[~～\-]\s*\d+\s*%?)?\s*[/／]\s*(\d+)\s*%?", txt)
    if m:
        got["cr"] = num(m.group(1))
        got["cd"] = num(m.group(2))
    else:
        m = re.search(r"暴击率\s*[：:]\s*(\d+)\s*%?", txt)
        if m:
            got["cr"] = num(m.group(1))
        m = re.search(r"暴击伤害\s*[：:]\s*(\d+)\s*%?", txt)
        if m:
            got["cd"] = num(m.group(1))
    # 合理性过滤（WIKI 偶有笔误）
    if got.get("cr") is not None and not (0 < got["cr"] <= 100):
        got.pop("cr")
    if got.get("cd") is not None and not (0 < got["cd"] <= 300):
        got.pop("cd")

    for key, label in (("em", "元素精通"), ("er", "元素充能效率"),
                       ("atk", "攻击力"), ("hp", "生命值"), ("def", "防御力")):
        if key in got:
            continue
        # 允许「实战」「约」等前缀，允许 130%~140% 取第一个数
        m = re.search(label + r"[^\d\n]{0,12}?(\d+)", txt)
        if m:
            val = num(m.group(1))
            if val is not None:
                got[key] = val
    return got


def main():
    pages = load_pages()
    result = {}
    for title, text in pages.items():
        if not title.endswith("/攻略"):
            continue
        name = title.split("/")[0]
        best = None
        for cell in candidate_cells(text, name):
            got = parse_cell(cell)
            if got and (best is None or len(got) > len(best[0])):
                best = (got, cell.strip())
        if not best:
            continue
        got, note = best
        note = re.sub(r"\s*\n\s*", " ", note)
        note = re.sub(r"\{\{.*?\}\}", "", note)
        note = note.strip(" |")
        result[name] = {
            "cr": got.get("cr"), "cd": got.get("cd"),
            "atk": got.get("atk"), "hp": got.get("hp"), "def": got.get("def"),
            "em": got.get("em"), "er": got.get("er"),
            "note": note[:160],
            "source_url": "https://wiki.biligame.com/ys/" + name + "/攻略",
        }

    with open(OUT, "w", encoding="utf-8") as f:
        json.dump(result, f, ensure_ascii=False, indent=1)
    total = len([t for t in pages if t.endswith("/攻略")])
    print("pages:", total, "panel ok:", len(result))
    from collections import Counter
    cnt = Counter()
    for v in result.values():
        for k in STAT_KEYS:
            if v.get(k):
                cnt[k] += 1
    print("stat coverage:", dict(cnt))
    for n in list(result)[:8]:
        print(n, result[n])


if __name__ == "__main__":
    main()
