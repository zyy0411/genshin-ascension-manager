# -*- coding: utf-8 -*-
"""解析 biligame 原神角色「名/攻略」页，抽取默认流派配装 → _wiki_tmp/wiki_builds.json。

用法：
    python parse_builds.py            # 用 _wiki_tmp/pages/batch_*.json 缓存解析
    python parse_builds.py --fetch    # 先重新抓取（2~3s 间隔）再解析
"""
import glob
import json
import os
import re
import sys
from datetime import datetime, timedelta, timezone

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(BASE, "_source"))
TMP = os.path.join(BASE, "_wiki_tmp")
API = "https://wiki.biligame.com/ys/api.php"
SITE = "https://wiki.biligame.com/ys/"
TZ = timezone(timedelta(hours=8))

# ---------------------------------------------------------------- 基础 wikitext 工具


def strip_comments(text):
    return re.sub(r"<!--.*?-->", "", text, flags=re.S)


def template_end(text, start):
    """text[start:start+2]=='{{'，返回匹配 '}}' 之后的下标。"""
    i, depth, n = start + 2, 1, len(text)
    while i < n - 1:
        if text.startswith("{{", i):
            depth += 1
            i += 2
        elif text.startswith("}}", i):
            depth -= 1
            i += 2
            if depth == 0:
                return i
        else:
            i += 1
    return n


def split_top(s, sep="|"):
    """按顶层分隔符切分，忽略 {{...}} 与 [[...]] 内部的分隔符。"""
    parts, buf, curly, square = [], [], 0, 0
    i, n = 0, len(s)
    while i < n:
        if s.startswith("{{", i):
            curly += 1
            buf.append("{{")
            i += 2
            continue
        if s.startswith("}}", i) and curly:
            curly -= 1
            buf.append("}}")
            i += 2
            continue
        if s.startswith("[[", i):
            square += 1
            buf.append("[[")
            i += 2
            continue
        if s.startswith("]]", i) and square:
            square -= 1
            buf.append("]]")
            i += 2
            continue
        ch = s[i]
        if ch == sep and curly == 0 and square == 0:
            parts.append("".join(buf))
            buf = []
        else:
            buf.append(ch)
        i += 1
    parts.append("".join(buf))
    return parts


def find_template(text, name):
    """找名字完全匹配的第一个顶层/嵌套模板，返回 (start, end, body) 或 None。"""
    i = 0
    while True:
        j = text.find("{{" + name, i)
        if j < 0:
            return None
        e = template_end(text, j)
        body = text[j + 2:e - 2]
        if split_top(body, "|")[0].strip() == name:
            return j, e, body
        i = j + 2


# ---------------------------------------------------------------- 模板清洗

ICON_TYPE_NAMES = {"武器", "圣遗物", "角色", "材料", "食物", "道具", "任务道具", "怪物"}


def render_icon(params):
    """{{图标|武器|金流监督}} / {{图标|圣遗物|逐影猎人|||4件套}} / {{图标/攻略|班尼特}}"""
    params = [p.strip() for p in params]
    if len(params) >= 2 and params[0] in ICON_TYPE_NAMES:
        return clean_text(params[1])
    if params:
        return clean_text(params[0])
    return ""


def clean_text(raw):
    """去 wiki/HTML 标记，只留可读文本。"""
    if not raw:
        return ""
    s = strip_comments(raw)

    out, i = [], 0
    while True:
        j = s.find("{{", i)
        if j < 0:
            out.append(s[i:])
            break
        e = template_end(s, j)
        inner = s[j + 2:e - 2]
        pieces = split_top(inner, "|")
        name = pieces[0].strip()
        params = pieces[1:]
        out.append(s[i:j])
        if name.startswith("#"):
            pass  # {{#info:...}} 之类注释性解析函数，丢弃
        elif name in ("图标", "图标/攻略"):
            out.append(render_icon(params))
        elif name in ("颜色", "效果", "黑幕", "悬浮", "注音", "强调", "tip", "color"):
            nonempty = [p for p in params if p.strip()]
            out.append(clean_text(nonempty[-1]) if nonempty else "")
        elif name == "ruby":
            out.append(clean_text(params[0]) if params else "")
        else:
            # 未知模板：取最后一个非空参数的可读文本
            nonempty = [p for p in params if p.strip()]
            out.append(clean_text(nonempty[-1]) if nonempty else "")
        i = e
    s = "".join(out)

    s = re.sub(r"<br\s*/?>", " ", s, flags=re.I)
    s = re.sub(r"<[^>]+>", "", s)
    s = s.replace("'''", "").replace("''", "")
    s = re.sub(r"&nbsp;", " ", s)
    s = re.sub(r"&amp;", "&", s)
    s = re.sub(r"&lt;", "<", s)
    s = re.sub(r"&gt;", ">", s)
    # [[页面|显示]] / [[页面]]
    s = re.sub(r"\[\[([^\[\]|]*)\|([^\[\]]*)\]\]", r"\2", s)
    s = re.sub(r"\[\[([^\[\]]*)\]\]", r"\1", s)
    s = re.sub(r"\[https?://[^\]]+\s+([^\]]+)\]", r"\1", s)
    # 空白归一
    s = s.replace("\u3000", " ")
    s = re.sub(r"[ \t\r\n]+", " ", s).strip()
    # 去重：图标模板渲染文本偶发重复两次
    m = re.fullmatch(r"(.{2,24}?)\1", s)
    if m and "、" not in m.group(1):
        s = m.group(1)
    return s


def parse_icon_templates(value):
    """从字段值中取全部 图标 模板 → [(名称, 套装件数), ...]。"""
    out = []
    i = 0
    while True:
        j = value.find("{{", i)
        if j < 0:
            break
        e = template_end(value, j)
        body = value[j + 2:e - 2]
        head = split_top(body, "|")[0].strip()
        if head == "图标":
            params = [p.strip() for p in split_top(body, "|")[1:]]
            if len(params) >= 2 and params[0] in ICON_TYPE_NAMES:
                name = clean_text(params[1])
                pieces = clean_text(params[4]) if len(params) > 4 else ""
            else:
                name = clean_text(params[0]) if params else ""
                pieces = clean_text(params[4]) if len(params) > 4 else ""
            if name:
                out.append((name, pieces))
        i = e
    if not any(p for _, p in out):
        m = re.search(r"[248]件套", clean_text(value))
        if m:
            out = [(n, m.group(0)) for n, _ in out]
    return out


def parse_icon_template(value):
    """从字段值中取首个 图标 模板 → (名称, 套装件数)。"""
    tpls = parse_icon_templates(value)
    return tpls[0] if tpls else ("", "")


def icon_slot_text(value):
    """图标字段 → 显示名与件数；同槽多个套装（2+2 散搭/备选）用「、」连接名称。"""
    tpls = parse_icon_templates(value)
    names, seen = [], set()
    for name, _ in tpls:
        if name and name not in seen:
            seen.add(name)
            names.append(name)
    pieces = next((p for _, p in tpls if p), "")
    return "、".join(names), pieces


# ---------------------------------------------------------------- 页面解析

def parse_simple_config(page_text):
    """返回 (variant_name, {参数名: 原文}) / None。默认流派 = 第一个「默认显示」后的首个简易配装。"""
    text = strip_comments(page_text)
    has_display = "{{角色/配装推荐|默认显示" in text
    variant = ""
    search_from = 0
    if has_display:
        m = re.search(r"\{\{角色/配装推荐\|默认显示\s*(?:\|([^}|]*))?\}\}", text)
        if m:
            variant = (m.group(1) or "").strip()
            search_from = m.end()
    seg = text[search_from:]
    t = find_template(seg, "简易配装")
    if not t:
        # 兜底：全页第一个简易配装
        seg = text
        search_from = 0
        t = find_template(seg, "简易配装")
        if not t:
            return None
    body = seg[t[0] + 2: t[1] - 2]
    parts = split_top(body, "|")
    params = {}
    for p in parts[1:]:
        if "=" in p:
            k, _, v = p.partition("=")
            k = k.strip()
            if k:
                params[k] = v.strip()
    return variant, params


TALENTS = [("普通攻击", "普攻"), ("元素战技", "战技"), ("元素爆发", "爆发")]


def parse_talent_order(text):
    found = []
    for full, short in TALENTS:
        idx = text.find(full)
        if idx >= 0:
            found.append((idx, short))
    return [s for _, s in sorted(found)]


def parse_character(name, rec, roster):
    out = {"name": name}
    if not rec or rec.get("missing"):
        out["status"] = "no_page"
        out["error"] = rec.get("error") if rec and rec.get("error") else "WIKI 无该角色「名/攻略」页"
        return out
    try:
        parsed = parse_simple_config(rec["content"])
        if parsed is None:
            out["status"] = "no_template"
            out["error"] = "未找到「简易配装」模板"
            return out
        variant, params = parsed
        if not any(v.strip() for v in params.values()):
            out["status"] = "no_template"
            out["error"] = "简易配装模板未填写（空占位页）"
            return out

        def txt(key):
            return clean_text(params.get(key, ""))

        artifacts = []
        for i in (1, 2, 3):
            raw = params.get("圣遗物%d" % i, "")
            if not raw.strip():
                continue
            aname, pieces = icon_slot_text(raw)
            if not aname:
                aname = txt("圣遗物%d" % i)
            artifacts.append({"name": aname, "pieces": pieces,
                              "desc": txt("圣遗物描述%d" % i)})

        optional_weapons, optional_descs = [], []
        for i in (1, 2, 3):
            raw = params.get("可选武器%d" % i, "")
            if not raw.strip():
                continue
            wname, _ = icon_slot_text(raw)
            optional_weapons.append(wname or txt("可选武器%d" % i))
            optional_descs.append(txt("可选武器描述%d" % i))

        best_weapon, _ = icon_slot_text(params.get("毕业武器", ""))
        if not best_weapon:
            best_weapon = txt("毕业武器")

        talent_text = txt("技能升级推荐")
        title = rec.get("title") or (name + "/攻略")
        out.update({
            "status": "ok",
            "variant": variant,
            "source_url": SITE + title.replace(" ", "_"),
            "rev_timestamp": rec.get("timestamp", ""),
            "artifacts": artifacts,
            "sand": txt("沙"),
            "goblet": txt("杯"),
            "circlet": txt("头"),
            "substats": txt("副词条"),
            "best_weapon": best_weapon,
            "best_weapon_desc": txt("毕业武器描述"),
            "optional_weapons": optional_weapons,
            "optional_weapon_descs": optional_descs,
            "talent_text": talent_text,
            "talent_order": parse_talent_order(talent_text),
            "talent_reason": txt("技能升级推荐理由"),
        })
        return out
    except Exception as e:  # noqa: BLE001
        out["status"] = "error"
        out["error"] = "%s: %s" % (type(e).__name__, e)
        return out


# ---------------------------------------------------------------- 主流程

def load_roster():
    with open(os.path.join(TMP, "roster.json"), encoding="utf-8") as f:
        return json.load(f)


def load_pages():
    pages = {}
    for path in sorted(glob.glob(os.path.join(TMP, "pages", "batch_*.json"))):
        with open(path, encoding="utf-8") as f:
            d = json.load(f)
        for req, rec in d["results"].items():
            pages[req[:-3] if req.endswith("/攻略") else req] = rec
    return pages


def main():
    roster = load_roster()
    if "--fetch" in sys.argv:
        import fetch_wiki
        fetch_wiki.fetch_all_guides()
    pages = load_pages()

    characters = {}
    stats = {"ok": 0, "no_page": 0, "no_template": 0, "error": 0}
    for name in roster:
        rec = pages.get(name)
        item = parse_character(name, rec, roster)
        stats[item["status"]] = stats.get(item["status"], 0) + 1
        characters[name] = item

    result = {
        "generated_at": datetime.now(TZ).isoformat(timespec="seconds"),
        "source_api": API,
        "total_roster": len(roster),
        "stats": stats,
        "characters": characters,
    }
    out_path = os.path.join(TMP, "wiki_builds.json")
    with open(out_path, "w", encoding="utf-8") as f:
        json.dump(result, f, ensure_ascii=False, indent=1)

    print("saved:", out_path)
    print("stats:", stats, "coverage: %.1f%%" % (100.0 * stats["ok"] / len(roster)))
    bad = [(n, c.get("status"), c.get("error", "")) for n, c in characters.items()
           if c["status"] != "ok"]
    print("not ok (%d):" % len(bad))
    for row in bad[:30]:
        print("  ", row)


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    main()
