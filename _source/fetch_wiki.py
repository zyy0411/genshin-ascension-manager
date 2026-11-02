# -*- coding: utf-8 -*-
"""从 biligame 原神 WIKI 拉取角色基础数据（SMW ask API）。"""
import json
import os
import random
import time
import urllib.parse
import urllib.request

API = "https://wiki.biligame.com/ys/api.php"
UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
      "(KHTML, like Gecko) Chrome/120.0 Safari/537.36")
BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
CACHE = os.path.join(BASE, "_wiki_tmp")
os.makedirs(CACHE, exist_ok=True)


def get(params, tries=4):
    url = API + "?" + urllib.parse.urlencode(params)
    last = None
    for i in range(tries):
        try:
            req = urllib.request.Request(url, headers={"User-Agent": UA, "Referer": "https://wiki.biligame.com/ys/"})
            with urllib.request.urlopen(req, timeout=60) as r:
                return json.loads(r.read().decode("utf-8"))
        except Exception as e:  # noqa: BLE001
            last = e
            time.sleep(8 * (i + 1))
    raise RuntimeError(f"fetch failed: {url} :: {last}")


def ask_all():
    props = ["稀有度", "武器类型", "元素属性", "实装版本", "实装日期", "所属", "TAG"]
    chars, offset = {}, 0
    while True:
        q = "[[分类:角色]]|" + "|".join("?" + p for p in props) + f"|limit=1000|offset={offset}"
        data = get({"action": "ask", "query": q, "format": "json"})
        res = data["query"]["results"]
        for name, rec in res.items():
            p = rec.get("printouts", {})
            chars[name] = {
                "name": rec.get("fulltext", name),
                "element": (p.get("元素属性") or [""])[0],
                "weapon": (p.get("武器类型") or [""])[0],
                "rarity": (p.get("稀有度") or [""])[0],
                "version": (p.get("实装版本") or [""])[0],
                "date": (p.get("实装日期") or [""])[0],
                "region": (p.get("所属") or [""])[0],
                "tag": (p.get("TAG") or [""])[0],
            }
        nxt = data.get("query-continue-offset")
        print("batch", len(res), "total", len(chars), "next", nxt)
        if nxt is None:
            break
        offset = int(nxt)
        time.sleep(1.5)
    return chars


def fetch_pages(titles, cache_dir=None, batch_size=50, rounds=3):
    """批量取页面 wikitext（含 redirects 解析、时间戳），按批缓存到 pages/batch_XX.json。

    返回 {requested_title: {title, missing, content, timestamp, redirected}}。
    网络失败按 8/16/32s 指数退避；每批之间随机停 2~3s 防止 567。
    """
    cache_dir = cache_dir or os.path.join(CACHE, "pages")
    os.makedirs(cache_dir, exist_ok=True)
    result = {}
    batches = [titles[i:i + batch_size] for i in range(0, len(titles), batch_size)]
    for bi, batch in enumerate(batches):
        path = os.path.join(cache_dir, "batch_%02d.json" % bi)
        if os.path.exists(path):
            with open(path, encoding="utf-8") as f:
                recorded = json.load(f)
        else:
            params = {
                "action": "query", "prop": "revisions", "titles": "|".join(batch),
                "rvprop": "content|timestamp", "rvslots": "main",
                "redirects": "1", "format": "json", "formatversion": "2",
            }
            data = None
            for r in range(rounds):
                try:
                    data = get(params)
                    break
                except Exception as e:  # noqa: BLE001
                    wait = 8 * (2 ** r)
                    print("batch %d retry %d after %ss: %s" % (bi, r + 1, wait, e))
                    time.sleep(wait)
            if data is None:
                recorded = {b: {"requested": b, "missing": True, "content": "",
                                "timestamp": "", "error": "fetch_failed"} for b in batch}
            else:
                recorded = _parse_page_response(data, batch)
            recorded = {"batch": bi, "requested": batch, "results": recorded}
            with open(path, "w", encoding="utf-8") as f:
                json.dump(recorded, f, ensure_ascii=False)
            print("batch %d fetched %d pages" % (bi, len(recorded["results"])))
        result.update(recorded["results"])
        if bi < len(batches) - 1:
            time.sleep(random.uniform(2.0, 3.0))
    return result


def _parse_page_response(data, batch):
    q = data.get("query", {})
    normalized = {n["from"]: n["to"] for n in q.get("normalized", [])}
    redirects = {r["from"]: r["to"] for r in q.get("redirects", [])}
    pages = {p["title"]: p for p in q.get("pages", [])}
    out = {}
    for req in batch:
        title = normalized.get(req, req)
        title = redirects.get(title, title)
        page = pages.get(title)
        if page is None:  # 理论上不会发生
            out[req] = {"requested": req, "title": title, "missing": True,
                        "content": "", "timestamp": "", "error": "no_page_in_response"}
            continue
        if page.get("missing"):
            out[req] = {"requested": req, "title": req, "missing": True,
                        "content": "", "timestamp": ""}
            continue
        rev = (page.get("revisions") or [{}])[0]
        content = ((rev.get("slots") or {}).get("main") or {}).get("content", "")
        if content.startswith("#REDIRECT") or content.startswith("#redirect"):
            out[req] = {"requested": req, "title": title, "missing": True,
                        "content": "", "timestamp": "", "error": "unresolved_redirect"}
            continue
        out[req] = {"requested": req, "title": page.get("title", title),
                    "missing": False, "content": content,
                    "timestamp": rev.get("timestamp", ""),
                    "redirected": page.get("title", title) != req}
    return out


def fetch_all_guides(roster_path=None):
    """取 roster 中全部角色的「名/攻略」页并缓存。返回 {角色名: page_record}。"""
    roster_path = roster_path or os.path.join(CACHE, "roster.json")
    with open(roster_path, encoding="utf-8") as f:
        roster = json.load(f)
    names = list(roster.keys())
    guides = fetch_pages([n + "/攻略" for n in names])
    out = {}
    for n in names:
        rec = guides.get(n + "/攻略")
        if rec:
            rec = dict(rec)
            rec["requested"] = n
            out[n] = rec
    return out


if __name__ == "__main__":
    import sys
    if "--guides" in sys.argv:
        pages = fetch_all_guides()
        print("guide pages:", len(pages), "ok:", sum(1 for p in pages.values() if not p.get("missing")))
    else:
        chars = ask_all()
        out = os.path.join(CACHE, "roster.json")
        with open(out, "w", encoding="utf-8") as f:
            json.dump(chars, f, ensure_ascii=False, indent=1)
        versions = sorted({c["version"] for c in chars.values() if c["version"]})
        print("saved", out, "chars", len(chars), "versions", versions)
