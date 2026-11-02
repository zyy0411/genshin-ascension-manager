# -*- coding: utf-8 -*-
"""米游社账号数据导入（本地运行，不上传任何数据）

用法:
  1) 浏览器登录米游社，F12 → Application → Cookies，复制含 ltuid/ltoken 的整段 Cookie
  2) 保存到项目根目录的「米游社cookie.txt」
  3) 双击「导入米游社数据.bat」，或运行:
       python _source/import_miyoushe.py --uid 123456789 --server cn_gf01

说明:
  - 脚本只访问米游社官方接口（api-takumi.mihoyo.com），只在本机读写文件。
  - 会读取账号角色列表 → 角色基础数据 → 角色详情，把等级/武器/天赋/圣遗物填入表格。
  - 直接修改 xlsx 内部 XML，保留图表与格式；最后可选调用 LibreOffice 重算公式。
"""
import argparse
import hashlib
import io
import glob
import json
import os
import random
import re
import shutil
import subprocess
import sys
import time
import urllib.error
import urllib.parse
import urllib.request
import zipfile
from datetime import datetime
from xml.sax.saxutils import escape

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
TMP = os.path.join(BASE, "_wiki_tmp", "miyoushe")
os.makedirs(TMP, exist_ok=True)

UA = ("Mozilla/5.0 (Linux; Android 13; M2101K9C Build/TKQ1.220829.002; wv) "
      "AppleWebKit/537.36 (KHTML, like Gecko) Version/4.0 Chrome/108.0.5359.128 "
      "Mobile Safari/537.36 miHoYoBBS/2.11.1")
CN_SALT = "xV8v4Qu54lUKrEYFZkJhB8cuOh9Asafs"
BASE_CN = "https://api-takumi.mihoyo.com"
RECORD_BASE = "https://api-takumi-record.mihoyo.com"

# 米游社不同客户端类型使用不同 salt 与 DS 算法（见 UIGF 文档）
# type 5: 4X salt + DS2（Android App）; type 4: LK2 salt + DS1（网页）
# type 2: K2 salt + DS1（Android 旧版）; type 1: K2 salt + DS1（iOS）
CLIENT_TYPES = {
    5: {"salt": CN_SALT, "ds": 2, "app_version": "2.11.1", "client_type": "5",
        "ua": UA, "referer": "https://webstatic.mihoyo.com", "origin": "https://webstatic.mihoyo.com",
        "extra": {}},
    4: {"salt": "EJncUPGnOHajenjLhBOsdpwEMZmiCmQX", "ds": 1, "app_version": "2.71.1", "client_type": "4",
        "ua": ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
               "(KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"),
        "referer": "https://www.miyoushe.com/", "origin": "https://www.miyoushe.com",
        "extra": {}},
    2: {"salt": "rtvTthKxEyreVXQCnhluFgLXPOFKPHlA", "ds": 1, "app_version": "2.71.1", "client_type": "2",
        "ua": ("Mozilla/5.0 (Linux; Android 13; M2101K9C Build/TKQ1.220829.002; wv) "
               "AppleWebKit/537.36 (KHTML, like Gecko) Version/4.0 Chrome/120.0.0.0 "
               "Mobile Safari/537.36 miHoYoBBS/2.71.1"),
        "referer": "https://app.mihoyo.com", "origin": "https://app.mihoyo.com",
        "extra": {"x-rpc-device_fp": "38d7f349e93d8", "x-rpc-verify_key": "bll8iq97cem8",
                  "x-rpc-channel": "miyoushe"}},
    1: {"salt": "NjqtetBQOv7BxuOmK0vxzBDQfDXwDoJR", "ds": 1, "app_version": "2.68.1", "client_type": "1",
        "ua": "Hyperion/434 CFNetwork/1492.0.1 Darwin/23.3.0",
        "referer": "https://app.mihoyo.com", "origin": "https://app.mihoyo.com",
        "extra": {"x-rpc-device_fp": "38d7f349e93d8", "x-rpc-verify_key": "bll8iq97cem8",
                  "x-rpc-channel": "appstore", "x-rpc-h265_supported": "0"}},
}
CLIENT = dict(CLIENT_TYPES[2])


def set_client(n):
    global CLIENT
    CLIENT = dict(CLIENT_TYPES.get(n, CLIENT_TYPES[5]))
    log("客户端类型 x-rpc-client_type={} salt={} DS{}".format(
        CLIENT["client_type"], CLIENT["salt"], CLIENT["ds"]))

PROPS = {
    10001: "生命值", 10002: "攻击力", 10003: "防御力",
    10004: "生命值%", 10005: "攻击力%", 10006: "防御力%",
    10007: "元素精通", 10008: "元素充能效率", 10009: "治疗加成",
    10010: "暴击率", 10011: "暴击伤害", 10012: "物理伤害加成",
    20001: "火元素伤害加成", 20002: "雷元素伤害加成", 20003: "水元素伤害加成",
    20004: "草元素伤害加成", 20005: "风元素伤害加成", 20006: "岩元素伤害加成",
    20007: "冰元素伤害加成", 20008: "物理伤害加成",
}
PROP_FROM_ENKA = {
    "FIGHT_PROP_HP": "生命值", "FIGHT_PROP_ATTACK": "攻击力", "FIGHT_PROP_DEFENSE": "防御力",
    "FIGHT_PROP_HP_PERCENT": "生命值%", "FIGHT_PROP_ATTACK_PERCENT": "攻击力%",
    "FIGHT_PROP_DEFENSE_PERCENT": "防御力%", "FIGHT_PROP_CRITICAL": "暴击率",
    "FIGHT_PROP_CRITICAL_HURT": "暴击伤害", "FIGHT_PROP_CHARGE_EFFICIENCY": "元素充能效率",
    "FIGHT_PROP_ELEMENT_MASTERY": "元素精通", "FIGHT_PROP_HEAL_ADD": "治疗加成",
    "FIGHT_PROP_PHYSICAL_ADD_HURT": "物理伤害加成", "FIGHT_PROP_FIRE_ADD_HURT": "火元素伤害加成",
    "FIGHT_PROP_ELEC_ADD_HURT": "雷元素伤害加成", "FIGHT_PROP_WATER_ADD_HURT": "水元素伤害加成",
    "FIGHT_PROP_GRASS_ADD_HURT": "草元素伤害加成", "FIGHT_PROP_WIND_ADD_HURT": "风元素伤害加成",
    "FIGHT_PROP_ROCK_ADD_HURT": "岩元素伤害加成", "FIGHT_PROP_ICE_ADD_HURT": "冰元素伤害加成",
}
NUM_PROPS = {
    2: "生命值", 5: "攻击力", 8: "防御力",
    3: "生命值%", 6: "攻击力%", 9: "防御力%",
    20: "暴击率", 22: "暴击伤害", 23: "元素充能效率", 26: "治疗加成", 28: "元素精通",
    40: "物理伤害加成", 41: "雷元素伤害加成", 42: "水元素伤害加成", 43: "火元素伤害加成",
    44: "风元素伤害加成", 45: "岩元素伤害加成", 46: "冰元素伤害加成", 47: "草元素伤害加成",
}
PIECES = {1: "生之花", 2: "死之羽", 3: "时之沙", 4: "空之杯", 5: "理之冠"}
PANEL_COLS = {"cr": "Z", "cd": "AA", "atk": "AB", "hp": "AC", "def": "AD", "em": "AE", "er": "AF"}


def log(msg):
    print(msg, flush=True)


# ---------------- HTTP ----------------
def md5(s):
    return hashlib.md5(s.encode("utf-8")).hexdigest()


def gen_ds(body=None, query=""):
    t = int(time.time())
    r = random.randint(100001, 200000)
    if CLIENT.get("ds", 2) == 1:
        raw = "salt={}&t={}&r={}".format(CLIENT["salt"], t, r)
    else:
        b = json.dumps(body) if body is not None else ""
        raw = "salt={}&t={}&r={}&b={}&q={}".format(CLIENT["salt"], t, r, b, query)
    return "{},{},{}".format(t, r, md5(raw))


_DEVICE_FILE = os.path.join(TMP, "device_id.txt")


def device_id():
    try:
        if os.path.exists(_DEVICE_FILE):
            v = io.open(_DEVICE_FILE, encoding="utf-8").read().strip()
            if v:
                return v
    except Exception:
        pass
    import uuid
    v = uuid.uuid4().hex.upper()
    try:
        io.open(_DEVICE_FILE, "w", encoding="utf-8").write(v)
    except Exception:
        pass
    return v


def http_json(url, method="GET", payload=None, cookie=None, extra=None):
    headers = {
        "User-Agent": CLIENT["ua"],
        "Referer": CLIENT["referer"],
        "Origin": CLIENT["origin"],
        "X-Requested-With": "com.mihoyo.hyperion",
        "x-rpc-app_version": CLIENT["app_version"],
        "x-rpc-client_type": CLIENT["client_type"],
        "x-rpc-device_id": device_id(),
        "x-rpc-device_name": "Xiaomi M2101K9C",
        "x-rpc-device_model": "M2101K9C",
        "x-rpc-sys_version": "13",
        "x-rpc-channel": "miyoushe",
        "x-rpc-language": "zh-cn",
        "Accept": "application/json",
    }
    headers.update(CLIENT.get("extra") or {})
    if extra:
        headers.update(extra)
    body = None
    if payload is not None:
        body = json.dumps(payload).encode("utf-8")
        headers["Content-Type"] = "application/json"
    if cookie:
        headers["Cookie"] = cookie
    req = urllib.request.Request(url, data=body, headers=headers, method=method)
    try:
        raw = urllib.request.urlopen(req, timeout=30).read().decode("utf-8", "replace")
    except urllib.error.HTTPError as e:
        raw = e.read().decode("utf-8", "replace")
        raise RuntimeError("HTTP {}: {}".format(e.code, raw[:400]))
    data = json.loads(raw)
    if data.get("retcode") != 0:
        raise RuntimeError("米游社返回 retcode={} message={}".format(data.get("retcode"), data.get("message")))
    return data.get("data")


def api_get(path, params=None, cookie=None, base=None):
    q = ""
    if params:
        q = "&".join("{}={}".format(k, urllib.parse.quote(str(v))) for k, v in sorted(params.items()))
    return http_json((base or BASE_CN) + path + ("?" + q if q else ""), "GET", None,
                     cookie, {"DS": gen_ds(None, q)})


def api_post(path, payload, cookie=None, base=None, params=None):
    q = ""
    if params:
        q = "&".join("{}={}".format(k, urllib.parse.quote(str(v))) for k, v in sorted(params.items()))
    return http_json((base or BASE_CN) + path + ("?" + q if q else ""), "POST", payload,
                     cookie, {"DS": gen_ds(payload, q)})


def cookie_guide():
    log("没找到可用的米游社 Cookie（需要 ltuid_v2 / ltoken_v2）。")
    log("获取方法（任选其一）：")
    log("  A. 浏览器登录 https://www.miyoushe.com/ 或 https://bbs.mihoyo.com/")
    log("     F12 → Application（应用）→ Storage → Cookies → 点开对应网站")
    log("     优先复制 .mihoyo.com 那一组 ltuid_v2、ltoken_v2（可带 ltmid_v2）")
    log("  B. 用 Cookie-Editor 扩展 → Export → JSON，把整段 JSON 粘贴进去")
    log("  C. 直接粘贴浏览器请求头里的 Cookie 整行（含 ltuid/ltoken 的 a=b; c=d 形式）")
    log("")
    log("把内容保存到项目根目录的《米游社cookie.txt》（覆盖里面的说明文字即可）。")
    log("也支持：Chrome Cookie-Editor 导出的 JSON、Netscape cookies.txt。")
    log("保存后可先双击《检查米游社Cookie.bat》验证，再双击《导入米游社数据.bat》。")


def _cookie_pairs(text):
    text = (text or "").strip()
    if not text:
        return {}
    pairs = {}
    # 1) JSON（Cookie-Editor / 浏览器扩展导出）
    if text.startswith("[") or text.startswith("{"):
        try:
            data = json.loads(text)
            items = data if isinstance(data, list) else data.get("cookies", [])
            for it in items:
                if isinstance(it, dict) and it.get("name"):
                    pairs[str(it["name"]).strip()] = str(it.get("value", "")).strip()
        except Exception:
            pass
    # 2) Netscape cookies.txt
    if not pairs:
        for ln in text.splitlines():
            ln = ln.strip()
            if not ln or ln.startswith("#"):
                continue
            parts = ln.split("	")
            if len(parts) >= 7 and ("mihoyo" in parts[0] or "miyoushe" in parts[0]):
                pairs[parts[5].strip()] = parts[6].strip()
    # 3) 普通 Cookie 头 / a=b; c=d
    if not pairs:
        norm = " ".join(text.splitlines())
        for seg in re.split(r";\s*", norm):
            if "=" in seg:
                k, v = seg.split("=", 1)
                if k.strip():
                    pairs[k.strip()] = v.strip().strip('"')
    keep = {}
    for k, v in pairs.items():
        lk = k.lower()
        if (lk.startswith("ltuid") or lk.startswith("ltoken") or lk.startswith("ltmid")
                or lk in ("account_id", "cookie_token", "account_id_v2", "cookie_token_v2")):
            if v:
                keep[k] = v
    return keep


def read_cookie():
    for p in [os.path.join(BASE, "米游社cookie.txt"), os.path.join(BASE, "_source", "miyoushe_cookie.txt")]:
        if os.path.exists(p):
            try:
                text = io.open(p, encoding="utf-8-sig").read()
            except Exception:
                continue
            pairs = _cookie_pairs(text)
            if pairs:
                return "; ".join("{}={}".format(k, v) for k, v in pairs.items())
    return ""


def get_roles(cookie):
    return api_get("/binding/api/getUserGameRolesByCookie", {"game_biz": "hk4e_cn"}, cookie).get("list", [])


def get_avatar_index(uid, server, cookie):
    return api_get("/game_record/app/genshin/api/index",
                   {"server": server, "role_id": uid}, cookie, RECORD_BASE).get("avatars", [])


def get_characters(uid, server, ids, cookie):
    out = []
    for i in range(0, len(ids), 20):
        data = api_post("/game_record/app/genshin/api/character",
                        {"character_ids": ids[i:i + 20]},
                        cookie, RECORD_BASE, {"server": server, "role_id": uid})
        out.extend(data.get("avatars", []))
    return out


def get_details(uid, server, ids, cookie):
    out = []
    for i in range(0, len(ids), 10):
        data = api_post("/game_record/app/genshin/api/character/detail",
                        {"character_ids": ids[i:i + 10]},
                        cookie, RECORD_BASE, {"server": server, "role_id": uid})
        out.extend(data.get("list", []))
    return out


# ---------------- 解析 ----------------
def prop_name(v):
    if isinstance(v, str):
        if v.isdigit():
            return PROPS.get(int(v), v)
        return PROP_FROM_ENKA.get(v, v)
    try:
        return PROPS.get(int(v), str(v))
    except Exception:
        return str(v)


def parse_avatar(av, det=None):
    out = {"name": None, "level": None, "weapon": None, "refine": None, "weapon_level": None,
           "talents": [None, None, None], "relics": {}, "panel": {}}
    src = det or av or {}
    bi = src.get("baseInfo") or src
    out["name"] = bi.get("name") or av.get("name")
    out["level"] = bi.get("level") or av.get("level")
    out["constellation"] = bi.get("actived_constellation_num")
    out["element"] = bi.get("element")

    w = src.get("weapon") or av.get("weapon") or {}
    out["weapon"] = w.get("name")
    out["weapon_level"] = w.get("level")
    ref = w.get("affix_level", w.get("affixLevel"))
    if ref is None:
        am = w.get("affixMap") or {}
        ref = list(am.values())[0] if am else None
    out["refine"] = ref

    for i, sk in enumerate((src.get("skillList") or [])[:3]):
        out["talents"][i] = sk.get("level")

    for a in (src.get("reliquaryList") or av.get("reliquaries") or []):
        try:
            pos = int(a.get("pos") or a.get("posId"))
        except Exception:
            continue
        flat = a.get("flat") or {}
        main = a.get("mainPropId")
        main_name = prop_name(main) if main not in (None, "") else None
        if not main_name and flat.get("reliquaryMainstat"):
            main_name = prop_name(flat["reliquaryMainstat"].get("mainPropId"))
        subs = []
        for s in (flat.get("reliquarySubstats") or []):
            subs.append((prop_name(s.get("appendPropId")), s.get("statValue")))
        if not subs:
            for sid in (a.get("appendPropIdList") or []):
                subs.append((prop_name(sid), None))
        setname = ((a.get("set") or {}).get("name") or a.get("setName") or "")
        out["relics"][pos] = {"level": a.get("level"), "set": setname,
                              "main": main_name, "subs": subs}

    # 面板：部分接口会给 fightPropMap / stats 列表，没有就留空
    fp = src.get("fightPropMap") or {}
    if fp:
        def g(k):
            v = fp.get(str(k))
            return float(v) if v is not None else None
        cr, cd, er = g(20), g(22), g(23)
        out["panel"] = {"hp": g(2), "atk": g(5), "def": g(8),
                        "cr": None if cr is None else round(cr * 100, 1),
                        "cd": None if cd is None else round(cd * 100, 1),
                        "er": None if er is None else round(er * 100, 1),
                        "em": g(28)}
    return out


# ---------------- 直接改 xlsx 内部 XML ----------------
def _unescape(s):
    return (s.replace("&amp;", "&").replace("&lt;", "<").replace("&gt;", ">")
             .replace("&quot;", '"').replace("&apos;", "'"))


def sheet_paths(zf):
    wb = zf.read("xl/workbook.xml").decode("utf-8")
    rels = zf.read("xl/_rels/workbook.xml.rels").decode("utf-8")
    rel_map = {}
    for rel in re.findall(r"<Relationship [^>]*>", rels):
        rid = re.search(r'Id="([^"]+)"', rel)
        tgt = re.search(r'Target="([^"]+)"', rel)
        if rid and tgt:
            rel_map[rid.group(1)] = tgt.group(1)
    out = {}
    for name, rid in re.findall(r'<sheet [^>]*?name="([^"]+)"[^>]*?r:id="([^"]+)"', wb):
        tgt = rel_map.get(rid, "")
        if tgt.startswith("/"):
            tgt = tgt[1:]
        elif not tgt.startswith("xl/"):
            tgt = "xl/" + tgt
        out[_unescape(name)] = tgt
    return out



def _col_idx(ref):
    letters = "".join(ch for ch in ref if ch.isalpha())
    n = 0
    for ch in letters:
        n = n * 26 + (ord(ch.upper()) - 64)
    return n


def _create_cell(xml, ref, val):
    """单元格不存在时创建：继承同一行 B 列（没有则 A 列）的样式，按列序插入。"""
    row_no = "".join(ch for ch in ref if ch.isdigit())
    rm = re.search(r'<row r="' + row_no + r'"[^>]*?(?:/>|>.*?</row>)', xml, re.S)
    if not rm:
        log("跳过不存在的行：{}".format(ref))
        return xml
    row = rm.group(0)
    if row.endswith("/>"):
        row = row[:-2] + "></row>"
    donor = ""
    for dcol in ("B", "A", "C"):
        dm = re.search(r'<c r="' + dcol + row_no + r'"[^>]*?\bs="(\d+)"', row)
        if dm:
            donor = ' s="{}"'.format(dm.group(1))
            break
    if val is None or val == "":
        new = '<c r="{0}"{1}/>'.format(ref, donor)
    elif isinstance(val, (int, float)) and not isinstance(val, bool):
        new = '<c r="{0}"{1} t="n"><v>{2}</v></c>'.format(ref, donor, val)
    else:
        new = ('<c r="{0}"{1} t="inlineStr"><is><t xml:space="preserve">{2}</t></is></c>'
               .format(ref, donor, escape(str(val))))
    target = _col_idx(ref)
    pos = None
    for cm in re.finditer(r'<c r="([A-Z]+)' + row_no + r'"', row):
        if _col_idx(cm.group(1) + row_no) > target:
            pos = cm.start()
            break
    if pos is None:
        inner = row[:-len("</row>")] + new + "</row>"
    else:
        inner = row[:pos] + new + row[pos:]
    return xml[:rm.start()] + inner + xml[rm.end():]


def edit_cells(src, dst, edits):
    """edits = {sheet: {ref: value}}；保留样式，字符串用 inlineStr，数字用数值。"""
    with zipfile.ZipFile(src) as zin:
        paths = sheet_paths(zin)
        payload = {n: zin.read(n) for n in zin.namelist()}
    for sheet, cells in edits.items():
        key = paths.get(sheet)
        if not key:
            log("找不到工作表：" + sheet)
            continue
        xml = payload[key].decode("utf-8")
        for ref, val in cells.items():
            m = re.search(r'<c r="' + re.escape(ref) + r'"[^>]*?(/>|>.*?</c>)', xml, re.S)
            if not m:
                xml = _create_cell(xml, ref, val)
                continue
            attrs = m.group(0)[:m.group(0).find(">")]
            st = re.search(r's="([^"]+)"', attrs)
            s_attr = ' s="{}"'.format(st.group(1)) if st else ""
            if val is None or val == "":
                new = '<c r="{0}"{1}/>'.format(ref, s_attr)
            elif isinstance(val, (int, float)) and not isinstance(val, bool):
                new = '<c r="{0}"{1} t="n"><v>{2}</v></c>'.format(ref, s_attr, val)
            else:
                new = ('<c r="{0}"{1} t="inlineStr"><is><t xml:space="preserve">{2}</t></is></c>'
                       .format(ref, s_attr, escape(str(val))))
            xml = xml[:m.start()] + new + xml[m.end():]
        payload[key] = xml.encode("utf-8")
    tmp = dst + ".tmp"
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for n, data in payload.items():
            zout.writestr(n, data)
    shutil.move(tmp, dst)


def recalc(src, dst):
    """可选重算工具由使用者配置，未配置时交给 Excel 打开后计算。"""
    node = os.environ.get("NODE_EXE") or shutil.which("node")
    cli = os.environ.get("LIBREOFFICE_KIT_CLI", "")
    if not (node and cli and os.path.isfile(cli)):
        log("跳过重算，Excel 打开会自动计算。")
        return False
    if os.path.exists(dst):
        os.remove(dst)
    try:
        result = subprocess.run([node, cli, "recalculate", "--input", src, "--output", dst],
                                check=False, timeout=300)
        return result.returncode == 0 and os.path.exists(dst)
    except (OSError, subprocess.TimeoutExpired):
        log("重算失败，Excel 打开会自动计算。")
        return False


# ---------------- 解析 character/detail ----------------
def _num(v):
    if v is None:
        return None
    if isinstance(v, (int, float)):
        return float(v)
    try:
        return float(str(v).replace("%", "").strip())
    except Exception:
        return None


def _prop(bucket, ptype):
    for it in bucket or []:
        if int(it.get("property_type", -1)) == ptype:
            return _num(it.get("final", it.get("value")))
    return None


def parse_detail(det):
    """把米游社 character/detail 的单个角色解析成模板需要的结构。"""
    base = det.get("base") or {}
    weapon = det.get("weapon") or base.get("weapon") or {}
    out = {"name": base.get("name"), "level": base.get("level"),
           "constellation": base.get("actived_constellation_num"),
           "element": base.get("element"),
           "weapon": weapon.get("name"), "weapon_level": weapon.get("level"),
           "refine": weapon.get("affix_level"),
           "talents": [None, None, None], "relics": {}, "panel": {}}
    # 天赋：skill_type=1 为战斗天赋，顺序 [普攻, 战技, (冲刺), 爆发]；取第 1、第 2、最后一个
    combat = [sk for sk in (det.get("skills") or []) if sk.get("skill_type") == 1 and sk.get("level") is not None]
    if len(combat) >= 3:
        out["talents"] = [combat[0].get("level"), combat[1].get("level"), combat[-1].get("level")]
    elif combat:
        out["talents"] = [sk.get("level") for sk in combat[:3]] + [None] * (3 - len(combat[:3]))
    # 最终面板
    sel = det.get("selected_properties") or []
    out["panel"] = {
        "cr": _prop(sel, 20), "cd": _prop(sel, 22),
        "hp": _prop(sel, 2000), "atk": _prop(sel, 2001), "def": _prop(sel, 2002),
        "em": _prop(sel, 28), "er": _prop(sel, 23),
    }
    # 圣遗物
    for r in det.get("relics") or []:
        try:
            pos = int(r.get("pos"))
        except Exception:
            continue
        mp = r.get("main_property") or {}
        main = NUM_PROPS.get(int(mp.get("property_type", -1)))
        subs = []
        for sp in r.get("sub_property_list") or []:
            nm = NUM_PROPS.get(int(sp.get("property_type", -1)))
            if nm:
                subs.append((nm, _num(sp.get("value"))))
        out["relics"][pos] = {"level": r.get("level"), "set": (r.get("set") or {}).get("name"),
                              "main": main, "main_value": _num(mp.get("value")), "subs": subs}
    return out


# ---------------- main ----------------
def fill(template, parsed):
    """把解析后的角色数据写入模板副本，返回输出路径（不依赖网络，便于测试）。"""
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    import openpyxl
    wb = openpyxl.load_workbook(template, read_only=True, data_only=False)
    ws_main, ws_det = wb["养成总表"], wb["圣遗物明细"]
    row_by_name, empty_rows = {}, []
    for r in range(3, 136):
        nm = ws_main.cell(r, 2).value
        if nm:
            row_by_name[str(nm).strip()] = r
        else:
            empty_rows.append(r)
    det_rows = {}
    for r in range(3, 668):
        nm, piece = ws_det.cell(r, 1).value, ws_det.cell(r, 2).value
        if nm and piece:
            det_rows[(str(nm).strip(), str(piece).strip())] = r
    wb.close()

    edits_main, edits_det, filled, skipped = {}, {}, [], []
    panel_missing = 0
    ELEM_CN = {"Anemo": "风", "Geo": "岩", "Electro": "雷", "Dendro": "草",
               "Hydro": "水", "Pyro": "火", "Cryo": "冰"}
    for p in parsed:
        nm = (p.get("name") or "").strip()
        if not nm:
            skipped.append("空名字")
            continue
        if nm.startswith("奇偶"):
            nm = "奇偶"
        if "旅行者" in nm:
            cn = ELEM_CN.get(str(p.get("element") or ""), "")
            if cn:
                nm = "旅行者/" + cn
            else:
                skipped.append(nm)
                continue
        r = row_by_name.get(nm)
        if r is None and empty_rows:
            r = empty_rows.pop(0)
            edits_main["B{}".format(r)] = nm
            row_by_name[nm] = r
        if r is None:
            skipped.append(nm)
            continue
        edits_main["AQ{}".format(r)] = "已有"
        if p.get("constellation") is not None:
            edits_main["AP{}".format(r)] = int(p["constellation"])
        if p.get("level"):
            edits_main["D{}".format(r)] = int(p["level"])
        edits_main["E{}".format(r)] = 90
        if p.get("weapon"):
            edits_main["H{}".format(r)] = p["weapon"]
        if p.get("refine"):
            edits_main["I{}".format(r)] = "精{}".format(min(5, max(1, int(p["refine"]))))
        if p.get("weapon_level"):
            edits_main["J{}".format(r)] = int(p["weapon_level"])
        for i, col in enumerate(["S", "T", "U"]):
            if p["talents"][i]:
                edits_main["{}{}".format(col, r)] = int(p["talents"][i])
        panel = p.get("panel") or {}
        has_panel = False
        for k, col in PANEL_COLS.items():
            v = panel.get(k)
            if v:
                edits_main["{}{}".format(col, r)] = round(float(v), 1)
                has_panel = True
        if not has_panel:
            panel_missing += 1
        for pos, relic in (p.get("relics") or {}).items():
            dr = det_rows.get((nm, PIECES.get(pos)))
            if not dr:
                continue
            if relic.get("set"):
                edits_det["C{}".format(dr)] = relic["set"]
            if relic.get("main"):
                edits_det["D{}".format(dr)] = relic["main"]
            if relic.get("level") is not None:
                edits_det["E{}".format(dr)] = int(relic["level"])
            for j, (sname, sval) in enumerate((relic.get("subs") or [])[:4]):
                if sname:
                    edits_det["{}{}".format("FGHI"[j], dr)] = sname
                if sval is not None:
                    edits_det["{}{}".format("JKLM"[j], dr)] = round(float(sval), 2)
        filled.append(nm)

    if not filled:
        log("没有可写入的角色（可能都失败了）。")
        return 1

    out = os.path.join(BASE, "原神角色养成管理表_我的数据_{}.xlsx".format(stamp))
    edit_cells(template, out, {"养成总表": edits_main, "圣遗物明细": edits_det})
    log("已写入 {} 个角色：{}".format(len(filled), "、".join(filled[:15]) + ("…" if len(filled) > 15 else "")))
    if skipped:
        log("跳过（旅行者/无匹配行）：{}".format("、".join([s for s in skipped if s][:12])))
    if panel_missing:
        log("提示：{} 个角色接口未返回最终面板（暴击率等），面板列留空，可在主表手填；"
            "填了面板才能触发「面板达标」判定。".format(panel_missing))

    calc = out.replace(".xlsx", "_已重算.xlsx")
    if recalc(out, calc):
        os.replace(calc, out)
        log("已用 LibreOffice 重算公式。")
    log("完成：" + out)
    return out


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--uid", default="")
    ap.add_argument("--server", default="")
    ap.add_argument("--cookie", default="")
    ap.add_argument("--template", default="")
    ap.add_argument("--update", default="", help="覆盖更新指定表格（自动备份到 旧版本/）")
    ap.add_argument("--check", action="store_true", help="只检查 Cookie 并列出账号角色，不写表格")
    ap.add_argument("--client-type", type=int, default=2, choices=[1, 2, 4, 5],
                    help="米游社客户端类型：5=App，4=网页，2=旧Android(默认)，1=iOS")
    args = ap.parse_args()
    set_client(args.client_type)

    template = args.template
    if not template and not args.check:
        cands = sorted([f for f in os.listdir(BASE)
                        if f.startswith("原神角色养成管理表_v1.0") and f.endswith(".xlsx")
                        and "我的数据" not in f])
        if not cands:
            log("找不到 v1.0 模板，请确认表格文件名。")
            return 1
        template = os.path.join(BASE, cands[0])
    if template:
        log("模板：" + template)

    cookie = args.cookie or read_cookie()
    if not cookie:
        cookie_guide()
        return 1
    if args.check:
        try:
            roles = get_roles(cookie)
        except Exception as e:
            log("Cookie 检查失败：{}".format(e))
            log("如果提示登录失效（retcode -100 / 10001），请重新登录米游社后复制新的 Cookie。")
            return 1
        if not roles:
            log("Cookie 可用，但账号下没有找到原神角色（可能未绑定国服角色）。")
            return 1
        log("Cookie 可用，检测到 {} 个原神角色：".format(len(roles)))
        for i, r in enumerate(roles, 1):
            log("  {}. {}  UID {}  服务器 {}".format(i, r.get("nickname"), r.get("game_uid"), r.get("region")))
        return 0

    roles = []
    try:
        roles = get_roles(cookie)
    except Exception as e:
        log("读取米游社角色列表失败：{}".format(e))
        log("如果 Cookie 已过期，请重新登录米游社复制。")
        if not args.uid:
            try:
                args.uid = input("也可以手动输入原神 UID 继续：").strip()
            except EOFError:
                args.uid = ""
        if not args.uid:
            return 1

    uid, server = str(args.uid), args.server
    if roles:
        log("检测到米游社角色：")
        for i, r in enumerate(roles, 1):
            log("  {}. {}  UID {}  服务器 {}".format(i, r.get("nickname"), r.get("game_uid"), r.get("region")))
        if not uid:
            if len(roles) == 1:
                r = roles[0]
            else:
                pick = input("选择要导入的角色序号（回车=1）：").strip() or "1"
                r = roles[int(pick) - 1]
            uid, server = str(r.get("game_uid")), r.get("region")
        else:
            for r in roles:
                if str(r.get("game_uid")) == uid and not server:
                    server = r.get("region")
    if not uid:
        return 1
    server = server or ("cn_gf01" if uid[:1] in ("1", "2") else "cn_qd01")
    log("UID={}  服务器={}".format(uid, server))

    details = []
    used_cache = False
    try:
        idx = get_avatar_index(uid, server, cookie)
    except Exception as e:
        log("读取角色一览失败：{}".format(e))
        cache = sorted(glob.glob(os.path.join(TMP, "raw_{}_*.json".format(uid))))
        if cache:
            log("改用本地缓存继续：" + cache[-1])
            details = json.load(io.open(cache[-1], encoding="utf-8")).get("details") or []
            used_cache = True
        else:
            log("常见原因：Cookie 已过期、账号未绑定原神角色、或服务器选错（cn_gf01=天空岛，cn_qd01=世界树）。")
            return 1
    if not used_cache:
        ids = [a.get("id") for a in idx if a.get("id")]
        log("角色一览：{} 个".format(len(ids)))
        # character/detail 已包含等级、武器、天赋、圣遗物与面板，直接用它
        try:
            details = get_details(uid, server, ids, cookie)
        except Exception as e:
            log("读取角色详情失败：{}".format(e))
            return 1
    log("角色详情：{} 个".format(len(details)))
    parsed = [parse_detail(d) for d in details if (d.get("base") or {}).get("name")]
    avatars = []

    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    raw_path = os.path.join(TMP, "raw_{}_{}.json".format(uid, stamp))
    io.open(raw_path, "w", encoding="utf-8").write(json.dumps(
        {"roles": roles, "avatars": avatars, "details": details}, ensure_ascii=False, indent=1))
    log("原始数据已保存：" + raw_path)

    res = fill(template, parsed)
    if not isinstance(res, str):
        return 1
    if args.update:
        upd = args.update
        if upd.strip().lower() == "auto":
            cands = sorted([f for f in os.listdir(BASE)
                            if f.startswith("原神角色养成管理表_v1.0") and f.endswith(".xlsx")
                            and "我的数据" not in f])
            if not cands:
                log("找不到主表格（原神角色养成管理表_v1.0.xlsx）")
                return 1
            upd = os.path.join(BASE, cands[0])
        update = os.path.abspath(upd)
        if not os.path.exists(update):
            log("找不到要更新的表格：" + update)
            return 1
        bak_dir = os.path.join(BASE, "旧版本")
        os.makedirs(bak_dir, exist_ok=True)
        bak = os.path.join(bak_dir, "更新前备份_{}.xlsx".format(stamp))
        shutil.copy2(update, bak)
        os.replace(res, update)
        log("已更新：" + update)
        log("原文件已备份：" + bak)
        return update
    return res


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    result = main()
    sys.exit(result if isinstance(result, int) else 0)
