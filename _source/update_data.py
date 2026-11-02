# -*- coding: utf-8 -*-
"""一键更新：重抓 biligame WIKI → 重建 Excel。

用法（在“原神角色养成管理”文件夹双击『更新数据.bat』，或在本目录运行）：
    python _source/update_data.py            # 抓取 + 解析 + 重建 + 重算缓存
    python _source/update_data.py --offline  # 只用本地缓存重建（不联网）

流程：
  1. 拉取角色名单（SMW ask API）→ _wiki_tmp/roster.json
  2. 拉取每个角色「名/攻略」页 → _wiki_tmp/pages/batch_*.json
  3. 解析默认流派配装 → _wiki_tmp/wiki_builds.json
  4. 复制 roster / builds 到 _source/，解析毕业面板 → 重建 原神角色养成管理表_v1.0.xlsx
  5. 用 LibreOffice Kit 重算公式缓存（找不到则跳过，Excel 打开时会自动重算）
"""
import os
import shutil
import time
import subprocess
import sys

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SRC = os.path.join(BASE, "_source")
TMP = os.path.join(BASE, "_wiki_tmp")
PY = sys.executable
os.environ.setdefault("PYTHONIOENCODING", "utf-8")

from import_miyoushe import recalc


def run(args, desc):
    print(f"\n[update] {desc}")
    print("  $", " ".join(args))
    r = subprocess.run(args, cwd=BASE)
    if r.returncode != 0:
        raise SystemExit(f"[update] 失败：{desc}（exit {r.returncode}）")


def main():
    offline = "--offline" in sys.argv
    os.makedirs(TMP, exist_ok=True)

    if not offline:
        run([PY, os.path.join(SRC, "fetch_wiki.py")], "1/5 拉取角色名单")
        run([PY, os.path.join(TMP, "parse_builds.py"), "--fetch"], "2/5 拉取攻略页并解析配装")
        run([PY, os.path.join(SRC, "extract_panels.py")], "3/5 解析毕业参考面板")
        for name, dst_name in (("roster.json", "wiki_roster.json"), ("wiki_builds.json", "wiki_builds.json")):
            source = os.path.join(TMP, name)
            if not os.path.exists(source):
                raise SystemExit(f"[update] 缺少 {source}")
            shutil.copyfile(source, os.path.join(SRC, dst_name))
    else:
        print("[update] 离线模式：使用随包附带的 _source/wiki_*.json，不重新解析网络缓存。")
    for name in ("wiki_roster.json", "wiki_builds.json", "wiki_panels.json"):
        if not os.path.exists(os.path.join(SRC, name)):
            raise SystemExit(f"[update] 缺少 _source/{name}")

    xlsx = os.path.join(BASE, "原神角色养成管理表_v1.0.xlsx")
    if os.path.exists(xlsx):
        bak_dir = os.path.join(BASE, "旧版本")
        os.makedirs(bak_dir, exist_ok=True)
        stamp = time.strftime("%Y%m%d_%H%M%S")
        bak = os.path.join(bak_dir, f"更新前备份_{stamp}.xlsx")
        shutil.copy2(xlsx, bak)
        print(f"[update] 已备份当前版本到 {bak}")

    run([PY, os.path.join(SRC, "assemble_v4.py")], "拼装生成脚本")
    run([PY, os.path.join(SRC, "build_tracker_v4.py")], "4/5 重建工作簿")
    calc = os.path.join(TMP, "v4_calc.xlsx")
    if os.path.exists(calc):
        os.remove(calc)
    if recalc(xlsx, calc):
        os.replace(calc, xlsx)
        print("[update] 已写回重算结果")

    print("\n[update] 完成。数据版本、角色数、抓取时间见『数据与更新』工作表。")


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    main()
