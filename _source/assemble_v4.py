"""从分模块源码拼装生成器，可在任意工作目录执行。"""
from pathlib import Path

HERE = Path(__file__).resolve().parent
parts = ["v4_common.py", "v4_build_a.py", "v4_build_b.py", "v4_build_c.py", "v4_build_d.py"]
(HERE / "build_tracker_v4.py").write_text("\n\n".join((HERE / p).read_text(encoding="utf-8") for p in parts), encoding="utf-8")
print("assembled build_tracker_v4.py")
