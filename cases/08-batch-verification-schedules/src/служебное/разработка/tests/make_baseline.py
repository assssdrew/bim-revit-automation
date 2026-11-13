"""Записать эталон режима «Всё» (запускать ТОЛЬКО на коде, поведение которого признано эталонным):
    PWSH=/путь/к/pwsh python tests/make_baseline.py            — режим «Всё» (tests/baseline)
    PWSH=/путь/к/pwsh python tests/make_baseline.py --partial  — «Выборочно», «Динамика», файлы состояния (tests/baseline_partial)
"""
import json
import os
import shutil
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import baseline_lib as bl  # noqa: E402

pwsh = os.environ.get("PWSH") or shutil.which("pwsh")
if not pwsh:
    sys.exit("нужен pwsh (PWSH=...)")
if "--partial" in sys.argv:
    # эталон «Выборочно»/«Динамики» и файлов состояния (снят на коде фазы 6)
    bl.save_partial_baseline(bl.partial_scenario(pwsh))
    print("эталон записан в", bl.BASELINE_PARTIAL)
else:
    bl.save_baseline(bl.all_scenarios(pwsh))
    print("эталон записан в", bl.BASELINE)
