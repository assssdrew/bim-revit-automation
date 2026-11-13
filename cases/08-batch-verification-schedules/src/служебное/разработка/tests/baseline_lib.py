"""Эталон режима «Всё»: сценарии сборки на тестовой выгрузке и структурный дамп результата.

Дамп не зависит от упаковки xlsx (значения, форматы, границы, ширины, группы, высоты) и от текущего времени
(время сборки задаётся BVS_FAKE_NOW). make_baseline.py записывает эталон, test_baseline.py сверяет с ним.
"""
import json
import os
import shutil
import subprocess
import sys
import tempfile
import warnings

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(os.path.dirname(os.path.dirname(HERE)))
TOOL = os.path.join(ROOT, "служебное", "скрипты")      # папка с .ps1/.py окна
BASELINE = os.path.join(HERE, "baseline")

# Сценарий A: сборка с замечаниями -> исправлено -> вернулось -> спецификацию убрали из списка
SCENARIO_A = [
    ("A1", "2026-01-10 10:00:00", []),
    ("A2", "2026-01-10 11:00:00", ["--fixed"]),
    ("A3", "2026-01-10 12:00:00", []),
    ("A4", "2026-01-11 09:30:00", ["--drop-s1"]),
]

# Старые файлы: история в старом формате (6 и 7 колонок, без BuildId) и журнал без колонки «вернулось»
LEGACY_HISTORY = (
    "2026-01-08\tEOM\tORG_K1_PROJ_IOS_EOM_R23\tЭОМ Проверка заполнения\tПроверить, чтобы была заполнена \"Длина\", ADSK_Размер_Длина\t5\n"
    "2026-01-08\tSS\tORG_K1_PROJ_IOS_SS_R23\tСС Проверка марки\tПроверить, чтобы была заполнена \"Длина\", ADSK_Размер_Длина\t3\n"
    "2026-01-09\tEOM\tORG_K1_PROJ_IOS_EOM_R23\tЭОМ Проверка заполнения\tПроверить, чтобы была заполнена \"Длина\", ADSK_Размер_Длина\t2\tADSK_Размер_Длина\n"
    "2026-01-09\tSS\tORG_K1_PROJ_IOS_SS_R23\tСС Проверка марки\tПроверить, чтобы была заполнена \"Длина\", ADSK_Размер_Длина\t1\n"
)
LEGACY_LEDGER = (
    "ORG_K1_PROJ_IOS_EOM_R23\tEOM\tЭОМ Проверка заполнения\tСветильник: тип А\tСветильник 1\tADSK_Размер_Длина\t\t2026-01-08 10:00:00\t2026-01-09 10:00:00\topen\t\t0\n"
    "ORG_K1_PROJ_IOS_SS_R23\tSS\tСС Проверка марки\tИзвещатель\tИП-1\tADSK_Размер_Длина\t\t2026-01-08 10:00:00\t2026-01-09 10:00:00\topen\t\t0\n"
    "ORG_K1_PROJ_IOS_EOM_R23\tEOM\tЭОМ Проверка заполнения\tСветильник: тип Б\tСветильник 2\tADSK_Размер_Длина\t007\t2026-01-08 10:00:00\t2026-01-09 10:00:00\topen\t\t0\n"
    "ORG_K2_PROJ_IOS_EOM_R23\tEOM\tЭОМ Проверка заполнения\tРозетка\t=1+1\tADSK_Размер_Длина\tпробел\t2026-01-08 10:00:00\t2026-01-09 10:00:00\topen\t\t0\n"
)
SCENARIO_L = [("L1", "2026-01-10 10:00:00", [])]


def dump_xlsx(path):
    import openpyxl
    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        wb = openpyxl.load_workbook(path)
    out = {"sheets": []}
    for ws in wb.worksheets:
        sheet = {"name": ws.title, "dims": [ws.max_row, ws.max_column], "freeze": ws.freeze_panes,
                 "merges": sorted(str(m) for m in ws.merged_cells.ranges), "cells": [], "cols": {}, "rows": {}}
        for k, v in sorted(ws.column_dimensions.items()):
            sheet["cols"][k] = round(v.width or 0, 1)
        for r, rd in sorted(ws.row_dimensions.items()):
            info = {}
            if rd.height:
                info["h"] = round(rd.height, 1)
            if rd.outlineLevel:
                info["lvl"] = rd.outlineLevel
            if rd.hidden:
                info["hidden"] = True
            if rd.collapsed:
                info["collapsed"] = True
            if info:
                sheet["rows"][str(r)] = info
        for row in ws.iter_rows():
            for c in row:
                fill = c.fill.fgColor.rgb if c.fill and c.fill.fill_type else None
                font_color = c.font.color.rgb if c.font and c.font.color and isinstance(c.font.color.rgb, str) else None
                borders = [getattr(c.border, s).style if getattr(c.border, s) is not None else None for s in ("left", "right", "top", "bottom")]
                al = c.alignment
                entry = {"v": c.value, "fill": fill, "b": bool(c.font.b), "fc": font_color,
                         "wrap": bool(al.wrap_text), "rot": al.textRotation or 0, "h": al.horizontal, "bd": borders}
                if entry["v"] is None and not fill and not entry["b"] and not font_color and borders == [None] * 4 and not entry["rot"]:
                    continue
                entry["ref"] = c.coordinate
                sheet["cells"].append(entry)
        out["sheets"].append(sheet)
    return out


def run_pwsh(pwsh, args, env_extra=None):
    env = dict(os.environ, DOTNET_SYSTEM_GLOBALIZATION_INVARIANT="1")
    env.update(env_extra or {})
    return subprocess.run([pwsh, "-NoProfile"] + args, capture_output=True, text=True, encoding="utf-8", env=env, cwd=TOOL, timeout=300)


def run_scenario(pwsh, steps, workdir, legacy=False, extra_build_args=None):
    """Выполняет сборки подряд в одну папку; возвращает {шаг: {xlsx, payload, history, ledger}}."""
    out = os.path.join(workdir, "out")
    os.makedirs(out, exist_ok=True)
    if legacy:
        with open(os.path.join(out, "задание_история.txt"), "w", encoding="utf-8-sig", newline="") as f:
            f.write(LEGACY_HISTORY)
        with open(os.path.join(out, "задание_журнал.txt"), "w", encoding="utf-8-sig", newline="") as f:
            f.write(LEGACY_LEDGER)
    result = {}
    for step, now, flags in steps:
        exp = os.path.join(workdir, "exp_" + step)
        subprocess.check_call([sys.executable, os.path.join(HERE, "make_sample_export.py"), exp, "--no-history"] + list(flags))
        xlsx = os.path.join(out, "task.xlsx")
        args = ["-File", os.path.join(TOOL, "build_assignment.ps1"), "-ExportDir", exp, "-OutXlsx", xlsx] + list(extra_build_args or [])
        r = run_pwsh(pwsh, args, {"BVS_FAKE_NOW": now})
        if r.returncode != 0:
            raise RuntimeError("сборка %s: %s %s" % (step, r.stdout, r.stderr))
        item = {"xlsx": dump_xlsx(xlsx), "stdout": r.stdout.strip().splitlines()[-1] if r.stdout.strip() else ""}
        with open(os.path.join(out, "google_push_payload.json"), encoding="utf-8") as f:
            item["payload"] = json.load(f)
        for key, name in (("history", "задание_история.txt"), ("ledger", "задание_журнал.txt")):
            with open(os.path.join(out, name), encoding="utf-8-sig") as f:
                text = f.read().replace("\r\n", "\n")
            if key == "history":
                # порядок строк внутри среза зависит от перебора hashtable и между запусками различается — сравниваем как множество
                text = "\n".join(sorted(l for l in text.split("\n") if l)) + "\n"
            item[key] = text
        result[step] = item
    return result


def normalize_stdout(line, workdir):
    return line.replace(workdir, "<WORK>").replace("\\", "/")


def all_scenarios(pwsh):
    res = {}
    for name, steps, legacy in (("A", SCENARIO_A, False), ("L", SCENARIO_L, True)):
        tmp = tempfile.mkdtemp(prefix="bvs_base_")
        try:
            r = run_scenario(pwsh, steps, tmp, legacy=legacy)
            for step, item in r.items():
                item["stdout"] = normalize_stdout(item["stdout"], tmp)
            res.update(r)
        finally:
            shutil.rmtree(tmp, True)
    return res


def save_baseline(res):
    os.makedirs(BASELINE, exist_ok=True)
    for step, item in res.items():
        d = os.path.join(BASELINE, step)
        os.makedirs(d, exist_ok=True)
        for key in ("xlsx", "payload"):
            with open(os.path.join(d, key + ".json"), "w", encoding="utf-8") as f:
                json.dump(item[key], f, ensure_ascii=False, indent=1, sort_keys=True)
        for key in ("history", "ledger"):
            with open(os.path.join(d, key + ".txt"), "w", encoding="utf-8", newline="\n") as f:
                f.write(item[key])
        with open(os.path.join(d, "stdout.txt"), "w", encoding="utf-8") as f:
            f.write(item["stdout"] + "\n")


def load_baseline(step):
    d = os.path.join(BASELINE, step)
    item = {}
    for key in ("xlsx", "payload"):
        with open(os.path.join(d, key + ".json"), encoding="utf-8") as f:
            item[key] = json.load(f)
    for key in ("history", "ledger"):
        with open(os.path.join(d, key + ".txt"), encoding="utf-8") as f:
            item[key] = f.read()
    with open(os.path.join(d, "stdout.txt"), encoding="utf-8") as f:
        item["stdout"] = f.read().strip()
    return item


# ---------- эталон «Выборочно» и «Динамики» (снят на коде фазы 6, до объединения файлов состояния) ----------
BASELINE_PARTIAL = os.path.join(HERE, "baseline_partial")
STATE_PARTS = ("buffer", "checks", "catalog")
STATE_OLD = {"buffer": "_буфер.txt", "checks": "_проверки.txt", "catalog": "_каталог.txt"}
STATE_NEW = "_состояние.txt"


def partial_steps():
    import test_partial as tp
    rows2 = {tp.S1: [("Светильник: тип А", "Светильник 1", "", tp.E1)], tp.S2: tp.ROWS[tp.S2]}
    rows3 = {tp.S1: tp.ROWS[tp.S1] + [("Розетка", "Розетка 9", "x", tp.E2)]}
    cat_all = (tp.ALL_CAT_SPECS, tp.ALL_CAT_MODELS)
    # (шаг, время, строки, пары | None, -Partial, (спецификации каталога, модели каталога) | None — без каталога, DynamicsOnly)
    return [
        ("P1", "2026-02-01 10:00:00", tp.ROWS, None, False, cat_all, False),
        ("P2", "2026-02-02 10:00:00", rows2, [(tp.E1, tp.S1)], True, cat_all, False),
        ("P3", "2026-02-03 10:00:00", rows3, [(tp.E1, tp.S1), (tp.E2, tp.S1)], True, cat_all, False),
        ("P4", "2026-02-04 10:00:00", {tp.S1: []}, [], True, (tp.ALL_CAT_SPECS, [(tp.E1, "unloaded"), (tp.E2, "loaded"), (tp.SSM, "loaded")]), False),
        ("P5", "2026-02-05 10:00:00", {tp.S1: tp.ROWS[tp.S1][:2]}, [(tp.E1, tp.S1)], True, ([tp.S1], [(tp.E1, "loaded"), (tp.E2, "loaded")]), False),
        ("P6", "2026-02-06 10:00:00", rows2, [(tp.E1, tp.S1)], True, None, False),         # выгрузка без каталога: прежний каталог остаётся
        ("P7", "2026-02-07 10:00:00", tp.ROWS, None, False, cat_all, False),
        ("P8", "2026-02-08 10:00:00", None, None, False, None, True),                         # «Собрать динамику»
    ]


def read_state(out):
    """Состояние «Выборочно» рядом с историей задание_история.txt: {часть: отсортированные строки | None — нет}. Любой формат."""
    res = {}
    merged = os.path.join(out, "задание_история" + STATE_NEW)
    parts = None
    if os.path.exists(merged):
        parts, cur = {}, None
        with open(merged, encoding="utf-8-sig") as f:
            for line in f.read().splitlines():
                if line in ("[buffer]", "[checks]", "[catalog]"):
                    cur = line[1:-1]
                    parts[cur] = []
                elif cur and line.strip() and not line.startswith("#"):
                    parts[cur].append(line)
    for part in STATE_PARTS:
        old = os.path.join(out, "задание_история" + STATE_OLD[part])
        if parts is not None and part in parts:
            res[part] = sorted(parts[part])
        elif os.path.exists(old):
            with open(old, encoding="utf-8-sig") as f:
                res[part] = sorted(l for l in f.read().splitlines() if l.strip())
        else:
            res[part] = None
    return res


def run_partial_step(pwsh, workdir, out, step):
    import test_partial as tp
    name, now, rows, pairs, flag, cat, dyn = step
    xlsx = os.path.join(out, "task.xlsx")
    if dyn:
        args = ["-File", os.path.join(TOOL, "build_assignment.ps1"), "-OutXlsx", xlsx, "-DynamicsOnly", "-DynamicsFrom", "2026-02-01", "-DynamicsTo", "2026-02-08",
                "-HistoryPath", os.path.join(out, "задание_история.txt")]
    else:
        exp = os.path.join(workdir, "exp_" + name)
        if cat is None:
            tp.write_export(exp, rows, pairs)
        else:
            tp.write_export(exp, rows, pairs, cat[0], cat[1])
        args = ["-File", os.path.join(TOOL, "build_assignment.ps1"), "-ExportDir", exp, "-OutXlsx", xlsx] + (["-Partial"] if flag else [])
    r = run_pwsh(pwsh, args, {"BVS_FAKE_NOW": now})
    if r.returncode != 0:
        raise RuntimeError("сборка %s: %s %s" % (name, r.stdout, r.stderr))
    item = {"xlsx": dump_xlsx(xlsx), "stdout": [normalize_stdout(l, workdir) for l in r.stdout.strip().splitlines() if l.startswith(("ISSUES=", "PARTIAL="))]}
    pl = os.path.join(out, "google_push_payload.json")
    if os.path.exists(pl) and not dyn:
        with open(pl, encoding="utf-8") as f:
            item["payload"] = json.load(f)
    for key, fname in (("history", "задание_история.txt"), ("ledger", "задание_журнал.txt")):
        with open(os.path.join(out, fname), encoding="utf-8-sig") as f:
            item[key] = sorted(l for l in f.read().splitlines() if l)     # порядок строк зависит от перебора hashtable
    item["state"] = read_state(out)
    return item


def partial_scenario(pwsh, start=0, seed=None, before_step=None):
    """Шаги partial_steps() с номера start; seed(out) готовит папку (например, файлы состояния прежнего формата)."""
    tmp = tempfile.mkdtemp(prefix="bvs_pbase_")
    try:
        out = os.path.join(tmp, "out")
        os.makedirs(out)
        if seed:
            seed(out)
        res = {}
        for step in partial_steps()[start:]:
            if before_step:
                before_step(step[0], out)
            res[step[0]] = run_partial_step(pwsh, tmp, out, step)
        return res
    finally:
        shutil.rmtree(tmp, True)


def save_partial_baseline(res):
    os.makedirs(BASELINE_PARTIAL, exist_ok=True)
    for step, item in res.items():
        with open(os.path.join(BASELINE_PARTIAL, step + ".json"), "w", encoding="utf-8") as f:
            json.dump(item, f, ensure_ascii=False, indent=1, sort_keys=True)


def load_partial_baseline(step):
    with open(os.path.join(BASELINE_PARTIAL, step + ".json"), encoding="utf-8") as f:
        return json.load(f)
