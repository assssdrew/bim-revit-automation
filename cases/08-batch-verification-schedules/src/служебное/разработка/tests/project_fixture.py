"""Данные для тестов: файл настройки проекта, собранный по прежним пакетам schedule_sets.txt (в коде инструмента этих таблиц нет —
распределение задаётся BIM-менеджером; здесь пакеты нужны только как готовый набор отметок для проверки потребителей)."""
import os

ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))))
TOOL = os.path.join(ROOT, "служебное", "скрипты")
SETS = os.path.join(os.path.dirname(os.path.abspath(__file__)), "fixtures", "schedule_sets.txt")

PACKS = {"OV": ("common", "ducts", "equipment", "geo"), "PD": ("common", "ducts"), "VK": ("common", "pipes", "equipment", "geo"),
         "OT": ("common", "pipes", "equipment", "geo", "thermal"), "HS": ("common", "pipes", "equipment", "geo", "thermal"),
         "PT": ("common", "pipes", "equipment", "geo"), "EOM": ("common", "trays", "conduits", "geo"), "SS": ("common", "trays", "conduits", "geo")}


def section_of(model):
    for code in PACKS:
        if "_%s_" % code in model.upper():
            return code
    return ""


def project_text(models=None, in_project=None):
    """Файл настройки: отмечены спецификации пакетов раздела. in_project=None — без раздела [inproject] (в проекте все)."""
    groups, cur = {}, None
    with open(SETS, encoding="utf-8") as f:
        for line in f.read().splitlines():
            t = line.strip()
            if t.startswith("[") and t.endswith("]"):
                cur = t[1:-1]
                groups[cur] = []
            elif t and cur:
                groups[cur].append(t.split("        ", 1)[-1] if t.startswith("П        ") else t)
    specs = []
    for g in groups.values():
        for n in g:
            if n not in specs:
                specs.append(n)
    out = ["[project]", "stage=П"]
    names = models or ["PROJ_IOS_%s_ORG_K01_R24" % c for c in PACKS]
    out.append("[models]")
    for m in names:
        out.append("%s\t%s" % (m, section_of(m)))
    out.append("[specs]")
    out += specs
    if in_project is not None:
        out.append("[inproject]")
        out += [s for s in specs if in_project(s)]
    out.append("[cells]")
    for m in names:
        code = section_of(m)
        on = {n for pk in PACKS.get(code, ()) for n in groups.get(pk, [])}
        for sp in specs:
            out.append("%s\t%s\t%d\ta\t" % (m, sp, 1 if sp in on else 0))
    return "\n".join(out) + "\n"
