"""Генерирует тестовую выгрузку (manifest.txt + tsv) и историю для build_assignment.ps1.

Использование: python make_sample_export.py <папка> [--empty] [--fixed] [--drop-s1] [--no-history]
"""
import datetime
import os
import sys


def main():
    out = sys.argv[1]
    empty = "--empty" in sys.argv
    os.makedirs(out, exist_ok=True)
    # как в реальных спецификациях: поле «Семейство и типоразмер» с переименованным заголовком,
    # проверяемый параметр виден в фильтре (4-й столбец manifest)
    head = ['Проверить, чтобы была заполнена "Длина"', "Наименование", "ADSK_Размер_Длина", "Количество", "Связанный файл"]
    fields = "|".join(["Семейство и типоразмер", "Наименование", "ADSK_Размер_Длина", "Количество", "Связанный файл"])
    filt = "ADSK_Размер_Длина"
    key_param = head[0] + ", " + head[2]  # ключ истории: заголовки колонок-параметров по прежним правилам
    models = [
        ("ORG_K1_PROJ_IOS_EOM_R23.rvt", "EOM"),
        ("ORG_K2_PROJ_IOS_EOM_R23.rvt", "EOM"),
        ("ORG_K1_PROJ_IOS_SS_R23.rvt", "SS"),
    ]
    manifest = []
    files = {
        "s1.tsv": ("ЭОМ Проверка заполнения", [
            ("Светильник: тип А", "Светильник 1", "", "1", models[0][0]),
            ("Светильник: тип Б", "Светильник 2", "007", "1", models[0][0]),
            ("Розетка", "=1+1", "  пробел ", "1", models[1][0]),
        ]),
        "s2.tsv": ("СС Проверка марки", [
            ("Извещатель", "ИП-1", "", "1", models[2][0]),
        ]),
    }
    fixed = "--fixed" in sys.argv        # s1 проверена и пуста: в manifest «*empty*», файла нет
    drop_s1 = "--drop-s1" in sys.argv    # s1 убрана из списка спецификаций
    for fn, (name, rows) in files.items():
        if fn == "s1.tsv" and drop_s1:
            continue
        if fn == "s1.tsv" and fixed:
            manifest.append("*empty*\t%s" % name)
            continue
        manifest.append("%s\t%s\t%s\t%s" % (fn, name, fields, filt))
        with open(os.path.join(out, fn), "w", encoding="utf-8-sig", newline="") as f:
            f.write("\t".join(head) + "\n")
            if not empty:
                for r in rows:
                    f.write("\t".join(r) + "\n")
    with open(os.path.join(out, "manifest.txt"), "w", encoding="utf-8-sig", newline="") as f:
        f.write("\n".join(manifest) + "\n")
    with open(os.path.join(out, "export_result.txt"), "w", encoding="utf-8") as f:
        f.write("exported=0\n" if empty else "exported=4\n")
    if "--no-history" in sys.argv:
        return
    # история: позавчера (старый 6-колоночный формат) и вчера (с 7-й колонкой — имя параметра у EOM)
    d1 = (datetime.date.today() - datetime.timedelta(days=1)).isoformat()
    d2 = (datetime.date.today() - datetime.timedelta(days=2)).isoformat()
    hist = os.path.join(out, "..", "задание_история.txt")
    with open(hist, "w", encoding="utf-8-sig", newline="") as f:
        f.write("%s\tEOM\tORG_K1_PROJ_IOS_EOM_R23\tЭОМ Проверка заполнения\t%s\t5\n" % (d2, key_param))
        f.write("%s\tSS\tORG_K1_PROJ_IOS_SS_R23\tСС Проверка марки\t%s\t3\n" % (d2, key_param))
        f.write("%s\tEOM\tORG_K1_PROJ_IOS_EOM_R23\tЭОМ Проверка заполнения\t%s\t2\tADSK_Размер_Длина\n" % (d1, key_param))
        f.write("%s\tSS\tORG_K1_PROJ_IOS_SS_R23\tСС Проверка марки\t%s\t1\n" % (d1, key_param))

main()
