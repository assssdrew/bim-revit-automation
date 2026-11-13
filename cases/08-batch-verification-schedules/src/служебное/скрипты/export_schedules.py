# -*- coding: utf-8 -*-
"""
Revit Batch Processor — выгрузка проверочных спецификаций из сводной модели.
Берутся спецификации с «Пустые» или «Без значения» в имени и явный список замечаний без такого суффикса
(LITERAL_ISSUE_PARAMS: «0_Ошибка типоразмера», «1_Классификатор_999», «16.0_Корпус_Кириллица», «17.0_Секция_Кириллица»).
Имя спецификации пишется целиком, как в Revit. Обзорные «*_Все» и «Уникальные_*» не выгружаются.

run.cfg: report_dir=папка, куда кладутся tsv и итог выгрузки export.txt
  Выборочная выгрузка (только если assign_scope=partial; иначе выгружается всё, как раньше):
    assign_scope=partial
    assign_specs=имя спецификации|имя спецификации|...   (разделитель | — в именах Revit его быть не может)
    assign_models=модель|модель|...                       (имена связей как в отчёте, без .rvt)
  Тогда выгружаются только выбранные спецификации и только строки выбранных моделей, а в раздел [manifest] файла export.txt добавляются
  строки «*scope*<TAB>partial» и «*pair*<TAB>модель<TAB>спецификация» — какие пары проверены в этом прогоне
  (включая пустые результаты): это отличает «исправлено» от «не проверялось».
Итог — один файл export.txt (пишется последним): [result] — счётчики (прежний export_result.txt), [manifest] — список tsv
и пар (прежний manifest.txt), [catalog] — все спецификации проверок и все связи сводной со статусом загрузки (прежний catalog.txt).
Читает bvs_export_result.ps1; выгрузки прежних версий с тремя файлами читаются так же.
"""

import clr
import os
import re
import time

clr.AddReference("RevitAPI")

from Autodesk.Revit.DB import FilteredElementCollector, ViewSchedule, SectionType

import revit_script_util

# Пути — только от расположения скриптов и профиля пользователя: без зашитых имён пользователей и серверов (папку можно копировать куда угодно, в т. ч. на UNC).
LOCAL_APPDATA = os.environ.get("LOCALAPPDATA") or os.path.join(os.path.expanduser("~"), "AppData", "Local")
LOCAL_CFG = os.path.join(LOCAL_APPDATA, r"BatchRvt\batch_verification_schedules\run.cfg")


def get_script_dir():
    """RBP запускает файл через exec, поэтому __file__ часто нет; тогда — текущая папка."""
    try:
        folder = os.path.dirname(os.path.abspath(__file__))
        if folder and file_exists(os.path.join(folder, "export_schedules.py")):
            return folder
    except Exception:
        pass
    return os.getcwd()



def as_text(value):
    if value is None:
        return u""
    return u"{}".format(value)


def lower_text(value):
    """Нижний регистр БЕЗ str.lower(): в IronPython 2.7 у str оно меняет только ASCII, кириллица остаётся как была
    («Ширина».lower() == «Ширина»), и сравнения с «ширина» молча не срабатывали. Меняем латиницу A-Z, кириллицу А-Я и Ё сами."""
    try:
        make_char = unichr  # IronPython 2 / Python 2
    except NameError:
        make_char = chr
    out = []
    for ch in as_text(value):
        code = ord(ch)
        if 65 <= code <= 90 or 1040 <= code <= 1071:
            ch = make_char(code + 32)
        elif code == 1025:
            ch = make_char(1105)
        out.append(ch)
    return u"".join(out)


def read_text(path):
    from System.IO import StreamReader
    reader = StreamReader(path, True)
    try:
        text = reader.ReadToEnd()
    finally:
        reader.Close()
    if text is None:
        return u""
    return u"{}".format(text)


def read_cfg():
    data = {}
    for path in (LOCAL_CFG, os.path.join(get_script_dir(), "run.cfg")):
        if not path or not file_exists(path):
            continue
        for line in read_text(path).splitlines():
            line = line.strip()
            if not line or line.startswith("#") or "=" not in line:
                continue
            key, value = line.split("=", 1)
            data[key.strip().lower()] = value.strip()
        if data.get("report_dir"):
            return data
    return data


def file_exists(path):
    """.NET вместо os.path.isfile: в IronPython 2.7 на путях с кириллицей (отчеты, сервер) os.path.* отвечает неверно."""
    try:
        from System.IO import File
        return bool(File.Exists(path))
    except ImportError:
        return os.path.isfile(path)


def ensure_dir(path):
    """Папка через .NET (os.path.isdir/os.makedirs в IronPython 2.7 на пути с кириллицей ошибаются: «папка уже существует»)."""
    try:
        from System.IO import Directory
        Directory.CreateDirectory(path)
        return
    except ImportError:
        pass
    if not os.path.isdir(path):
        os.makedirs(path)


def split_list(text):
    return [x.strip() for x in as_text(text).split(u"|") if x.strip()]


def norm_model(text):
    """Имя модели как в сборке (Get-ModelFileName): без пути и без .rvt; для сравнения — без регистра."""
    t = as_text(text).strip()
    t = re.split(u"[\\\\/]", t)[-1]
    t = re.sub(u"(?i)\\.rvt.*$", u"", t)        # всё после «.rvt» (например « : 1» у экземпляра связи) отбрасываем
    return lower_text(t.strip())


def link_column(headers):
    idx = len(headers) - 1
    for i, h in enumerate(headers):
        if u"Связанный файл" in as_text(h):
            idx = i          # как в сборке: берётся последняя такая колонка
    return idx


def collect_links(document):
    """[(имя связи без .rvt, 'loaded'|'unloaded')] по типам связей сводной модели. Ошибка не должна ломать выгрузку."""
    out = []
    try:
        from Autodesk.Revit.DB import RevitLinkType, LinkedFileStatus
        for lt in FilteredElementCollector(document).OfClass(RevitLinkType):
            name = re.sub(u"(?i)\\.rvt$", u"", as_text(lt.Name).strip())
            try:
                status = u"loaded" if lt.GetLinkedFileStatus() == LinkedFileStatus.Loaded else u"unloaded"
            except Exception:
                status = u"unloaded"
            out.append((name, status))
    except Exception:
        pass
    return out


# Опечатки в именах спецификаций: (было, стало). «Cантехприборы» с латинской C -> «Сантехприборы» с кириллической С (регистр первой буквы сохраняется)
SPEC_NAME_FIXES = (
    (u"значния", u"значения"),
    (u"\u0043антехприборы", u"\u0421антехприборы"),
    (u"\u0063антехприборы", u"\u0441антехприборы"),
)


def canon_text(name):
    """Каноническое имя: исправляются известные опечатки (SPEC_NAME_FIXES)."""
    text = as_text(name)
    for old, new in SPEC_NAME_FIXES:
        text = text.replace(old, new)
    return text


def canon_spec(name):
    """Старая опечатка «Без значния» = «Без значения»: выбранные в окне имена (из истории) сравниваются с именами в Revit по каноническому виду."""
    return canon_text(name)


# Замечания без суффикса «Пустые» / «Без значения». Ключ — номер и название без пробелов и «_»
# (префикс стадии, вид пробелов и опечатка «кирилица» не важны). Параметр взят из реальных фильтров сводной
# (tests/fixtures/align_dumps): не проверка пустоты, поэтому is_check_filter его не видит.
# Не добавлять обзорные *_Все и Уникальные_*.
LITERAL_ISSUE_PARAMS = (
    (u"0ошибкатипоразмера", u"ADSK_Наименование"),          # Equal «Ошибка типоразмера»
    (u"1классификатор999", u"Код по классификатору"),       # BeginsWith 999
    (u"16.0корпускириллица", u"ADSK_Номер здания"),         # HasValue
    (u"17.0секциякириллица", u"ADSK_Номер секции"),         # HasValue
)


def schedule_tail(name):
    """Имя без префикса стадии («П», «РД», «ПП», в том числе без пробела: «П0_…») и без лишних пробелов."""
    text = lower_text(as_text(name)).replace(u"\u00a0", u" ")
    text = u" ".join(text.split())
    for prefix in (u"рд ", u"пп ", u"п "):
        if text.startswith(prefix):
            rest = text[len(prefix):].lstrip()
            if rest[:1].isdigit():
                return rest
    for prefix in (u"рд", u"пп", u"п"):
        if text.startswith(prefix) and len(text) > len(prefix) and text[len(prefix)].isdigit():
            return text[len(prefix):]
    return text


def issue_name_key(name):
    """Ключ списка замечаний: хвост без пробелов и «_», «кирилица» = «кириллица»."""
    text = schedule_tail(name).replace(u" ", u"").replace(u"_", u"")
    return text.replace(u"кирилица", u"кириллица")


def literal_issue_param(name):
    """Проверяемый параметр спецификации из LITERAL_ISSUE_PARAMS или пустая строка."""
    key = issue_name_key(name)
    for item, param in LITERAL_ISSUE_PARAMS:
        if key == item:
            return param
    return u""


def is_issue_schedule(name):
    text = as_text(name)
    if (u"Пустые" in text) or (u"Без знач" in text):
        return True
    return bool(literal_issue_param(name))


def cell_text(view, section, row, col):
    try:
        return as_text(view.GetCellText(section, row, col)).replace(u"\t", u" ").replace(u"\r", u" ").replace(u"\n", u" ")
    except Exception:
        return u""


EXPORT_RESULT_FILE = u"export.txt"      # итог выгрузки: [result] / [manifest] / [catalog]


def write_utf8(path, text):
    from System.IO import StreamWriter
    from System.Text import UTF8Encoding
    writer = StreamWriter(path, False, UTF8Encoding(True))
    try:
        writer.Write(text)
    finally:
        writer.Close()


def write_live(index, total, name, exported, skipped):
    folder = os.path.join(os.environ.get("LOCALAPPDATA", ""), "BatchRvt", "batch_verification_schedules")
    try:
        ensure_dir(folder)
    except Exception:
        return
    path = os.path.join(folder, "live_status.txt")
    body = u"\r\n".join([
        u"index={}".format(index),
        u"total={}".format(total),
        u"name={}".format(as_text(name)),
        u"inserted={}".format(exported),
        u"skipped={}".format(skipped),
        u"params=0",
    ]) + u"\r\n"
    try:
        write_utf8(path, body)
    except Exception:
        pass


def header_cells(view, cols):
    headers = []
    header = None
    try:
        header = view.GetTableData().GetSectionData(SectionType.Header)
    except Exception:
        header = None
    hrow = -1
    if header is not None:
        try:
            if header.NumberOfRows > 0:
                hrow = header.NumberOfRows - 1
        except Exception:
            hrow = -1
    for col in range(cols):
        if hrow >= 0:
            headers.append(cell_text(view, SectionType.Header, hrow, col))
        else:
            headers.append(u"")
    return headers


def field_names(view, cols):
    """Имена видимых полей спецификации (= имена параметров) по колонкам таблицы.

    Заголовок колонки в Revit можно переименовать («Проверить, чтобы ...»), а имя поля — это сам параметр.
    Скрытые поля колонок не дают. Если число видимых полей не совпало с числом колонок — пустые имена.
    """
    names = []
    try:
        definition = view.Definition
        for i in range(definition.GetFieldCount()):
            field = definition.GetField(i)
            try:
                if field.IsHidden:
                    continue
            except Exception:
                pass
            names.append(as_text(field.GetName()))
    except Exception:
        names = []
    if len(names) != cols:
        return [u""] * cols
    return [n.replace(u"\t", u" ").replace(u"|", u"/") for n in names]


EMPTY_MARK = u"*empty*"
# служебные поля — по точному имени (кроме «Связанный файл: …»); иначе «ADSK_Материал наименование» принимался бы за служебное
SERVICE_FIELD_RE = re.compile(u"^Связанный файл|^Семейство и типоразмер$|^Наименование$|^Количество$|^Категория$|^Тип$|^Семейство$", re.I)


# Типы условий фильтра, которые НЕ проверяют пустоту (это отбор элементов): «имеет значение», «содержит», «не равно» и т.д.
SELECTION_FILTER_TYPES = (
    u"hasvalue", u"hasparameter", u"notequal", u"contains", u"notcontains", u"beginswith", u"notbeginswith",
    u"endswith", u"notendswith", u"greater", u"less", u"associated",
)
EXPORT_STAMP = time.strftime(u"%Y-%m-%d %H:%M:%S")   # когда сделана выгрузка: по этой колонке видно, свежий ли файл
def _lower_diag():
    """Что делает родной str.lower() на этом Python и что даёт lower_text: видно в фильтры_спецификаций.txt."""
    native = as_text(u"\u0428\u0438\u0440\u0438\u043d\u0430").lower()
    return u"lower(): {}; lower_text: {}; тип имени: {}".format(
        u"работает" if native == u"\u0448\u0438\u0440\u0438\u043d\u0430" else u"НЕ меняет кириллицу",
        u"ok" if lower_text(u"\u0428\u0438\u0440\u0438\u043d\u0430") == u"\u0448\u0438\u0440\u0438\u043d\u0430" else u"ОШИБКА",
        type(as_text(u"x")).__name__)


LOWER_DIAG = _lower_diag()
FILTER_DEBUG = []  # строки для фильтры_спецификаций.txt: что за фильтры и какой выбран проверяемым


def filter_infos(view):
    """Все фильтры спецификации: [(имя поля, тип условия, строковое значение или None)]."""
    infos = []
    try:
        definition = view.Definition
        for i in range(definition.GetFilterCount()):
            flt = definition.GetFilter(i)
            name = as_text(definition.GetField(flt.FieldId).GetName()).strip()
            try:
                ftype = as_text(flt.FilterType)
            except Exception:
                ftype = u""
            value = None
            try:
                value = as_text(flt.GetStringValue())
            except Exception:
                value = None
            infos.append((name, ftype, value))
    except Exception:
        pass
    return infos


def is_check_filter(ftype, value):
    """Фильтр, который проверяет пустоту: «не имеет значения» (HasNoValue) или «равно» пустой строке."""
    t = as_text(ftype).lower()
    if u"novalue" in t:
        return True
    if t == u"equal" and value is not None and value.strip() == u"":
        return True
    return False


# Тема спецификации (второй сегмент имени «П   5.1_Ширина_Детали возд-в_Без значения» -> «Ширина») -> подстроки имён полей.
# Нужна, когда фильтров «не имеет значения» несколько (в «Деталях воздуховодов» это размер и диаметр): проверяет тот, что назван в имени.
TOPIC_FIELDS = {
    u"классификатор": [u"код по классификатору"],
    u"клас-я системы": [u"классификация систем"],
    u"тип системы": [u"имя системы"],          # Revit не фильтрует по «Типу системы» (ссылка на элемент): проверяется «Имя системы»
    u"имя системы": [u"имя системы"],
    u"корпус": [u"номер здания"],
    u"секция": [u"номер секции"],
    u"этаж": [u"этаж"],
    u"мощность тепловая": [u"тепловая мощность"],
}
MARK_FIELD = u"adsk_марка изделия"            # в спецификациях систем — отбор (HasNoValue), а не проверка
CLASSIFIER_FIELD = u"код по классификатору"   # «NotEqual 999» — отбор; проверяемым бывает только при «равно пусто»/«не имеет значения»


def spec_parts(schedule_name):
    """(номер, тема) спецификации без регулярных выражений: «П   5.1_Ширина_Детали возд-в_Без значения» -> («5.1», «ширина»).
    Только строковые операции: регулярные выражения с lookbehind в IronPython 2.7 на живом Revit работали не так, как в CPython
    (тема не находилась). Префикс («П», регистр, вид пробелов) не важен: номер — последнее «слово» перед первым «_»,
    без нецифровых символов в начале («П13» -> «13»)."""
    parts = as_text(schedule_name).split(u"_")
    if len(parts) < 2:
        return u"", u""
    words = parts[0].split()
    token = words[-1] if words else u""
    i = 0
    while i < len(token) and not token[i].isdigit():
        i += 1
    token = token[i:]
    if not token or token.endswith(u".") or any(c not in u"0123456789." for c in token):
        return u"", u""
    return token, lower_text(u" ".join(parts[1].split()))


def spec_topic(schedule_name):
    return spec_parts(schedule_name)[1]


def topic_matches(topic, field_name):
    low = lower_text(field_name)
    for part in TOPIC_FIELDS.get(topic, []):
        if part in low:
            return True
    if topic in TOPIC_FIELDS:
        return False
    for word in topic.split():
        if len(word) >= 4 and word[:5] in low:
            return True
    return False


def choose_check_params(infos, schedule_name=u""):
    """Проверяемый параметр — поле фильтра, проверяющего пустоту. Остальные фильтры (отбор: «Код по классификатору»,
    материал, «ADSK_Марка изделия» и т.п.) в «Параметр» не попадают.

    1. Проверяют пустоту «не имеет значения» (HasNoValue) и «равно» пустой строке. «ADSK_Марка изделия» при другом
       проверяющем фильтре — отбор, его отбрасываем.
    2. Если проверяющих несколько — оставляем тот, чьё имя совпадает с темой из имени спецификации (Диаметр, Ширина…);
       нет совпадения — все.
    3. Если проверяющих нет — запасной вариант: не служебные фильтры без явного отбора (HasValue, содержит, не равно…),
       затем все не служебные. «Код по классификатору» в запасных вариантах не берётся никогда.
    Возвращает (имена, [помечено ли выбранным для каждого фильтра])."""
    def usable(name):
        return bool(name) and not SERVICE_FIELD_RE.search(name)

    def fallback_ok(name):
        return usable(name) and CLASSIFIER_FIELD not in lower_text(name)

    # Спецификации из LITERAL_ISSUE_PARAMS: параметр задан списком (фильтр не про пустоту).
    # «1_Классификатор_999» иначе не получит «Код по классификатору»: BeginsWith — отбор, а классификатор в запасном пути запрещён.
    forced = literal_issue_param(schedule_name)
    if forced:
        want = lower_text(forced)
        hit = [i for i, (n, t, v) in enumerate(infos) if lower_text(as_text(n).strip()) == want]
        eq = [i for i in hit if as_text(infos[i][1]).lower() == u"equal"]
        chosen_forced = eq or hit
        names = []
        for i in chosen_forced:
            if infos[i][0] not in names:
                names.append(infos[i][0])
        if not names:
            names = [forced]
        return names, [i in chosen_forced for i in range(len(infos))]

    chosen = [i for i, (n, t, v) in enumerate(infos) if usable(n) and is_check_filter(t, v)]
    if len(chosen) > 1:
        rest = [i for i in chosen if lower_text(as_text(infos[i][0]).strip()) != MARK_FIELD]
        if rest:
            chosen = rest
    if len(chosen) > 1:
        topic = spec_topic(schedule_name)
        if topic:
            hit = [i for i in chosen if topic_matches(topic, infos[i][0])]
            if hit:
                chosen = hit
    if not chosen:
        chosen = [i for i, (n, t, v) in enumerate(infos)
                  if fallback_ok(n) and not any(k in as_text(t).lower() for k in SELECTION_FILTER_TYPES)]
    if not chosen:
        chosen = [i for i, (n, t, v) in enumerate(infos) if fallback_ok(n)]
    names = []
    for i in chosen:
        if infos[i][0] not in names:
            names.append(infos[i][0])
    return names, [i in chosen for i in range(len(infos))]


def filter_param_names(view, schedule_name=u""):
    """Параметры, которые проверяет спецификация (для колонки «Параметр» отчёта, например ADSK_Размер_Ширина)."""
    infos = filter_infos(view)
    names, marks = choose_check_params(infos, schedule_name)
    for (n, t, v), mark in zip(infos, marks):
        topic = spec_topic(schedule_name)
        match = u""
        if is_check_filter(t, v) and topic:
            match = u"да" if topic_matches(topic, n) else u"нет"
        FILTER_DEBUG.append(u"\t".join([schedule_name, n, t, u"" if v is None else v, u"ПРОВЕРЯЕМЫЙ" if mark else u"",
                                        topic, EXPORT_STAMP, match, LOWER_DIAG]))
    return [n.replace(u"\t", u" ").replace(u"|", u"/") for n in names]


def main():
    document = revit_script_util.GetScriptDocument()
    cfg = read_cfg()
    out_dir = cfg.get("report_dir", u"")
    if not out_dir:
        raise Exception(u"В run.cfg нет report_dir")
    ensure_dir(out_dir)

    schedules = []
    collector = FilteredElementCollector(document).OfClass(ViewSchedule)
    for view in collector:
        try:
            if view.IsTemplate:
                continue
        except Exception:
            pass
        name = as_text(view.Name)
        if is_issue_schedule(name):
            schedules.append(view)

    partial = cfg.get("assign_scope", u"").strip().lower() == u"partial"
    sel_specs = split_list(cfg.get("assign_specs", u"")) if partial else []
    sel_models = split_list(cfg.get("assign_models", u"")) if partial else []
    # assign_pairs_file: точные пары «модель<TAB>спецификация» (окно, режим «Выборочно»): выгружаются только они, а не пересечение списков
    pair_map = None
    pairs_file = cfg.get("assign_pairs_file", u"") if partial else u""
    if pairs_file:
        if not file_exists(pairs_file):
            raise Exception(u"Нет файла пар: " + pairs_file)
        pair_map = {}
        pair_models = []
        pair_specs = []
        for pline in read_text(pairs_file).splitlines():
            pparts = pline.rstrip(u"\r\n").split(u"\t", 1)
            if len(pparts) < 2 or not pparts[0].strip() or not pparts[1].strip():
                continue
            pm, ps = pparts[0].strip(), pparts[1].strip()
            pair_map.setdefault(canon_spec(ps), set()).add(norm_model(pm))
            if pm not in pair_models:
                pair_models.append(pm)
            if ps not in pair_specs:
                pair_specs.append(ps)
        sel_models = pair_models
        sel_specs = pair_specs
    sel_norm = set(norm_model(m) for m in sel_models)
    catalog_specs = [as_text(v.Name) for v in schedules]
    if partial:
        # только выбранные спецификации; номер в «index» и «total» — среди выбранных
        sel_canon = set(canon_spec(n) for n in sel_specs)
        schedules = [v for v in schedules if canon_spec(v.Name) in sel_canon]
    seen_models = set()           # модели, у которых нашлись строки: связь точно загружена
    checked_specs = []            # выбранные спецификации, которые нашлись и проверены (в т.ч. пустые)

    total = len(schedules)
    exported = 0
    skipped = 0
    lines = []
    write_live(0, total, u"", 0, 0)

    for index, view in enumerate(schedules, 1):
        name = as_text(view.Name)
        write_live(index, total, name, exported, skipped)
        body = None
        try:
            body = view.GetTableData().GetSectionData(SectionType.Body)
        except Exception:
            body = None
        if body is None:
            skipped += 1
            continue
        if body.NumberOfRows < 1 or body.NumberOfColumns < 1:
            # спецификация проверена и пуста (замечаний нет): без этой строки нельзя отличить «исправлено» от «не проверялось»
            lines.append(EMPTY_MARK + u"\t" + name)
            filter_param_names(view, name)
            checked_specs.append(name)
            skipped += 1
            continue
        cols = body.NumberOfColumns
        headers = header_cells(view, cols)
        link_idx = link_column(headers)
        file_name = u"{0:04d}.tsv".format(exported + 1)
        out_path = os.path.join(out_dir, file_name)
        from System.IO import StreamWriter
        from System.Text import UTF8Encoding
        writer = None
        try:
            for row_index in range(body.NumberOfRows):
                row = [cell_text(view, SectionType.Body, row_index, col) for col in range(cols)]
                if not any(piece.strip() for piece in row):
                    continue
                if partial:
                    key = norm_model(row[link_idx]) if link_idx < len(row) else u""
                    if key not in sel_norm:
                        continue          # строка модели, которую не выбрали
                    if pair_map is not None and key not in pair_map.get(canon_spec(name), ()):
                        continue          # эта пара «модель × спецификация» не выбрана
                    seen_models.add(key)
                if writer is None:
                    writer = StreamWriter(out_path, False, UTF8Encoding(True))
                    writer.Write(u"\t".join(headers))
                    writer.Write(u"\r\n")
                writer.Write(u"\t".join(row))
                writer.Write(u"\r\n")
        finally:
            # файл не должен оставаться открытым, если спецификация упала на середине
            if writer is not None:
                writer.Close()
        if writer is None:
            lines.append(EMPTY_MARK + u"\t" + name)
            filter_param_names(view, name)
            checked_specs.append(name)
            skipped += 1
            continue
        checked_specs.append(name)
        lines.append(file_name + u"\t" + name + u"\t" + u"|".join(field_names(view, cols)) + u"\t" + u"|".join(filter_param_names(view, name)))
        exported += 1

    links = collect_links(document)
    loaded_norm = set(norm_model(n) for n, st in links if st == u"loaded")
    pairs = []
    unloaded = []
    if partial:
        # пара проверена, если спецификация найдена и выбрана, а связь загружена (или у неё есть строки)
        for m in sel_models:
            if norm_model(m) in loaded_norm or norm_model(m) in seen_models:
                for sp in checked_specs:
                    if pair_map is not None and norm_model(m) not in pair_map.get(canon_spec(sp), ()):
                        continue          # пара не выбрана: не проверялась, замечания не закрываются
                    pairs.append((m, sp))
            else:
                unloaded.append(m)
        lines.append(u"*scope*\tpartial")
        for m, sp in pairs:
            lines.append(u"*pair*\t" + re.sub(u"(?i)\\.rvt$", u"", m) + u"\t" + sp)      # имя как в отчёте (без .rvt)
    cat = None
    try:
        cat = [u"spec\t" + n for n in catalog_specs] + [u"model\t" + n + u"\t" + st for n, st in links]
    except Exception:
        cat = None
    write_utf8(
        os.path.join(out_dir, u"фильтры_спецификаций.txt"),
        u"\r\n".join([u"Спецификация\tПоле фильтра\tУсловие\tЗначение\tРоль\tТема спеки (распознана)\tВыгружено\tТема совпала с полем\tДиагностика"] + FILTER_DEBUG) + u"\r\n",
    )
    result = [u"matched={0}".format(total), u"exported={0}".format(exported), u"skipped={0}".format(skipped)]
    if partial:
        catalog_canon = set(canon_spec(n) for n in catalog_specs)
        missing = [n for n in sel_specs if canon_spec(n) not in catalog_canon]
        result += [u"partial=1", u"pairs={0}".format(len(pairs)), u"missing_specs={0}".format(u"|".join(missing)),
                   u"unloaded_models={0}".format(u"|".join(unloaded))]
    # один итоговый файл вместо прежних manifest.txt + export_result.txt + catalog.txt (читает bvs_export_result.ps1);
    # пишется последним: есть export.txt — выгрузка дошла до конца
    out = [u"[result]"] + result + [u"[manifest]"] + lines
    if cat is not None:
        out += [u"[catalog]"] + cat
    write_utf8(os.path.join(out_dir, EXPORT_RESULT_FILE), u"\r\n".join(out) + u"\r\n")
    write_live(total, total, u"готово", exported, skipped)
    try:
        revit_script_util.Output(u"Exported {0} skipped {1} matched {2}".format(exported, skipped, total))
    except Exception:
        pass


main()
