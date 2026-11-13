"""Подставной Revit API для тестов вставки спецификаций.

Классы раньше жили в test_align_schedules.py. Их импортирует test_insert_project.py.
Скрипт приведения спецификаций к эталону удалён; эта заглушка остаётся только для тестов.
"""
import sys
import types

LINK = "Связанный файл: Имя файла"
NO_VALUE_KINDS = ("HasNoValue", "HasValue", "HasParameter")

class Field:
    def __init__(self, fid, name):
        self.FieldId, self.name = fid, name
        self.ParameterId = None

    def GetName(self):
        return self.name

    @property
    def ColumnHeading(self):
        return self.name


class Flt:
    def __init__(self, fid, kind, value=""):
        self.FieldId, self.FilterType, self.value = fid, kind, value

    def GetStringValue(self):
        if self.FilterType in NO_VALUE_KINDS or not isinstance(self.value, str):
            raise TypeError("no string value")
        return self.value


class Definition:
    def __init__(self):
        self.fields, self.filters = [], []       # fields: [Field]; id = индекс
        self.IncludeLinkedFiles = True

    def add_field(self, name):
        for f in self.fields:
            if f.name == name:
                return f.FieldId
        self.fields.append(Field(len(self.fields), name))
        return len(self.fields) - 1

    def GetFieldOrder(self):
        return [f.FieldId for f in self.fields]

    def GetFieldCount(self):
        return len(self.fields)

    def GetField(self, key):
        return self.fields[key]

    def GetFilterCount(self):
        return len(self.filters)

    def GetFilter(self, i):
        return self.filters[i]

    def RemoveFilter(self, i):
        del self.filters[i]

    def AddFilter(self, f):
        self.filters.append(f)

    def RemoveField(self, fid):
        self.fields = [f for f in self.fields if f.FieldId != fid]


class View:
    _next = [1]

    def __init__(self, name, definition, template=False):
        self.Name, self.Definition, self.IsTemplate = name, definition, template
        self.Id = View._next[0]
        View._next[0] += 1
        self.Parameters = []


class Doc:
    IsWorkshared = False
    IsDetached = False

    def __init__(self, views, refs=None):
        self.views, self.saved, self.closed, self.refs, self.opened = views, 0, None, refs or {}, []
        self.committed = 0
        self.Application = self

    def Save(self):
        self.saved += 1

    def Close(self, save):
        self.closed = save

    def GetElement(self, eid):
        for v in self.views:
            if v.Id == eid:
                return v

    def OpenDocumentFile(self, path, opts=None):
        self.opened.append(path)
        ref = self.refs.get(path)
        if ref is None:
            raise Exception("Не удалось открыть эталон")
        return ref



def install(doc, path):
    class Collector:
        def __init__(self, d):
            self.d = d

        def OfClass(self, _):
            return list(self.d.views)

    class Transaction:
        def __init__(self, d, name):
            self.d = d

        def Start(self):
            pass

        def Commit(self):
            self.d.committed += 1

        def RollBack(self):
            pass

    class IdList(list):
        Count = property(lambda self: len(self))
        Add = list.append

    class GenericList:
        def __class_getitem__(cls, item):
            return IdList

    class Utils:
        @staticmethod
        def CopyElements(source, ids, target, transform, options):
            result = IdList()
            for sid in ids:
                src = [v for v in source.views if v.Id == sid][0]
                d = Definition()
                d.fields = [Field(f.FieldId, f.name) for f in src.Definition.fields]
                d.filters = [Flt(f.FieldId, f.FilterType, f.value) for f in src.Definition.filters]
                d.IncludeLinkedFiles = src.Definition.IncludeLinkedFiles
                new = View(src.Name, d)
                target.views.append(new)
                result.Add(new.Id)
            return result

    db = types.ModuleType("Autodesk.Revit.DB")
    db.FilteredElementCollector = Collector
    db.ViewSchedule = object
    db.Transaction = Transaction
    db.ElementTransformUtils = Utils
    db.CopyPasteOptions = type("CopyPasteOptions", (), {"SetDuplicateTypeNamesHandler": lambda self, h: None})
    db.Transform = types.SimpleNamespace(Identity=None)
    db.IDuplicateTypeNamesHandler = object
    db.DuplicateTypeAction = types.SimpleNamespace(UseDestinationTypes=1)
    db.ElementId = int
    db.ScheduleFilter = Flt
    db.ScheduleFilterType = types.SimpleNamespace(**{k: k for k in (
        "Equal", "NotEqual", "Contains", "NotContains", "HasValue", "HasNoValue", "GreaterThan")})
    db.BuiltInParameter = types.SimpleNamespace(RVT_LINK_FILE_NAME_PARAM=-1)
    db.WorksharingUtils = types.SimpleNamespace(GetCheckoutStatus=lambda d, i: "Owned")
    db.CheckoutStatus = types.SimpleNamespace(OwnedByOtherUser="Other")
    db.StorageType = types.SimpleNamespace(String="String")
    db.OpenOptions = type("OpenOptions", (), {})
    db.DetachFromCentralOption = types.SimpleNamespace(DetachAndPreserveWorksets=1)
    db.ModelPathUtils = types.SimpleNamespace(ConvertUserVisiblePathToModelPath=lambda p: p)
    for name in ("SynchronizeWithCentralOptions", "TransactWithCentralOptions", "RelinquishOptions"):
        setattr(db, name, type(name, (), {"__init__": lambda self, *a: None}))
    revit = types.ModuleType("Autodesk.Revit")
    revit.DB = db
    autodesk = types.ModuleType("Autodesk")
    autodesk.Revit = revit
    clr = types.ModuleType("clr")
    clr.AddReference = lambda n: None
    rsu = types.ModuleType("revit_script_util")
    rsu.GetScriptDocument = lambda: doc
    rsu.GetRevitFilePath = lambda: path
    rsu.LOG = []
    rsu.Output = rsu.LOG.append
    gen = types.ModuleType("System.Collections.Generic")
    gen.List = GenericList
    mods = {"clr": clr, "Autodesk": autodesk, "Autodesk.Revit": revit, "Autodesk.Revit.DB": db,
            "revit_script_util": rsu, "System.Collections.Generic": gen}
    saved = {k: sys.modules.get(k) for k in mods}
    sys.modules.update(mods)
    return saved, rsu


def restore(saved):
    for k, v in saved.items():
        if v is None:
            sys.modules.pop(k, None)
        else:
            sys.modules[k] = v


# Заглушка для тестов дампа фильтров (раньше в test_filter_replace.py).
# ---------- подставной Revit API ----------
class FakeField:
    def __init__(self, name):
        self.name = name

    @property
    def ColumnHeading(self):
        return self.name

    def GetName(self):
        return self.name


class FakeFilter:
    def __init__(self, field_id, kind, value):
        self.FieldId = field_id
        self.FilterType = kind
        self.value = value

    def GetStringValue(self):
        if not isinstance(self.value, str):
            raise TypeError("not a string filter")
        return self.value

    def SetValue(self, v):
        if getattr(self, "fail", False):
            raise Exception("element is owned by another user")
        self.value = v


class FakeDefinition:
    def __init__(self, fields, filters):
        self.fields = fields      # id -> имя
        self.filters = filters
        self.set_calls = 0

    def GetFilterCount(self):
        return len(self.filters)

    def GetFilter(self, i):
        f = self.filters[i]
        return FakeFilter(f.FieldId, f.FilterType, f.value) if not getattr(f, "fail", False) else f

    def GetField(self, fid):
        return FakeField(self.fields[fid])

    def SetFilter(self, i, f):
        self.set_calls += 1
        self.filters[i] = f


class FakeView:
    def __init__(self, name, definition, template=False):
        self.Name = name
        self.Definition = definition
        self.IsTemplate = template
        self.Id = name


class FakeDoc:
    IsWorkshared = False
    IsDetached = False

    def __init__(self, views):
        self.views = views
        self.saved = 0
        self.transactions = []

    def Save(self):
        self.saved += 1


def install_fake_revit(doc, path):
    class Collector:
        def __init__(self, d):
            self.d = d

        def OfClass(self, _):
            return list(self.d.views)

    class Transaction:
        def __init__(self, d, name):
            self.d = d
            self.state = "new"

        def Start(self):
            self.state = "started"
            self.d.transactions.append(self)

        def Commit(self):
            self.state = "committed"

        def RollBack(self):
            self.state = "rolledback"

    db = types.ModuleType("Autodesk.Revit.DB")
    db.FilteredElementCollector = Collector
    db.ViewSchedule = object
    db.Transaction = Transaction
    db.WorksharingUtils = types.SimpleNamespace(GetCheckoutStatus=lambda d, i: "Owned")
    db.CheckoutStatus = types.SimpleNamespace(OwnedByOtherUser="Other")
    for name in ("SynchronizeWithCentralOptions", "TransactWithCentralOptions", "RelinquishOptions"):
        setattr(db, name, type(name, (), {"__init__": lambda self, *a: None}))
    revit = types.ModuleType("Autodesk.Revit")
    revit.DB = db
    autodesk = types.ModuleType("Autodesk")
    autodesk.Revit = revit
    clr = types.ModuleType("clr")
    clr.AddReference = lambda n: None
    rsu = types.ModuleType("revit_script_util")
    rsu.GetScriptDocument = lambda: doc
    rsu.GetRevitFilePath = lambda: path
    rsu.Output = lambda t: None

    class _Rd:
        def __init__(self, p, detect):
            self.f = open(p, "r", encoding="utf-8-sig")

        def ReadToEnd(self):
            return self.f.read()

        def Close(self):
            self.f.close()

    class _Wr:
        def __init__(self, p, append, enc):
            self.f = open(p, "a" if append else "w", encoding="utf-8", newline="")

        def Write(self, t):
            self.f.write(t)

        def Close(self):
            self.f.close()

    system = types.ModuleType("System")
    sio = types.ModuleType("System.IO")
    sio.StreamReader = _Rd
    sio.StreamWriter = _Wr
    stx = types.ModuleType("System.Text")
    stx.UTF8Encoding = lambda b=False: None
    system.IO = sio
    system.Text = stx
    mods = {"clr": clr, "Autodesk": autodesk, "Autodesk.Revit": revit, "Autodesk.Revit.DB": db,
            "revit_script_util": rsu, "System": system, "System.IO": sio, "System.Text": stx}
    saved = {k: sys.modules.get(k) for k in mods}
    sys.modules.update(mods)
    return saved


def restore(saved):
    for k, v in saved.items():
        if v is None:
            sys.modules.pop(k, None)
        else:
            sys.modules[k] = v

