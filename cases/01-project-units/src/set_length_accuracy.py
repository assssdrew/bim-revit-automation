# -*- coding: utf-8 -*-
# Name: set_length_accuracy.py
# Version: 1.0
# What it does: RBP task: set project Length display accuracy (mm) on disk, UNC, and RSN:// models.
# Inputs: rvt_list.txt, accuracy.cfg; workshared via Create New Local, Detach off.
# Outputs: Updated models (Sync+Relinquish or SaveAs); RBP log output.
# How to run: Run as Revit Batch Processor script on rvt_list.txt after choose_accuracy.cmd.
# Notes: SynchronizeWithCentral for workshared; non-workshared uses SaveAs in place.

import clr
import System
import os

clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")
from Autodesk.Revit.DB import *

import revit_script_util
from revit_script_util import Output

FORCE_MILLIMETERS = True
DEFAULT_ACCURACY_MM = 0.1
TOOL_DIR_FALLBACK = r""  # optional: path to deployed copy on share


def get_script_dir():
    try:
        return os.path.dirname(os.path.abspath(__file__))
    except Exception:
        pass

    candidates = [TOOL_DIR_FALLBACK, os.getcwd()]
    for folder in candidates:
        if folder and os.path.isfile(os.path.join(folder, "accuracy.cfg")):
            return folder
    return TOOL_DIR_FALLBACK


SCRIPT_DIR = get_script_dir()
CONFIG_PATH = os.path.join(SCRIPT_DIR, "accuracy.cfg")


def load_accuracy_mm():
    if not os.path.isfile(CONFIG_PATH):
        Output("No accuracy.cfg — using default {}".format(DEFAULT_ACCURACY_MM))
        return DEFAULT_ACCURACY_MM

    with open(CONFIG_PATH, "r") as f:
        for line in f:
            line = line.strip()
            if not line or line.startswith("#"):
                continue
            if "=" in line:
                line = line.split("=", 1)[1].strip()
            return float(line.replace(",", "."))

    Output("accuracy.cfg empty — using default {}".format(DEFAULT_ACCURACY_MM))
    return DEFAULT_ACCURACY_MM


TARGET_ACCURACY_MM = load_accuracy_mm()

doc = revit_script_util.GetScriptDocument()
revitFilePath = revit_script_util.GetRevitFilePath()

Output()
Output("Script dir: {}".format(SCRIPT_DIR))
Output("Config: {}".format(CONFIG_PATH))
Output("File: {}".format(revitFilePath))
Output("Target Length accuracy: {} mm".format(TARGET_ACCURACY_MM))
Output(
    "Workshared={}, Detached={}".format(
        doc.IsWorkshared, getattr(doc, "IsDetached", False)
    )
)


def is_rsn_path(path):
    if not path:
        return False
    return path.upper().startswith("RSN://")


def clear_readonly_attribute(path):
    if is_rsn_path(path):
        return
    try:
        attrs = System.IO.File.GetAttributes(path)
        readonly = System.IO.FileAttributes.ReadOnly
        if (attrs & readonly) == readonly:
            System.IO.File.SetAttributes(path, attrs & ~readonly)
            Output("Cleared Windows ReadOnly on file.")
    except Exception as ex:
        Output("WARN clear ReadOnly: {}".format(ex))


def make_relinquish_all_options():
    """Освободить всё, что занято текущим пользователем."""
    opts = RelinquishOptions(True)
    try:
        opts.StandardWorksets = True
        opts.ViewWorksets = True
        opts.FamilyWorksets = True
        opts.UserWorksets = True
        opts.CheckedOutElements = True
    except Exception:
        pass
    return opts


def relinquish_everything(document):
    if not document.IsWorkshared:
        return
    try:
        Output("Relinquishing all worksets/elements...")
        WorksharingUtils.RelinquishOwnership(
            document, make_relinquish_all_options(), TransactWithCentralOptions()
        )
        Output("Relinquish done.")
    except Exception as ex:
        Output("WARN RelinquishOwnership: {}".format(ex))


def set_length_accuracy(document, accuracy_mm):
    units = document.GetUnits()
    opts = units.GetFormatOptions(SpecTypeId.Length)

    if FORCE_MILLIMETERS:
        opts.SetUnitTypeId(UnitTypeId.Millimeters)

    opts.Accuracy = float(accuracy_mm)

    try:
        opts.SuppressTrailingZeros = True
    except Exception:
        pass

    units.SetFormatOptions(SpecTypeId.Length, opts)

    t = Transaction(document, "Set Length Accuracy {} mm".format(accuracy_mm))
    t.Start()
    document.SetUnits(units)
    t.Commit()


def sync_with_relinquish(document, path):
    is_detached = False
    try:
        is_detached = document.IsDetached
    except Exception:
        is_detached = False

    if is_detached:
        raise Exception(
            "Document is DETACHED. For RSN:// and shared centrals use "
            "Create New Local in Batch Processor (Detach = OFF)."
        )

    if not document.IsWorkshared:
        raise Exception(
            "Expected workshared document for sync path: {}".format(path)
        )

    try:
        central = ModelPathUtils.ConvertModelPathToUserVisiblePath(
            document.GetWorksharingCentralModelPath()
        )
        Output("Sync to central: {}".format(central))
    except Exception:
        Output("Sync to central...")

    sync_opts = SynchronizeWithCentralOptions()
    sync_opts.SetRelinquishOptions(make_relinquish_all_options())
    try:
        sync_opts.Comment = "Batch: set Length accuracy {} mm".format(
            TARGET_ACCURACY_MM
        )
    except Exception:
        pass
    try:
        sync_opts.SaveLocalFile = True
    except Exception:
        pass

    document.SynchronizeWithCentral(TransactWithCentralOptions(), sync_opts)
    Output("SynchronizedWithCentral + relinquish OK")
    return path


def save_or_sync(document, path):
    """
    RSN:// или workshared (не detach) -> Sync + relinquish.
    Обычный файл / detach на диске -> SaveAs in place.
    """
    is_detached = False
    try:
        is_detached = document.IsDetached
    except Exception:
        is_detached = False

    # Revit Server и любые живые локалы от хранилища — только Sync
    if is_rsn_path(path):
        Output("Path type: Revit Server (RSN)")
        return sync_with_relinquish(document, path)

    if (
        document.IsWorkshared
        and (not is_detached)
        and document.GetWorksharingCentralModelPath() is not None
    ):
        Output("Path type: workshared local/UNC")
        return sync_with_relinquish(document, path)

    # Detach / non-workshared на диске
    if is_detached and document.IsWorkshared:
        Output(
            "WARN: Detached workshared file — SaveAs in place. "
            "Prefer Create New Local for shared models."
        )

    clear_readonly_attribute(path)

    opts = SaveAsOptions()
    opts.OverwriteExistingFile = True
    opts.Compact = False

    if document.IsWorkshared:
        wopts = WorksharingSaveAsOptions()
        wopts.SaveAsCentral = True
        opts.SetWorksharingOptions(wopts)

    document.SaveAs(path, opts)
    Output("SavedAs in place: {}".format(path))

    relinquish_everything(document)

    try:
        document.Save()
        Output("Saved after relinquish.")
    except Exception as ex:
        Output("WARN save after relinquish: {}".format(ex))
        try:
            document.SaveAs(path, opts)
            Output("SaveAs after relinquish OK")
        except Exception as ex2:
            Output("WARN final SaveAs: {}".format(ex2))

    return path


try:
    before = doc.GetUnits().GetFormatOptions(SpecTypeId.Length)
    Output(
        "Before: unit={}, accuracy={}".format(
            before.GetUnitTypeId().TypeId, before.Accuracy
        )
    )

    set_length_accuracy(doc, TARGET_ACCURACY_MM)

    after = doc.GetUnits().GetFormatOptions(SpecTypeId.Length)
    Output(
        "After:  unit={}, accuracy={}".format(
            after.GetUnitTypeId().TypeId, after.Accuracy
        )
    )

    saved = save_or_sync(doc, revitFilePath)
    Output("Done: {}".format(saved))
    Output("OK")
except Exception as ex:
    Output("ERROR: {}".format(ex))
    raise
