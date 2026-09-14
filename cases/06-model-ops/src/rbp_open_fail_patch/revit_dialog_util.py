#
# Revit Batch Processor
#
# Copyright (c) 2020 Dan Rumery, BVN
#
# Patched: OPEN_FAIL_PATCH_v2
# Dismiss dialog first, then log (Cyrillic captions must not block OverrideResult).
#

import clr
import System

from System import EventHandler
from System.Text import StringBuilder

clr.AddReference("RevitAPIUI")
from Autodesk.Revit.UI.Events import DialogBoxShowingEventArgs, TaskDialogShowingEventArgs, MessageBoxShowingEventArgs

import exception_util

IDOK = 1
IDCANCEL = 2
IDYES = 6
IDNO = 7
IDCLOSE = 8

def Try(action):
    result = None
    try:
        result = action()
    except:
        pass
    return result

def DialogShowingEventHandler(sender, eventArgs, output):
    try:
        dialogResult = IDOK
        if isinstance(eventArgs, TaskDialogShowingEventArgs):
            if eventArgs.DialogId == "TaskDialog_Missing_Third_Party_Updater":
                dialogResult = 1001
            elif eventArgs.DialogId == "TaskDialog_Location_Position_Changed":
                dialogResult = 1002
        eventArgs.OverrideResult(dialogResult)  # OPEN_FAIL_PATCH_v2 dismiss first
        try:
            msg = StringBuilder()
            msg.AppendLine()
            msg.AppendLine("Dialog box shown:")
            msg.AppendLine()
            if isinstance(eventArgs, TaskDialogShowingEventArgs):
                msg.AppendLine("\tMessage: " + str(eventArgs.Message))
            elif isinstance(eventArgs, MessageBoxShowingEventArgs):
                msg.AppendLine("\tMessage: " + str(eventArgs.Message))
                msg.AppendLine("\tDialogType: " + str(eventArgs.DialogType))
            dialogId = Try(lambda: eventArgs.DialogId)
            if dialogId is not None:
                msg.AppendLine("\tDialogId: " + str(dialogId))
            helpId = Try(lambda: eventArgs.HelpId)
            if helpId is not None:
                msg.AppendLine("\tHelpId: " + str(helpId))
            output(msg.ToString())
        except Exception:
            pass
    except Exception, e:
        try:
            errorMsg = StringBuilder()
            errorMsg.AppendLine()
            errorMsg.AppendLine("Caught exception in dialog event handler!")
            errorMsg.AppendLine("Exception message: " + e.message)
            output(errorMsg.ToString())
            exception_util.LogOutputErrorDetails(e, output)
        except Exception:
            pass
    return

def WithDialogBoxShowingHandler(uiapp, action, output):
    result = None
    dialogShowingEventHandler = (
        EventHandler[DialogBoxShowingEventArgs](
            lambda sender, eventArgs: DialogShowingEventHandler(sender, eventArgs, output)
        )
    )
    uiapp.DialogBoxShowing += dialogShowingEventHandler
    try:
        result = action()
    finally:
        uiapp.DialogBoxShowing -= dialogShowingEventHandler
    return result
