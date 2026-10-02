#
# Revit Batch Processor
#
# Copyright (c) 2020 Dan Rumery, BVN
# Source: https://github.com/bvn-architecture/RevitBatchProcessor (GPL-3.0).
# This file is a MODIFIED copy; see NOTICE and LICENSE-GPL-3.0.txt.
#
# Patched: OPEN_FAIL_PATCH_v2
# - never str() Cyrillic failure text (IronPython ascii)
# - always DeleteWarning / ResolveFailure even if logging fails
# - unresolvable Error => Continue (do not cancel Open)
#

import clr
import System

clr.AddReference("System.Core")
clr.ImportExtensions(System.Linq)

from System import EventHandler

clr.AddReference("RevitAPI")
from Autodesk.Revit.DB import *
from Autodesk.Revit.DB.Events import FailuresProcessingEventArgs

import global_test_mode
import exception_util

REVIT_WARNINGS_MESSAGE_HANDLER_PREFIX = "[ REVIT WARNINGS HANDLER ]"

# OPEN_FAIL_PATCH_v2
def _rbp_safe_text(val):
    if val is None:
        return ""
    try:
        if isinstance(val, str):
            return val
    except Exception:
        pass
    try:
        s = unicode(val)
        try:
            return s.encode("utf-8")
        except Exception:
            return s.encode("ascii", "replace")
    except Exception:
        return "[unprintable]"

def ElementIdsToSemicolonDelimitedText(elementIds):
    values = []
    for elementId in elementIds:
        if hasattr(elementId, "Value"):
            values.append(str(elementId.Value))
        else:
            values.append(str(elementId.IntegerValue))
    return "; ".join(values)

def ReportFailureWarning(failure, failureDefinition, output):
    failureSeverity = failure.GetSeverity()
    output()
    output(
        "\t" +
        _rbp_safe_text(failureSeverity) +
        " - " +
        _rbp_safe_text(failure.GetDescriptionText()) +
        " - " +
        "(" + "GUID: " + str(failure.GetFailureDefinitionId().Guid) + ")"
    )

    if failureSeverity == FailureSeverity.Error or failureSeverity == FailureSeverity.Warning:
        failingElementIds = failure.GetFailingElementIds()
        if len(failingElementIds) > 0:
            output()
            output("\t" + "Failing element ids: " + ElementIdsToSemicolonDelimitedText(failingElementIds))
        additionalElementIds = failure.GetAdditionalElementIds()
        if len(additionalElementIds) > 0:
            output()
            output("\t" + "Additional element ids: " + ElementIdsToSemicolonDelimitedText(additionalElementIds))

    if failureSeverity == FailureSeverity.Error:
        if failure.HasResolutions():
            output()
            output("\t" + "Applicable resolution types:")
            output()
            defaultResolutionType = failureDefinition.GetDefaultResolutionType()
            for resolutionType in failureDefinition.GetApplicableResolutionTypes():
                caption = _rbp_safe_text(failureDefinition.GetResolutionCaption(resolutionType))
                output(
                    "\t\t" +
                    str(resolutionType) +
                    (" (Default)" if (resolutionType == defaultResolutionType) else str.Empty) +
                    " - " +
                    "'" + caption + "'"
                )
        else:
            output()
            output("\t" + "WARNING: no resolutions available")

    return

def ProcessFailures(failuresAccessor, output, rollBackOnWarning=False):
    try:
        result = FailureProcessingResult.Continue
        doc = failuresAccessor.GetDocument()
        app = doc.Application
        failureReg = app.GetFailureDefinitionRegistry()
        failures = failuresAccessor.GetFailureMessages()
        if failures.Any():
            output()
            output("Processing Revit document warnings / failures (" + str(failures.Count) + "):")
            for failure in failures:
                failureDefinition = failureReg.FindFailureDefinition(failure.GetFailureDefinitionId())
                try:
                    ReportFailureWarning(failure, failureDefinition, output)
                except Exception:
                    try:
                        output("\t[failure description skipped]")
                    except Exception:
                        pass
                failureSeverity = failure.GetSeverity()
                if failureSeverity == FailureSeverity.Warning and not rollBackOnWarning:
                    failuresAccessor.DeleteWarning(failure)
                elif (
                    failureSeverity == FailureSeverity.Error
                    and
                    failure.HasResolutions()
                    and
                    result != FailureProcessingResult.ProceedWithRollBack
                    and
                    not rollBackOnWarning
                ):
                    if failure.HasResolutionOfType(FailureResolutionType.UnlockConstraints):
                        failure.SetCurrentResolutionType(FailureResolutionType.UnlockConstraints)
                    elif failureDefinition.IsResolutionApplicable(FailureResolutionType.UnlockConstraints):
                        output()
                        output("\t" + "WARNING: UnlockConstraints is not a valid resolution for this failure despite the definition reporting that it is an applicable resolution!")
                    elif failure.HasResolutionOfType(FailureResolutionType.DetachElements):
                        failure.SetCurrentResolutionType(FailureResolutionType.DetachElements)
                    elif failure.HasResolutionOfType(FailureResolutionType.SkipElements):
                        failure.SetCurrentResolutionType(FailureResolutionType.SkipElements)
                    output()
                    output("\t" + "Attempting to resolve error using resolution: " + str(failure.GetCurrentResolutionType()))
                    failuresAccessor.ResolveFailure(failure)
                    result = FailureProcessingResult.ProceedWithCommit
                else:
                    # OPEN_FAIL_PATCH_v2 do not cancel document Open
                    try:
                        output()
                        output("\t" + "WARNING: unresolvable error - continuing open instead of rollback")
                    except Exception:
                        pass
                    result = FailureProcessingResult.Continue
        else:
            result = FailureProcessingResult.Continue
    except Exception, e:
        output()
        output("ERROR: the failure handler generated an error!")
        exception_util.LogOutputErrorDetails(e, output)
        result = FailureProcessingResult.Continue
    return result

class FailuresPreprocessor(IFailuresPreprocessor):
    def __init__(self, output):
        self.output = output
        return

    def PreprocessFailures(self, failuresAccessor):
        result = ProcessFailures(failuresAccessor, self.output)
        return result

def SetTransactionFailureOptions(transaction, output):
    failureOptions = transaction.GetFailureHandlingOptions()
    failureOptions.SetForcedModalHandling(True)
    failureOptions.SetClearAfterRollback(True)
    failureOptions.SetFailuresPreprocessor(FailuresPreprocessor(output))
    transaction.SetFailureHandlingOptions(failureOptions)
    return

def SetFailuresAccessorFailureOptions(failuresAccessor):
    failureOptions = failuresAccessor.GetFailureHandlingOptions()
    failureOptions.SetForcedModalHandling(True)
    failureOptions.SetClearAfterRollback(True)
    failuresAccessor.SetFailureHandlingOptions(failureOptions)
    return

def FailuresProcessingEventHandler(sender, args, output):
    app = sender
    failuresAccessor = args.GetFailuresAccessor()
    SetFailuresAccessorFailureOptions(failuresAccessor)
    result = ProcessFailures(failuresAccessor, output)
    args.SetProcessingResult(result)
    return

def WithFailuresProcessingHandler(app, action, output_):
    output = global_test_mode.PrefixedOutputForGlobalTestMode(output_, REVIT_WARNINGS_MESSAGE_HANDLER_PREFIX)
    result = None
    failuresProcessingEventHandler = (
        EventHandler[FailuresProcessingEventArgs](
            lambda sender, args: FailuresProcessingEventHandler(sender, args, output)
        )
    )
    app.FailuresProcessing += failuresProcessingEventHandler
    try:
        result = action()
    finally:
        app.FailuresProcessing -= failuresProcessingEventHandler
    return result
