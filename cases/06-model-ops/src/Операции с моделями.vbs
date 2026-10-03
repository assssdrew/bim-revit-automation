' Name: Операции с моделями.vbs
' Version: 1.0
' What it does: Primary hidden STA launcher for the model-ops presets window.
' Inputs: presets.ps1 beside script.
' Outputs: PowerShell presets UI.
' How to run: Double-click in src\.
' Notes: ShellExecute with hidden PowerShell.
Set fso = CreateObject("Scripting.FileSystemObject")
Set app = CreateObject("Shell.Application")
tool = fso.GetParentFolderName(Wscript.ScriptFullName)
If Right(tool, 1) <> "\" Then tool = tool & "\"
ps1 = tool & "presets.ps1"
If Not fso.FileExists(ps1) Then
  MsgBox "Cannot find presets.ps1" & vbCrLf & ps1, 16, "batch_model_ops"
  Wscript.Quit 1
End If
args = "-NoProfile -ExecutionPolicy Bypass -STA -WindowStyle Hidden -File """ & ps1 & """"
app.ShellExecute "powershell.exe", args, tool, "", 1