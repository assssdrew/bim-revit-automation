' Name: Сжатие.vbs
' Version: 1.0
' What it does: Hidden launcher for the compact-save WinForms UI (no console).
' Inputs: ui_compact.ps1 beside this script.
' Outputs: Starts PowerShell STA hidden.
' How to run: Double-click Сжатие.vbs in src\.
' Notes: Primary entry point for operators.
Set sh = CreateObject("Wscript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")
tool = fso.GetParentFolderName(Wscript.ScriptFullName)
If Right(tool, 1) <> "\" Then tool = tool & "\"
ps1 = tool & "ui_compact.ps1"
cmd = "powershell.exe -NoProfile -ExecutionPolicy Bypass -STA -WindowStyle Hidden -File """ & ps1 & """"
sh.Run cmd, 0, False
