Set sh = CreateObject("Wscript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")
tool = fso.GetParentFolderName(Wscript.ScriptFullName)
If Right(tool, 1) <> "\" Then tool = tool & "\"
ps1 = tool & "ui_compact.ps1"
cmd = "powershell.exe -NoProfile -ExecutionPolicy Bypass -STA -WindowStyle Hidden -File """ & ps1 & """"
sh.Run cmd, 0, False
