Set fso = CreateObject("Scripting.FileSystemObject")
Set sh = CreateObject("WScript.Shell")
tool = fso.GetParentFolderName(Wscript.ScriptFullName)
If Right(tool, 1) <> "\" Then tool = tool & "\"
' Code lives in a subfolder (sluzhebnoe). Its Cyrillic name is not written here:
' a .vbs is read as ANSI, so the subfolder is found by its presets.ps1.
ps1 = ""
For Each d In fso.GetFolder(tool).SubFolders
  If fso.FileExists(d.Path & "\presets.ps1") Then ps1 = d.Path & "\presets.ps1"
Next
If ps1 = "" Then ps1 = tool & "presets.ps1"
If Not fso.FileExists(ps1) Then
  MsgBox "Cannot find presets.ps1" & vbCrLf & ps1, 16, "batch_model_ops"
  Wscript.Quit 1
End If
cmd = "powershell.exe -NoProfile -ExecutionPolicy Bypass -STA -WindowStyle Hidden -File """ & ps1 & """"
' 0 = hidden from the start: no console / Windows Terminal tab. The WinForms window still shows.
sh.CurrentDirectory = tool
sh.Run cmd, 0, False