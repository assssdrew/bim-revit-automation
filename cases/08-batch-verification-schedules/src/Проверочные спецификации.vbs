Set sh = CreateObject("Wscript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")
' All paths are derived from the location of this file: any folder works, including UNC (\\server\folder), Cyrillic names and spaces.
' ui_schedules.ps1 lies in a subfolder of the service folder (two levels down); folder names are not hard-coded here (this file must stay ASCII).
here = fso.GetParentFolderName(Wscript.ScriptFullName)
ps1 = ""
If fso.FileExists(here & "\ui_schedules.ps1") Then ps1 = here & "\ui_schedules.ps1"
If ps1 = "" Then
    For Each f1 In fso.GetFolder(here).SubFolders
        If fso.FileExists(f1.Path & "\ui_schedules.ps1") Then
            ps1 = f1.Path & "\ui_schedules.ps1"
        Else
            For Each f2 In f1.SubFolders
                If ps1 = "" Then
                    If fso.FileExists(f2.Path & "\ui_schedules.ps1") Then ps1 = f2.Path & "\ui_schedules.ps1"
                End If
            Next
        End If
        If ps1 <> "" Then Exit For
    Next
End If
If ps1 = "" Then
    MsgBox "ui_schedules.ps1 not found in the service folder next to this file: " & here, 16, "Verification schedules"
    Wscript.Quit 2
End If
cmd = "powershell.exe -NoProfile -ExecutionPolicy Bypass -STA -WindowStyle Hidden -File """ & ps1 & """"
sh.Run cmd, 0, False
