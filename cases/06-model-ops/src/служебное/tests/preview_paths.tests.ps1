# Name: preview_paths.tests.ps1
# Version: 3.1.0
# What it does: See the case README.
# Notes: Paths and project names in this portfolio copy are placeholders.

# New-JobPreviewRows: "Новый путь" is always a full absolute path (UNC, disk or RSN://) - for every
# placement mode (same / segment / prefix), the buffer and the Revit year change - or the row carries
# a clear path_error and the window refuses to start.
#   powershell -NoProfile -ExecutionPolicy Bypass -File tests\preview_paths.tests.ps1

$ToolSrc = Split-Path -Parent $PSScriptRoot
. (Join-Path $ToolSrc "job_lib.ps1")
. (Join-Path $ToolSrc "ui_logic.ps1")

$script:Fail = 0
$script:Pass = 0
function Check([bool]$Cond, [string]$Msg) {
    if ($Cond) { $script:Pass++ }
    else { $script:Fail++; Write-Host "  FAIL: $Msg" -ForegroundColor Red }
}
function Eq($Got, $Want, [string]$Msg) { Check ($Got -ceq $Want) ("{0}: got '{1}', want '{2}'" -f $Msg, $Got, $Want) }

# the default buffer folder is <profile>\Documents\doc\script\rbp\upgrade_buffer; a fixed one keeps the test the same on any OS
function Get-DefaultBufferDir { return 'D:\Users\t\Documents\doc\script\rbp\upgrade_buffer' }

$proj = -join ([char[]](0x41F, 0x440, 0x43E, 0x435, 0x43A, 0x442))      # Проект
$UNC = "\\fileserver.example.local\public\$proj\04\PROJ_IOS_EOM_ORG_K01_R24.rvt"
$UNCFolder = "\\fileserver.example.local\public\$proj\04"
$DISK = 'D:\Models\04\PROJ_IOS_EOM_ORG_K01_R24.rvt'
$RSN = "RSN://revit-server-2022.example.local/proj $proj/04/PROJ_IOS_EOM_ORG_K01_R24.rvt"

function Row([string]$Old, [string]$Mode = "same", [string]$OldSeg = "", [string]$NewSeg = "", [string]$OldPref = "", [string]$NewPref = "",
    [string]$Src = "", [string]$Dst = "", [string]$Buf = "", $Ops = @(), [bool]$Rename = $false, $Manual = @{}) {
    $r = @(New-JobPreviewRows -OldPaths @($Old) -Ops $Ops -DoRename $Rename -LocationMode $Mode -OldSegment $OldSeg -NewSegment $NewSeg `
            -OldPrefix $OldPref -NewPrefix $NewPref -SourceYear $Src -TargetYear $Dst -BufferDir $Buf -NameOverrides @{} -ManualRows $Manual)
    return $r[0]
}
function Ok($Row, [string]$Msg)  { Check ($Row.path_error -eq "" -and (Test-NewModelPath $Row.new_path).Ok) ("{0}: expected a full path, got [{1}] error [{2}]" -f $Msg, $Row.new_path, $Row.path_error) }
function Bad($Row, [string]$Msg, [string]$Needle) {
    Check ($Row.path_error -ne "" -and $Row.path_error.Contains($Needle)) ("{0}: expected an error containing '{1}', got [{2}] for path [{3}]" -f $Msg, $Needle, $Row.path_error, $Row.new_path)
}

# ------------------------------------------------------------------ same place
foreach ($p in @($UNC, $DISK, $RSN)) { $r = Row $p; Ok $r "same place $p"; Eq $r.new_path $p "same place keeps the path" }

# ------------------------------------------------------------ "Поменять кусок пути"
$r = Row $UNC "segment" "$proj\04" "$proj\05";            Ok $r "segment: one folder level"; Eq $r.new_path "\\fileserver.example.local\public\$proj\05\PROJ_IOS_EOM_ORG_K01_R24.rvt" "segment: result"
$r = Row $UNC "segment" $UNCFolder "D:\New";              Ok $r "segment: whole folder -> disk"; Eq $r.new_path "D:\New\PROJ_IOS_EOM_ORG_K01_R24.rvt" "segment: whole folder -> disk, result"
$r = Row $UNC "segment" "\public\$proj" "";               Ok $r "segment: remove a middle level (the rest is still a share path)"
$r = Row $UNC "segment" $UNCFolder "";                    Bad $r "segment: whole folder, 'стало' empty" "Новый путь не полный"
Check ($r.path_error.Contains("«куска пути»") -and $r.path_error.Contains($UNCFolder)) "segment: the message names the 'кусок пути' and what was typed"
Eq $r.new_path "\PROJ_IOS_EOM_ORG_K01_R24.rvt" "segment: whole folder, 'стало' empty -> bare name (the reported case)"
$r = Row $UNC "segment" "\\fileserver.example.local\public" "";      Bad $r "segment: share removed" "Новый путь не полный"
$r = Row $UNC "segment" $UNCFolder "x";                   Bad $r "segment: relative 'стало'" "Новый путь не полный"
$r = Row $UNC "segment" "\\nowhere\x" "\\other";          Ok $r "segment: 'было' not found -> path unchanged"; Eq $r.new_path $UNC "segment: not found keeps the path"
$r = Row $UNC "segment" "" "\\other";                     Ok $r "segment: empty 'было' -> path unchanged"
$r = Row $RSN "segment" "/04/" "/05/";                    Ok $r "segment on RSN"
$r = Row $RSN "segment" "RSN://revit-server-2022.example.local/proj $proj/04" "";  Bad $r "segment on RSN: folder removed" "Новый путь"
$r = Row $DISK "segment" "D:\Models" "E:\Archive";        Ok $r "segment on a disk path"
$r = Row $DISK "segment" "D:\Models\04" "";               Bad $r "segment on a disk path: folder removed" "Новый путь не полный"
$r = Row $UNC "segment" $UNCFolder "\\srv2\share\a b\c";  Ok $r "segment: spaces in the new folder"

# -------------------------------------------- "Другой сервер или корень шары" (prefix)
$r = Row $UNC "prefix" "" "" "\\fileserver.example.local\public" "\\srv2\share";   Ok $r "prefix: UNC -> UNC"; Eq $r.new_path "\\srv2\share\$proj\04\PROJ_IOS_EOM_ORG_K01_R24.rvt" "prefix: result"
$r = Row $UNC "prefix" "" "" "\\fileserver.example.local\public" "";               Bad $r "prefix: new prefix empty" "Новый путь не полный"
Check ($r.path_error.Contains("начала пути")) "prefix: the message names the start of the path"
$r = Row $UNC "prefix" "" "" "\\fileserver.example.local\public" "D:\mirror";      Ok $r "prefix: UNC -> disk"
$r = Row $UNC "prefix" "" "" "\\fileserver.example.local\public" "mirror";         Bad $r "prefix: relative new prefix" "Новый путь"
$r = Row $UNC "prefix" "" "" "\\other\share" "\\srv2\share";            Ok $r "prefix: old prefix does not match -> unchanged"; Eq $r.new_path $UNC "prefix: no match keeps the path"
$r = Row $UNC "prefix" "" "" "" "\\srv2\share";                         Ok $r "prefix: empty old prefix -> unchanged"
$r = Row $RSN "prefix" "" "" "RSN://revit-server-2022.example.local" "RSN://revit-server-2024.example.local"; Ok $r "prefix: RSN host change"; Eq $r.new_path "RSN://revit-server-2024.example.local/proj $proj/04/PROJ_IOS_EOM_ORG_K01_R24.rvt" "prefix: RSN host change, result"
$r = Row $RSN "prefix" "" "" "RSN://revit-server-2022.example.local" "";                   Bad $r "prefix: RSN host removed" "Новый путь не полный"
$r = Row $RSN "prefix" "" "" "RSN://revit-server-2022.example.local" "\\srv\share";        Ok $r "prefix: RSN -> UNC"
$r = Row $UNC "prefix" "" "" "\\fileserver.example.local\public" "RSN://h/p";      Ok $r "prefix: UNC -> RSN"
$r = Row $DISK "prefix" "" "" "D:\Models" "E:\Archive";                 Ok $r "prefix: disk -> disk"
$r = Row $DISK "prefix" "" "" "D:\Models" "";                           Bad $r "prefix: disk prefix removed" "Новый путь не полный"

# ------------------------------------------------ year change and the buffer (RSN only)
$r = Row $RSN "same" "" "" "" "" "2022" "2024" ""
Ok $r "year change, default buffer: new path"
Check ($r.needs_buffer -and $r.staging_path -ceq 'D:\Users\t\Documents\doc\script\rbp\upgrade_buffer\PROJ_IOS_EOM_ORG_K01_R24.rvt') "year change, default buffer: staging is [$($r.staging_path)]"
$r = Row $RSN "same" "" "" "" "" "2022" "2024" 'D:\buf';        Check ($r.path_error -eq "" -and $r.staging_path -ceq 'D:\buf\PROJ_IOS_EOM_ORG_K01_R24.rvt') "buffer on a disk: [$($r.staging_path)]"
$r = Row $RSN "same" "" "" "" "" "2022" "2024" 'D:\buf\';       Check ($r.path_error -eq "" -and $r.staging_path -ceq 'D:\buf\PROJ_IOS_EOM_ORG_K01_R24.rvt') "buffer with a trailing slash: [$($r.staging_path)]"
$r = Row $RSN "same" "" "" "" "" "2022" "2024" '\\pc\buf';      Check ($r.path_error -eq "") "buffer on a share"
$r = Row $RSN "same" "" "" "" "" "2022" "2024" 'buf';           Bad $r "relative buffer folder" "Путь в буфере не полный"
$r = Row $RSN "same" "" "" "" "" "2022" "2024" '\buf';          Bad $r "buffer folder without a drive" "Путь в буфере не полный"
$r = Row $RSN "same" "" "" "" "" "2022" "2024" 'D:';            Check ($r.path_error -eq "" -and $r.staging_path -ceq 'D:\PROJ_IOS_EOM_ORG_K01_R24.rvt') "buffer = the root of a disk is a valid full path"
$r = Row $RSN "prefix" "" "" "RSN://revit-server-2022.example.local" "RSN://revit-server-2024.example.local" "2022" "2024" 'D:\buf'
Check ($r.path_error -eq "" -and $r.new_path.StartsWith("RSN://revit-server-2024.example.local/")) "year change + prefix + disk buffer"
$r = Row $RSN "prefix" "" "" "RSN://revit-server-2022.example.local" "RSN://revit-server-2024.example.local" "2022" "2024" 'buf';  Bad $r "year change + prefix + relative buffer" "Путь в буфере не полный"
$r = Row $RSN "prefix" "" "" "RSN://revit-server-2022.example.local" "" "2022" "2024" 'buf';                     Bad $r "both broken: the new path is reported first" "Новый путь не полный"
$r = Row $UNC "same" "" "" "" "" "2022" "2024" 'buf';            Check ($r.path_error -eq "" -and $r.staging_path -eq "") "UNC model: no buffer needed, a relative buffer folder is not an error"
$r = Row $RSN "same" "" "" "" "" "2024" "2024" 'buf';            Check ($r.path_error -eq "" -and -not $r.needs_buffer) "same year: no buffer"

# ---------------------------------------------------------------- old path, names, manual
$r = Row '\PROJ_IOS_EOM_ORG_K01_R24.rvt';                       Bad $r "old path without a folder" "Старый путь не полный"
$r = Row 'PROJ_IOS_EOM_ORG_K01_R24.rvt';                        Bad $r "old path is a bare file name" "Старый путь не полный"
$ops = @(New-OpsFromSample 'PROJ_IOS_EOM_ORG_K01_R24' 'PROJ_IOS_EOM_KLWT_K01_R24')
$r = Row $UNC "same" "" "" "" "" "" "" "" $ops $true;             Ok $r "rename"; Eq $r.new_name "PROJ_IOS_EOM_KLWT_K01_R24" "rename: name"; Eq $r.new_path "$UNCFolder\PROJ_IOS_EOM_KLWT_K01_R24.rvt" "rename: path keeps the folder"
$r = Row $UNC "segment" "$proj\04" "$proj\05" "" "" "" "" "" $ops $true;  Ok $r "rename + segment"; Eq $r.new_path "\\fileserver.example.local\public\$proj\05\PROJ_IOS_EOM_KLWT_K01_R24.rvt" "rename + segment: result"
$bad = @(New-OpsFromSample 'A_ORG_B' 'A_TR?ST_B')
$r = Row "\\s\p\PROJ_ORG_B.rvt" "same" "" "" "" "" "" "" "" $bad $true;  Bad $r "a rule that makes an invalid file name" "недопустим"
$m = @{}; $m[(Get-ManualRowKey $UNC)] = @{ name = $null; path = 'relative\x.rvt' }
$r = Row $UNC "same" "" "" "" "" "" "" "" @() $false $m;          Bad $r "manual path that is not a full path" "Новый путь не полный"
Check (-not $r.path_error.Contains("куска пути")) "manual path: no placement-mode hint"
$m = @{}; $m[(Get-ManualRowKey $UNC)] = @{ name = $null; path = 'D:\Own\x.rvt' }
$r = Row $UNC "segment" $UNCFolder "" "" "" "" "" "" @() $false $m;   Ok $r "manual full path beats a broken placement rule"; Eq $r.manual "path" "marked as manual"

# ------------------------------------------------------------------ messages for the window
$rows = @(
    (Row $UNC "segment" $UNCFolder ""),
    (Row $UNC "same"),
    (Row $RSN "prefix" "" "" "RSN://revit-server-2022.example.local" "")
)
$lines = @(Get-PathErrorLines $rows 5)
Check ($lines.Count -eq 2) "Get-PathErrorLines: only the broken rows (got $($lines.Count))"
Check ($lines[0].StartsWith("PROJ_IOS_EOM_ORG_K01_R24: Новый путь не полный")) "Get-PathErrorLines: '<model>: <reason>'"
Eq @(Get-PathErrorLines $rows 1).Count 1 "Get-PathErrorLines: honours the limit"
Eq @(Get-PathErrorLines @(Row $UNC "same") 5).Count 0 "Get-PathErrorLines: nothing for good rows"
Eq @(Get-PathErrorLines @() 5).Count 0 "Get-PathErrorLines: no rows"
Check ((Get-CellTipText $rows[0] 2).Contains("Новый путь не полный")) "tooltip of 'Состояние' shows the reason"
Check ((Get-CellTipText $rows[0] 5).Contains("ОШИБКА:")) "tooltip of 'Новый путь' shows the reason"
Eq (Get-CellTipText $rows[1] 2) "" "no tooltip on 'Состояние' for a good row"
Check (-not (Get-CellTipText $rows[1] 5).Contains("ОШИБКА")) "no error in the tooltip of a good row"
$o = Row '\x.rvt'
Check ((Get-CellTipText $o 4).Contains("Старый путь не полный")) "tooltip of 'Старый путь' shows the old-path reason"

Write-Host ""
Write-Host ("Checks passed: {0}, failed: {1}" -f $script:Pass, $script:Fail)
if ($script:Fail -gt 0) { exit 1 }
exit 0
