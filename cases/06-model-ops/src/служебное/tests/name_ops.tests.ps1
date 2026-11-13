# Name: name_ops.tests.ps1
# Version: 3.1.0
# What it does: See the case README.
# Notes: Paths and project names in this portfolio copy are placeholders.

# Tests for New-OpsFromSample / Invoke-NameOps (job_lib.ps1).
# No dependencies, Windows PowerShell 5.1 or 7:
#   powershell -NoProfile -ExecutionPolicy Bypass -File tests\name_ops.tests.ps1

$ToolSrc = Split-Path -Parent $PSScriptRoot
. (Join-Path $ToolSrc "job_lib.ps1")

$script:Fail = 0
$script:Pass = 0
function Check([bool]$Cond, [string]$Msg) {
    if ($Cond) { $script:Pass++ }
    else { $script:Fail++; Write-Host "  FAIL: $Msg" -ForegroundColor Red }
}

function Apply-Sample([string]$From, [string]$To, [string]$Target) {
    $ops = @(New-OpsFromSample $From $To)
    return (Invoke-NameOps -Name $Target -Ops $ops)
}

# The sample itself must be reproduced exactly.
$pairs = @(
    @("PROJ_IOS_EOM_ORG_K01_R24", "PROJ_IOS_EOM_KLWT_K01_R24"),
    @("A_B_C_D", "A_B_X_D"),
    @("A_B_C_D", "A_C_D"),
    @("A_B_D", "A_B_C_D"),
    @("A_OLD_B_R24", "A_NEW_B_R24"),
    @("PROJ_IOS_EOM_ORG_01_K01_R24", "PROJ_IOS_EOM_ORG_K01_R24"),
    @("A_B_C", "X_Y_Z"),
    @("A_B_C_D", "A_X_D"),
    @("A_D", "A_B_C_D"),
    @("A_B_C_D", "A_D"),
    @("same_name", "same_name")
)
foreach ($p in $pairs) {
    $got = Apply-Sample $p[0] $p[1] $p[0]
    Check ($got -ceq $p[1]) ("sample {0} -> {1} reproduced, got {2}" -f $p[0], $p[1], $got)
}

# The reported case: ORG -> KLWT must stay in place, not move to the end.
$ex = Apply-Sample "PROJ_IOS_EOM_ORG_K01_R24" "PROJ_IOS_EOM_KLWT_K01_R24" "PROJ_IOS_EOM_ORG_K01_R24"
Check ($ex -ceq "PROJ_IOS_EOM_KLWT_K01_R24") "example: $ex"
$ops = @(New-OpsFromSample "PROJ_IOS_EOM_ORG_K01_R24" "PROJ_IOS_EOM_KLWT_K01_R24")
Check ($ops.Count -eq 1 -and $ops[0].kind -eq "replace" -and $ops[0].from -eq "ORG" -and $ops[0].to -eq "KLWT") "single replace ORG->KLWT"

# Same rule applied to other models, including compound K-parts.
$sample = @("PROJ_IOS_EOM_ORG_K01_R24", "PROJ_IOS_EOM_KLWT_K01_R24")
$targets = @(
    @("PROJ_IOS_EOM_ORG_K03-K05-K07-K10-ST_R24", "PROJ_IOS_EOM_KLWT_K03-K05-K07-K10-ST_R24"),
    @("PROJ_IOS_EOM_ORG_K03-K05-K07-K10-ST1-ST2_R24", "PROJ_IOS_EOM_KLWT_K03-K05-K07-K10-ST1-ST2_R24"),
    @("PROJ_IOS_EOM_ORG_K02_R24", "PROJ_IOS_EOM_KLWT_K02_R24"),
    @("proj_ios_eom_org_k04_r24", "proj_ios_eom_KLWT_k04_r24"),
    @("PROJ_IOS_EOM_KLWT_K09_R24", "PROJ_IOS_EOM_KLWT_K09_R24"),
    @("OTHER_NAME_R24", "OTHER_NAME_R24")
)
foreach ($t in $targets) {
    $got = Apply-Sample $sample[0] $sample[1] $t[0]
    Check ($got -ceq $t[1]) ("{0} -> {1}, got {2}" -f $t[0], $t[1], $got)
}

# Insert keeps its place after the previous token of the NEW name.
$got = Apply-Sample "A_B_D" "A_B_C_D" "X_B_D"
Check ($got -ceq "X_B_C_D") "insert after previous token: $got"
# Replace followed by an insert must not anchor on a replaced-away token.
$got = Apply-Sample "A_OLD_D" "A_NEW_C_D" "A_OLD_D"
Check ($got -ceq "A_NEW_C_D") "replace + insert: $got"

# Randomised edits (replace / delete / insert, no reordering, unique tokens):
# the rule built from the sample must reproduce the sample exactly.
$rnd = New-Object System.Random 20260101
$bad = 0
for ($n = 0; $n -lt 400; $n++) {
    $len = $rnd.Next(2, 8)
    $a = @(for ($k = 0; $k -lt $len; $k++) { "T{0}x{1}" -f $n, $k })
    $b = New-Object System.Collections.Generic.List[string]
    $fresh = 0
    foreach ($t in $a) {
        switch ($rnd.Next(0, 5)) {
            0 { }                                                             # delete
            1 { $b.Add(("N{0}y{1}" -f $n, $fresh)); $fresh++ }                 # replace
            2 { $b.Add($t); $b.Add(("N{0}y{1}" -f $n, $fresh)); $fresh++ }     # keep + insert after
            default { $b.Add($t) }                                            # keep
        }
    }
    if ($b.Count -eq 0) { continue }
    $from = $a -join "_"
    $to = $b -join "_"
    $got = Apply-Sample $from $to $from
    if ($got -cne $to) { $bad++; if ($bad -le 3) { Write-Host "  random miss: $from -> $to, got $got" } }
}
Check ($bad -eq 0) "randomised samples reproduced (misses: $bad)"

Write-Host ""
Write-Host ("Checks passed: {0}, failed: {1}" -f $script:Pass, $script:Fail)
if ($script:Fail -gt 0) { exit 1 }
exit 0
