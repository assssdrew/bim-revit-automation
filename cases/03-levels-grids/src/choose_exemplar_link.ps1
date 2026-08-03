# Pick ONE exemplar Revit model (same UI as choose_models).
# Writes exemplar_model.cfg — audit/apply match the link by file name/path.

$ToolDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$Picker = Join-Path $ToolDir "choose_models.ps1"

if (-not (Test-Path -LiteralPath $Picker)) {
    Write-Host "ERROR: choose_models.ps1 not found."
    exit 1
}

& $Picker -Purpose Exemplar
exit $LASTEXITCODE
