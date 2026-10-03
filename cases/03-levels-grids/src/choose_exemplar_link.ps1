# Name: choose_exemplar_link.ps1
# Version: 1.0
# What it does: Select one exemplar Revit model for levels/grids comparison.
# Inputs: choose_models.ps1 in Exemplar mode.
# Outputs: exemplar_model.cfg (and link cfg as documented in case README).
# How to run: choose_exemplar_link.cmd.
# Notes: Audit/apply match the link by file name/path.
$ToolDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$Picker = Join-Path $ToolDir "choose_models.ps1"

if (-not (Test-Path -LiteralPath $Picker)) {
    Write-Host "ERROR: choose_models.ps1 not found."
    exit 1
}

& $Picker -Purpose Exemplar
exit $LASTEXITCODE
