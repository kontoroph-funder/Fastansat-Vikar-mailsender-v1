param([switch]$SkipBuild)
$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
Set-Location -LiteralPath $root
$python = Join-Path $root '.venv\Scripts\python.exe'
if (-not $SkipBuild) {
    & $python -m PyInstaller FastansatVikar.spec --noconfirm
    if ($LASTEXITCODE -ne 0) { throw 'Windows-build fejlede.' }
}
& $python scripts\package.py
if ($LASTEXITCODE -ne 0) { throw 'Pakning fejlede.' }
