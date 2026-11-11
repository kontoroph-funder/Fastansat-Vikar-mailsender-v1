$ErrorActionPreference = 'Stop'
$installRoot = Join-Path $env:LOCALAPPDATA 'FastansatVikar'
Write-Host 'Installerer Fastansatte & Vikarer. Python og appens biblioteker er inkluderet.'
Write-Host 'Manglende LibreOffice hentes via WinGet. Windows kan bede om administratorgodkendelse.'
$appRoot = Join-Path $installRoot 'application'
$bundle = Join-Path $PSScriptRoot 'FastansatVikar'
$exe = Join-Path $appRoot 'FastansatVikar.exe'
if (-not (Test-Path -LiteralPath (Join-Path $bundle 'FastansatVikar.exe'))) { throw 'Pak hele ZIP-filen ud, før du starter installationen.' }
if (Get-Process -Name FastansatVikar -ErrorAction SilentlyContinue) { throw 'Luk Fastansatte & Vikarer før installation eller opdatering. Brug Afslut appen i brugerfladen.' }
New-Item -ItemType Directory -Force -Path $appRoot | Out-Null
Copy-Item -Path (Join-Path $bundle '*') -Destination $appRoot -Recurse -Force
$shell = New-Object -ComObject WScript.Shell
foreach ($shortcutFolder in @([Environment]::GetFolderPath('Desktop'), [Environment]::GetFolderPath('Programs'))) {
    $shortcut = $shell.CreateShortcut((Join-Path $shortcutFolder 'Fastansatte og Vikarer.lnk'))
    $shortcut.TargetPath = $exe
    $shortcut.WorkingDirectory = $appRoot
    $shortcut.Description = 'Timeoversigter til fastansatte og vikarer'
    $shortcut.Save()
}
$reportPath = Join-Path $installRoot 'runtime-check.json'
function Get-RuntimeReport {
    if (Test-Path -LiteralPath $reportPath) { Remove-Item -LiteralPath $reportPath -Force }
    $process = Start-Process -FilePath $exe -ArgumentList @('--check-runtime', ('"' + $reportPath + '"')) -WindowStyle Hidden -Wait -PassThru
    if ($process.ExitCode -ne 0 -or -not (Test-Path -LiteralPath $reportPath)) {
        throw 'PDF-programmer kunne ikke kontrolleres. Brug kontrollen under Indstillinger i appen.'
    }
    return Get-Content -LiteralPath $reportPath -Raw -Encoding UTF8 | ConvertFrom-Json
}
$dependenciesReady = $false
try {
    $report = Get-RuntimeReport
    if (-not $report.libreoffice) {
        $winget = Get-Command winget.exe -ErrorAction SilentlyContinue
        if ($winget) {
            Write-Host 'Henter og installerer LibreOffice med Python/UNO. Det kan tage flere minutter.'
            & $winget.Source install --id TheDocumentFoundation.LibreOffice --exact --source winget --architecture x64 --silent --accept-package-agreements --accept-source-agreements --disable-interactivity
            $wingetExit = $LASTEXITCODE
            if ($wingetExit -ne 0) { Write-Warning "WinGet returnerede fejlkode $wingetExit. Kontrollerer installationen igen." }
            $report = Get-RuntimeReport
        } else {
            Write-Warning 'WinGet mangler. Installer App Installer fra Microsoft Store, og kør install.cmd igen. Alternativt: https://www.libreoffice.org/download/download-libreoffice/'
        }
    }
    if (-not $report.excel) { Write-Warning 'Microsoft Excel mangler. Installer og aktivér desktop-Excel med din Microsoft 365-/Office-licens for at bruge Vikarer. Excel Online er ikke tilstrækkeligt.' }
    if (-not $report.libreoffice) { Write-Warning ('LibreOffice/UNO er ikke klar: ' + $report.libreoffice_error + '. Ved en brugerdefineret installation: angiv programstien under Indstillinger i appen.') }
    $dependenciesReady = $report.excel -and $report.libreoffice
} catch { Write-Warning $_.Exception.Message }
try {
    $action = New-ScheduledTaskAction -Execute $exe -Argument '--cleanup' -WorkingDirectory $appRoot
    $trigger = New-ScheduledTaskTrigger -Once -At (Get-Date).AddMinutes(5) -RepetitionInterval (New-TimeSpan -Hours 1)
    $taskSettings = New-ScheduledTaskSettingsSet -StartWhenAvailable -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries
    $principal = New-ScheduledTaskPrincipal -UserId ([Security.Principal.WindowsIdentity]::GetCurrent().Name) -LogonType Interactive -RunLevel Limited
    $taskName = 'FastansatVikar-Cleanup-' + [Security.Principal.WindowsIdentity]::GetCurrent().User.Value
    Register-ScheduledTask -TaskName $taskName -Action $action -Trigger $trigger -Settings $taskSettings -Principal $principal -Force | Out-Null
} catch { Write-Warning 'Automatisk oprydning kunne ikke registreres. Oprydning sker ved næste appstart og mens appen er åben.' }
Write-Host 'Installeret. Åbn Fastansatte og Vikarer via genvejen på skrivebordet.'
if ($dependenciesReady) {
    Write-Host 'Begge PDF-programmer er klar. Forbind Microsoft 365 under Indstillinger i appen.'
} else {
    Write-Warning 'Appen er installeret, men mindst ét PDF-program kræver opsætning. Se beskederne ovenfor og START-HER.txt.'
}
Write-Host 'Afslut gamle lønkørsler og stop den tidligere automatisering før første rigtige udsendelse.'
if (-not $dependenciesReady) { exit 2 }
