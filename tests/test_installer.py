"""Exercise dependency decisions without installing software or changing the PC."""
import base64
import os
from pathlib import Path
import subprocess

import pytest


@pytest.mark.skipif(os.name != 'nt', reason='Windows installer')
@pytest.mark.parametrize('scenario,ready,downloads', [
    ('installed', True, 0), ('download', True, 1),
    ('missing_winget', False, 0), ('failed_download', False, 1),
    ('missing_excel', False, 0),
])
def test_dependency_setup(scenario, ready, downloads):
    source = (Path(__file__).resolve().parents[1] / 'install.ps1').read_text(encoding='utf-8-sig')
    start = source.index('$dependenciesReady = $false')
    end = source.index('    $action = New-ScheduledTaskAction', start)
    logic = source[start:end].rsplit('try {', 1)[0]
    script = r'''
$ErrorActionPreference = 'Stop'
$WarningPreference = 'SilentlyContinue'
$script:downloads = 0
function Write-Host { param($Object) }
function Get-Command {
    param($Name, $ErrorAction)
    if ($scenario -ne 'missing_winget') { return @{Source='Invoke-FakeWinget'} }
}
function Invoke-FakeWinget {
    $script:downloads++
    $global:LASTEXITCODE = 0
    if ($scenario -eq 'failed_download') { $global:LASTEXITCODE = 1 }
}
function Get-RuntimeReport {
    $lo = $scenario -in @('installed', 'missing_excel') -or ($scenario -eq 'download' -and $script:downloads -eq 1)
    return @{excel=($scenario -ne 'missing_excel'); libreoffice=$lo; libreoffice_error='test failure'}
}
'''
    script += f"\n$scenario = '{scenario}'\n" + logic
    script += f"\nif ($dependenciesReady -ne ${str(ready).lower()} -or $script:downloads -ne {downloads}) {{ throw 'Unexpected installer result' }}"
    encoded = base64.b64encode(script.encode('utf-16-le')).decode('ascii')
    result = subprocess.run(['powershell.exe', '-NoProfile', '-EncodedCommand', encoded], capture_output=True, timeout=30)
    assert result.returncode == 0, result.stderr.decode(errors='replace')


@pytest.mark.skipif(os.name != 'nt', reason='Windows file locking')
def test_installer_detects_source_app_lock(tmp_path):
    import msvcrt
    source = (Path(__file__).resolve().parents[1] / 'install.ps1').read_text(encoding='utf-8-sig')
    block = source[source.index('$runningAppLock ='):source.index('New-Item -ItemType Directory -Force -Path $appRoot')]
    (tmp_path / 'data').mkdir()
    lock_path = tmp_path / 'data' / 'app.lock'
    lock_path.write_bytes(b'0')
    script = "$ErrorActionPreference='Stop'; $installRoot='" + str(tmp_path).replace("'", "''") + "'\n" + block
    encoded = base64.b64encode(script.encode('utf-16-le')).decode('ascii')
    command = ['powershell.exe', '-NoProfile', '-EncodedCommand', encoded]
    with lock_path.open('r+b') as stream:
        msvcrt.locking(stream.fileno(), msvcrt.LK_NBLCK, 1)
        try:
            assert subprocess.run(command, capture_output=True, timeout=15).returncode != 0
        finally:
            stream.seek(0)
            msvcrt.locking(stream.fileno(), msvcrt.LK_UNLCK, 1)
    assert subprocess.run(command, capture_output=True, timeout=15).returncode == 0
