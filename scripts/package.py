"""Create the redistributable ZIP from an explicit artifact allowlist."""
import hashlib
import os
from pathlib import Path
import time
import uuid
import zipfile

ROOT = Path(__file__).resolve().parents[1]


def read_artifact(path):
    for attempt in range(10):
        try:
            return path.read_bytes()
        except PermissionError:
            if attempt == 9:
                raise
            time.sleep(.25)


def main():
    bundle = ROOT / 'dist' / 'FastansatVikar'
    if not (bundle / 'FastansatVikar.exe').is_file():
        raise RuntimeError('Build the Windows executable first.')
    destination = ROOT / 'dist' / 'FastansatVikar-Windows.zip'
    temporary = destination.with_name('package-' + uuid.uuid4().hex + '.tmp')
    with zipfile.ZipFile(temporary, 'w', zipfile.ZIP_DEFLATED, compresslevel=6) as archive:
        for path in sorted(bundle.rglob('*')):
            if path.is_symlink():
                raise RuntimeError('Symbolic links cannot be packaged.')
            if path.is_file():
                archive.writestr('FastansatVikar/' + path.relative_to(bundle).as_posix(), read_artifact(path))
        for name in ('install.cmd', 'install.ps1', 'START-HER.txt'):
            archive.writestr(name, (ROOT / name).read_bytes())
    with zipfile.ZipFile(temporary) as archive:
        assert archive.testzip() is None
        assert 'FastansatVikar/_internal/app/static/index.html' in archive.namelist()
        assert 'FastansatVikar/_internal/app/pdf_worker.py' in archive.namelist()
        assert not any(name.endswith(('.sqlite3','.dpapi')) for name in archive.namelist())
    os.replace(temporary, destination)
    sha = hashlib.sha256(destination.read_bytes()).hexdigest()
    destination.with_suffix('.sha256').write_text(sha + '  ' + destination.name + '\n', encoding='ascii')
    print(str(destination))
    print(f'{destination.stat().st_size:,} bytes; SHA256 {sha}')


if __name__ == '__main__':
    main()
