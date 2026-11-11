# Build with: .venv\Scripts\python.exe -m PyInstaller FastansatVikar.spec --noconfirm
from PyInstaller.utils.hooks import collect_data_files

a = Analysis(
    ['run.py'], pathex=[], binaries=[],
    datas=[('app/static', 'app/static'), ('app/pdf_worker.py', 'app')] + collect_data_files('tzdata'),
    hiddenimports=['uvicorn.logging', 'uvicorn.loops.asyncio', 'uvicorn.protocols.http.h11_impl',
                   'uvicorn.lifespan.on', 'win32timezone', 'pythoncom', 'pywintypes', 'app.excel_worker', 'app.outlook_worker'],
    hookspath=[], hooksconfig={}, runtime_hooks=[],
    excludes=['pytest', 'playwright', 'pymupdf', 'fitz', 'numpy', 'onnxruntime', 'tkinter', 'matplotlib'],
    noarchive=False,
)
pyz = PYZ(a.pure)
exe = EXE(pyz, a.scripts, [], exclude_binaries=True, name='FastansatVikar', debug=False,
          bootloader_ignore_signals=False, strip=False, upx=False, console=False)
coll = COLLECT(exe, a.binaries, a.datas, strip=False, upx=False, name='FastansatVikar')
