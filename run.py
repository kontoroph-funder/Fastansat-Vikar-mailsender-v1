"""Entry point for the packaged Windows executable and isolated workers."""
import json
import os
from pathlib import Path
import sys


def main():
    if '--excel-worker' in sys.argv:
        from app.excel_worker import export
        export(*sys.argv[2:])
        return
    if '--check-runtime' in sys.argv:
        from app.domain import defaults
        from app.runtime import check_runtime
        report = check_runtime(defaults())
        Path(sys.argv[2]).write_text(json.dumps(report, ensure_ascii=False), encoding='utf-8')
        return
    from app.__main__ import main as start
    start()


if __name__ == '__main__':
    try:
        main()
    except Exception:
        if '--excel-worker' in sys.argv or '--check-runtime' in sys.argv or '--cleanup' in sys.argv:
            raise
        if os.name == 'nt':
            import ctypes
            ctypes.windll.user32.MessageBoxW(None, 'Appen kunne ikke startes. Kontrollér, at port 9000 er ledig, og at installationen er komplet.', 'Fastansatte & Vikarer', 0x10)
        raise
