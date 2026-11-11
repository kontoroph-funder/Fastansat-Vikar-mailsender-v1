import os
import json
from pathlib import Path
import socket
import sys
import threading
import webbrowser

import uvicorn

from .api import create_app


def main():
    port = int(os.environ.get('PAYROLL_PORT', '9000'))
    if not 1024 <= port <= 65535:
        raise ValueError('Ugyldig lokal port.')
    root = Path(os.environ.get('PAYROLL_DATA_DIR') or Path(os.environ.get('LOCALAPPDATA', str(Path.home()))) / 'FastansatVikar' / 'data')
    root.mkdir(parents=True, exist_ok=True)
    lock = (root / 'app.lock').open('a+b')
    if os.name == 'nt':
        import msvcrt
        lock.seek(0)
        lock.write(b'0')
        lock.flush()
        lock.seek(0)
        try:
            msvcrt.locking(lock.fileno(), msvcrt.LK_NBLCK, 1)
        except OSError:
            if '--cleanup' not in sys.argv:
                webbrowser.open(f'http://127.0.0.1:{port}')
            return
    if '--cleanup' in sys.argv:
        from .service import Service
        service = Service(root)
        service.cleanup()
        service.close()
        lock.close()
        return
    def open_browser():
        import time
        for _ in range(100):
            try:
                with socket.create_connection(('127.0.0.1', port), timeout=.2):
                    webbrowser.open(f'http://127.0.0.1:{port}')
                    return
            except OSError:
                time.sleep(.1)
    if os.environ.get('PAYROLL_NO_BROWSER') != '1':
        threading.Thread(target=open_browser, daemon=True).start()
    try:
        app = create_app(root)
        server = uvicorn.Server(uvicorn.Config(app, host='127.0.0.1', port=port, access_log=False, log_level='warning', log_config=None))
        app.state.shutdown = lambda: setattr(server, 'should_exit', True)
        server.run()
    finally:
        lock.close()


if __name__ == '__main__':
    main()
