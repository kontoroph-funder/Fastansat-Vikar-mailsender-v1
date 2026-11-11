from pathlib import Path
import os

from .payroll_pdf import check_runtime as check_libreoffice


def check_runtime(settings):
    excel = False
    if os.name == 'nt':
        import winreg
        try:
            with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, r'Excel.Application\CLSID'):
                excel = True
        except OSError:
            pass
    try:
        check_libreoffice(settings)
        libreoffice, error = True, ''
    except Exception as exc:
        libreoffice, error = False, str(exc)
    return {'excel': excel, 'libreoffice': libreoffice, 'libreoffice_error': error}
