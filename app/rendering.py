from __future__ import annotations

import hashlib
import os
from pathlib import Path
import re
import subprocess
import sys

from pypdf import PdfReader, PdfWriter

from .files import render_directory
from .payroll_pdf import _render, PdfValidationError


def render(source: Path, output: Path, row: dict, kind: str, config: dict):
    output.mkdir(parents=True, exist_ok=True)
    destination = output / (row['id'] + '.pdf')
    with render_directory(output) as temporary:
        work = Path(temporary)
        if kind == 'fastansat':
            segments = _render(source, work, row['_export'], config)
            writer = PdfWriter()
            identity = False
            try:
                for segment in segments:
                    pdf = PdfReader(segment, strict=True)
                    if pdf.is_encrypted or not pdf.pages:
                        raise PdfValidationError('Tomt eller krypteret PDF-afsnit.')
                    text = ' '.join(' '.join((p.extract_text() or '').split()) for p in pdf.pages)
                    if not text.strip():
                        raise PdfValidationError('PDF-afsnit indeholder ingen læsbar tekst.')
                    expected = ' '.join(row['name'].casefold().split())
                    identity |= bool(re.search(r'(?<!\w)' + re.escape(expected) + r'(?!\w)', text.casefold()))
                    for page in pdf.pages:
                        writer.add_page(page)
                if not identity:
                    raise PdfValidationError('Medarbejdernavnet findes ikke i PDF’en.')
                with (work / 'result.pdf').open('xb') as handle:
                    writer.write(handle)
            finally:
                writer.close()
        else:
            command = [sys.executable, '--excel-worker'] if getattr(sys, 'frozen', False) else [sys.executable, '-m', 'app.excel_worker']
            process = subprocess.Popen([*command, str(source), row['sheet'], str(work / 'result.pdf')],
                                       stdout=subprocess.PIPE, stderr=subprocess.PIPE,
                                       creationflags=getattr(subprocess, 'CREATE_NO_WINDOW', 0))
            try:
                _, stderr = process.communicate(timeout=180)
            except subprocess.TimeoutExpired:
                process.kill()
                process.communicate()
                pidfile = work / 'excel.pid'
                if pidfile.exists() and os.name == 'nt':
                    pid = pidfile.read_text().strip()
                    if pid.isdigit():
                        subprocess.run(['taskkill.exe', '/PID', pid, '/T', '/F'], capture_output=True,
                                       creationflags=subprocess.CREATE_NO_WINDOW, timeout=20)
                raise PdfValidationError('Excel-eksport overskred tidsgrænsen.') from None
            if process.returncode:
                raise PdfValidationError('Excel-eksport fejlede. Kontrollér at Microsoft Excel er installeret og kan åbnes.')
        result = work / 'result.pdf'
        pdf = PdfReader(result, strict=True)
        if pdf.is_encrypted or not pdf.pages or not any((p.extract_text() or '').strip() for p in pdf.pages):
            raise PdfValidationError('PDF’en er tom eller ulæselig.')
        if kind == 'vikar' and len(pdf.pages) != 2:
            raise PdfValidationError('Vikar-PDF skal indeholde to sider. Kontrollér arkets layout.')
        raw = result.read_bytes()
        with destination.open('xb') as handle:
            handle.write(raw)
        return {'pdf': destination.name, 'sha256': hashlib.sha256(raw).hexdigest(), 'pages': len(pdf.pages)}
