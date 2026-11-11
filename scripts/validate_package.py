"""Smoke-test the packaged executable with real PDF engines, without mail login."""
from pathlib import Path
import json
import os
import socket
import subprocess
import sys
import time
import uuid

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(ROOT / 'tests'))
from conftest import fast_workbook, vikar_workbook
import requests
from pypdf import PdfReader
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.workbook.defined_name import DefinedName


def main():
    with socket.socket() as probe:
        if probe.connect_ex(('127.0.0.1', 9000)) == 0:
            raise RuntimeError('Port 9000 er optaget. Luk appen før pakketesten.')
    run = ROOT / '.test-data' / 'package' / uuid.uuid4().hex[:8]
    run.mkdir(parents=True)
    exe = ROOT / 'dist' / 'FastansatVikar' / 'FastansatVikar.exe'
    report_file = run / 'runtime.json'
    result = subprocess.run([str(exe),'--check-runtime',str(report_file)], timeout=60, creationflags=subprocess.CREATE_NO_WINDOW)
    assert result.returncode == 0
    report = json.loads(report_file.read_text(encoding='utf-8'))
    assert report['excel'] and report['libreoffice'], report
    env = dict(os.environ, PAYROLL_DATA_DIR=str(run / 'data'), PAYROLL_NO_BROWSER='1')
    process = subprocess.Popen([str(exe)], cwd=run, env=env, creationflags=subprocess.CREATE_NO_WINDOW)
    session = requests.Session()
    origin = 'http://127.0.0.1:9000'
    try:
        for _ in range(120):
            if process.poll() is not None:
                raise RuntimeError('Packaged app exited: ' + str(process.returncode))
            try:
                response = session.get(origin+'/api/session', timeout=1)
                if response.ok:
                    break
            except requests.RequestException:
                pass
            time.sleep(.5)
        else:
            raise RuntimeError('Packaged server did not start')
        session.headers['X-CSRF-Token'] = response.json()['csrf']
        assert response.json()['account'] is None
        assert 'id="cc"' in session.get(origin+'/', timeout=10).text
        report['pdfs'] = {}
        for kind in ('fastansat','vikar'):
            path = (fast_workbook if kind == 'fastansat' else vikar_workbook)(run/(kind+'.xlsx'),people=('Anna Pakke','Bo Pakke'))
            wb = load_workbook(path)
            wb.defined_names.add(DefinedName('OldExternalLink', attr_text="'[Old.xlsx]Sheet1'!A1"))
            wb['Registrering 2' if kind == 'fastansat' else 'Bo Pakke']['C10'] = "='[Old.xlsx]Sheet1'!A1"
            wb.save(path)
            wb.close()
            with path.open('rb') as file:
                response = session.post(origin+'/api/jobs', data={'kind':kind,'month':'10','year':'2026','cc':'copy@example.com'},files={'file':(path.name,file)},timeout=30)
            response.raise_for_status()
            job = response.json()
            for _ in range(120):
                job = session.get(origin+'/api/jobs/'+job['id'],timeout=10).json()
                if job['state']!='ANALYZING':
                    break
                time.sleep(2)
            assert job['state']=='DRAFT', job
            assert job['cc'] == 'copy@example.com'
            assert job['warnings']
            assert not job['recipients'][1]['sendable']
            response = session.patch(origin+'/api/jobs/'+job['id']+'/cc', json={'cc':'changed@example.com'}, timeout=10)
            response.raise_for_status()
            assert response.json()['cc'] == 'changed@example.com'
            row = job['recipients'][0]
            assert row['sendable'], row['errors']
            response = session.get(origin+f"/api/jobs/{job['id']}/recipients/{row['id']}/pdf",timeout=15)
            response.raise_for_status()
            pages=len(PdfReader(BytesIO(response.content)).pages)
            assert pages == (13 if kind=='fastansat' else 2)
            report['pdfs'][kind]=pages
            print(f'Packaged {kind}: {pages} pages OK',flush=True)
        response=session.post(origin+'/api/shutdown',timeout=10)
        response.raise_for_status()
        process.wait(timeout=30)
        assert process.returncode == 0
        (run/'report.json').write_text(json.dumps(report,ensure_ascii=False,indent=2),encoding='utf-8')
        print(json.dumps({'passed':True,'report':report,'artifacts':str(run)},ensure_ascii=False),flush=True)
    finally:
        if process.poll() is None:
            process.terminate()
            process.wait(timeout=30)


if __name__ == '__main__':
    main()
