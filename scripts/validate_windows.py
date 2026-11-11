"""Real Excel/LibreOffice regression on synthetic workbooks; never sends email."""
from pathlib import Path
import json
import shutil
import sys
import types
import zipfile

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(ROOT / 'tests'))
from conftest import fast_workbook, vikar_workbook
from app.adapters import analyze_fastansat, analyze_vikar
from app.domain import defaults
from app.payroll_pdf import sanitize_workbook
from app.rendering import render


def main():
    import pymupdf
    target = ROOT / '.test-data' / 'windows-validation'
    target.mkdir(parents=True, exist_ok=True)
    run = target / __import__('uuid').uuid4().hex[:8]
    run.mkdir()
    old_payroll_zip, old_vikar_zip = map(Path, sys.argv[1:3])
    with zipfile.ZipFile(old_payroll_zip) as archive:
        original_payroll = types.ModuleType('app.original_payroll')
        original_payroll.__package__ = 'app'
        original_payroll.__file__ = str(ROOT / 'app' / 'payroll_pdf.py')
        exec(compile(archive.read('payroll_local/pdf.py'), 'original_payroll.py', 'exec'), original_payroll.__dict__)
    with zipfile.ZipFile(old_vikar_zip) as archive:
        original_vikar = types.ModuleType('original_vikar')
        source = archive.read('Mail sender/src/pdf_exporter.py').decode('utf-8-sig')
        # Original layout, isolated COM instance so existing Excel windows are untouched.
        source = source.replace('win32com.client.Dispatch(', 'win32com.client.DispatchEx(')
        exec(compile(source, 'original_vikar.py', 'exec'), original_vikar.__dict__)
    report = {}
    for kind in ('fastansat', 'vikar'):
        folder = run / kind
        folder.mkdir()
        source = (fast_workbook if kind == 'fastansat' else vikar_workbook)(folder / 'input.xlsx', people=('Anna Eksempel',))
        clean = folder / 'sanitized.xlsx'
        sanitize_workbook(source, clean)
        row = (analyze_fastansat if kind == 'fastansat' else analyze_vikar)(clean)[0]
        assert not row['errors'], row['errors']
        print('Rendering new ' + kind, flush=True)
        rendered = render(clean, folder / 'new', row, kind, defaults())
        print('Rendering original ' + kind, flush=True)
        if kind == 'fastansat':
            original_rows = original_payroll.generate_documents(source, folder / 'original', defaults())
            original_path = folder / 'original' / original_rows[0]['pdf_file']
        else:
            original_source = folder / 'original-input.xlsx'
            shutil.copyfile(source, original_source)
            original_path = folder / 'original.pdf'
            original_vikar.export_sheet_to_pdf(original_source, row['sheet'], original_path)
        new_path = folder / 'new' / rendered['pdf']
        original_doc, new_doc = pymupdf.open(original_path), pymupdf.open(new_path)
        assert len(original_doc) == len(new_doc), (kind, len(original_doc), len(new_doc))
        differences = []
        for index in range(len(new_doc)):
            a, b = original_doc[index].get_pixmap(), new_doc[index].get_pixmap()
            if a.width != b.width or a.height != b.height or a.samples != b.samples:
                differences.append(index + 1)
            if index in (0, len(new_doc) - 1):
                a.save(str(folder / f'original-page-{index+1}.png'))
                b.save(str(folder / f'new-page-{index+1}.png'))
        report[kind] = {'pages': len(new_doc), 'pixel_differences': differences, 'new_pdf': str(new_path), 'original_pdf': str(original_path)}
        assert not differences, (kind, differences)
        original_doc.close()
        new_doc.close()
    (run / 'report.json').write_text(json.dumps(report, indent=2, ensure_ascii=False), encoding='utf-8')
    print(json.dumps(report, indent=2, ensure_ascii=False), flush=True)


if __name__ == '__main__':
    main()
