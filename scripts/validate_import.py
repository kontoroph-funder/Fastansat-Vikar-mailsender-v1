"""Read a supplied workbook, report counts only, optionally render one PDF.

All private working files are confined to a temporary directory and removed.
Does not log in, send mail, or modify the supplied workbook or live jobs.
"""
import argparse
from collections import Counter
import hashlib
import json
from pathlib import Path
import sys
import time

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from app.adapters import analyze_fastansat, analyze_vikar
from app.domain import defaults
from app.files import render_directory
from app.payroll_pdf import sanitize_workbook
from app.rendering import render


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('source', type=Path)
    parser.add_argument('kind', choices=('fastansat', 'vikar'))
    parser.add_argument('--render-first', action='store_true')
    args = parser.parse_args()
    started = time.monotonic()
    original_hash = hashlib.sha256(args.source.read_bytes()).hexdigest()
    workspace = ROOT / '.test-data'
    workspace.mkdir(exist_ok=True)
    with render_directory(workspace) as temporary:
        clean = Path(temporary) / 'sanitized.xlsx'
        report = sanitize_workbook(args.source, clean)
        rows = (analyze_fastansat if args.kind == 'fastansat' else analyze_vikar)(clean)
        bad = {i['sheet'] for i in report.issues}
        valid = [r for r in rows if not r['errors'] and not {r['sheet'], r.get('overview')} & bad]
        result = {'sheets': len(rows), 'valid_before_pdf': len(valid), 'frozen': report.frozen,
                  'warnings': len(report.warnings), 'data_errors': len(report.issues),
                  'adapter_errors': dict(Counter(e['code'] for r in rows for e in r['errors']))}
        print(json.dumps(result), flush=True)
        if args.render_first:
            assert valid, 'No eligible employee for PDF export'
            exported = render(clean, Path(temporary) / 'output', valid[0], args.kind, defaults())
            print(json.dumps({'pdf_pages': exported['pages']}), flush=True)
    assert hashlib.sha256(args.source.read_bytes()).hexdigest() == original_hash
    print(json.dumps({'original_unchanged': True, 'temporary_files_removed': True,
                      'seconds': round(time.monotonic() - started, 1)}), flush=True)


if __name__ == '__main__':
    main()
