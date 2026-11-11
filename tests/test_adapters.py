from openpyxl import load_workbook
import pytest

from app.adapters import analyze_fastansat, analyze_vikar, number
from app.payroll_pdf import sanitize_workbook, PdfValidationError
from conftest import fast_workbook, vikar_workbook


def test_mismatch_is_local_and_diagnostics_name_both_sheets(tmp_path):
    path = fast_workbook(tmp_path / 'mixed.xlsx', mismatch=True)
    rows = analyze_fastansat(path)
    assert not rows[0]['errors']
    bad = rows[1]
    assert {'EMAIL_MISMATCH'} <= {e['code'] for e in bad['errors']}
    cells = [c for e in bad['errors'] for c in e['cells']]
    assert {'sheet': 'Registrering 2', 'cell': 'R2', 'value': 'person2@example.com'} in cells
    assert {'sheet': 'Oversigt 2', 'cell': 'S2', 'value': 'forkert@example.com'} in cells
    assert any(e['code'] == 'ORPHAN_OVERVIEW' for r in rows for e in r['errors'])


def test_duplicate_registration_email_blocks_all_parties(tmp_path):
    path = fast_workbook(tmp_path / 'duplicates.xlsx')
    wb = load_workbook(path)
    wb['Registrering 2']['R2'] = 'person1@example.com'
    wb.save(path)
    rows = analyze_fastansat(path)
    assert all(any(e['code'] == 'DUPLICATE_EMAIL' for e in r['errors']) for r in rows[:2])


def test_conflicting_formula_blocks_both_otherwise_matched_employees(tmp_path):
    path = fast_workbook(tmp_path / 'conflict.xlsx')
    wb = load_workbook(path)
    wb['Oversigt 1']['C4'] = "='Registrering 2'!B7"
    wb.save(path)
    rows = analyze_fastansat(path)
    assert all(any(e['code'] == 'CONFLICTING_REFERENCE' for e in r['errors']) for r in rows[:2])


def test_missing_and_duplicate_overview_and_month_are_local(tmp_path):
    path = fast_workbook(tmp_path / 'layout.xlsx', people=('Anna', 'Bo', 'Carl'))
    wb = load_workbook(path)
    wb['Registrering 2']['A71'] = 'Ikke december'
    wb['Oversigt 3']['S2'] = 'person1@example.com'
    wb.save(path)
    rows = analyze_fastansat(path)
    assert 'AMBIGUOUS_MATCH' in {e['code'] for e in rows[0]['errors']}
    assert 'MONTH_LAYOUT' in {e['code'] for e in rows[1]['errors']}
    assert 'EMAIL_MISMATCH' in {e['code'] for e in rows[2]['errors']}


@pytest.mark.parametrize('value,expected', [(7.5,7.5),('7,5',7.5),('7.5',7.5),('1.234,50',1234.5),('0',0),(None,None),('tekst',None)])
def test_numeric_hours(value, expected):
    assert number(value) == expected


def test_vikar_uses_browser_cells_and_fallback(tmp_path):
    path = vikar_workbook(tmp_path / 'vikar.xlsx', hours='7,5')
    wb = load_workbook(path)
    wb['Anna Vikar']['B2'] = None
    wb['Anna Vikar']['B1'] = 'fallback@example.com'
    wb['Bo Vikar']['B5'] = 0
    wb['Bo Vikar']['G8'] = 10  # Desktop rule must not be used.
    wb.save(path)
    rows = analyze_vikar(path)
    assert rows[0]['email'] == 'fallback@example.com'
    assert rows[0]['hours'] == 7.5
    assert not rows[0]['errors']
    assert 'NO_HOURS' in {e['code'] for e in rows[1]['errors']}


def test_missing_email_is_visible(tmp_path):
    path = fast_workbook(tmp_path / 'empty.xlsx')
    wb = load_workbook(path)
    wb['Registrering 2']['R2'] = None
    wb.save(path)
    rows = analyze_fastansat(path)
    assert rows[1]['email'] == ''
    assert 'EMAIL_INVALID' in {e['code'] for e in rows[1]['errors']}
    assert not rows[0]['errors']


def test_external_formula_without_saved_value_is_reported_per_sheet(tmp_path):
    path = vikar_workbook(tmp_path / 'external.xlsx')
    wb = load_workbook(path)
    wb.active['A10'] = "='[External.xlsx]Ark1'!A1"
    wb.save(path)
    report = sanitize_workbook(path, tmp_path / 'sanitized.xlsx')
    assert {i['sheet'] for i in report.issues} == {'Anna Vikar'}
    assert report.issues[0]['cell'] == 'A10'
    clean = load_workbook(tmp_path / 'sanitized.xlsx')
    assert clean.active['A10'].value == '#REF!'
    clean.close()
