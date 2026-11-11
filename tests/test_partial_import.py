import zipfile
import xml.etree.ElementTree as ET

import pytest
from openpyxl import load_workbook
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.formatting.rule import FormulaRule

from app.payroll_pdf import MAIN, NS, PdfValidationError, sanitize_workbook
from conftest import create_job, fast_workbook, vikar_workbook


def rewrite_part(path, part, change):
    with zipfile.ZipFile(path) as archive:
        contents = {name: archive.read(name) for name in archive.namelist()}
    contents[part] = change(contents[part])
    with zipfile.ZipFile(path, 'w', zipfile.ZIP_DEFLATED) as archive:
        for name, raw in contents.items():
            archive.writestr(name, raw)


def test_unused_external_names_do_not_block_import(service, tmp_path):
    path = vikar_workbook(tmp_path / 'names.xlsx')
    wb = load_workbook(path)
    wb.defined_names.add(DefinedName('OldLink', attr_text="'[Old.xlsx]Sheet1'!$A$1"))
    wb.save(path)
    original = path.read_bytes()
    job = create_job(service, path, 'vikar')
    assert job['state'] == 'DRAFT'
    assert all(r['sendable'] for r in job['recipients'])
    assert any(w['cell'] == 'OldLink' for w in job['warnings'])
    assert path.read_bytes() == original
    clean = load_workbook(service.directory(job['id']) / 'sanitized.xlsx')
    assert 'OldLink' not in clean.defined_names
    clean.close()


@pytest.mark.parametrize('reference', ['=OldLink', '=Alias', '=SUM(OldLink)'])
def test_used_external_names_only_block_dependent_employee(service, tmp_path, reference):
    path = vikar_workbook(tmp_path / 'used-name.xlsx')
    wb = load_workbook(path)
    wb.defined_names.add(DefinedName('OldLink', attr_text="'[Old.xlsx]Sheet1'!$A$1"))
    wb.defined_names.add(DefinedName('Alias', attr_text='OldLink'))
    wb['Anna Vikar']['C10'] = reference
    wb.save(path)
    job = create_job(service, path, 'vikar')
    assert [r['sendable'] for r in job['recipients']] == [False, True]
    assert any('navngivet område' in e['message'] for e in job['recipients'][0]['errors'])


def test_local_names_do_not_block_other_scope(service, tmp_path):
    path = vikar_workbook(tmp_path / 'scoped.xlsx')
    wb = load_workbook(path)
    wb['Anna Vikar'].defined_names.add(DefinedName('Hours', attr_text="'[Old.xlsx]Sheet1'!$A$1"))
    wb['Bo Vikar'].defined_names.add(DefinedName('Hours', attr_text='7.5'))
    for ws in wb:
        ws['C10'] = '=Hours'
    wb.save(path)
    job = create_job(service, path, 'vikar')
    assert [r['sendable'] for r in job['recipients']] == [False, True]


@pytest.mark.parametrize('dependency', ["='anna vikar'!C10", '=INDIRECT("Anna Vikar!C10")', "=SUM('Anna Vikar:Bo Vikar'!C10)"])
def test_transitive_dependencies_are_blocked(service, tmp_path, dependency):
    path = vikar_workbook(tmp_path / 'dependencies.xlsx', people=('Anna Vikar', 'Bo Vikar', 'Carl Vikar', 'Dorte Vikar'))
    wb = load_workbook(path)
    wb['Anna Vikar']['C10'] = "='[Old.xlsx]Sheet1'!A1"
    wb['Bo Vikar']['C10'] = dependency
    wb['Carl Vikar']['C10'] = "='Bo Vikar'!C10"
    wb.save(path)
    job = create_job(service, path, 'vikar')
    assert [r['sendable'] for r in job['recipients']] == [False, False, False, True]


@pytest.mark.parametrize('malformation', ['xml', 'number', 'missing'])
@pytest.mark.parametrize('kind', ['vikar', 'fastansat'])
def test_unreadable_sheet_does_not_stop_good_employees(service, tmp_path, kind, malformation):
    path = (vikar_workbook if kind == 'vikar' else fast_workbook)(tmp_path / 'broken.xlsx')
    def damage(raw):
        if malformation == 'xml':
            return b'<worksheet><broken'
        root = ET.fromstring(raw)
        cell = root.find('.//m:c[@r="B5"]', NS) if kind == 'vikar' else root.find('.//m:c[@r="B7"]', NS)
        cell.find('m:v', NS).text = 'not-a-number'
        return ET.tostring(root)
    if malformation == 'missing':
        with zipfile.ZipFile(path) as archive:
            parts = {n: archive.read(n) for n in archive.namelist() if n != 'xl/worksheets/sheet1.xml'}
        with zipfile.ZipFile(path, 'w') as archive:
            for n, raw in parts.items():
                archive.writestr(n, raw)
    else:
        rewrite_part(path, 'xl/worksheets/sheet1.xml', damage)
    job = create_job(service, path, kind)
    assert job['state'] == 'DRAFT', job['error']
    good = [r for r in job['recipients'] if r['sendable']]
    assert len(good) == 1
    assert good[0]['email'] == ('vikar2@example.com' if kind == 'vikar' else 'person2@example.com')
    assert any(e['code'] == 'WORKBOOK_DATA' for r in job['recipients'] for e in r['errors'])


def test_overview_formula_error_blocks_only_its_employee(service, tmp_path):
    path = fast_workbook(tmp_path / 'overview.xlsx')
    wb = load_workbook(path)
    wb['Oversigt 1']['B10'] = "='[Old.xlsx]Sheet1'!A1"
    wb.save(path)
    job = create_job(service, path)
    assert [r['sendable'] for r in job['recipients']] == [False, True]
    assert any(c['sheet'] == 'Oversigt 1' and c['cell'] == 'B10' for e in job['recipients'][0]['errors'] for c in e['cells'])


def test_external_conditional_formatting_requires_pdf_review(service, tmp_path):
    path = vikar_workbook(tmp_path / 'format.xlsx')
    wb = load_workbook(path)
    wb.active.conditional_formatting.add('A1:A10', FormulaRule(formula=["'[Old.xlsx]Sheet1'!A1>0"]))
    wb.save(path)
    job = create_job(service, path, 'vikar')
    first, second = job['recipients']
    assert first['sendable'] and first['needs_pdf_review']
    assert second['sendable'] and not second.get('needs_pdf_review')
    job = service.edit(job['id'], first['id'], {'body': 'Ændret tekst'})
    with pytest.raises(ValueError):
        service.approve(job['id'], [first['id']], job['revision'])
    job = service.edit(job['id'], first['id'], {'reviewed': True})
    service.approve(job['id'], [first['id']], job['revision'])


@pytest.mark.parametrize('expression', ["'[Old.xlsx]Sheet1'!A1", 'WEBSERVICE("https://example.invalid")'])
def test_cached_formulas_become_values_and_original_is_unchanged(service, tmp_path, expression):
    path = vikar_workbook(tmp_path / 'cache.xlsx')
    wb = load_workbook(path)
    wb.active['C10'] = '=' + expression
    wb.save(path)
    def add_cache(raw):
        root = ET.fromstring(raw)
        root.find('.//m:c[@r="C10"]/m:v', NS).text = '12.5'
        return ET.tostring(root)
    rewrite_part(path, 'xl/worksheets/sheet1.xml', add_cache)
    before = path.read_bytes()
    job = create_job(service, path, 'vikar')
    assert all(r['sendable'] for r in job['recipients'])
    clean = load_workbook(service.directory(job['id']) / 'sanitized.xlsx')
    assert clean.active['C10'].value == 12.5
    assert path.read_bytes() == before
    clean.close()


def test_shared_external_formula_followers_cannot_escape_validation(service, tmp_path):
    path = vikar_workbook(tmp_path / 'shared.xlsx')
    def shared(raw):
        root = ET.fromstring(raw)
        data = root.find('m:sheetData', NS)
        for number in (76, 77):
            row = ET.SubElement(data, '{%s}row' % MAIN, r=str(number))
            cell = ET.SubElement(row, '{%s}c' % MAIN, r=f'A{number}')
            formula = ET.SubElement(cell, '{%s}f' % MAIN, t='shared', si='0')
            if number == 76:
                formula.text = "'[Old.xlsx]Sheet1'!A1"
                ET.SubElement(cell, '{%s}v' % MAIN).text = '10'
        return ET.tostring(root)
    rewrite_part(path, 'xl/worksheets/sheet1.xml', shared)
    job = create_job(service, path, 'vikar')
    assert [r['sendable'] for r in job['recipients']] == [False, True]
    assert any(c['cell'] == 'A77' for e in job['recipients'][0]['errors'] for c in e['cells'])


def test_macros_still_rejected_as_global_package_problem(tmp_path):
    path = vikar_workbook(tmp_path / 'macro.xlsx')
    with zipfile.ZipFile(path, 'a') as archive:
        archive.writestr('xl/vbaProject.bin', b'not-a-real-macro')
    with pytest.raises(PdfValidationError, match='Macros'):
        sanitize_workbook(path, tmp_path / 'clean.xlsx')


def test_existing_cached_excel_error_allows_pdf_generation_with_review(service, tmp_path):
    path = vikar_workbook(tmp_path / 'cached-error.xlsx')
    def cached_error(raw):
        root = ET.fromstring(raw)
        cell = root.find('.//m:c[@r="A1"]', NS)
        cell.clear()
        cell.set('r', 'A1')
        cell.set('t', 'e')
        ET.SubElement(cell, '{%s}f' % MAIN).text = '1+"text"'
        ET.SubElement(cell, '{%s}v' % MAIN).text = '#VALUE!'
        return ET.tostring(root)
    rewrite_part(path, 'xl/worksheets/sheet1.xml', cached_error)
    job = create_job(service, path, 'vikar')
    assert all(r['sendable'] for r in job['recipients'])
    assert job['recipients'][0]['needs_pdf_review']
    assert '#VALUE!' in job['recipients'][0]['warnings'][0]
    with pytest.raises(ValueError):
        service.approve(job['id'], [job['recipients'][0]['id']], job['revision'])


def test_failed_import_can_be_retried_after_update(service, tmp_path):
    path = vikar_workbook(tmp_path / 'retry.xlsx')
    job = create_job(service, path, 'vikar')
    previous = service.store.get(job['id'])
    previous.update(state='ERROR', recipients=[], error='Old sanitizer rejected file')
    service.store.put(previous)
    retried = create_job(service, path, 'vikar')
    assert retried['id'] != job['id'] and retried['state'] == 'DRAFT'
