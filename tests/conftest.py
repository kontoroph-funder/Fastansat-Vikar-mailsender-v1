import hashlib
import io
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from pypdf import PdfWriter
import pytest

from app.domain import MONTHS
from app.service import Service


def pytest_configure(config):
    (Path(__file__).resolve().parents[1] / '.test-data').mkdir(exist_ok=True)


def fast_workbook(path, people=('Anna Eksempel', 'Bo Eksempel'), mismatch=False):
    wb = Workbook()
    wb.remove(wb.active)
    for index, name in enumerate(people):
        reg = wb.create_sheet('Registrering ' + str(index + 1))
        reg['F1'] = 'Timeregistrering ' + name
        reg['R2'] = f'person{index+1}@example.com'
        reg['A2'] = name
        for month_index, month in enumerate(MONTHS):
            start = 5 + month_index * 6
            reg.cell(start, 1, month.title())
            reg.cell(start, 1).font = Font(bold=True, size=14)
            reg.cell(start+1, 1, 'Dato')
            reg.cell(start+1, 2, 'Timer')
            reg.cell(start+2, 1, '01. ' + month)
            reg.cell(start+2, 2, 7.5)
            reg.cell(start+4, 1, 'I alt')
            reg.cell(start+4, 2, 7.5)
        reg.column_dimensions['A'].width = 24
        reg.column_dimensions['B'].width = 14
        overview = wb.create_sheet('Oversigt ' + str(index+1))
        overview['S2'] = f'person{index+1}@example.com' if not mismatch or index == 0 else 'forkert@example.com'
        overview['A1'] = name
        overview['A1'].font = Font(size=16, bold=True)
        overview['A2'] = 'Timebank og timeregistrering'
        overview['A3'] = 'Måned'
        overview['C3'] = 'Timer'
        for i, month in enumerate(MONTHS):
            overview.cell(i+4, 1, month.title())
            overview.cell(i+4, 3, f"='{reg.title}'!B{7+i*6}")
        overview.column_dimensions['A'].width = 27
        overview.column_dimensions['C'].width = 16
    wb.save(path)
    wb.close()
    return path


def vikar_workbook(path, people=('Anna Vikar', 'Bo Vikar'), hours=7.5):
    wb = Workbook()
    wb.remove(wb.active)
    for i, name in enumerate(people):
        ws = wb.create_sheet(name)
        ws['A1'] = 'Timeoversigt – ' + name
        ws['A1'].font = Font(size=16, bold=True)
        ws['A2'], ws['B2'] = 'E-mail', f'vikar{i+1}@example.com'
        ws['A4'], ws['B4'] = 'Dato', 'Timer'
        ws['A5'], ws['B5'] = '16. september', hours
        for number in range(6, 26):
            ws.cell(number, 1, str(number) + '. oktober')
            ws.cell(number, 2, 0)
        ws['A39'], ws['A75'] = 'Opsummering – ' + name, 'Samlet timeoversigt'
        ws['B75'] = '=SUM(B5:B25)'
        for number in range(1, 76):
            ws.row_dimensions[number].height = 14
        ws.column_dimensions['A'].width = 25
        ws.column_dimensions['B'].width = 35
        for col in 'CDEFGHIJK':
            ws.column_dimensions[col].width = 5
    wb.save(path)
    wb.close()
    return path


class FakeGraph:
    login = None
    def __init__(self):
        self.sent = []
        self.fail_at = None
        self.identity = {'id': 'test-account', 'email': 'administrator@example.com', 'name': 'Test Administrator'}
    def account(self):
        return self.identity
    def send(self, *args):
        if self.fail_at is not None and len(self.sent) == self.fail_at:
            from app.graph import UnknownDelivery
            raise UnknownDelivery('Simuleret timeout')
        self.sent.append(args)
    def logout(self):
        self.identity = None


def fake_renderer(source, output, row, kind, config):
    output.mkdir(parents=True, exist_ok=True)
    path = output / (row['id'] + '.pdf')
    writer = PdfWriter()
    writer.add_blank_page(595, 842)
    writer.write(path)
    writer.close()
    return {'pdf': path.name, 'sha256': hashlib.sha256(path.read_bytes()).hexdigest(), 'pages': 1}


def wait(service):
    for future in list(service.futures):
        future.result(timeout=30)


@pytest.fixture
def service(tmp_path):
    graph = FakeGraph()
    result = Service(tmp_path / 'data', graph=graph, renderer=fake_renderer)
    yield result
    result.close()


def create_job(service, path, kind='fastansat', month=10, year=2026):
    job = service.create(kind, month, year, path.name, path.read_bytes())
    wait(service)
    return service.get(job['id'])
