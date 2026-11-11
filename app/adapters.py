"""Workbook-specific analysis. A bad employee never invalidates a good match."""
from __future__ import annotations

from collections import defaultdict
from pathlib import Path
import math
import re
import uuid

from openpyxl import Workbook, load_workbook
from openpyxl.utils.cell import get_column_letter

from .payroll_pdf import EMAIL, MONTHS, PdfValidationError, _references, _print_area, sheet_cells


def employee(name, sheet, email):
    return {'id': uuid.uuid4().hex, 'name': name, 'sheet': sheet, 'email': email,
            'errors': [], 'warnings': [], 'sendable': False, 'selected': False,
            'pdf': '', 'sha256': '', 'pages': 0, 'status': 'BLOCKED', 'custom': False,
            'needs_review': False, 'subject': '', 'body': ''}


def issue(row, code, text, cells=None):
    row['errors'].append({'employee_id': row['id'], 'code': code, 'message': text,
                          'scope': 'employee', 'cells': cells or [],
                          'sheets': list(dict.fromkeys(c['sheet'] for c in (cells or []))) or [row['sheet']]})


def address_cell(sheet, cell):
    return {'sheet': sheet.title, 'cell': cell, 'value': str(sheet[cell].value or '').strip()}


def number(value):
    if isinstance(value, bool):
        return None
    if isinstance(value, (float, int)):
        return float(value) if math.isfinite(value) else None
    if isinstance(value, str):
        text = value.strip().replace('\u00a0', '').replace(' ', '')
        if ',' in text:
            text = text.replace('.', '').replace(',', '.')
        try:
            result = float(text)
            return result if math.isfinite(result) else None
        except ValueError:
            pass
    return None


def analyze_vikar(source: Path):
    wb = load_workbook(source, data_only=True, read_only=True, keep_links=False)
    rows = []
    try:
        for sheet in wb:
            try:
                values = {i: values[0] for i, values in enumerate(sheet.iter_rows(min_row=1, max_row=25, min_col=2, max_col=2, values_only=True), 1)}
                email = str(values.get(2) or values.get(1) or '').strip().casefold()
                row = employee(sheet.title, sheet.title, email)
                hours = [number(values.get(i)) for i in range(5, 26)]
            except Exception:
                row = employee(sheet.title, sheet.title, '')
                issue(row, 'SHEET_READ_ERROR', 'Arket kunne ikke læses. De øvrige ark behandles videre.')
                rows.append(row)
                continue
            row['hours'] = sum(n for n in hours if n is not None)
            if not EMAIL.fullmatch(email):
                issue(row, 'EMAIL_INVALID', 'Mangler gyldig e-mail i B2 eller B1.',
                      [{'sheet': sheet.title, 'cell': f'B{i}', 'value': str(values.get(i) or '')} for i in (2, 1)])
            if not any(n is not None and n > 0 for n in hours):
                issue(row, 'NO_HOURS', 'Ingen positive timer i B5:B25.')
            row['_export'] = {'sheet': sheet.title}
            rows.append(row)
        duplicates = defaultdict(list)
        for row in rows:
            if row['email']:
                duplicates[row['email']].append(row)
        for group in duplicates.values():
            if len(group) > 1:
                cells = [{'sheet': r['sheet'], 'cell': 'B2/B1', 'value': r['email']} for r in group]
                for row in group:
                    issue(row, 'DUPLICATE_EMAIL', 'Samme e-mail findes på flere medarbejderark.', cells)
        return rows
    finally:
        wb.close()


def analyze_fastansat(source: Path):
    # Read sheets independently; irrelevant display metadata must not make
    # openpyxl abort every employee before matching begins.
    original = load_workbook(source, data_only=False, read_only=True, keep_links=False)
    wb = Workbook()
    wb.remove(wb.active)
    unreadable = []
    try:
        for source_sheet in original:
            sheet = wb.create_sheet(source_sheet.title)
            try:
                for cell in sheet_cells(source_sheet):
                    if cell['value'] is not None:
                        sheet.cell(cell['row'], cell['column'], cell['value'])
            except Exception:
                wb.remove(sheet)
                row = employee(source_sheet.title, source_sheet.title, '')
                issue(row, 'SHEET_READ_ERROR', 'Arket kunne ikke læses. De øvrige ark behandles videre.')
                unreadable.append(row)
    finally:
        original.close()
    try:
        registrations = {}
        for sheet in wb:
            match = re.fullmatch(r'Timeregistrering\s+(.+)', str(sheet['F1'].value or '').strip(), re.I)
            if match:
                registrations[sheet.title] = (sheet, employee(match[1].strip(), sheet.title,
                    str(sheet['R2'].value or '').strip().casefold()))
        if not registrations:
            if unreadable:
                return unreadable
            raise PdfValidationError('Ingen registreringsark med medarbejdernavn i F1 blev fundet.')

        overviews, refs, by_email = {}, {}, defaultdict(list)
        for sheet in wb:
            if sheet.title in registrations:
                continue
            references = set().union(*(_references(sheet[f'C{i}'].value) for i in range(4, 16))) & registrations.keys()
            email = str(sheet['S2'].value or '').strip().casefold()
            if email or references:
                overviews[sheet.title] = sheet
                refs[sheet.title] = references
                if email:
                    by_email[email].append(sheet)

        # Block all parties in contradictory links, including otherwise valid pairs.
        matched_overviews = set()
        for title, (sheet, row) in registrations.items():
            cells = [address_cell(sheet, 'R2')]
            if not EMAIL.fullmatch(row['email']):
                issue(row, 'EMAIL_INVALID', 'Mangler gyldig e-mail i R2.', cells)
            matches = by_email.get(row['email'], []) if row['email'] else []
            if len(matches) > 1:
                matched_overviews.update(s.title for s in matches)
                issue(row, 'AMBIGUOUS_MATCH', 'Flere mulige match: samme S2-adresse findes på flere oversigtsark.',
                      cells + [address_cell(s, 'S2') for s in matches])
            elif not matches:
                candidates = [s for t, s in overviews.items() if title in refs[t]]
                issue(row, 'EMAIL_MISMATCH' if candidates else 'MISSING_OVERVIEW',
                      'Mailadresser matcher ikke.' if candidates else 'Mangler oversigtsark; modparten kan ikke identificeres entydigt.',
                      cells + [address_cell(s, 'S2') for s in candidates])
            else:
                overview = matches[0]
                matched_overviews.add(overview.title)
                row['overview'] = overview.title
                if refs[overview.title] and refs[overview.title] != {title}:
                    involved = {title} | refs[overview.title]
                    conflict_cells = [address_cell(overview, 'S2')] + [address_cell(registrations[t][0], 'R2') for t in sorted(involved)]
                    for affected in involved:
                        issue(registrations[affected][1], 'CONFLICTING_REFERENCE',
                              'Oversigtsarkets formler henviser til en anden eller flere medarbejdere.', conflict_cells)

            months = {}
            for (row_number, col), cell in sheet._cells.items():
                if col != 1:
                    continue
                text = str(cell.value or '').strip().casefold()
                for month in MONTHS:
                    if re.match(r'^' + month + r'\b', text):
                        if month in months:
                            issue(row, 'DUPLICATE_MONTH', f'Måneden {month} findes flere gange.',
                                  [{'sheet': title, 'cell': f'A{row_number}', 'value': text}])
                        months[month] = row_number
                        break
            if len(months) != 12 or [months.get(m, 0) for m in MONTHS] != sorted(months.values()):
                issue(row, 'MONTH_LAYOUT', 'Månedsafsnit mangler eller står i forkert rækkefølge.')
            elif len(matches) == 1:
                try:
                    last_row, last_col = _print_area(sheet, months)
                    ranges = [{'month': month, 'area': 'A%d:%s%d' % (months[month], get_column_letter(last_col),
                        months[MONTHS[index + 1]] - 1 if index < 11 else last_row)} for index, month in enumerate(MONTHS)]
                    row['_export'] = {'overview': matches[0].title, 'data': title, 'overview_area': 'A1:G16', 'months': ranges}
                except PdfValidationError as exc:
                    issue(row, 'PRINT_LAYOUT', f'Udskriftsområdet kunne ikke valideres: {exc}')

        # An overview explicitly referring to X must not silently pair with Y by email.
        for title, overview in overviews.items():
            for target in refs[title]:
                sheet, row = registrations[target]
                if str(overview['S2'].value or '').strip().casefold() != row['email']:
                    if not any(e['code'] == 'EMAIL_MISMATCH' and title in e['sheets'] for e in row['errors']):
                        issue(row, 'EMAIL_MISMATCH', 'Mailadresser matcher ikke på det refererende oversigtsark.',
                              [address_cell(sheet, 'R2'), address_cell(overview, 'S2')])

        rows = [r for _, r in registrations.values()]
        for key, code in [('email', 'DUPLICATE_EMAIL'), ('name', 'DUPLICATE_NAME')]:
            groups = defaultdict(list)
            for row in rows:
                if row[key]:
                    groups[row[key].casefold()].append(row)
            for group in groups.values():
                if len(group) > 1:
                    cells = [{'sheet': r['sheet'], 'cell': 'R2' if key == 'email' else 'F1', 'value': r[key]} for r in group]
                    for row in group:
                        issue(row, code, 'Dubleret medarbejderadresse eller medarbejdernavn.', cells)
        for title, sheet in overviews.items():
            if title not in matched_overviews:
                row = employee(f'Oversigtsark: {title}', title, str(sheet['S2'].value or '').strip())
                issue(row, 'ORPHAN_OVERVIEW', 'Oversigtsark uden et entydigt match til et registreringsark.', [address_cell(sheet, 'S2')])
                rows.append(row)
        return rows + unreadable
    finally:
        wb.close()
