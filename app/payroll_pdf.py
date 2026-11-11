"""Local, fail-closed XLSX to employee PDFs. No mail or service dependencies.

The workbook is sanitized before it reaches LibreOffice. We retain its XML
formatting and internal formulas; external formulas become their saved scalar
values. Originals are never modified. PDF output follows the existing layout:
one overview followed by twelve monthly registration sections.
"""
from __future__ import annotations

import json
import io
import math
import os
import posixpath
from dataclasses import dataclass, field
from pathlib import Path, PurePosixPath
import re
import subprocess
import tempfile
import xml.etree.ElementTree as ET
import zipfile

from openpyxl import load_workbook
from openpyxl.formula import Tokenizer
from openpyxl.worksheet._reader import WorkSheetParser
from openpyxl.utils.cell import get_column_letter, range_boundaries

MONTHS = ('januar', 'februar', 'marts', 'april', 'maj', 'juni', 'juli',
          'august', 'september', 'oktober', 'november', 'december')
MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
REL = 'http://schemas.openxmlformats.org/package/2006/relationships'
DOCREL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
NS = {'m': MAIN}
EMAIL = re.compile(r"[A-Za-z0-9.!#$%&'*+/=?^_`{|}~-]+@[A-Za-z0-9](?:[A-Za-z0-9.-]*[A-Za-z0-9])?\.[A-Za-z]{2,63}")
EXTERNAL = re.compile(r'\[[^\]]+\][^!]*!')
UNSAFE_FORMULA = re.compile(r'(?i)(?:\b(?:WEBSERVICE|DDE|RTD|CALL|REGISTER|EXEC|IMAGE|HYPERLINK)\s*\(|://|\\\\|\|)')


class PdfValidationError(ValueError):
    """Input cannot safely be matched, rendered or published."""


@dataclass
class WorkbookReport:
    frozen: int = 0
    issues: list = field(default_factory=list)
    warnings: list = field(default_factory=list)
    dependencies: dict = field(default_factory=dict)

    def error(self, sheet, location, message, expression=''):
        self.issues.append({'sheet': sheet, 'cell': location, 'message': message, 'value': expression})

    def warning(self, sheet, location, message, review=False):
        self.warnings.append({'sheet': sheet, 'cell': location, 'message': message, 'review': review})


def _formula_dependencies(expression, sheet, titles, names):
    """Resolve sheet/name dependencies, including scoped names and 3-D ranges.

    INDIRECT/OFFSET cannot be resolved statically: conservatively connect them
    to all sheets/names so missing data cannot silently become a valid PDF.
    """
    title_map = {title.casefold(): title for title in titles}
    result = set()
    if re.search(r'(?i)\b(?:INDIRECT|OFFSET)\s*\(', expression):
        return set(titles) | set(names.values())
    try:
        tokens = Tokenizer('=' + expression.lstrip('=')).items
    except Exception:
        # Unsupported syntax is conservatively dependent on any damaged data.
        return set(titles) | set(names.values())
    for token in tokens:
        if token.type != 'OPERAND' or token.subtype != 'RANGE':
            continue
        value = token.value
        scope = sheet
        if '!' in value:
            prefix, value = value.rsplit('!', 1)
            prefix = prefix.strip("'").replace("''", "'")
            if ':' in prefix:
                first, last = (title_map.get(p.strip("'").casefold()) for p in prefix.split(':', 1))
                if first and last:
                    a, b = sorted((titles.index(first), titles.index(last)))
                    result.update(titles[a:b + 1])
            else:
                scope = title_map.get(prefix.casefold())
                if scope:
                    result.add(scope)
        key = names.get((scope, value.casefold()), names.get((None, value.casefold())))
        if key:
            result.add(key)
    return result


def sheet_cells(sheet):
    """Stream actual cells without expanding gaps to millions of empty cells.

    Isolate openpyxl's sparse parser here (the packaged version is pinned).
    """
    wb = sheet.parent
    with sheet._get_source() as source:
        parser = WorkSheetParser(source, sheet._shared_strings, data_only=wb.data_only,
                                 epoch=wb.epoch, date_formats=wb._date_formats,
                                 timedelta_formats=wb._timedelta_formats)
        for _, cells in parser.parse():
            yield from cells


def _xml(data: bytes) -> ET.Element:
    if b'<!DOCTYPE' in data.upper() or b'<!ENTITY' in data.upper():
        raise PdfValidationError('XML entities are not supported')
    try:
        return ET.fromstring(data)
    except ET.ParseError as exc:
        raise PdfValidationError('Invalid workbook XML') from exc


def _serialize(root: ET.Element) -> bytes:
    # ElementTree reassigns namespace prefixes. Ignore only compatibility hints
    # and alternate-content extensions, whose prefix-list values would otherwise
    # point at removed namespace declarations. Standard worksheet data stays.
    compatibility = 'http://schemas.openxmlformats.org/markup-compatibility/2006'
    for parent in root.iter():
        parent.attrib.pop('{%s}Ignorable' % compatibility, None)
        for child in list(parent):
            if child.tag == '{%s}AlternateContent' % compatibility:
                parent.remove(child)
    return ET.tostring(root, encoding='utf-8', xml_declaration=True)


def _freeze_formula(cell: ET.Element, strings: list[str]) -> None:
    value = cell.find('m:v', NS)
    if value is None or value.text is None:
        raise PdfValidationError('External formula has no cached value; save the workbook in Excel first')
    kind, text = cell.get('t', 'n'), value.text
    if kind == 'e':
        raise PdfValidationError('External formula cache contains an Excel error')
    if kind in ('str', 's'):
        if kind == 's':
            try:
                text = strings[int(text)]
            except (ValueError, IndexError) as exc:
                raise PdfValidationError('Invalid cached shared string') from exc
        cell.set('t', 'inlineStr')
        cell.remove(value)
        inline = ET.SubElement(cell, '{%s}is' % MAIN)
        node = ET.SubElement(inline, '{%s}t' % MAIN)
        node.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
        node.text = text
    elif kind == 'b':
        if text not in ('0', '1'):
            raise PdfValidationError('Invalid cached boolean')
    elif kind == 'n':
        try:
            if not math.isfinite(float(text)):
                raise ValueError()
        except ValueError as exc:
            raise PdfValidationError('Invalid cached numeric value') from exc
    else:
        raise PdfValidationError('Unsupported external formula cache type')
    for formula in list(cell.findall('m:f', NS)):
        cell.remove(formula)


def sanitize_workbook(source: Path, destination: Path) -> WorkbookReport:
    """Write an inert copy and report recoverable problems by sheet.

    Only XLSX packages without macros, embedded objects, data connections or
    linked media are accepted. External workbook parts and hyperlinks are
    removed. Shared external formulas are frozen including their followers.
    """
    source, destination = Path(source), Path(destination)
    if source.resolve() == destination.resolve() or source.suffix.lower() != '.xlsx':
        raise PdfValidationError('Expected a separate .xlsx source and temporary destination')
    if source.stat().st_size > 40 * 1024 * 1024:
        raise PdfValidationError('Workbook exceeds 40 MB')
    try:
        with zipfile.ZipFile(source) as package:
            infos = package.infolist()
            names = [item.filename for item in infos]
            if len(names) != len(set(names)) or len(names) > 5000:
                raise PdfValidationError('Duplicate or excessive workbook parts')
            if sum(item.file_size for item in infos) > 150 * 1024 * 1024:
                raise PdfValidationError('Uncompressed workbook exceeds 150 MB')
            for item in infos:
                parts = PurePosixPath(item.filename).parts
                lower = item.filename.lower()
                if '..' in parts or item.filename.startswith('/') or '\\' in item.filename or item.flag_bits & 1:
                    raise PdfValidationError('Unsafe workbook package path')
                if any(token in lower for token in ('vbaproject', 'activex/', 'embeddings/', 'querytables/', 'connections.xml', 'macrosheets/', 'dialogsheets/')):
                    raise PdfValidationError('Macros, objects or data connections are unsupported')
            if 'xl/workbook.xml' not in names or '[Content_Types].xml' not in names:
                raise PdfValidationError('Not an XLSX workbook')
            contents = {item.filename: package.read(item) for item in infos if not item.is_dir()}
    except zipfile.BadZipFile as exc:
        raise PdfValidationError('Invalid XLSX package') from exc
    report = WorkbookReport()
    # Workbook metadata must be readable; individual worksheets may be replaced
    # with inert placeholders, preserving sheet identities and dependencies.
    workbook = _xml(contents['xl/workbook.xml'])
    relationships = _xml(contents.get('xl/_rels/workbook.xml.rels', b'<Relationships/>'))
    targets = {r.get('Id'): posixpath.normpath('xl/' + r.get('Target', ''))
               if not r.get('Target', '').startswith('/') else r.get('Target').lstrip('/')
               for r in relationships}
    sheets = {targets.get(s.get('{%s}id' % DOCREL)): s.get('name')
              for s in workbook.findall('m:sheets/m:sheet', NS)}
    titles = list(sheets.values())
    if not titles or None in sheets or any(not t for t in titles):
        raise PdfValidationError('Projektmappens arkoversigt kunne ikke læses.')
    empty_sheet = '<worksheet xmlns="%s"><sheetData/></worksheet>' % MAIN
    roots = {}
    changed = set()
    for name, value in contents.items():
        if not name.endswith(('.xml', '.rels', '.vml')):
            continue
        try:
            roots[name] = _xml(value)
        except PdfValidationError:
            if name not in sheets or b'<!DOCTYPE' in value.upper() or b'<!ENTITY' in value.upper():
                raise
            roots[name] = ET.fromstring(empty_sheet)
            changed.add(name)
            report.error(sheets[name], 'Arkets XML', 'Arket kunne ikke læses og er udeladt.')
    for name, title in sheets.items():
        if name not in roots:
            roots[name] = ET.fromstring(empty_sheet)
            contents[name] = empty_sheet.encode()
            changed.add(name)
            report.error(title, 'Arkets XML', 'Arkets indhold mangler i filen og er udeladt.')
    workbook = roots['xl/workbook.xml']
    types = roots['[Content_Types].xml']
    if any('macroenabled' in item.get('ContentType', '').lower() or 'vba' in item.get('ContentType', '').lower()
           for item in types):
        raise PdfValidationError('Macro-enabled workbook content is unsupported')
    strings = []
    if 'xl/sharedStrings.xml' in roots:
        strings = [''.join(item.itertext()) for item in roots['xl/sharedStrings.xml']]
    removed = {name for name in contents if name.startswith('xl/externalLinks/') or name == 'xl/calcChain.xml'}
    defined_names = workbook.find('m:definedNames', NS)
    name_keys = {}
    definitions = []
    for element in list(defined_names) if defined_names is not None else []:
        local = element.get('localSheetId')
        try:
            scope = titles[int(local)] if local is not None else None
        except (ValueError, IndexError):
            raise PdfValidationError('Et navngivet område har ugyldig arktilknytning.') from None
        key = ('name', scope, element.get('name', '').casefold())
        name_keys[(scope, key[2])] = key
        definitions.append((element, scope, key))
    graph = {title: set() for title in titles}
    broken_names = {}
    for element, scope, key in definitions:
        expression = element.text or ''
        graph[key] = _formula_dependencies(expression, scope, titles, name_keys)
        if EXTERNAL.search(expression) or UNSAFE_FORMULA.search(expression) or '#REF!' in expression.upper():
            defined_names.remove(element)
            changed.add('xl/workbook.xml')
            report.warning(scope, element.get('name', ''), 'Ekstern eller ugyldig reference i navngivet område fjernet fra arbejdskopien.',
                           review=key[2] == '_xlnm.print_titles')
            # Print areas are overridden by our exporters. Missing print titles
            # affect appearance only and require review, not a failed import.
            if key[2] not in ('_xlnm.print_area', '_xlnm.print_titles', '_xlnm._filterdatabase'):
                broken_names[key] = element.get('name', '')

    # Determine owners of chart/table parts through package relationships.
    owners = {part: {title} for part, title in sheets.items()}
    links = {}
    for part, root in roots.items():
        if not part.endswith('.rels') or '/_rels/' not in part:
            continue
        directory, filename = part.rsplit('/_rels/', 1)
        owner = directory + '/' + filename[:-5]
        links[owner] = {posixpath.normpath(directory + '/' + r.get('Target', ''))
                        if not r.get('Target', '').startswith('/') else r.get('Target').lstrip('/')
                        for r in root if r.get('TargetMode', '').casefold() != 'external'}
    for _ in range(len(links) + 1):
        updated = False
        for owner, children in links.items():
            for child in children:
                before = len(owners.get(child, set()))
                owners.setdefault(child, set()).update(owners.get(owner, set()))
                updated |= len(owners[child]) != before
        if not updated:
            break

    for name, root in roots.items():
        if name in removed:
            continue
        if name.endswith('.rels'):
            for relationship in list(root):
                kind = relationship.get('Type', '').rsplit('/', 1)[-1]
                if kind in ('externalLink', 'calcChain', 'hyperlink'):
                    root.remove(relationship)
                    changed.add(name)
                elif relationship.get('TargetMode', '').casefold() == 'external':
                    raise PdfValidationError('External linked media or objects are unsupported')
        if name.startswith('xl/worksheets/') and name.endswith('.xml'):
            title = sheets.get(name)
            if title is None:
                raise PdfValidationError('Et regneark mangler tilknytning til projektmappen.')
            shared = {}
            for formula in root.findall('.//m:c/m:f', NS):
                if formula.get('t') == 'shared' and formula.text:
                    shared[formula.get('si')] = formula.text
            for cell in root.findall('.//m:c', NS):
                formula = cell.find('m:f', NS)
                if formula is None:
                    continue
                expression = formula.text or shared.get(formula.get('si'), '')
                graph[title].update(_formula_dependencies(expression, title, titles, name_keys))
                if not expression or EXTERNAL.search(expression) or UNSAFE_FORMULA.search(expression) or '#REF!' in expression.upper():
                    try:
                        if not expression or formula.get('t') not in (None, 'normal', 'shared') or '#REF!' in expression.upper():
                            raise PdfValidationError('unsupported expression')
                        _freeze_formula(cell, strings)
                        report.frozen += 1
                        report.warning(title, cell.get('r', '?'), 'Formel erstattet med Excels gemte værdi. Kontrollér tallet i PDF’en.')
                    except PdfValidationError:
                        # An error value, never a blank/zero that looks correct.
                        for child in list(cell):
                            cell.remove(child)
                        cell.set('t', 'e')
                        ET.SubElement(cell, '{%s}v' % MAIN).text = '#REF!'
                        report.error(title, cell.get('r', '?'), 'Formlen kan ikke bruges, og en pålidelig gemt værdi mangler.', expression)
                    changed.add(name)
            for cell in root.findall('.//m:c', NS):
                if cell.get('t') == 'e' and not any(i['sheet'] == title and i['cell'] == cell.get('r') for i in report.issues):
                    report.warning(title, cell.get('r', '?'),
                        'Excel har gemt en fejl i cellen (' + cell.findtext('m:v', '', NS) + '). Kontrollér cellen i PDF’en efter genberegning.', review=True)
            for linkgroup in list(root.findall('m:hyperlinks', NS)):
                root.remove(linkgroup)
                changed.add(name)
        # Input validation is irrelevant to PDF output. Removing conditional
        # formatting can affect appearance and requires a PDF review.
        parents = {child: parent for parent in root.iter() for child in parent}
        for element in list(root.iter()):
            local_name = element.tag.rsplit('}', 1)[-1]
            if local_name in ('definedName', 'formula', 'formula1', 'formula2', 'f', 'calculatedColumnFormula', 'totalsRowFormula'):
                expression = element.text or ''
                # Non-cell formulas and names can themselves refer to bad names.
                affected = owners.get(name, set())
                if not (EXTERNAL.search(expression) or UNSAFE_FORMULA.search(expression)):
                    for title in affected:
                        graph[title].update(_formula_dependencies(expression, title, titles, name_keys))
                    continue
                ancestor = element
                while ancestor in parents and ancestor.tag.rsplit('}', 1)[-1] not in ('conditionalFormatting', 'dataValidation'):
                    ancestor = parents[ancestor]
                kind = ancestor.tag.rsplit('}', 1)[-1]
                if kind in ('conditionalFormatting', 'dataValidation') and ancestor in parents:
                    if ancestor in list(parents[ancestor]):
                        parents[ancestor].remove(ancestor)
                    for title in affected:
                        report.warning(title, ancestor.get('sqref', kind),
                            'Ekstern regel for betinget formatering fjernet. Gennemgå PDF’ens udseende.' if kind == 'conditionalFormatting'
                            else 'Ekstern regel for dataindtastning fjernet fra arbejdskopien.', review=kind == 'conditionalFormatting')
                else:
                    if not affected:
                        raise PdfValidationError(f'En ekstern reference i {name} kunne ikke knyttes til et bestemt ark.')
                    element.text = '#REF!'
                    for title in affected:
                        report.error(title, name, 'Ekstern formel uden en brugbar gemt værdi i arkets indhold.', expression)
                changed.add(name)
        for group in root.findall('m:dataValidations', NS):
            group.set('count', str(len(group)))
    for references in list(workbook.findall('m:externalReferences', NS)):
        workbook.remove(references)
        changed.add('xl/workbook.xml')
    for item in list(types):
        if item.get('PartName', '').lstrip('/') in removed:
            types.remove(item)
            changed.add('[Content_Types].xml')
    def write_package(target, mode):
        with zipfile.ZipFile(target, mode, compression=zipfile.ZIP_DEFLATED) as package:
            for part, value in contents.items():
                if part not in removed:
                    package.writestr(part, _serialize(roots[part]) if part in changed else value)

    # Lazy reading isolates malformed cell values to one sheet, before either
    # adapter or PDF engine opens the complete workbook.
    temporary = io.BytesIO()
    write_package(temporary, 'w')
    temporary.seek(0)
    wb = load_workbook(temporary, read_only=True, data_only=False, keep_links=False)
    try:
        for part, title in sheets.items():
            try:
                ws = wb[title]
                for _ in sheet_cells(ws):
                    pass
            except Exception:
                roots[part] = ET.fromstring(empty_sheet)
                changed.add(part)
                report.error(title, 'Arkets celler', 'Arket indeholder ulæselige celledata og er udeladt.')
    finally:
        wb.close()

    bad = {i['sheet'] for i in report.issues} | set(broken_names)
    while True:
        more = {node for node, dependencies in graph.items() if dependencies & bad} - bad
        if not more:
            break
        for node in more:
            if isinstance(node, str):
                dependencies = sorted(str(broken_names.get(d, d if isinstance(d, str) else d[2])) for d in graph[node] & bad)
                report.error(node, 'Formelreference', 'Arket bruger data fra et fejlramt ark eller navngivet område: ' + ', '.join(dependencies))
        bad.update(more)
    report.dependencies = {title: sorted(d for d in graph[title] if isinstance(d, str)) for title in titles}
    destination.parent.mkdir(parents=True, exist_ok=True)
    write_package(destination, 'x')
    return report


def _references(formula: object) -> set[str]:
    if not isinstance(formula, str) or not formula.startswith('='):
        return set()
    quoted = re.findall(r"'((?:[^']|'')+)'!", formula)
    masked = re.sub(r"'(?:[^']|'')+'!", ' ', formula)
    plain = re.findall(r'(?<![\w.])([A-Za-z_][\w.]*)!', masked)
    return {item.replace("''", "'") for item in quoted} | set(plain)


def _print_area(sheet, months: dict) -> tuple[int, int]:
    last_row, last_col = max(months.values()), 1
    for (row, col), cell in sheet._cells.items():
        if cell.value is None or isinstance(cell.value, str) and not cell.value.strip():
            continue
        if row > 6000 or col > 60:
            raise PdfValidationError('Registration content outside supported print bounds')
        last_row = max(last_row, row)
        if row >= months['januar']:
            last_col = max(last_col, col)
    if last_row <= months['december'] or last_col < 2:
        raise PdfValidationError('Registration contains no December data or total')
    try:
        area = str(sheet.print_area).split('!')[-1].replace('$', '')
        mincol, minrow, maxcol, maxrow = range_boundaries(area)
        if mincol == 1 and minrow <= months['januar'] and maxrow <= 6000 and maxcol <= 60:
            last_row, last_col = max(last_row, maxrow), max(last_col, maxcol)
    except (ValueError, TypeError, AttributeError):
        pass
    return last_row, last_col


def libreoffice_runtime(config: dict) -> tuple[Path, Path, dict]:
    """Resolve native runtime explicitly; never use the service user's PATH."""
    program = Path(config.get('libreoffice_program') or r'C:\Program Files\LibreOffice\program').resolve()
    if program.name.casefold() == 'soffice.exe':
        program = program.parent
    soffice = program / 'soffice.exe'
    candidates = [Path(config['libreoffice_python'])] if config.get('libreoffice_python') else [program / 'python.exe', *program.glob('python-core-*/bin/python.exe')]
    python = next((item for item in candidates if item.is_file()), None)
    if not soffice.is_file() or python is None:
        raise PdfValidationError('Install native LibreOffice with its Python/UNO runtime; configure libreoffice_program and optionally libreoffice_python')
    env = os.environ.copy()
    env['PATH'] = str(program) + os.pathsep + env.get('PATH', '')
    env['PYTHONPATH'] = str(program)
    env.pop('PYTHONHOME', None)
    env['URE_BOOTSTRAP'] = 'vnd.sun.star.pathname:' + str(program / 'fundamental.ini')
    env['PYTHONIOENCODING'] = 'utf-8'
    return soffice, python.resolve(), env


def check_runtime(config: dict) -> dict:
    soffice, python, env = libreoffice_runtime(config)
    result = subprocess.run([str(python), '-c', 'import uno; print("UNO_READY")'], env=env,
                            capture_output=True, text=True, timeout=30,
                            creationflags=getattr(subprocess, 'CREATE_NO_WINDOW', 0))
    if result.returncode or 'UNO_READY' not in result.stdout:
        raise PdfValidationError('Configured LibreOffice Python cannot import UNO')
    return {'soffice': str(soffice), 'python': str(python), 'uno': True}


def _render(sanitized: Path, work: Path, export: dict, config: dict) -> list[Path]:
    soffice, python, env = libreoffice_runtime(config)
    manifest = work / 'manifest.json'
    spec = {'source': str(sanitized), 'tempdir': str(work), 'manifest': str(manifest),
            'soffice': str(soffice), **export}
    configfile = work / 'render.json'
    configfile.write_text(json.dumps(spec), encoding='utf-8')
    process = subprocess.Popen([str(python), str(Path(__file__).with_name('pdf_worker.py')), str(configfile)],
                               env=env, stdout=subprocess.PIPE, stderr=subprocess.PIPE,
                               creationflags=getattr(subprocess, 'CREATE_NO_WINDOW', 0))
    try:
        process.communicate(timeout=240)
    except subprocess.TimeoutExpired as exc:
        # Terminate only this worker's process tree, including its isolated
        # LibreOffice instance; do not leave an orphan holding payroll files.
        if os.name == 'nt':
            subprocess.run(['taskkill.exe', '/PID', str(process.pid), '/T', '/F'],
                           capture_output=True, timeout=20,
                           creationflags=getattr(subprocess, 'CREATE_NO_WINDOW', 0))
        if process.poll() is None:
            process.kill()
        process.communicate()
        raise PdfValidationError('LibreOffice export exceeded its time limit') from exc
    if process.returncode or not manifest.is_file():
        raise PdfValidationError('LibreOffice export failed; no document was published')
    segments = [Path(item).resolve() for item in json.loads(manifest.read_text(encoding='utf-8'))['pdfs']]
    if len(segments) != 13 or len(set(segments)) != 13 or any(item.parent != work.resolve() for item in segments):
        raise PdfValidationError('Expected overview and exactly twelve monthly PDF sections')
    return segments


