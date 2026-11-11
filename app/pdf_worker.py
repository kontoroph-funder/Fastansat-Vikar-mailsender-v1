"""Private worker executed by LibreOffice's native, UNO-compatible Python.

No third-party Python packages are required in this interpreter. The parent
supplies only a sanitized local workbook. UNO constants follow the official
MediaDescriptor API: UpdateDocMode=NO_UPDATE and MacroExecutionMode=NEVER_EXECUTE.
"""
import json
import os
from pathlib import Path
import socket
import subprocess
import sys
import time
import uuid

import uno
from com.sun.star.beans import PropertyValue


def pv(name, value):
    item = PropertyValue()
    item.Name, item.Value = name, value
    return item


def connect(port, process):
    local = uno.getComponentContext()
    resolver = local.ServiceManager.createInstanceWithContext('com.sun.star.bridge.UnoUrlResolver', local)
    for _ in range(100):
        try:
            context = resolver.resolve('uno:socket,host=127.0.0.1,port=%d;urp;StarOffice.ComponentContext' % port)
            return context.ServiceManager.createInstanceWithContext('com.sun.star.frame.Desktop', context)
        except Exception:
            if process.poll() is not None and process.returncode:
                break
            time.sleep(.25)
    raise RuntimeError('LibreOffice connection failed')


def set_style(document, sheet, landscape):
    # Separate style per export sheet: a shared source PageStyle must not let
    # the landscape registration change the portrait overview unexpectedly.
    styles = document.StyleFamilies.getByName('PageStyles')
    name = 'Payroll_' + uuid.uuid4().hex
    source = styles.getByName(sheet.PageStyle)
    style = document.createInstance('com.sun.star.style.PageStyle')
    styles.insertByName(name, style)
    # Preserve headers/footers and other printable source attributes where
    # writable, then override only pagination and privacy-related properties.
    for item in source.getPropertySetInfo().Properties:
        if item.Attributes & 16:
            continue
        try:
            style.setPropertyValue(item.Name, source.getPropertyValue(item.Name))
        except Exception:
            pass
    sheet.PageStyle = name
    style.IsLandscape = bool(landscape)
    style.Width = 29700 if landscape else 21000
    style.Height = 21000 if landscape else 29700
    style.ScaleToPagesX, style.ScaleToPagesY = 1, 0
    style.LeftMargin = style.RightMargin = 900
    style.TopMargin = style.BottomMargin = 1000
    style.PrintAnnotations = False
    return style


def export(document, output):
    document.storeToURL(uno.systemPathToFileUrl(str(output)), (
        pv('FilterName', 'calc_pdf_Export'), pv('Overwrite', False),
        pv('FilterData', uno.Any('[]com.sun.star.beans.PropertyValue', (
            pv('ExportBookmarks', False), pv('ExportNotes', False),
            pv('ExportHiddenSheets', False), pv('SinglePageSheets', False),
        ))),
    ))
    if not output.is_file() or output.stat().st_size < 100:
        raise RuntimeError('No PDF generated')


def main(config_path):
    config = json.loads(Path(config_path).read_text(encoding='utf-8'))
    source = Path(config['source']).resolve()
    directory = Path(config['tempdir']).resolve()
    profile = directory / 'profile'
    with socket.socket() as listener:
        listener.bind(('127.0.0.1', 0))
        port = listener.getsockname()[1]
    process = subprocess.Popen([
        config['soffice'], '-env:UserInstallation=' + profile.as_uri(),
        '--headless', '--nologo', '--nodefault', '--nofirststartwizard', '--norestore',
        '--accept=socket,host=127.0.0.1,port=%d;urp;StarOffice.ComponentContext' % port,
    ], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
       creationflags=getattr(subprocess, 'CREATE_NO_WINDOW', 0))
    document, desktop = None, None
    try:
        desktop = connect(port, process)
        document = desktop.loadComponentFromURL(uno.systemPathToFileUrl(str(source)), '_blank', 0, (
            pv('Hidden', True), pv('ReadOnly', True), pv('UpdateDocMode', 0),
            pv('MacroExecutionMode', 0),
        ))
        if document is None:
            raise RuntimeError('Workbook could not be loaded')
        sheets = document.Sheets
        overview, registration = sheets.getByName(config['overview']), sheets.getByName(config['data'])
        overview.IsVisible = True
        for sheet in sheets:
            sheet.setPrintAreas(())
            sheet.IsVisible = sheet.Name == overview.Name
            for index in range(sheet.Annotations.Count):
                sheet.Annotations.getByIndex(index).IsVisible = False
        styles = document.StyleFamilies.getByName('PageStyles')
        for index in range(styles.Count):
            style = styles.getByIndex(index)
            if style.getPropertySetInfo().hasPropertyByName('PrintAnnotations'):
                style.PrintAnnotations = False
        set_style(document, overview, False)
        overview.setPrintAreas((overview.getCellRangeByName(config['overview_area']).RangeAddress,))
        output = directory / '00_overview.pdf'
        export(document, output)
        outputs = [str(output)]
        registration.IsVisible = True
        overview.IsVisible = False
        overview.setPrintAreas(())
        set_style(document, registration, True)
        for index, month in enumerate(config['months']):
            registration.setPrintAreas((registration.getCellRangeByName(month['area']).RangeAddress,))
            output = directory / ('%02d.pdf' % (index + 1))
            export(document, output)
            outputs.append(str(output))
        Path(config['manifest']).write_text(json.dumps({'pdfs': outputs}), encoding='utf-8')
    finally:
        if document is not None:
            try:
                document.close(True)
            except Exception:
                document.dispose()
        if desktop is not None:
            try:
                desktop.terminate()
            except Exception:
                pass
        if process.poll() is None:
            process.terminate()
            try:
                process.wait(timeout=10)
            except subprocess.TimeoutExpired:
                process.kill()
                process.wait(timeout=10)


if __name__ == '__main__':
    main(sys.argv[1])
