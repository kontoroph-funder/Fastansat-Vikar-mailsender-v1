"""End-to-end UI test with synthetic PDFs and an in-memory mail transport."""
import json
import socket
from pathlib import Path
import sys
import threading
import time
import uuid

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(ROOT / 'tests'))

from playwright.sync_api import sync_playwright, expect
import uvicorn

from app.api import create_app
from app.service import Service
from conftest import FakeGraph, fake_renderer, fast_workbook, vikar_workbook
from openpyxl import load_workbook
from openpyxl.workbook.defined_name import DefinedName


def main():
    with socket.socket() as probe:
        if probe.connect_ex(('127.0.0.1', 9000)) == 0:
            raise RuntimeError('Port 9000 er optaget. Luk appen før browsertesten.')
    run = ROOT / '.test-data' / 'browser' / uuid.uuid4().hex[:8]
    run.mkdir(parents=True)
    workbook = fast_workbook(run / 'oktober-2026.xlsx', mismatch=True)
    vikar = vikar_workbook(run / 'vikar.xlsx', hours='7,5')
    wb = load_workbook(vikar)
    wb.defined_names.add(DefinedName('OldExternalLink', attr_text="'[Old.xlsx]Sheet1'!A1"))
    wb['Bo Vikar']['C10'] = "='[Old.xlsx]Sheet1'!A1"
    wb.save(vikar)
    wb.close()
    graph = FakeGraph()
    service = Service(run / 'data', graph=graph, renderer=fake_renderer)
    server = uvicorn.Server(uvicorn.Config(create_app(service=service), host='127.0.0.1', port=9000, access_log=False, log_level='warning'))
    thread = threading.Thread(target=server.run, daemon=True)
    thread.start()
    errors = []
    try:
        for _ in range(100):
            if server.started:
                break
            time.sleep(.1)
        with sync_playwright() as p:
            browser = p.chromium.launch_persistent_context(str(run / 'profile'), headless=True,
                executable_path=r'C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe', viewport={'width':1500,'height':1000})
            page = browser.pages[0]
            page.on('pageerror', lambda error: errors.append(str(error)))
            page.goto('http://127.0.0.1:9000')
            expect(page.locator('#month option')).to_have_count(12)
            expect(page.locator('#year option')).to_have_count(11)
            page.locator('#month').select_option('1')
            page.locator('#year').select_option('2027')
            page.locator('#cc').fill('copy@example.com')
            page.locator('#fileInput').set_input_files(workbook)
            page.locator('#analyze').click()
            expect(page.locator('#statReady')).to_have_text('1', timeout=15000)
            expect(page.locator('#errorPanel')).to_be_visible()
            expect(page.locator('#errorList')).to_contain_text('Registrering 2!R2')
            expect(page.locator('#errorList')).to_contain_text('Oversigt 2!S2')
            expect(page.locator('#employees input[type=checkbox]').nth(1)).to_be_disabled()
            expect(page.locator('#periodLabel')).to_contain_text('december 2026')
            page.locator('#employees .employee-name').first.click()
            page.locator('#subject').fill('En personlig hilsen')
            page.locator('#body').fill('Hej Anna\n\nHer er din personlige besked.')
            page.locator('#saveMessage').click()
            expect(page.locator('#saveStatus')).to_have_text('Gemt')
            page.locator('#month').select_option('2')
            expect(page.locator('#reviewBox')).to_be_visible()
            expect(page.locator('#subject')).to_have_value('En personlig hilsen')
            page.locator('#markReviewed').click()
            expect(page.locator('#reviewBox')).not_to_be_visible()
            page.evaluate('window.scrollTo(0,0)')
            page.screenshot(path=str(run / 'overview-desktop.png'), full_page=True)
            page.set_viewport_size({'width':390,'height':844})
            page.screenshot(path=str(run / 'overview-mobile.png'), full_page=True)
            assert page.evaluate('document.documentElement.scrollWidth <= window.innerWidth'), 'Mobile horizontal overflow'
            page.set_viewport_size({'width':1500,'height':1000})
            page.reload()
            page.get_by_role('button', name='Historik & kladder').click()
            page.get_by_role('button', name='Åbn kørsel').click()
            expect(page.locator('#month')).to_have_value('2')
            expect(page.locator('#cc')).to_have_value('copy@example.com')
            page.locator('#cc').fill('new-copy@example.com')
            page.locator('#employees .employee-name').first.click()
            expect(page.locator('#subject')).to_have_value('En personlig hilsen')
            page.locator('#sendSelected').click()
            expect(page.locator('#approvalDialog')).to_be_visible()
            expect(page.locator('#approvalRecipients')).to_contain_text('Anna Eksempel')
            expect(page.locator('#approvalExcluded')).to_contain_text('Bo Eksempel')
            expect(page.locator('#approvalCc')).to_contain_text('new-copy@example.com')
            page.locator('#confirmSend').click()
            expect(page.locator('#systemStatus')).to_have_text('Afsluttet', timeout=15000)
            assert len(graph.sent) == 1 and graph.sent[0][0] == 'person1@example.com'
            assert graph.sent[0][1] == 'En personlig hilsen'
            assert graph.sent[0][6] == 'new-copy@example.com'
            page.locator('#newJob').click()
            page.locator('#kind').select_option('vikar')
            page.locator('#cc').fill('test-excluded@example.com')
            page.locator('#fileInput').set_input_files(vikar)
            page.locator('#analyze').click()
            expect(page.locator('#statReady')).to_have_text('1', timeout=15000)
            expect(page.locator('#statError')).to_have_text('1')
            expect(page.locator('#warningPanel')).to_be_visible()
            expect(page.locator('#errorList')).to_contain_text('Bo Vikar!C10')
            page.locator('#employees .employee-name').first.click()
            page.once('dialog', lambda dialog: dialog.accept())
            page.locator('#testMail').click()
            expect(page.locator('#progress')).to_contain_text('Testmail: Accepteret af Outlook', timeout=15000)
            assert graph.sent[-1][0] == 'administrator@example.com'
            assert len(graph.sent[-1]) == 6
            page.get_by_role('button', name='Indstillinger', exact=True).click()
            expect(page.locator('#templateFastBody')).to_be_visible()
            page.evaluate('window.scrollTo(0,0)')
            page.screenshot(path=str(run / 'settings.png'), full_page=True)
            assert not errors, errors
            browser.close()
        print(json.dumps({'passed':True,'browser_errors':errors,'simulated_emails':len(graph.sent),'artifacts':str(run)},ensure_ascii=False))
    finally:
        server.should_exit = True
        thread.join(timeout=30)


if __name__ == '__main__':
    main()
