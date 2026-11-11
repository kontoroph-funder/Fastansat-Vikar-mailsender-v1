import copy
import time

import pytest

from app.domain import period_text
from app.service import Service
from conftest import create_job, fast_workbook, vikar_workbook, wait, FakeGraph, fake_renderer


def test_partial_workbook_sends_only_valid_employee(service, tmp_path):
    job = create_job(service, fast_workbook(tmp_path / 'mixed.xlsx', mismatch=True))
    good, bad = job['recipients'][:2]
    assert good['sendable'] and not bad['sendable']
    with pytest.raises(ValueError):
        service.edit(job['id'], bad['id'], {'selected': True})
    with pytest.raises(ValueError):
        service.approve(job['id'], [bad['id']], job['revision'])
    approval = service.approve(job['id'], [good['id']], job['revision'])
    service.send(job['id'], approval['id'])
    wait(service)
    assert len(service.graph.sent) == 1
    assert service.graph.sent[0][0] == 'person1@example.com'
    receipt = service.get(job['id'])
    assert receipt['state'] == 'COMPLETE'
    assert receipt['cleaned'] and not receipt['recipients']
    assert not service.directory(job['id']).exists()


def test_edit_period_restart_and_reset(service, tmp_path):
    job = create_job(service, fast_workbook(tmp_path / 'period.xlsx'))
    row = job['recipients'][0]
    service.edit(job['id'], row['id'], {'subject': 'Personligt emne', 'body': 'Hej Anna, en individuel besked.'})
    changed = service.period(job['id'], 1, 2027)
    assert changed['recipients'][0]['subject'] == 'Personligt emne'
    assert changed['recipients'][0]['needs_review']
    assert '16. december 2026 – 15. januar 2027' in changed['recipients'][1]['subject']
    with pytest.raises(ValueError):
        service.approve(job['id'], [row['id']], changed['revision'])
    service.close()
    reopened = Service(service.root, graph=service.graph, renderer=fake_renderer)
    try:
        assert reopened.get(job['id'])['recipients'][0]['body'] == 'Hej Anna, en individuel besked.'
        reset = reopened.edit(job['id'], row['id'], {'reset': True})
        assert not reset['recipients'][0]['custom']
        assert not reset['recipients'][0]['needs_review']
        assert 'januar 2027' in reset['recipients'][0]['subject']
    finally:
        reopened.close()


def test_approval_invalidated_by_edit_pdf_change_or_account(service, tmp_path):
    job = create_job(service, fast_workbook(tmp_path / 'approve.xlsx'))
    first = job['recipients'][0]
    approval = service.approve(job['id'], [first['id']], job['revision'])
    job = service.edit(job['id'], first['id'], {'body': 'Ny tekst'})
    with pytest.raises(ValueError):
        service.send(job['id'], approval['id'])
    approval = service.approve(job['id'], [first['id']], job['revision'])
    service.graph.identity = {'id': 'other', 'email': 'other@example.com'}
    with pytest.raises(ValueError):
        service.send(job['id'], approval['id'])
    path = service.directory(job['id']) / 'output' / first['pdf']
    path.write_bytes(b'%PDF-changed')
    with pytest.raises(ValueError, match='ændret'):
        service.approve(job['id'], [first['id']], job['revision'])


def test_double_click_and_partial_send_recovery(service, tmp_path):
    job = create_job(service, fast_workbook(tmp_path / 'send.xlsx', people=('Anna','Bo','Carl')))
    ids = [r['id'] for r in job['recipients']]
    service.graph.fail_at = 1
    approval = service.approve(job['id'], ids, job['revision'])
    service.send(job['id'], approval['id'])
    with pytest.raises(ValueError):
        service.send(job['id'], approval['id'])
    wait(service)
    paused = service.get(job['id'])
    assert [r['status'] for r in paused['recipients']] == ['ACCEPTED','UNKNOWN','READY']
    assert len(service.graph.sent) == 1
    with pytest.raises(ValueError):
        service.approve(job['id'], [ids[2]], paused['revision'])
    service.graph.fail_at = None
    job = service.reconcile(job['id'], ids[1], 'retry')
    approval = service.approve(job['id'], ids[1:], job['revision'])
    service.send(job['id'], approval['id'])
    wait(service)
    assert len(service.graph.sent) == 3
    assert [s[0] for s in service.graph.sent].count('person1@example.com') == 1


def test_restart_turns_inflight_into_unknown_without_sending(service, tmp_path):
    job = create_job(service, fast_workbook(tmp_path / 'crash.xlsx'))
    stored = service.store.get(job['id'])
    stored['state'] = 'SENDING'
    stored['recipients'][0]['status'] = 'SENDING'
    service.store.put(stored)
    service.close()
    reopened = Service(service.root, graph=service.graph, renderer=fake_renderer)
    try:
        job = reopened.get(job['id'])
        assert job['state'] == 'PAUSED'
        assert job['recipients'][0]['status'] == 'UNKNOWN'
        assert not service.graph.sent
    finally:
        reopened.close()


def test_testmail_is_to_self_and_does_not_change_delivery(service, tmp_path):
    job = create_job(service, vikar_workbook(tmp_path / 'test.xlsx'), kind='vikar')
    service.test_mail(job['id'], job['recipients'][0]['id'])
    wait(service)
    assert service.graph.sent[0][0] == 'administrator@example.com'
    assert service.graph.sent[0][1].startswith('[TEST]')
    assert service.get(job['id'])['recipients'][0]['status'] == 'READY'


def test_template_changes_affect_only_future_jobs(service, tmp_path):
    job = create_job(service, fast_workbook(tmp_path / 'template.xlsx'))
    settings = service.store.settings()
    settings['templates']['fastansat']['subject'] = 'Ændret standardskabelon'
    service.store.save_settings(settings)
    reset = service.edit(job['id'], job['recipients'][0]['id'], {'reset':True})
    assert reset['recipients'][0]['subject'] != 'Ændret standardskabelon'


def test_expiry_scrubs_files_and_content(service, tmp_path):
    job = create_job(service, fast_workbook(tmp_path / 'expire.xlsx'))
    stored = service.store.get(job['id'])
    stored['expires'] = time.time() - 1
    service.store.put(stored)
    service.cleanup()
    receipt = service.get(job['id'])
    assert receipt['state'] == 'EXPIRED'
    assert 'Anna' not in str(receipt)
    assert not service.directory(job['id']).exists()


def test_pdf_failure_only_blocks_affected_person(service, tmp_path):
    def render(source, output, row, kind, config):
        if row['name'].startswith('Anna'):
            raise ValueError('Simuleret PDF-fejl')
        return fake_renderer(source, output, row, kind, config)
    service.renderer = render
    job = create_job(service, fast_workbook(tmp_path / 'pdf.xlsx'))
    assert not job['recipients'][0]['sendable']
    assert job['recipients'][1]['sendable']


def test_original_input_unchanged_and_same_pdf_sent(service, tmp_path):
    path = fast_workbook(tmp_path / 'unchanged.xlsx')
    original = path.read_bytes()
    job = create_job(service, path)
    first = job['recipients'][0]
    pdf = service.pdf_path(job['id'], first).read_bytes()
    job = service.edit(job['id'], first['id'], {'subject':'Individuelt emne','body':'Individuel besked'})
    approval = service.approve(job['id'], [first['id']], job['revision'])
    service.send(job['id'], approval['id'])
    wait(service)
    sent = service.graph.sent[0]
    assert sent[1:3] == ('Individuelt emne','Individuel besked')
    assert sent[4] == pdf
    assert path.read_bytes() == original


def test_unknown_recipient_and_reimport_are_rejected(service, tmp_path):
    path = fast_workbook(tmp_path / 'repeat.xlsx')
    job = create_job(service, path)
    with pytest.raises(ValueError):
        service.approve(job['id'], ['not-a-recipient'], job['revision'])
    with pytest.raises(ValueError, match='allerede'):
        service.create('fastansat',10,2026,path.name,path.read_bytes())


def test_interrupted_upload_directory_expires(service):
    import os
    orphan = service.directory('a' * 32)
    orphan.mkdir(parents=True)
    (orphan / 'input.xlsx').write_bytes(b'partial upload')
    old = time.time() - 49 * 3600
    os.utime(orphan, (old, old))
    service.cleanup()
    assert not orphan.exists()


def test_expired_unsent_draft_can_be_imported_again(service, tmp_path):
    path = fast_workbook(tmp_path / 'expired-unsent.xlsx')
    job = create_job(service, path)
    stored = service.store.get(job['id'])
    stored['expires'] = time.time() - 1
    service.store.put(stored)
    service.cleanup()
    new_job = create_job(service, path)
    assert new_job['id'] != job['id']
    assert new_job['state'] == 'DRAFT'


def test_cc_saved_invalidates_approval_and_is_excluded_from_test(service, tmp_path):
    job = create_job(service, vikar_workbook(tmp_path / 'cc.xlsx'), 'vikar')
    first = job['recipients'][0]
    approval = service.approve(job['id'], [first['id']], job['revision'])
    job = service.cc(job['id'], 'copy@example.com')
    assert service.store.get(job['id'])['cc'] == 'copy@example.com'
    with pytest.raises(ValueError):
        service.send(job['id'], approval['id'])
    with pytest.raises(ValueError):
        service.cc(job['id'], 'copy@example.com\r\nBcc: hidden@example.com')
    service.test_mail(job['id'], first['id'])
    wait(service)
    assert service.graph.sent[0][0] == 'administrator@example.com'
    assert len(service.graph.sent[0]) == 6  # No Cc argument for test mail.
    approval = service.approve(job['id'], [first['id']], job['revision'])
    service.send(job['id'], approval['id'])
    wait(service)
    assert service.graph.sent[1][0] == 'vikar1@example.com'
    assert service.graph.sent[1][6] == 'copy@example.com'
    service.close()
    reopened = Service(service.root, graph=service.graph, renderer=fake_renderer)
    try:
        assert reopened.get(job['id'])['cc'] == 'copy@example.com'
    finally:
        reopened.close()


def test_legacy_draft_without_cc_can_send(service, tmp_path):
    job = create_job(service, fast_workbook(tmp_path / 'legacy.xlsx'))
    old = service.store.get(job['id'])
    old.pop('cc')
    service.store.put(old)
    approval = service.approve(job['id'], [job['recipients'][0]['id']], job['revision'])
    service.send(job['id'], approval['id'])
    wait(service)
    assert service.graph.sent[0][6] == ''
