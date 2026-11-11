import json
import subprocess
from pathlib import Path
from types import SimpleNamespace as NS
from unittest.mock import Mock

import pytest
from fastapi.testclient import TestClient

from app.api import create_app
from app.graph import DefiniteFailure, UnknownDelivery
from app.outlook import Outlook
from app.outlook_worker import execute, identity
from conftest import create_job, fast_workbook, wait


class Collection:
    def __init__(self, items):
        self.items = items
        self.Count = len(items)
    def Item(self, index):
        return self.items[index - 1]


def application():
    sender = NS(SmtpAddress='sender@example.com', DisplayName='Sender', DeliveryStore=NS(StoreID='store'))
    session = NS(CurrentProfileName='profile', Accounts=Collection([sender]), Offline=False)
    mail = Mock()
    mail.SendUsingAccount = sender
    mail.Recipients.ResolveAll.return_value = True
    app = NS(Session=session, CreateItem=Mock(return_value=mail))
    request = dict(operation='send', account=identity(session, sender), email='employee@example.com',
                   cc='copy@example.com', subject='Personlig besked', body='Tekst', pdf='C:/synthetic.pdf')
    return app, mail, request


def test_worker_sends_exact_account_and_attachment():
    app, mail, request = application()
    assert execute(app, request) == {'status': 'accepted'}
    assert mail.Subject == request['subject'] and mail.Body == request['body']
    assert [c.args[0] for c in mail.Recipients.Add.call_args_list] == ['employee@example.com', 'copy@example.com']
    mail.Attachments.Add.assert_called_once_with('C:/synthetic.pdf', 1)
    assert mail._oleobj_.Invoke.call_args.args[-1] is app.Session.Accounts.Item(1)
    mail.Send.assert_called_once()


@pytest.mark.parametrize('change', ['missing', 'duplicate', 'profile', 'store', 'offline'])
def test_worker_never_sends_from_fallback_account(change):
    app, mail, request = application()
    account = app.Session.Accounts.Item(1)
    if change == 'missing': app.Session.Accounts = Collection([])
    if change == 'duplicate': app.Session.Accounts = Collection([account, account])
    if change == 'profile': app.Session.CurrentProfileName = 'another profile'
    if change == 'store': account.DeliveryStore.StoreID = 'other store'
    if change == 'offline': app.Session.Offline = True
    assert execute(app, request)['status'] == 'failed'
    mail.Send.assert_not_called()


def test_uncertain_send_does_not_retry():
    app, mail, request = application()
    mail.Send.side_effect = RuntimeError('COM disconnected')
    assert execute(app, request)['status'] == 'unknown'
    assert mail.Send.call_count == 1


def test_unresolved_recipient_is_not_sent():
    app, mail, request = application()
    mail.Recipients.ResolveAll.return_value = False
    with pytest.raises(ValueError): execute(app, request)
    mail.Send.assert_not_called()


def configure(service):
    _, _, request = application()
    account = request['account']
    fake = Mock()
    fake.account.return_value = account
    fake.verify.return_value = account
    fake.accounts.return_value = [account]
    service.outlook = fake
    config = service.store.settings()
    config.update(mail_method='outlook', outlook_account=account)
    service.save_settings(config)
    return fake, account


def test_outlook_job_send_receipt_and_no_graph(service, tmp_path):
    fake, account = configure(service)
    job = create_job(service, fast_workbook(tmp_path/'new.xlsx'))
    assert job['mail_method'] == 'outlook'
    approval = service.approve(job['id'], [r['id'] for r in job['recipients']], job['revision'])
    service.send(job['id'], approval['id'])
    wait(service)
    assert fake.send.call_count == 2
    assert not service.graph.sent
    receipt = service.get(job['id'])
    assert receipt['cleaned'] and receipt['mail_method'] == 'outlook'


def test_unknown_stops_remaining_recipients(service, tmp_path):
    fake, _ = configure(service)
    fake.send.side_effect = UnknownDelivery('timeout')
    job = create_job(service, fast_workbook(tmp_path/'unknown.xlsx'))
    approval = service.approve(job['id'], [r['id'] for r in job['recipients']], job['revision'])
    service.send(job['id'], approval['id'])
    wait(service)
    job = service.get(job['id'])
    assert job['state'] == 'PAUSED'
    assert [r['status'] for r in job['recipients']] == ['UNKNOWN', 'READY']
    assert fake.send.call_count == 1 and not job['cleaned']


def test_existing_job_cannot_silently_change_transport(service, tmp_path):
    job = create_job(service, fast_workbook(tmp_path/'old.xlsx'))
    stored = service.store.get(job['id'])
    stored.pop('mail_method')  # A draft from the previous release.
    service.store.put(stored)
    configure(service)
    job = service.get(job['id'])
    with pytest.raises(ValueError, match='Graph'):
        service.approve(job['id'], [job['recipients'][0]['id']], job['revision'])
    assert service.get(job['id'])['mail_method'] == 'graph'


def test_setting_changes_invalidate_approval_and_test_uses_selected_account(service, tmp_path):
    fake, account = configure(service)
    job = create_job(service, fast_workbook(tmp_path/'settings.xlsx'))
    service.test_mail(job['id'], job['recipients'][0]['id'])
    wait(service)
    args = fake.send.call_args.args
    assert args[0] == account['email'] and len(args) == 6  # No CC for tests.
    approval = service.approve(job['id'], [job['recipients'][0]['id']], job['revision'])
    config = service.store.settings()
    config['mail_method'] = 'graph'
    service.save_settings(config)
    config['mail_method'] = 'outlook'
    service.save_settings(config)
    with pytest.raises(ValueError, match='godkendelsen|Godkendelsen'):
        service.send(job['id'], approval['id'])


def test_api_account_discovery_and_invalid_account(service):
    fake, _ = configure(service)
    with TestClient(create_app(service=service), base_url='http://127.0.0.1:9000') as client:
        assert client.post('/api/outlook/accounts').status_code == 403
        client.headers['X-CSRF-Token'] = client.get('/api/session').json()['csrf']
        assert len(client.post('/api/outlook/accounts').json()['accounts']) == 1
        config = client.get('/api/settings').json()
        config['outlook_account']['id'] = 'b'*64
        fake.verify.side_effect = DefiniteFailure('Konto mangler')
        assert client.put('/api/settings', json=config).status_code == 400


@pytest.mark.parametrize('operation,error', [('send', UnknownDelivery), ('accounts', DefiniteFailure)])
def test_worker_timeout_is_bounded_and_kills_only_worker(tmp_path, monkeypatch, operation, error):
    worker = Mock()
    worker.wait.side_effect = [subprocess.TimeoutExpired('worker', 1), 0]
    monkeypatch.setattr(subprocess, 'Popen', Mock(return_value=worker))
    transport = Outlook(tmp_path, lambda: {})
    with pytest.raises(error): transport.call({'operation': operation})
    worker.kill.assert_called_once()


def test_temporary_pdf_removed_even_on_unknown_result(tmp_path):
    account = {'id': 'test'}
    transport = Outlook(tmp_path, lambda: {'outlook_account': account})
    def call(request):
        assert Path(request['pdf']).read_bytes() == b'%PDF-test'
        raise UnknownDelivery('timeout')
    transport.call = call
    with pytest.raises(UnknownDelivery):
        transport.send('a@example.com', 'Emne', 'Tekst', 'Timeoversigt.pdf', b'%PDF-test', account)
    assert not list(transport.root.iterdir())
