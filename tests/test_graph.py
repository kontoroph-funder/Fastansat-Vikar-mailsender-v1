from unittest.mock import Mock

import pytest
import requests

from app.domain import defaults
from app.graph import Graph, DefiniteFailure, UnknownDelivery, payload
from app.vault import SecretStore


@pytest.mark.parametrize('status,exception', [(202,None),(400,DefiniteFailure),(401,DefiniteFailure),(403,DefiniteFailure),(429,DefiniteFailure),(500,UnknownDelivery),(302,UnknownDelivery)])
def test_graph_outcomes_and_exact_payload(tmp_path, monkeypatch, status, exception):
    config = defaults()
    config['client_id'] = '00000000-0000-0000-0000-000000000001'
    graph = Graph(tmp_path, lambda: config)
    account = {'id':'account','email':'sender@example.com','name':'Sender'}
    graph.account = lambda: account
    graph.vault = Mock()
    app, cache = Mock(), Mock()
    app.get_accounts.return_value = [{'home_account_id':'account'}]
    app.acquire_token_silent.return_value = {'access_token':'fake-token'}
    cache.serialize.return_value = '{}'
    graph._app = lambda settings: (app, cache)
    response = Mock(status_code=status)
    post = Mock(return_value=response)
    monkeypatch.setattr(requests, 'post', post)
    args = ('person@example.com','Emne','Personlig tekst','test.pdf',b'%PDF-test',account,'copy@example.com')
    if exception:
        with pytest.raises(exception):
            graph.send(*args)
    else:
        graph.send(*args)
    data = __import__('json').loads(post.call_args.kwargs['data'])
    assert data['message']['body']['content'] == 'Personlig tekst'
    assert data['message']['toRecipients'][0]['emailAddress']['address'] == 'person@example.com'
    assert data['message']['ccRecipients'] == [{'emailAddress': {'address': 'copy@example.com'}}]
    assert post.call_args.kwargs['allow_redirects'] is False
    assert post.call_count == 1


def test_cc_validation_and_duplicate_recipient():
    import json
    args = ('a@example.com', 'Emne', 'Tekst', 'a.pdf', b'%PDF-ok')
    with pytest.raises(DefiniteFailure, match='Cc'):
        payload(*args, 'not-an-email')
    assert json.loads(payload(*args, 'A@example.com'))['message']['ccRecipients'] == []
    assert json.loads(payload(*args))['message']['ccRecipients'] == []


def test_payload_rejects_header_injection_and_large_pdf():
    with pytest.raises(DefiniteFailure):
        payload('a@example.com','Emne\r\nBcc: b@example.com','Text','a.pdf',b'%PDF-ok')
    with pytest.raises(DefiniteFailure):
        payload('a@example.com','Emne','Text','a.pdf',b'%PDF-'+b'0'*3_000_000)


@pytest.mark.parametrize('stage', ['discovery', 'device_flow'])
def test_login_connection_failure_is_actionable_and_retryable(tmp_path, stage):
    graph = Graph(tmp_path, defaults)
    app = Mock()
    failure = requests.ConnectionError('blocked socket')
    graph._app = Mock(side_effect=failure) if stage == 'discovery' else Mock(return_value=(app, Mock()))
    app.initiate_device_flow.side_effect = failure
    for _ in range(2):
        with pytest.raises(ValueError, match='login.microsoftonline.com'):
            graph.start_login()
        assert graph.login is None


def test_login_missing_client_id_is_validation_error(tmp_path):
    graph = Graph(tmp_path, defaults)
    with pytest.raises(ValueError, match='client-id'):
        graph.start_login()


def test_dpapi_roundtrip(tmp_path):
    import os
    if os.name != 'nt':
        pytest.skip('Windows DPAPI only')
    vault = SecretStore(tmp_path)
    vault.set('test', {'token':'synthetic-test-token'})
    assert b'synthetic-test-token' not in (tmp_path/'test.dpapi').read_bytes()
    assert vault.get('test') == {'token':'synthetic-test-token'}
    vault.delete('test')
    assert vault.get('test') is None
