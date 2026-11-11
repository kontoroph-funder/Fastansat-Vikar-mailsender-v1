from fastapi.testclient import TestClient

from app.api import create_app
from conftest import create_job, fast_workbook


def test_api_enforces_csrf_origin_and_blocked_recipient(service, tmp_path):
    job = create_job(service, fast_workbook(tmp_path / 'api.xlsx', mismatch=True))
    with TestClient(create_app(service=service), base_url='http://127.0.0.1:9000') as client:
        assert client.get('/').status_code == 200
        assert client.get('/api/session', headers={'Origin':'https://attacker.example'}).status_code == 403
        assert client.get('/api/jobs', headers={'Host':'attacker.example'}).status_code == 403
        assert client.post('/api/jobs/'+job['id']+'/cancel').status_code == 403
        csrf = client.get('/api/session').json()['csrf']
        client.headers['X-CSRF-Token'] = csrf
        bad = job['recipients'][1]
        response = client.patch(f"/api/jobs/{job['id']}/recipients/{bad['id']}",json={'selected':True})
        assert response.status_code == 400
        response = client.post(f"/api/jobs/{job['id']}/approve",json={'ids':[bad['id']],'revision':job['revision']})
        assert response.status_code == 400
        assert not service.graph.sent
        assert client.get('/state.sqlite3').status_code == 404
        assert client.get('/static/../store.py').status_code == 404


def test_api_month_year_and_unknown_fields(service, tmp_path):
    job = create_job(service, fast_workbook(tmp_path / 'date.xlsx'))
    with TestClient(create_app(service=service), base_url='http://127.0.0.1:9000') as client:
        client.headers['X-CSRF-Token'] = client.get('/api/session').json()['csrf']
        response = client.patch(f"/api/jobs/{job['id']}/period",json={'month':1,'year':2027})
        assert response.status_code == 200
        assert 'december 2026' in response.json()['period']
        row = job['recipients'][0]
        response = client.patch(f"/api/jobs/{job['id']}/recipients/{row['id']}",json={'email':'other@example.com'})
        assert response.status_code == 422
        assert client.patch(f"/api/jobs/{job['id']}/period",json={'month':13,'year':2027}).status_code == 422


def test_cc_upload_edit_clear_and_validation(service, tmp_path):
    path = fast_workbook(tmp_path / 'cc-api.xlsx')
    with TestClient(create_app(service=service), base_url='http://127.0.0.1:9000') as client:
        client.headers['X-CSRF-Token'] = client.get('/api/session').json()['csrf']
        response = client.post('/api/jobs', data={'kind':'fastansat','month':10,'year':2026,'cc':'copy@example.com'},
                               files={'file':('cc-api.xlsx', path.read_bytes())})
        assert response.status_code == 200
        job = response.json()
        assert job['cc'] == 'copy@example.com'
        from conftest import wait
        wait(service)
        endpoint = f"/api/jobs/{job['id']}/cc"
        assert client.patch(endpoint, json={'cc':'invalid'}).status_code == 400
        assert client.patch(endpoint, json={'cc':''}).json()['cc'] == ''
        assert client.patch(endpoint, json={'cc':'new@example.com'}).json()['cc'] == 'new@example.com'
