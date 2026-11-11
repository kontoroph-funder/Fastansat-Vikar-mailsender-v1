from __future__ import annotations

import asyncio
from contextlib import asynccontextmanager, suppress
import os
from pathlib import Path
import secrets
import threading
from typing import Literal

from fastapi import FastAPI, File, Form, Request, UploadFile
from fastapi.responses import FileResponse, JSONResponse
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel, ConfigDict, Field, field_validator

from .domain import local_now
from .service import Service

STATIC = Path(__file__).parent / 'static'


class Strict(BaseModel):
    model_config = ConfigDict(extra='forbid')


class Period(Strict):
    month: int = Field(ge=1, le=12)
    year: int = Field(ge=1900, le=2200)


class Edit(Strict):
    subject: str | None = Field(default=None, max_length=250)
    body: str | None = Field(default=None, max_length=50000)
    selected: bool | None = None
    reset: bool = False
    reviewed: bool = False


class Approval(Strict):
    ids: list[str] = Field(min_length=1, max_length=5000)
    revision: int


class Send(Strict):
    approval_id: str


class Reconcile(Strict):
    outcome: Literal['accepted', 'retry']


class CopyRecipient(Strict):
    cc: str = Field(max_length=254)


class Template(Strict):
    subject: str = Field(min_length=1, max_length=250)
    body: str = Field(min_length=1, max_length=50000)


class Templates(Strict):
    fastansat: Template
    vikar: Template


class Settings(Strict):
    client_id: str = Field(max_length=36, pattern=r'^(?:[0-9a-fA-F-]{36})?$')
    tenant_id: str = Field(min_length=1, max_length=150, pattern=r'^[a-zA-Z0-9.-]+$')
    libreoffice_program: str = Field(max_length=500)
    libreoffice_python: str = Field(max_length=500)
    templates: Templates

    @field_validator('templates')
    @classmethod
    def subjects(cls, value):
        for template in (value.fastansat, value.vikar):
            if '\n' in template.subject or '\r' in template.subject or not template.subject.strip() or not template.body.strip():
                raise ValueError('Udfyld skabeloner uden linjeskift i emnet.')
        return value


def create_app(root=None, service=None):
    root = Path(root or os.environ.get('PAYROLL_DATA_DIR') or Path(os.environ.get('LOCALAPPDATA', str(Path.home()))) / 'FastansatVikar' / 'data')
    csrf = secrets.token_urlsafe(32)

    @asynccontextmanager
    async def lifespan(app):
        app.state.service = service or Service(root)
        async def housekeeping():
            while True:
                await asyncio.sleep(60)
                await asyncio.to_thread(app.state.service.cleanup)
        task = asyncio.create_task(housekeeping())
        yield
        task.cancel()
        with suppress(asyncio.CancelledError):
            await task
        await asyncio.to_thread(app.state.service.close)

    app = FastAPI(title='Fastansatte & Vikarer', lifespan=lifespan, docs_url=None, redoc_url=None, openapi_url=None)

    @app.middleware('http')
    async def local_only(request: Request, call_next):
        host = request.headers.get('host', '')
        if host not in ('127.0.0.1:9000', 'localhost:9000'):
            return JSONResponse({'error': 'Ugyldig lokal vært.'}, status_code=403)
        origin = request.headers.get('origin')
        expected_origin = 'http://' + host
        if origin and origin != expected_origin:
            return JSONResponse({'error': 'Kald fra andre websteder er ikke tilladt.'}, status_code=403)
        if request.headers.get('sec-fetch-site') == 'cross-site':
            return JSONResponse({'error': 'Åbn appen via startgenvejen.'}, status_code=403)
        if request.method not in ('GET', 'HEAD'):
            if not secrets.compare_digest(request.headers.get('x-csrf-token', ''), csrf):
                return JSONResponse({'error': 'Sessionen er udløbet. Genindlæs siden.'}, status_code=403)
        try:
            length = int(request.headers.get('content-length', '0'))
        except ValueError:
            return JSONResponse({'error': 'Ugyldig filstørrelse.'}, status_code=400)
        if length > 26 * 1024 * 1024:
            return JSONResponse({'error': 'Filen er for stor (maks. 25 MB).'}, status_code=413)
        response = await call_next(request)
        response.headers['Cache-Control'] = 'no-store'
        response.headers['X-Content-Type-Options'] = 'nosniff'
        response.headers['Referrer-Policy'] = 'no-referrer'
        response.headers['Content-Security-Policy'] = "default-src 'self'; script-src 'self'; style-src 'self'; img-src 'self' data:; frame-src 'self'; object-src 'none'; base-uri 'none'; frame-ancestors 'none'; form-action 'self'"
        return response

    @app.exception_handler(ValueError)
    async def value_error(request, exc):
        return JSONResponse({'error': str(exc)}, status_code=400)

    @app.get('/')
    def index():
        return FileResponse(STATIC / 'index.html')

    def svc():
        return app.state.service

    @app.get('/api/session')
    def session():
        now = local_now()
        return {'csrf': csrf, 'month': now.month, 'year': now.year, 'account': svc().graph.account()}

    @app.get('/api/settings')
    def settings():
        return svc().store.settings()

    @app.put('/api/settings')
    def save_settings(body: Settings):
        with svc().lock:
            svc().ensure_account_change_allowed()
            if svc().graph.login and svc().graph.login['status'] == 'waiting':
                raise ValueError('Vent til Microsoft-login er afsluttet.')
            svc().store.save_settings(body.model_dump())
        return {'saved': True}

    @app.get('/api/jobs')
    def jobs():
        return svc().list()

    @app.post('/api/jobs')
    async def create(kind: str = Form(), month: int = Form(), year: int = Form(), file: UploadFile = File(), cc: str = Form(default='')):
        raw = await file.read(25 * 1024 * 1024 + 1)
        await file.close()
        return svc().create(kind, month, year, file.filename or '', raw, cc)

    @app.patch('/api/jobs/{job_id}/cc')
    def cc(job_id: str, body: CopyRecipient):
        return svc().cc(job_id, body.cc)

    @app.get('/api/jobs/{job_id}')
    def job(job_id: str):
        return svc().get(job_id)

    @app.patch('/api/jobs/{job_id}/period')
    def period(job_id: str, body: Period):
        return svc().period(job_id, body.month, body.year)

    @app.patch('/api/jobs/{job_id}/recipients/{row_id}')
    def edit(job_id: str, row_id: str, body: Edit):
        return svc().edit(job_id, row_id, body.model_dump(exclude_none=True, exclude_unset=True))

    @app.get('/api/jobs/{job_id}/recipients/{row_id}/pdf')
    def pdf(job_id: str, row_id: str, download: bool = False):
        with svc().lock:
            job = svc().store.get(job_id)
            if job.get('cleaned') or job['expires'] <= __import__('time').time():
                raise ValueError('PDF’en er udløbet eller slettet.')
            row = svc().row(job, row_id)
            if not row['sendable']:
                raise ValueError('Medarbejderens PDF er ikke valideret.')
            path = svc().pdf_path(job_id, row)
            return FileResponse(path, media_type='application/pdf', filename='Timeoversigt.pdf',
                                content_disposition_type='attachment' if download else 'inline')

    @app.post('/api/jobs/{job_id}/approve')
    def approve(job_id: str, body: Approval):
        return svc().approve(job_id, body.ids, body.revision)

    @app.post('/api/jobs/{job_id}/send')
    def send(job_id: str, body: Send):
        return svc().send(job_id, body.approval_id)

    @app.post('/api/jobs/{job_id}/cancel')
    def cancel(job_id: str):
        return svc().cancel(job_id)

    @app.post('/api/jobs/{job_id}/recipients/{row_id}/reconcile')
    def reconcile(job_id: str, row_id: str, body: Reconcile):
        return svc().reconcile(job_id, row_id, body.outcome)

    @app.post('/api/jobs/{job_id}/recipients/{row_id}/test')
    def test(job_id: str, row_id: str):
        return svc().test_mail(job_id, row_id)

    @app.post('/api/auth/connect')
    def connect():
        with svc().lock:
            svc().ensure_account_change_allowed()
            return svc().graph.start_login()

    @app.get('/api/auth/status')
    def auth_status():
        return {'account': svc().graph.account(), 'login': svc().graph.login}

    @app.post('/api/auth/logout')
    def logout():
        with svc().lock:
            svc().ensure_account_change_allowed()
            svc().graph.logout()
        return {'ok': True}

    @app.get('/api/runtime')
    def runtime():
        from .runtime import check_runtime
        return check_runtime(svc().store.settings())

    @app.post('/api/shutdown')
    def shutdown():
        with svc().lock:
            if any(j['state'] in ('ANALYZING', 'SENDING') or j.get('test_status') == 'SENDING' for j in svc().store.all()):
                raise ValueError('Vent til analyse eller udsendelse er afsluttet, før du lukker appen.')
            callback = getattr(app.state, 'shutdown', None)
            if callback:
                threading.Timer(.5, callback).start()
        return {'stopped': True}

    app.mount('/static', StaticFiles(directory=STATIC), name='static')
    return app
