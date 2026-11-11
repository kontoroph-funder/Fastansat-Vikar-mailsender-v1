from __future__ import annotations

import copy
from concurrent.futures import ThreadPoolExecutor
import hashlib
from pathlib import Path
import re
import threading
import time
import uuid

from .adapters import analyze_fastansat, analyze_vikar, employee, issue
from .domain import approval_digest, message, period_text, validate_period, validate_cc
from .files import remove_tree
from .graph import Graph, DefiniteFailure, payload
from .payroll_pdf import sanitize_workbook
from .rendering import render
from .store import Store
from .outlook import Outlook

TERMINAL = {'COMPLETE', 'CANCELLED', 'EXPIRED'}
LOCKED_ROWS = {'SENDING', 'ACCEPTED', 'UNKNOWN', 'FAILED'}


class Service:
    def __init__(self, root, graph=None, renderer=render, outlook=None):
        self.root = Path(root).resolve()
        self.store = Store(self.root)
        self.lock = threading.RLock()
        self.worker = ThreadPoolExecutor(max_workers=1, thread_name_prefix='payroll')
        self.graph = graph or Graph(self.root, self.store.settings)
        self.outlook = outlook or Outlook(self.root, self.store.settings)
        self.renderer = renderer
        self.futures = []
        for job in self.store.all():
            if job.get('cleaned'):
                continue
            if job['state'] == 'ANALYZING':
                job.update(state='ERROR', error='Analysen blev afbrudt. Annullér kørslen og upload filen igen.')
            elif job['state'] == 'SENDING':
                job.update(state='PAUSED', error='Afsendelsen blev afbrudt. Kontrollér status før fortsættelse.')
            for row in job['recipients']:
                if row['status'] == 'SENDING':
                    row['status'] = 'UNKNOWN'
            if job.get('test_status') == 'SENDING':
                job['test_status'] = 'UNKNOWN'
                job['test_error'] = 'Testmail blev afbrudt. Kontrollér Udbakke og Sendt post.'
            job['approval'] = None
            self.store.put(job)
        self.cleanup()

    def method(self):
        return self.store.settings().get('mail_method', 'graph')

    def transport(self, job=None):
        method = self.method()
        if job is not None and job.get('mail_method', 'graph') != method:
            raise ValueError('Denne kørsel bruger ' + ('klassisk Outlook' if job.get('mail_method') == 'outlook' else 'Microsoft Graph') + '. Vælg samme afsendelsesmetode under Indstillinger.')
        return self.outlook if method == 'outlook' else self.graph

    def account(self, job=None, verify=False):
        transport = self.transport(job)
        account = transport.account()
        if verify and self.method() == 'outlook':
            try:
                self.outlook.verify(account)
            except DefiniteFailure as exc:
                raise ValueError(str(exc)) from None
        return account

    def validate_message(self, job, *args):
        return payload(*args, enforce_graph_limit=job.get('mail_method', 'graph') == 'graph')

    def save_settings(self, settings):
        with self.lock:
            self.ensure_account_change_allowed()
            if self.graph.login and self.graph.login['status'] == 'waiting':
                raise ValueError('Vent til Microsoft-login er afsluttet.')
            old = self.store.settings()
            if settings['mail_method'] == 'outlook' and settings.get('outlook_account') and settings.get('outlook_account') != old.get('outlook_account'):
                try:
                    self.outlook.verify(settings['outlook_account'])
                except DefiniteFailure as exc:
                    raise ValueError(str(exc)) from None
            if any(old.get(k) != settings.get(k) for k in ('mail_method', 'outlook_account', 'client_id', 'tenant_id')):
                for job in self.store.all():
                    if not job.get('cleaned'):
                        self.touch(job)
                        self.store.put(job)
            self.store.save_settings(settings)

    def submit(self, fn, *args):
        self.futures = [f for f in self.futures if not f.done()]
        future = self.worker.submit(fn, *args)
        self.futures.append(future)
        return future

    def close(self):
        self.worker.shutdown(wait=True)

    def directory(self, job_id):
        if not re.fullmatch(r'[a-f0-9]{32}', job_id):
            raise ValueError('Ugyldigt kørsels-ID.')
        return self.root / 'jobs' / job_id

    def public(self, job):
        result = copy.deepcopy(job)
        result.pop('approval', None)
        result.pop('template', None)
        for row in result.get('recipients', []):
            for key in list(row):
                if key.startswith('_'):
                    row.pop(key)
        result['mail_method'] = job.get('mail_method', 'graph')
        result['period'] = period_text(result['month'], result['year'])
        return result

    def get(self, job_id):
        with self.lock:
            return self.public(self.store.get(job_id))

    def list(self):
        with self.lock:
            return [self.public(j) for j in self.store.all()]

    def create(self, kind, month, year, filename, raw, cc=''):
        if kind not in ('fastansat', 'vikar'):
            raise ValueError('Vælg Fastansatte eller Vikarer.')
        validate_period(month, year)
        cc = validate_cc(cc)
        if not filename.lower().endswith('.xlsx') or not 0 < len(raw) <= 25 * 1024 * 1024:
            raise ValueError('Upload én .xlsx-fil på højst 25 MB.')
        with self.lock:
            self.cleanup()
            source_hash = hashlib.sha256(raw).hexdigest()
            for old in self.store.all():
                expired_unsent = old['state'] == 'EXPIRED' and not any(old.get('counts', {}).get(s, 0) for s in ('ACCEPTED', 'UNKNOWN'))
                failed_import = old['state'] == 'ERROR' and not any(r['status'] in ('ACCEPTED', 'UNKNOWN', 'SENDING') for r in old.get('recipients', []))
                if old.get('source_hash') == source_hash and old['kind'] == kind and old['month'] == month and old['year'] == year and old['state'] != 'CANCELLED' and not expired_unsent and not failed_import:
                    raise ValueError('Denne fil og periode er allerede indlæst. Åbn den eksisterende kørsel i Historik.')
            job_id = uuid.uuid4().hex
            directory = self.directory(job_id)
            directory.mkdir(parents=True)
            (directory / 'input.xlsx').write_bytes(raw)
            job = {'id': job_id, 'mail_method': self.method(), 'kind': kind, 'month': month, 'year': year, 'cc': cc, 'warnings': [],
                   'filename': Path(filename).name, 'source_hash': source_hash, 'created': time.time(),
                   'expires': time.time() + 48 * 3600, 'revision': 1, 'state': 'ANALYZING', 'error': '',
                   'progress': 'Kontrollerer Excel-filen…', 'recipients': [], 'approval': None,
                   'template': self.store.settings()['templates'][kind], 'test_status': '', 'cleaned': False}
            self.store.put(job)
            self.submit(self.analyze, job_id)
            return self.public(job)

    def analyze(self, job_id):
        job = self.store.get(job_id)
        directory = self.directory(job_id)
        try:
            sanitized = directory / 'sanitized.xlsx'
            report = sanitize_workbook(directory / 'input.xlsx', sanitized)
            job['warnings'] = report.warnings
            try:
                rows = (analyze_fastansat if job['kind'] == 'fastansat' else analyze_vikar)(sanitized)
            except ValueError:
                if job['kind'] != 'fastansat' or not report.issues:
                    raise
                rows = []
            for problem in report.issues:
                affected = [r for r in rows if problem['sheet'] in (r['sheet'], r.get('overview'))]
                if not affected:
                    affected = [employee(problem['sheet'], problem['sheet'], '')]
                    rows.extend(affected)
                for row in affected:
                    issue(row, 'WORKBOOK_DATA', problem['message'], [{k: problem[k] for k in ('sheet', 'cell', 'value')}])
            if not rows:
                raise ValueError('Ingen medarbejderark blev fundet.')
            config = self.store.settings()
            for index, row in enumerate(rows):
                row.update(message(job['template'], row, job['month'], job['year']))
                for warning in report.warnings:
                    if warning['sheet'] is None or warning['sheet'] in (row['sheet'], row.get('overview')):
                        row['warnings'].append(f"{warning['sheet'] or 'Projektmappe'}!{warning['cell']}: {warning['message']}")
                        if warning.get('review'):
                            row['needs_pdf_review'] = True
                if not row['errors']:
                    try:
                        row.update(self.renderer(sanitized, directory / 'output', row, job['kind'], config))
                        self.validate_message(job, row['email'], row['subject'], row['body'], 'timeoversigt.pdf',
                                (directory / 'output' / row['pdf']).read_bytes())
                        row.update(sendable=True, selected=True, status='READY')
                    except Exception as exc:
                        issue(row, 'PDF_ERROR', f'PDF kunne ikke klargøres: {exc}')
                row.pop('_export', None)
                with self.lock:
                    job['progress'] = f'Behandlet {index + 1} af {len(rows)} ark'
                    job['recipients'] = rows[:index + 1]
                    self.store.put(job)
            job.update(state='DRAFT', progress='', recipients=rows)
        except Exception as exc:
            job.update(state='ERROR', error=f'Filen kunne ikke behandles: {exc}', progress='')
        with self.lock:
            self.store.put(job)

    def editable(self, job):
        if job['state'] not in ('DRAFT', 'PAUSED') or job.get('cleaned') or job['expires'] <= time.time():
            raise ValueError('Kørslen kan ikke ændres nu.')
        if job.get('test_status') == 'SENDING':
            raise ValueError('Vent til testmailen er afsluttet.')

    @staticmethod
    def touch(job):
        job['revision'] += 1
        job['approval'] = None

    @staticmethod
    def row(job, row_id):
        row = next((r for r in job['recipients'] if r['id'] == row_id), None)
        if row is None:
            raise ValueError('Medarbejderen findes ikke i kørslen.')
        return row

    def period(self, job_id, month, year):
        validate_period(month, year)
        with self.lock:
            job = self.store.get(job_id)
            self.editable(job)
            if any(r['status'] in LOCKED_ROWS for r in job['recipients']):
                raise ValueError('Perioden kan ikke ændres efter et afsendelsesforsøg. Opret en ny kørsel.')
            if (month, year) == (job['month'], job['year']):
                return self.public(job)
            job.update(month=month, year=year)
            for row in job['recipients']:
                if row['custom']:
                    row['needs_review'] = True
                else:
                    row.update(message(job['template'], row, month, year))
            self.touch(job)
            self.store.put(job)
            return self.public(job)

    def edit(self, job_id, row_id, changes):
        with self.lock:
            job = self.store.get(job_id)
            self.editable(job)
            row = self.row(job, row_id)
            if row['status'] in LOCKED_ROWS:
                raise ValueError('Denne afsendelse er låst. Afklar eventuelle fejl først.')
            if changes.get('selected') and not row['sendable']:
                raise ValueError('Medarbejderen har fejl og kan ikke vælges til udsendelse.')
            if changes.get('reset'):
                row.update(message(job['template'], row, job['month'], job['year']))
                row.update(custom=False, needs_review=False)
            elif 'subject' in changes or 'body' in changes:
                subject = changes.get('subject', row['subject']).strip()
                body = changes.get('body', row['body']).strip()
                if not subject or '\n' in subject or '\r' in subject or not body or len(subject) > 250 or len(body) > 50000:
                    raise ValueError('Udfyld emne (højst 250 tegn uden linjeskift) og besked (højst 50.000 tegn).')
                row.update(subject=subject, body=body, custom=True, needs_review=False)
            if changes.get('reviewed'):
                row['needs_review'] = False
                row['needs_pdf_review'] = False
            if 'selected' in changes:
                row['selected'] = bool(changes['selected'])
            self.touch(job)
            self.store.put(job)
            return self.public(job)

    def cc(self, job_id, value):
        value = validate_cc(value)
        with self.lock:
            job = self.store.get(job_id)
            self.editable(job)
            if value != job.get('cc', ''):
                job['cc'] = value
                self.touch(job)
                self.store.put(job)
            return self.public(job)

    def pdf_path(self, job_id, row):
        root = (self.directory(job_id) / 'output').resolve()
        path = (root / row['pdf']).resolve()
        if path.parent != root or path.name != row['id'] + '.pdf':
            raise ValueError('Ugyldig PDF-reference.')
        if not path.is_file() or hashlib.sha256(path.read_bytes()).hexdigest() != row['sha256']:
            raise ValueError('PDF’en mangler eller er ændret. Opret en ny kørsel.')
        return path

    def approve(self, job_id, ids, revision):
        with self.lock:
            job = self.store.get(job_id)
            self.editable(job)
            if revision != job['revision']:
                raise ValueError('Kørslen er ændret. Genindlæs og gennemgå den igen.')
            account = self.account(job, verify=True)
            if not account:
                raise ValueError('Vælg en afsenderkonto under Indstillinger før godkendelse.')
            if not ids or len(ids) != len(set(ids)):
                raise ValueError('Vælg mindst én modtager uden dubletter.')
            if any(r['status'] in ('UNKNOWN', 'FAILED') for r in job['recipients']):
                raise ValueError('Afklar tidligere afsendelsesfejl før næste udsendelse.')
            for row_id in ids:
                row = self.row(job, row_id)
                if not row['sendable'] or row['errors'] or row['status'] != 'READY' or row['needs_review'] or row.get('needs_pdf_review'):
                    raise ValueError('Alle valgte medarbejdere skal være fejlfrie og gennemgået.')
                self.validate_message(job, row['email'], row['subject'], row['body'], 'timeoversigt.pdf', self.pdf_path(job_id, row).read_bytes(), job.get('cc', ''))
            approval = {'id': uuid.uuid4().hex, 'ids': ids, 'account': account,
                        'digest': approval_digest(job, ids, account), 'created': time.time()}
            job['approval'] = approval
            self.store.put(job)
            return approval

    def send(self, job_id, approval_id):
        with self.lock:
            job = self.store.get(job_id)
            self.editable(job)
            approval = job.get('approval')
            if not approval or approval['id'] != approval_id or approval['account'] != self.account(job) or approval['digest'] != approval_digest(job, approval['ids'], approval['account']):
                raise ValueError('Godkendelsen er ikke længere gyldig.')
            if time.time() - approval['created'] > 600:
                raise ValueError('Godkendelsen er udløbet. Gennemgå kørslen igen.')
            job.update(state='SENDING', error='')
            self.store.put(job)
            self.submit(self.send_worker, job_id, approval)
            return self.public(job)

    def send_worker(self, job_id, approval):
        for row_id in approval['ids']:
            with self.lock:
                job = self.store.get(job_id)
                row = self.row(job, row_id)
                if row['status'] != 'READY':
                    continue
                try:
                    raw = self.pdf_path(job_id, row).read_bytes()
                    if approval['digest'] != approval_digest(job, approval['ids'], approval['account']):
                        raise ValueError('Det godkendte indhold er ændret.')
                except Exception as exc:
                    job.update(state='PAUSED', error=str(exc), approval=None)
                    self.store.put(job)
                    return
                row['status'] = 'SENDING'
                self.store.put(job)  # Claim is durable before any external request.
            try:
                filename = re.sub(r'[<>:"/\\|?*\x00-\x1f]', '_', row['name'])[:100] + ' - Timeoversigt.pdf'
                self.transport(job).send(row['email'], row['subject'], row['body'], filename, raw, approval['account'], job.get('cc', ''))
                result, error = 'ACCEPTED', ''
            except DefiniteFailure as exc:
                result, error = 'FAILED', str(exc)
            except Exception:
                result, error = 'UNKNOWN', 'Usikker afsendelse. Kontrollér Udbakke og Sendt post før et nyt forsøg.'
            with self.lock:
                job = self.store.get(job_id)
                self.row(job, row_id)['status'] = result
                if error:
                    job.update(state='PAUSED', error=error, approval=None)
                self.store.put(job)
                if error:
                    return
        with self.lock:
            job = self.store.get(job_id)
            job.update(state='DRAFT', approval=None)
            self.store.put(job)
            if all(r['status'] == 'ACCEPTED' for r in job['recipients'] if r['sendable']):
                self.finish(job, 'COMPLETE')

    def reconcile(self, job_id, row_id, outcome):
        with self.lock:
            job = self.store.get(job_id)
            self.editable(job)
            row = self.row(job, row_id)
            if row['status'] not in ('UNKNOWN', 'FAILED') or outcome not in ('accepted', 'retry'):
                raise ValueError('Denne post afventer ikke manuel afklaring.')
            row['status'] = 'ACCEPTED' if outcome == 'accepted' else 'READY'
            row['selected'] = outcome == 'retry'
            job.update(state='DRAFT', error='')
            self.touch(job)
            self.store.put(job)
            if all(r['status'] == 'ACCEPTED' for r in job['recipients'] if r['sendable']):
                self.finish(job, 'COMPLETE')
            return self.public(self.store.get(job_id))

    def test_mail(self, job_id, row_id):
        with self.lock:
            job = self.store.get(job_id)
            self.editable(job)
            row = self.row(job, row_id)
            account = self.account(job, verify=True)
            if not account or not row['sendable'] or row['needs_review'] or row.get('needs_pdf_review'):
                raise ValueError('Vælg en afsenderkonto og en fejlfri, gennemgået medarbejder.')
            raw = self.pdf_path(job_id, row).read_bytes()
            self.validate_message(job, account['email'], '[TEST] ' + row['subject'], row['body'], 'Test - Timeoversigt.pdf', raw)
            job.update(test_status='SENDING', test_error='')
            self.store.put(job)
            def work():
                try:
                    self.transport(job).send(account['email'], '[TEST] ' + row['subject'], row['body'], 'Test - Timeoversigt.pdf', raw, account)
                    status, error = 'ACCEPTED', ''
                except DefiniteFailure as exc:
                    status, error = 'FAILED', str(exc)
                except Exception:
                    status, error = 'UNKNOWN', 'Testmailens status er ukendt. Kontrollér Udbakke og Sendt post før en ny test.'
                with self.lock:
                    current = self.store.get(job_id)
                    current.update(test_status=status, test_error=error)
                    self.store.put(current)
            self.submit(work)
            return self.public(job)

    def finish(self, job, state):
        # Called under service lock, never during rendering or an in-flight send.
        counts = {s: sum(r['status'] == s for r in job['recipients']) for s in ('ACCEPTED', 'BLOCKED', 'READY', 'UNKNOWN', 'FAILED')}
        directory = self.directory(job['id'])
        if directory.exists():
            remove_tree(directory, within=self.root / 'jobs')
        receipt = {key: job[key] for key in ('id', 'kind', 'month', 'year', 'created', 'expires', 'revision', 'source_hash')}
        receipt.update(mail_method=job.get('mail_method', 'graph'), state=state, cleaned=True, closed_at=time.time(), counts=counts, recipients=[], error='', filename='', progress='')
        self.store.put(receipt)
        self.store.prune()

    def cancel(self, job_id):
        with self.lock:
            job = self.store.get(job_id)
            if job['state'] in ('ANALYZING', 'SENDING') or job.get('test_status') == 'SENDING':
                raise ValueError('Vent til den igangværende behandling er afsluttet.')
            if not job.get('cleaned'):
                self.finish(job, 'CANCELLED')
            return self.public(self.store.get(job_id))

    def cleanup(self):
        with self.lock:
            jobs = self.store.all()
            for job in jobs:
                if not job.get('cleaned') and job['expires'] <= time.time() and job['state'] not in ('ANALYZING', 'SENDING') and job.get('test_status') != 'SENDING':
                    self.finish(job, 'EXPIRED')
            # An interrupted upload may have written a private directory before
            # its SQLite record was committed. Expire these unreferenced files too.
            known = {job['id'] for job in jobs}
            base = self.root / 'jobs'
            if base.exists():
                for path in base.iterdir():
                    if path.is_dir() and re.fullmatch(r'[a-f0-9]{32}', path.name) and path.name not in known and path.stat().st_mtime < time.time() - 48 * 3600:
                        remove_tree(path, within=base)
            self.store.prune()
            self.outlook.cleanup()

    def ensure_account_change_allowed(self):
        if any(j['state'] == 'SENDING' or j.get('test_status') == 'SENDING' for j in self.store.all()):
            raise ValueError('Afsenderkontoen kan ikke ændres under en udsendelse.')
