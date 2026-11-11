"""Outlook transport; the subprocess owns COM, the parent owns attachments."""
import json
import os
from pathlib import Path
import subprocess
import sys
import tempfile
import time
import threading

from .graph import DefiniteFailure, UnknownDelivery, payload


class Outlook:
    def __init__(self, root, settings):
        self.root = Path(root) / 'outlook-work'
        self.settings = settings
        self.lock = threading.Lock()

    def account(self):
        return self.settings().get('outlook_account')

    def call(self, request):
        if os.name != 'nt':
            raise DefiniteFailure('Klassisk Outlook kræver Windows.')
        command = [sys.executable, '--outlook-worker'] if getattr(sys, 'frozen', False) else [sys.executable, '-m', 'app.outlook_worker']
        uncertain = request['operation'] == 'send'
        failure = UnknownDelivery if uncertain else DefiniteFailure
        self.root.mkdir(parents=True, exist_ok=True)
        # Files also work in the windowless PyInstaller executable, where stdio is None.
        with self.lock, tempfile.TemporaryDirectory(prefix='mail-', dir=self.root) as temporary:
            request_path = Path(temporary) / 'request.json'
            response_path = Path(temporary) / 'response.json'
            request_path.write_text(json.dumps(request), encoding='utf-8')
            try:
                process = subprocess.Popen([*command, str(request_path.resolve()), str(response_path.resolve())],
                                           stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
                                           creationflags=getattr(subprocess, 'CREATE_NO_WINDOW', 0))
            except OSError:
                raise DefiniteFailure('Outlook-worker kunne ikke startes.') from None
            try:
                process.wait(timeout=180 if uncertain else 45)
            except (subprocess.TimeoutExpired, OSError):
                process.kill()  # Only our worker, never the user's Outlook process.
                process.wait()
                raise failure('Outlook svarede ikke i tide. Kontrollér Outlook-dialoger, Udbakke og Sendt post.') from None
            try:
                result = json.loads(response_path.read_text(encoding='utf-8'))
                if process.returncode or not isinstance(result, dict):
                    raise ValueError()
            except (OSError, ValueError, UnicodeError):
                raise failure('Outlook-worker blev afbrudt. Kontrollér Udbakke og Sendt post.') from None
        if result.get('status') == 'failed':
            raise DefiniteFailure(result['error'])
        if result.get('status') == 'unknown':
            raise UnknownDelivery(result['error'])
        return result

    def accounts(self):
        return self.call({'operation': 'accounts'})['accounts']

    def verify(self, account):
        if not account:
            raise DefiniteFailure('Vælg og gem en afsenderkonto under Indstillinger.')
        return self.call({'operation': 'verify', 'account': account})['account']

    def send(self, email, subject, body, filename, raw, expected_account, cc=''):
        payload(email, subject, body, filename, raw, cc, enforce_graph_limit=False)
        if self.account() != expected_account:
            raise DefiniteFailure('Afsenderkontoen er ændret. Godkend kørslen igen.')
        self.root.mkdir(parents=True, exist_ok=True)
        with tempfile.TemporaryDirectory(prefix='mail-', dir=self.root) as temporary:
            attachment = Path(temporary) / filename
            if attachment.name != filename or not filename.lower().endswith('.pdf'):
                raise DefiniteFailure('Ugyldigt PDF-filnavn.')
            attachment.write_bytes(raw)
            result = self.call({'operation': 'send', 'account': expected_account, 'email': email,
                                'subject': subject, 'body': body, 'cc': cc, 'pdf': str(attachment.resolve())})
            if result.get('status') != 'accepted':
                raise UnknownDelivery('Outlook gav ikke et entydigt svar. Kontrollér Udbakke og Sendt post.')

    def cleanup(self):
        from .files import remove_tree
        if self.root.exists():
            for path in self.root.iterdir():
                if path.is_dir() and path.name.startswith('mail-') and path.stat().st_mtime < time.time() - 48 * 3600:
                    remove_tree(path, within=self.root)
