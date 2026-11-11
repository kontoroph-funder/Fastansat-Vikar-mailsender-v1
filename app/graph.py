from __future__ import annotations

import base64
import json
import threading
import uuid

import msal
import requests

from .payroll_pdf import EMAIL
from .vault import SecretStore
from .domain import validate_cc


class DefiniteFailure(RuntimeError):
    pass


class UnknownDelivery(RuntimeError):
    pass


def payload(email, subject, body, filename, raw, cc='', enforce_graph_limit=True):
    try:
        cc = validate_cc(cc)
    except ValueError as exc:
        raise DefiniteFailure(str(exc)) from None
    if not EMAIL.fullmatch(email) or not subject.strip() or any(c in subject for c in '\r\n') or not body.strip():
        raise DefiniteFailure('Ugyldig modtager, emne eller besked.')
    if not raw.startswith(b'%PDF-'):
        raise DefiniteFailure('Vedhæftningen er ikke en PDF.')
    result = json.dumps({'message': {'subject': subject, 'body': {'contentType': 'Text', 'content': body},
        'toRecipients': [{'emailAddress': {'address': email}}],
        'ccRecipients': [{'emailAddress': {'address': cc}}] if cc and cc.casefold() != email.casefold() else [],
        'attachments': [{'@odata.type': '#microsoft.graph.fileAttachment', 'name': filename,
                         'contentType': 'application/pdf', 'contentBytes': base64.b64encode(raw).decode()}]},
        'saveToSentItems': True}, ensure_ascii=False).encode()
    if enforce_graph_limit and len(result) > 3_800_000:
        raise DefiniteFailure('PDF og besked er for store til direkte Outlook-afsendelse (maks. 3,8 MB samlet forespørgsel).')
    return result


class Graph:
    scopes = ['User.Read', 'Mail.Send']

    def __init__(self, root, settings):
        self.vault = SecretStore(root / 'credentials')
        self.settings = settings
        self.lock = threading.RLock()
        self.login = None

    def _app(self, config):
        client_id = config['client_id'].strip()
        if not client_id:
            raise DefiniteFailure('Udfyld Microsoft client-id under Indstillinger.')
        cache = msal.SerializableTokenCache()
        saved = self.vault.get('microsoft') or {}
        if saved.get('client_id') == client_id and saved.get('tenant_id') == config['tenant_id']:
            cache.deserialize(saved.get('cache', '{}'))
        app = msal.PublicClientApplication(client_id, authority='https://login.microsoftonline.com/' + config['tenant_id'], token_cache=cache)
        return app, cache

    def account(self):
        saved = self.vault.get('microsoft') or {}
        config = self.settings()
        if saved.get('client_id') == config['client_id'] and saved.get('tenant_id') == config['tenant_id']:
            return saved.get('account')
        return None

    def start_login(self):
        with self.lock:
            if self.login and self.login['status'] == 'waiting':
                raise ValueError('Et Microsoft-login er allerede i gang.')
            config = self.settings()
            try:
                app, cache = self._app(config)
                flow = app.initiate_device_flow(scopes=self.scopes)
            except requests.RequestException:
                raise ValueError('Kunne ikke oprette forbindelse til Microsoft-login. Kontrollér internetforbindelsen, og at appen har adgang til login.microsoftonline.com.') from None
            except DefiniteFailure as exc:
                raise ValueError(str(exc)) from None
            if 'user_code' not in flow:
                raise ValueError('Microsoft-login kunne ikke startes. Kontrollér app-registreringen.')
            self.login = {'id': uuid.uuid4().hex, 'status': 'waiting', 'user_code': flow['user_code'],
                          'url': flow.get('verification_uri', 'https://microsoft.com/devicelogin'), 'error': ''}
            def complete():
                try:
                    result = app.acquire_token_by_device_flow(flow)
                    if 'access_token' not in result:
                        raise ValueError('Login afbrudt eller udløbet. Kontrollér også “Allow public client flows” i Entra.')
                    response = requests.get('https://graph.microsoft.com/v1.0/me?$select=id,displayName,mail,userPrincipalName',
                        headers={'Authorization': 'Bearer ' + result['access_token']}, timeout=30)
                    response.raise_for_status()
                    profile = response.json()
                    username = result.get('id_token_claims', {}).get('preferred_username', '')
                    accounts = app.get_accounts(username=username) if username else app.get_accounts()
                    if len(accounts) != 1:
                        raise ValueError('Kontoen kunne ikke identificeres entydigt. Log ud og prøv igen.')
                    account = {'id': accounts[0]['home_account_id'], 'email': profile.get('mail') or profile.get('userPrincipalName'),
                               'name': profile.get('displayName', '')}
                    if not EMAIL.fullmatch(account['email'] or ''):
                        raise ValueError('Microsoft-kontoen har ingen gyldig mailadresse.')
                    with self.lock:
                        self.vault.set('microsoft', {'client_id': config['client_id'], 'tenant_id': config['tenant_id'],
                                                    'cache': cache.serialize(), 'account': account})
                        self.login['status'] = 'connected'
                except Exception as exc:
                    with self.lock:
                        self.login['status'] = 'error'
                        self.login['error'] = str(exc) if isinstance(exc, ValueError) else 'Microsoft-login fejlede. Kontrollér forbindelsen og prøv igen.'
            threading.Thread(target=complete, daemon=True).start()
            return dict(self.login)

    def logout(self):
        with self.lock:
            if self.login and self.login['status'] == 'waiting':
                raise ValueError('Afslut det igangværende login først.')
            self.vault.delete('microsoft')
            self.login = None

    def send(self, email, subject, body, filename, raw, expected_account, cc=''):
        encoded = payload(email, subject, body, filename, raw, cc)
        with self.lock:
            config = self.settings()
            if self.account() != expected_account:
                raise DefiniteFailure('Afsenderkontoen er ændret. Godkend kørslen igen.')
            try:
                app, cache = self._app(config)
                accounts = [a for a in app.get_accounts() if a['home_account_id'] == expected_account['id']]
                if len(accounts) != 1:
                    raise DefiniteFailure('Log ind med Microsoft igen.')
                token = app.acquire_token_silent(self.scopes, account=accounts[0])
                if not token or 'access_token' not in token:
                    raise DefiniteFailure('Microsoft-login er udløbet. Log ind igen.')
                self.vault.set('microsoft', {'client_id': config['client_id'], 'tenant_id': config['tenant_id'],
                                            'cache': cache.serialize(), 'account': expected_account})
            except DefiniteFailure:
                raise
            except Exception:
                raise DefiniteFailure('Microsoft-login kunne ikke fornyes. Ingen mail blev forsøgt sendt.') from None
            try:
                response = requests.post('https://graph.microsoft.com/v1.0/me/sendMail', data=encoded,
                    headers={'Authorization': 'Bearer ' + token['access_token'], 'Content-Type': 'application/json'},
                    timeout=60, allow_redirects=False)
            except requests.RequestException:
                raise UnknownDelivery('Outlook svarede ikke entydigt. Kontrollér Sendt post før et nyt forsøg.') from None
            if response.status_code == 202:
                return
            if response.status_code in (400, 401, 403, 404, 413, 422, 429):
                raise DefiniteFailure(f'Outlook afviste afsendelsen (HTTP {response.status_code}).')
            raise UnknownDelivery(f'Usikkert svar fra Outlook (HTTP {response.status_code}). Kontrollér Sendt post.')
