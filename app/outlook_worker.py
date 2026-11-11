"""One bounded Outlook COM operation in the interactive Windows session.
Never closes Outlook, changes its security settings, or retries Send.
"""
import hashlib
import json
import sys
from pathlib import Path


def identity(session, account):
    email = str(account.SmtpAddress or '').strip().lower()
    from .payroll_pdf import EMAIL
    if not EMAIL.fullmatch(email):
        raise ValueError('Outlook-kontoen mangler en gyldig SMTP-adresse.')
    profile = str(session.CurrentProfileName)
    store = str(account.DeliveryStore.StoreID)
    key = hashlib.sha256((profile + '\0' + store + '\0' + email).encode()).hexdigest()
    return {'id': key, 'email': email, 'name': str(account.DisplayName), 'method': 'outlook'}


def accounts(session):
    result = []
    for index in range(1, session.Accounts.Count + 1):
        account = session.Accounts.Item(index)
        try:
            result.append((identity(session, account), account))
        except Exception:
            continue  # Only explicitly identifiable accounts can be selected.
    return result


def execute(application, request):
    session = application.Session
    found = accounts(session)
    if request['operation'] == 'accounts':
        return {'accounts': [item for item, _ in found]}
    expected = request['account']
    matches = [account for item, account in found if item == expected]
    if len(matches) != 1:
        return {'status': 'failed', 'error': 'Den valgte Outlook-konto findes ikke entydigt i den aktive profil. Hent konti igen.'}
    if request['operation'] == 'verify':
        return {'account': expected}
    if request['operation'] != 'send':
        raise ValueError('Ukendt Outlook-handling.')
    if session.Offline:
        return {'status': 'failed', 'error': 'Outlook er offline. Opret forbindelse, og prøv igen.'}
    mail = prepare(application, request, matches[0])
    try:
        mail.Send()
    except Exception:
        return {'status': 'unknown', 'error': 'Outlook gav ikke et entydigt svar. Kontrollér Udbakke og Sendt post før et nyt forsøg.'}
    return {'status': 'accepted'}


def prepare(application, request, account):
    session = application.Session
    expected = request['account']
    mail = application.CreateItem(0)
    # Dynamic dispatch needs a PROPERTYPUTREF for the Account COM object.
    import pythoncom
    mail._oleobj_.Invoke(64209, 0, pythoncom.DISPATCH_PROPERTYPUTREF, 0, account)
    if identity(session, mail.SendUsingAccount) != expected:
        raise ValueError('Outlook kunne ikke vælge den godkendte afsenderkonto.')
    mail.BodyFormat = 1  # Plain text, matching Graph messages.
    mail.Subject = request['subject']
    mail.Body = request['body']
    for address, recipient_type in [(request['email'], 1), (request.get('cc', ''), 2)]:
        if address and (recipient_type == 1 or address.casefold() != request['email'].casefold()):
            recipient = mail.Recipients.Add(address)
            recipient.Type = recipient_type
    if not mail.Recipients.ResolveAll():
        raise ValueError('Outlook kunne ikke finde alle modtagere.')
    mail.Attachments.Add(request['pdf'], 1)  # Copy PDF by value before handing off.
    mail.DeleteAfterSubmit = False
    # Recheck after potentially interactive recipient resolution.
    if identity(session, mail.SendUsingAccount) != expected:
        raise ValueError('Afsenderkontoen blev ændret i Outlook.')
    return mail


def main():
    request = json.loads(Path(sys.argv[-2]).read_text(encoding='utf-8'))
    import pythoncom
    from win32com.client import dynamic
    pythoncom.CoInitialize()
    try:
        application = dynamic.Dispatch('Outlook.Application')
        try:
            result = execute(application, request)
        except Exception:
            result = {'status': 'failed', 'error': 'Outlook kunne ikke klargøre mailen. Kontrollér profil, konto og eventuelle dialoger i klassisk Outlook.'}
    except Exception:
        result = {'status': 'failed', 'error': 'Klassisk Outlook kunne ikke åbnes. Åbn Outlook og log ind på afsenderkontoen.'}
    finally:
        pythoncom.CoUninitialize()
    Path(sys.argv[-1]).write_text(json.dumps(result, ensure_ascii=False), encoding='utf-8')


if __name__ == '__main__':
    main()
