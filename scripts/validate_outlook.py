"""Read-only Outlook probe and an unsent, discarded synthetic message. Never calls Send."""
from pathlib import Path
import sys
import tempfile

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
from app.outlook import Outlook
from app.outlook_worker import accounts, identity, prepare
from app.domain import defaults


def main():
    import pythoncom
    from win32com.client import dynamic
    from pypdf import PdfWriter
    with tempfile.TemporaryDirectory(prefix='FV-Outlook-validation-') as temp:
        transport = Outlook(Path(temp), defaults)
        listed = transport.accounts()
        assert listed, 'Ingen egnede konti i klassisk Outlook.'
        transport.verify(listed[0])
        pythoncom.CoInitialize()
        mail = None
        try:
            app = dynamic.Dispatch('Outlook.Application')
            selected = [(item, account) for item, account in accounts(app.Session) if item == listed[0]]
            assert len(selected) == 1
            account, com_account = selected[0]
            pdf = Path(temp) / 'Syntetisk-test.pdf'
            writer = PdfWriter()
            writer.add_blank_page(595, 842)
            writer.write(pdf)
            writer.close()
            mail = prepare(app, dict(account=account, email=account['email'], cc='', subject='Syntetisk test - sendes ikke', body='Kun kontrol af Outlook-klargøring.', pdf=str(pdf)), com_account)
            assert identity(app.Session, mail.SendUsingAccount) == account
            assert mail.Attachments.Count == 1 and mail.Recipients.Count == 1
            print('Outlook account, recipient and PDF attachment verified. No mail sent.')
        finally:
            if mail is not None:
                mail.Close(1)  # olDiscard
            pythoncom.CoUninitialize()


if __name__ == '__main__':
    main()
