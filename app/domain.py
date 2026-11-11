from __future__ import annotations

import copy
import hashlib
import json
from datetime import datetime
from zoneinfo import ZoneInfo

MONTHS = ('januar', 'februar', 'marts', 'april', 'maj', 'juni', 'juli',
          'august', 'september', 'oktober', 'november', 'december')
DEFAULTS = {
    'client_id': '', 'tenant_id': 'organizations',
    'libreoffice_program': r'C:\Program Files\LibreOffice\program',
    'libreoffice_python': '',
    'templates': {
        'fastansat': {
            'subject': 'Timeoversigt for lønperioden {period}',
            'body': 'Hej {name}\n\nVedhæftet finder du din personlige oversigt over Timebank og Timeregistrering for {period}.\n\nKontakt administrationen, hvis du har spørgsmål eller finder fejl.\n\nMed venlig hilsen\nOpholdssteder Administration',
        },
        'vikar': {
            'subject': 'Timeoversigt for lønperioden {period}',
            'body': 'Hej {name},\n\nDu modtager hermed din timeoversigt for lønperioden {period}.\n\nTimeoversigten viser de timer, der er registreret i den pågældende lønperiode, og er vedhæftet denne mail.\n\nHvis du har spørgsmål eller opdager fejl i timeoversigten, bedes du kontakte administrationen.\n\nMed venlig hilsen\nOpholdssteder Administration',
        },
    },
}


def local_now():
    return datetime.now(ZoneInfo('Europe/Copenhagen'))


def validate_period(month, year):
    if not isinstance(month, int) or not 1 <= month <= 12 or not isinstance(year, int) or not 1900 <= year <= 2200:
        raise ValueError('Vælg en gyldig måned og et gyldigt år.')


def period_text(month, year):
    validate_period(month, year)
    previous = 12 if month == 1 else month - 1
    start_year = f' {year - 1}' if month == 1 else ''
    return f'16. {MONTHS[previous - 1]}{start_year} – 15. {MONTHS[month - 1]} {year}'


def message(template, row, month, year):
    return {key: template[key].replace('{name}', row['name']).replace('{period}', period_text(month, year))
            for key in ('subject', 'body')}


def defaults():
    return copy.deepcopy(DEFAULTS)


def validate_cc(value):
    from .payroll_pdf import EMAIL
    if not isinstance(value, str) or len(value) > 254 or '\r' in value or '\n' in value:
        raise ValueError('Cc skal være én gyldig mailadresse eller være tomt.')
    value = value.strip()
    if value and not EMAIL.fullmatch(value):
        raise ValueError('Cc skal være én gyldig mailadresse eller være tomt.')
    return value


def digest(value):
    return hashlib.sha256(json.dumps(value, sort_keys=True, ensure_ascii=False, separators=(',', ':')).encode()).hexdigest()


def approval_digest(job, ids, account):
    return digest({'revision': job['revision'], 'account': account, 'month': job['month'], 'year': job['year'], 'cc': job.get('cc', ''),
                   'recipients': [{k: r[k] for k in ('id', 'email', 'subject', 'body', 'pdf', 'sha256')}
                                  for r in job['recipients'] if r['id'] in ids]})
