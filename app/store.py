from contextlib import contextmanager
import json
from pathlib import Path
import sqlite3
import time

from .domain import defaults


class Store:
    def __init__(self, root):
        self.root = Path(root).resolve()
        self.root.mkdir(parents=True, exist_ok=True)
        self.path = self.root / 'state.sqlite3'
        with self.connect() as db:
            db.executescript('''
                CREATE TABLE IF NOT EXISTS jobs (id TEXT PRIMARY KEY, updated REAL NOT NULL, data TEXT NOT NULL);
                CREATE TABLE IF NOT EXISTS settings (id INTEGER PRIMARY KEY CHECK(id=1), data TEXT NOT NULL);
            ''')

    @contextmanager
    def connect(self):
        db = sqlite3.connect(self.path, timeout=30)
        db.execute('PRAGMA secure_delete=ON')
        db.execute('PRAGMA journal_mode=WAL')
        db.execute('PRAGMA synchronous=FULL')
        try:
            yield db
            db.commit()
        except BaseException:
            db.rollback()
            raise
        finally:
            db.close()

    def put(self, job):
        with self.connect() as db:
            db.execute('INSERT INTO jobs VALUES(?,?,?) ON CONFLICT(id) DO UPDATE SET updated=excluded.updated,data=excluded.data',
                       (job['id'], time.time(), json.dumps(job, ensure_ascii=False)))

    def get(self, job_id):
        with self.connect() as db:
            row = db.execute('SELECT data FROM jobs WHERE id=?', (job_id,)).fetchone()
        if row is None:
            raise ValueError('Kørslen findes ikke.')
        return json.loads(row[0])

    def all(self):
        with self.connect() as db:
            return [json.loads(r[0]) for r in db.execute('SELECT data FROM jobs ORDER BY updated DESC')]

    def settings(self):
        with self.connect() as db:
            row = db.execute('SELECT data FROM settings WHERE id=1').fetchone()
        return json.loads(row[0]) if row else defaults()

    def save_settings(self, settings):
        with self.connect() as db:
            db.execute('INSERT INTO settings VALUES(1,?) ON CONFLICT(id) DO UPDATE SET data=excluded.data',
                       (json.dumps(settings, ensure_ascii=False),))

    def prune(self):
        for job in self.all():
            if job.get('cleaned') and job.get('closed_at', time.time()) < time.time() - 365 * 86400:
                with self.connect() as db:
                    db.execute('DELETE FROM jobs WHERE id=?', (job['id'],))
        with self.connect() as db:
            db.execute('PRAGMA wal_checkpoint(TRUNCATE)')
