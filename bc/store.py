"""The tool's own record: one row per reply file, plus his decision. SQLite, one file."""
import datetime as dt
import json
import sqlite3

SCHEMA = """
create table if not exists items(
  id integer primary key,
  path text not null,
  view_path text,
  file_hash text not null,
  content_key text unique,
  seen_in text not null default '[]',   -- which sources had this reply, e.g. ["Shared folder", "His mail"]
  found_at text not null,
  text text,
  mail_from text, mail_subject text, mail_date text,
  cust_no text, cust_how text, name_ok integer, candidates text,
  attachments text default '[]',
  suggested_theirs real, suggested_status text, needs_typing integer, reason text, flags text,
  state text not null default 'new',         -- new / saved / skipped / not_a_reply
  theirs real, status text, remarks text, saved_at text, saved_by text
);
create table if not exists acks(
  no text, code text, by text, at text, primary key(no, code)
);
create table if not exists log(
  id integer primary key, at text, item_id integer, action text, detail text
);
"""


class Store:
    def __init__(self, path):
        import threading
        self._path = str(path)
        self._local = threading.local()
        self.db.executescript(SCHEMA)
        try:  # older databases
            self.db.execute("alter table items add column attachments text default '[]'")
        except sqlite3.OperationalError:
            pass

    @property
    def db(self):
        """One SQLite connection per web-server thread (a shared one breaks when two people save at once)."""
        c = getattr(self._local, "c", None)
        if c is None:
            c = sqlite3.connect(self._path, timeout=30)
            c.row_factory = sqlite3.Row
            c.execute("pragma journal_mode=wal")
            c.execute("pragma busy_timeout=30000")
            self._local.c = c
        return c

    def has_key(self, key):
        return self.db.execute("select 1 from items where content_key=?", (key,)).fetchone() is not None

    def add_seen(self, key, source):
        """Same reply found again (another source or another copy): remember where, do not add a new item."""
        r = self.db.execute("select id, seen_in from items where content_key=?", (key,)).fetchone()
        seen = json.loads(r["seen_in"] or "[]")
        if source not in seen:
            seen.append(source)
            self.db.execute("update items set seen_in=? where id=?", (json.dumps(seen), r["id"]))
            self.db.commit()
        return r["id"]

    def all_items(self):
        return [_row(r) for r in self.db.execute("select * from items where state<>'not_a_reply' order by found_at, id")]

    def add(self, **kw):
        kw.setdefault("found_at", now())
        for k in ("candidates", "flags", "seen_in", "attachments"):
            kw[k] = json.dumps(kw.get(k) or [])
        cols = ",".join(kw)
        cur = self.db.execute(f"insert into items({cols}) values({','.join('?' * len(kw))})", list(kw.values()))
        self.db.commit()
        return cur.lastrowid

    def get(self, item_id):
        r = self.db.execute("select * from items where id=?", (item_id,)).fetchone()
        return _row(r)

    def next_new(self, after_id=0):
        r = self.db.execute("select * from items where state='new' and id>? order by found_at, id limit 1",
                            (after_id,)).fetchone()
        return _row(r)

    def new_items(self):
        return [_row(r) for r in self.db.execute("select * from items where state='new' order by found_at, id")]

    def count(self, state):
        return self.db.execute("select count(*) from items where state=?", (state,)).fetchone()[0]

    def saved_today(self):
        return self.db.execute("select count(*) from items where state='saved' and saved_at like ?",
                               (dt.date.today().isoformat() + "%",)).fetchone()[0]

    def saved_for_customer(self, no, exclude_id):
        return [_row(r) for r in self.db.execute(
            "select * from items where state='saved' and cust_no=? and id<>?", (no, exclude_id))]

    def save(self, item_id, no, theirs, status, remarks, user):
        self.db.execute("update items set state='saved', cust_no=?, theirs=?, status=?, remarks=?, saved_at=?, saved_by=? "
                        "where id=?", (no, theirs, status, remarks, now(), user, item_id))
        self.log(item_id, "saved", dict(no=no, theirs=theirs, status=status, remarks=remarks, by=user))

    def set_state(self, item_id, state, user):
        self.db.execute("update items set state=? where id=?", (state, item_id))
        self.log(item_id, state, dict(by=user))

    def reopen(self, item_id, user):
        self.db.execute("update items set state='new' where id=?", (item_id,))
        self.log(item_id, "reopened", dict(by=user))

    def decisions(self):
        """Latest saved decision per customer (a later reply replaces an earlier one)."""
        rows = self.db.execute("select * from items where state='saved' order by saved_at").fetchall()
        latest = {}
        for r in rows:
            latest[r["cust_no"]] = dict(no=r["cust_no"], theirs=r["theirs"], status=r["status"],
                                        remarks=r["remarks"], saved_at=r["saved_at"], received_on=None)
        return list(latest.values())

    def acked(self):
        return {(r["no"], r["code"]) for r in self.db.execute("select no, code from acks")}

    def ack(self, no, code, user):
        self.db.execute("insert or replace into acks values(?,?,?,?)", (no, code, user, now()))
        self.log(None, "check_accepted", dict(no=no, code=code, by=user))

    def manual_decision(self, no, theirs, status, remarks, user):
        """A correction made on the Checks page. Stored like a saved reply so it is logged and written the same way."""
        key = f"manual:{no}:{now()}"
        i = self.add(path="(correction on Checks page)", file_hash=key, content_key=key, seen_in=["Checks page"], text="",
                     cust_no=no, state="saved", theirs=theirs, status=status, remarks=remarks, saved_at=now(), saved_by=user)
        self.log(i, "corrected", dict(no=no, theirs=theirs, status=status, by=user))
        return i

    def recent_saved(self, limit=30):
        return [_row(r) for r in self.db.execute(
            "select * from items where state='saved' order by saved_at desc limit ?", (limit,))]

    def log(self, item_id, action, detail):
        self.db.execute("insert into log(at,item_id,action,detail) values(?,?,?,?)",
                        (now(), item_id, action, json.dumps(detail)))
        self.db.commit()


def now():
    return dt.datetime.now().isoformat(timespec="seconds")


def _row(r):
    if r is None:
        return None
    d = dict(r)
    for k in ("candidates", "flags", "seen_in", "attachments"):
        d[k] = json.loads(d.get(k) or "[]")
    return d
