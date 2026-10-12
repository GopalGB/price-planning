"""Small SQL boundary: server-owned SQLite WAL, optional external PostgreSQL."""
import contextlib, os, sqlite3
from pathlib import Path

ROOT=Path(__file__).resolve().parents[1]
DATA=Path(os.environ.get('DATA_DIR',ROOT/'data'))
DATA.mkdir(parents=True,exist_ok=True)
DSN=os.environ.get('DATABASE_URL','')
SCHEMA='''
CREATE TABLE IF NOT EXISTS meta (key TEXT PRIMARY KEY,value TEXT NOT NULL);
CREATE TABLE IF NOT EXISTS sheets (id INTEGER PRIMARY KEY,name TEXT UNIQUE NOT NULL,state TEXT,max_row INTEGER,max_col INTEGER,path TEXT,metadata TEXT);
CREATE TABLE IF NOT EXISTS cells (sheet INTEGER NOT NULL,r INTEGER NOT NULL,c INTEGER NOT NULL,value TEXT,formula TEXT,style INTEGER,editable INTEGER NOT NULL DEFAULT 0,PRIMARY KEY(sheet,r,c));
CREATE TABLE IF NOT EXISTS shared_formulas (sheet INTEGER,si INTEGER,r INTEGER,c INTEGER,formula TEXT,PRIMARY KEY(sheet,si));
CREATE TABLE IF NOT EXISTS users (id TEXT PRIMARY KEY,email TEXT UNIQUE NOT NULL,name TEXT NOT NULL,password TEXT NOT NULL,admin INTEGER NOT NULL DEFAULT 0,created REAL NOT NULL);
CREATE TABLE IF NOT EXISTS sessions (token TEXT PRIMARY KEY,user_id TEXT NOT NULL,csrf TEXT NOT NULL,expires REAL NOT NULL);
CREATE TABLE IF NOT EXISTS scenarios (id TEXT PRIMARY KEY,name TEXT NOT NULL,owner TEXT NOT NULL,revision INTEGER NOT NULL DEFAULT 0,created REAL NOT NULL,updated REAL NOT NULL);
CREATE TABLE IF NOT EXISTS members (scenario TEXT NOT NULL,user_id TEXT NOT NULL,role TEXT NOT NULL,PRIMARY KEY(scenario,user_id));
CREATE TABLE IF NOT EXISTS edits (scenario TEXT NOT NULL,sheet INTEGER NOT NULL,r INTEGER NOT NULL,c INTEGER NOT NULL,value TEXT NOT NULL,version INTEGER NOT NULL,actor TEXT NOT NULL,updated REAL NOT NULL,PRIMARY KEY(scenario,sheet,r,c));
CREATE TABLE IF NOT EXISTS events (id TEXT PRIMARY KEY,scenario TEXT NOT NULL,revision INTEGER NOT NULL,actor TEXT NOT NULL,kind TEXT NOT NULL,payload TEXT NOT NULL,created REAL NOT NULL);
CREATE INDEX IF NOT EXISTS events_sequence ON events(scenario,revision,created);
CREATE TABLE IF NOT EXISTS imports (id TEXT PRIMARY KEY,scenario TEXT,sheet INTEGER,actor TEXT,filename TEXT,sha256 TEXT,rows_count INTEGER,created REAL,status TEXT);
CREATE TABLE IF NOT EXISTS login_limits (key TEXT PRIMARY KEY,attempts INTEGER NOT NULL,until_time REAL NOT NULL);
CREATE TABLE IF NOT EXISTS source_cells (sheet INTEGER NOT NULL,r INTEGER NOT NULL,c INTEGER NOT NULL,value TEXT NOT NULL,PRIMARY KEY(sheet,r,c));
CREATE TABLE IF NOT EXISTS source_versions (sheet INTEGER PRIMARY KEY,sha256 TEXT,filename TEXT,updated REAL,actor TEXT,rows_count INTEGER);
'''

class DB:
 def __init__(self):
  self.pg=bool(DSN)
  if self.pg:
   import psycopg
   from psycopg.rows import dict_row
   self.con=psycopg.connect(DSN,row_factory=dict_row)
   self.con.isolation_level=psycopg.IsolationLevel.REPEATABLE_READ
  else:
   self.con=sqlite3.connect(DATA/'planning.db',timeout=30)
   self.con.row_factory=sqlite3.Row
   self.con.execute('PRAGMA journal_mode=WAL');self.con.execute('PRAGMA busy_timeout=30000')
   self.con.execute('BEGIN')
 def execute(self,sql,args=()):return self.con.execute(sql.replace('?','%s') if self.pg else sql,args)
 def many(self,sql,rows):
  with self.con.cursor() if self.pg else contextlib.nullcontext(self.con) as cur:cur.executemany(sql.replace('?','%s') if self.pg else sql,rows)
 def one(self,sql,args=()):
  r=self.execute(sql,args).fetchone();return dict(r) if r else None
 def all(self,sql,args=()):return [dict(x) for x in self.execute(sql,args).fetchall()]
 def begin_write(self):
  self.con.rollback()
  if self.pg:self.con.execute('SET TRANSACTION ISOLATION LEVEL READ COMMITTED')
  else:self.con.execute('BEGIN IMMEDIATE')
 def lock_scenario(self,sid):
   # PostgreSQL needs the row lock; SQLite takes it from BEGIN IMMEDIATE. Both branches
   # are fixed text, and sid stays a bound parameter either way.
   sql='SELECT * FROM scenarios WHERE id=? FOR UPDATE' if self.pg else 'SELECT * FROM scenarios WHERE id=?'
   return self.one(sql,(sid,))
 def commit(self):self.con.commit()
 def rollback(self):self.con.rollback()
 def close(self):self.con.close()
 def __enter__(self):return self
 def __exit__(self,typ,exc,tb):
  if exc:self.rollback()
  self.close()
def init_schema(db):
 for sql in (SCHEMA.replace(' REAL',' DOUBLE PRECISION') if db.pg else SCHEMA).split(';'):
  if sql.strip():db.execute(sql)
 db.commit()
