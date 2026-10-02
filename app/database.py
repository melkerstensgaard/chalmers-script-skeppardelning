import sqlite3
from pathlib import Path
from datetime import datetime, timezone

SCHEMA = """
PRAGMA foreign_keys = ON;
CREATE TABLE IF NOT EXISTS projects(
 id INTEGER PRIMARY KEY, input_root TEXT UNIQUE NOT NULL, output_root TEXT NOT NULL,
 created_at TEXT NOT NULL
);
CREATE TABLE IF NOT EXISTS volumes(
 id INTEGER PRIMARY KEY, project_id INTEGER NOT NULL REFERENCES projects(id),
 signum TEXT NOT NULL, source_path TEXT UNIQUE NOT NULL, source_name TEXT NOT NULL,
 source_size INTEGER NOT NULL, source_mtime_ns INTEGER NOT NULL, sha256 TEXT,
 page_count INTEGER NOT NULL, analyzed_pages INTEGER NOT NULL DEFAULT 0,
 analysis_status TEXT NOT NULL DEFAULT 'pending', analysis_error TEXT,
 export_status TEXT NOT NULL DEFAULT 'not_exported'
);
CREATE TABLE IF NOT EXISTS pages(
 id INTEGER PRIMARY KEY, volume_id INTEGER NOT NULL REFERENCES volumes(id),
 page_index INTEGER NOT NULL, ocr_text TEXT NOT NULL DEFAULT '', suggested_personnummer TEXT,
 confidence_score INTEGER NOT NULL DEFAULT 0, confidence_reasons TEXT NOT NULL DEFAULT '',
 page_class TEXT NOT NULL DEFAULT 'uncertain', final_personnummer TEXT,
 review_comment TEXT NOT NULL DEFAULT '', review_status TEXT NOT NULL DEFAULT 'unreviewed',
 reviewed_at TEXT, UNIQUE(volume_id,page_index)
);
CREATE TABLE IF NOT EXISTS candidates(
 id INTEGER PRIMARY KEY, page_id INTEGER NOT NULL REFERENCES pages(id) ON DELETE CASCADE,
 raw_value TEXT NOT NULL, normalized_value TEXT NOT NULL, source TEXT NOT NULL,
 ocr_confidence REAL NOT NULL DEFAULT 0, score INTEGER NOT NULL, reasons TEXT NOT NULL
);
"""

class Database:
    def __init__(self, path):
        self.path = Path(path)
        self.conn = sqlite3.connect(self.path)
        self.conn.row_factory = sqlite3.Row
        self.conn.executescript(SCHEMA)
    def close(self): self.conn.close()
    def add_project(self, input_root, output_root):
        i, o = str(Path(input_root).resolve()), str(Path(output_root).resolve())
        self.conn.execute("INSERT INTO projects(input_root,output_root,created_at) VALUES(?,?,?) ON CONFLICT(input_root) DO UPDATE SET output_root=excluded.output_root", (i,o,datetime.now(timezone.utc).isoformat()))
        self.conn.commit()
        return self.conn.execute("SELECT id FROM projects WHERE input_root=?", (i,)).fetchone()[0]
    def projects(self): return self.conn.execute("SELECT * FROM projects ORDER BY id").fetchall()
    def project(self, project_id): return self.conn.execute("SELECT * FROM projects WHERE id=?", (project_id,)).fetchone()
    def add_volume(self, project_id, signum, source_path, page_count, sha256):
        p=Path(source_path).resolve(); s=p.stat()
        self.conn.execute("INSERT INTO volumes(project_id,signum,source_path,source_name,source_size,source_mtime_ns,sha256,page_count) VALUES(?,?,?,?,?,?,?,?) ON CONFLICT(source_path) DO UPDATE SET project_id=excluded.project_id,signum=excluded.signum", (project_id,signum,str(p),p.name,s.st_size,s.st_mtime_ns,sha256,page_count))
        self.conn.commit()
        return self.conn.execute("SELECT id FROM volumes WHERE source_path=?", (str(p),)).fetchone()[0]
    def volumes(self, project_id=None):
        if project_id is None: return self.conn.execute("SELECT * FROM volumes ORDER BY signum,source_name").fetchall()
        return self.conn.execute("SELECT * FROM volumes WHERE project_id=? ORDER BY signum,source_name", (project_id,)).fetchall()
    def set_volume_status(self, volume_id, status, analyzed_pages=None, error=None):
        if analyzed_pages is None:
            self.conn.execute("UPDATE volumes SET analysis_status=?,analysis_error=? WHERE id=?",(status,error,volume_id))
        else:
            self.conn.execute("UPDATE volumes SET analysis_status=?,analyzed_pages=?,analysis_error=? WHERE id=?",(status,analyzed_pages,error,volume_id))
        self.conn.commit()
    def save_page(self, volume_id, page_index, text, suggested, score, reasons):
        self.conn.execute("INSERT INTO pages(volume_id,page_index,ocr_text,suggested_personnummer,confidence_score,confidence_reasons) VALUES(?,?,?,?,?,?) ON CONFLICT(volume_id,page_index) DO UPDATE SET ocr_text=excluded.ocr_text,suggested_personnummer=excluded.suggested_personnummer,confidence_score=excluded.confidence_score,confidence_reasons=excluded.confidence_reasons",(volume_id,page_index,text,suggested,score,reasons))
        self.conn.commit()
        return self.conn.execute("SELECT id FROM pages WHERE volume_id=? AND page_index=?",(volume_id,page_index)).fetchone()[0]
    def replace_candidates(self,page_id,candidates):
        self.conn.execute("DELETE FROM candidates WHERE page_id=?",(page_id,))
        self.conn.executemany("INSERT INTO candidates(page_id,raw_value,normalized_value,source,ocr_confidence,score,reasons) VALUES(?,?,?,?,?,?,?)",[(page_id,*c) for c in candidates]); self.conn.commit()
    def pages(self, volume_id): return self.conn.execute("SELECT * FROM pages WHERE volume_id=? ORDER BY page_index",(volume_id,)).fetchall()
    def candidates(self,page_id): return self.conn.execute("SELECT * FROM candidates WHERE page_id=? ORDER BY score DESC",(page_id,)).fetchall()
    def review(self,page_id,page_class,personnummer,comment,status='reviewed'):
        self.conn.execute("UPDATE pages SET page_class=?,final_personnummer=?,review_comment=?,review_status=?,reviewed_at=? WHERE id=?",(page_class,personnummer or None,comment,status,datetime.now(timezone.utc).isoformat(),page_id));self.conn.commit()
    def set_exported(self,volume_id): self.conn.execute("UPDATE volumes SET export_status='exported' WHERE id=?",(volume_id,));self.conn.commit()
