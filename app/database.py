import sqlite3
from pathlib import Path

SCHEMA = """
CREATE TABLE IF NOT EXISTS project(id INTEGER PRIMARY KEY, source_root TEXT);
CREATE TABLE IF NOT EXISTS pages(
 id INTEGER PRIMARY KEY,
 source_pdf TEXT NOT NULL,
 relative_pdf TEXT NOT NULL,
 page_number INTEGER NOT NULL,
 page_count INTEGER NOT NULL,
 ocr_text TEXT DEFAULT '',
 personnummer TEXT DEFAULT '',
 script_personnummer TEXT DEFAULT '',
 personnummer_source TEXT DEFAULT '',
 score INTEGER DEFAULT 0,
 score_reasons TEXT DEFAULT '',
 classification TEXT DEFAULT '',
 reviewed INTEGER DEFAULT 0,
 UNIQUE(source_pdf,page_number)
);
    
CREATE TABLE IF NOT EXISTS session_state (
    id INTEGER PRIMARY KEY CHECK(id = 1),
    current_page_id INTEGER,
    updated_at TEXT DEFAULT CURRENT_TIMESTAMP
);
"""

class WorkDatabase:
    def __init__(self, path):
        self.path = Path(path)
        self.conn = sqlite3.connect(path, check_same_thread=False)
        self.conn.row_factory = sqlite3.Row
        self.conn.executescript(SCHEMA)
        self._migrate()
        self.conn.commit()

    def _migrate(self):
        columns = {row["name"] for row in self.conn.execute("PRAGMA table_info(pages)")}
        for name in ("script_personnummer", "personnummer_source"):
            if name not in columns:
                self.conn.execute(f"ALTER TABLE pages ADD COLUMN {name} TEXT DEFAULT ''")
        self.conn.execute("UPDATE pages SET script_personnummer=personnummer WHERE script_personnummer='' AND personnummer<>''")

    def close(self): self.conn.close()
    def set_root(self, root):
        self.conn.execute("INSERT INTO project VALUES(1,?) ON CONFLICT(id) DO UPDATE SET source_root=excluded.source_root", (str(root),))
        self.conn.commit()
    def root(self):
        row = self.conn.execute("SELECT source_root FROM project WHERE id=1").fetchone()
        return row[0] if row else ""

    def upsert(self, data):
        self.conn.execute(
            """
            INSERT INTO pages (source_pdf,
                               relative_pdf,
                               page_number,
                               page_count,
                               ocr_text,
                               personnummer,
                               script_personnummer,
                               personnummer_source,
                               score,
                               score_reasons)
            VALUES (:source_pdf,
                    :relative_pdf,
                    :page_number,
                    :page_count,
                    :ocr_text,
                    :personnummer,
                    :script_personnummer,
                    :personnummer_source,
                    :score,
                    :score_reasons) ON CONFLICT(
                source_pdf,
                page_number
            )
            DO
            UPDATE SET
                ocr_text = excluded.ocr_text,
                score = excluded.score,
                score_reasons = excluded.score_reasons
            """,
            data
        )

    def save_current_position(self, page_id):
        """
        Sparar vilken sida användaren för närvarande arbetar med.
        Det finns alltid högst en rad i session_state.
        """

        self.conn.execute(
            """
            INSERT INTO session_state (id,
                                       current_page_id,
                                       updated_at)
            VALUES (1,
                    ?,
                    CURRENT_TIMESTAMP) ON CONFLICT(id) DO
            UPDATE SET
                current_page_id = excluded.current_page_id,
                updated_at = CURRENT_TIMESTAMP
            """,
            (page_id,)
        )

        self.conn.commit()

    def get_current_position(self):
        """
        Returnerar ID för den sida där användaren senast arbetade.
        Returnerar None om ingen position har sparats.
        """

        row = self.conn.execute(
            """
            SELECT current_page_id
            FROM session_state
            WHERE id = 1
            """
        ).fetchone()

        if row is None:
            return None

        return row["current_page_id"]

    def review_statistics(self):
        """
        Returnerar antal granskade och totala sidor.
        En sida räknas som granskad när reviewed = 1.
        """

        row = self.conn.execute(
            """
            SELECT COUNT(*)                   AS total_pages,
                   COALESCE(SUM(reviewed), 0) AS reviewed_pages
            FROM pages
            """
        ).fetchone()

        return {
            "reviewed_pages": row["reviewed_pages"],
            "total_pages": row["total_pages"],
        }
    def pages(self):
        return self.conn.execute("SELECT * FROM pages ORDER BY relative_pdf COLLATE NOCASE,page_number").fetchall()
    def save(self, page_id, classification, personnummer):
        row = self.conn.execute("SELECT script_personnummer FROM pages WHERE id=?", (page_id,)).fetchone()
        value = personnummer.strip()
        source = "" if not value else ("skript" if value == (row[0] or "") else "människa")
        self.conn.execute("UPDATE pages SET classification=?,personnummer=?,personnummer_source=?,reviewed=1 WHERE id=?", (classification,value,source,page_id))
        self.conn.commit()
