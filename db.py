"""
SQLite storage layer for VER Data Extraction Tool.
Stores extracted village records persistently in a local ver_data.db file.
"""
import sqlite3
import json
from pathlib import Path
from collections import OrderedDict
from datetime import datetime
from comprehensive_extract import MASTER_FIELDS

DB_PATH = Path(__file__).parent / "ver_data.db"


def get_connection():
    """Get a SQLite connection with WAL mode for better concurrent reads."""
    conn = sqlite3.connect(str(DB_PATH))
    conn.execute("PRAGMA journal_mode=WAL")
    conn.execute("PRAGMA foreign_keys=ON")
    conn.row_factory = sqlite3.Row
    return conn


def init_db():
    """Create the villages table if it doesn't exist."""
    conn = get_connection()
    conn.execute("""
        CREATE TABLE IF NOT EXISTS villages (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            village_name TEXT,
            state TEXT,
            data_json TEXT NOT NULL,
            created_at TEXT NOT NULL,
            updated_at TEXT NOT NULL
        )
    """)
    conn.commit()
    conn.close()


def save_village(record: dict) -> int:
    """Save a village record to the database. Returns the row id."""
    conn = get_connection()
    now = datetime.now().isoformat()
    data_json = json.dumps(record, ensure_ascii=False, default=str)
    cursor = conn.execute(
        "INSERT INTO villages (village_name, state, data_json, created_at, updated_at) VALUES (?, ?, ?, ?, ?)",
        (record.get("village_name", ""), record.get("state", ""), data_json, now, now),
    )
    conn.commit()
    row_id = cursor.lastrowid
    conn.close()
    return row_id


def load_all_villages() -> list[dict]:
    """Load all village records from the database."""
    conn = get_connection()
    rows = conn.execute("SELECT id, data_json, created_at FROM villages ORDER BY created_at DESC").fetchall()
    conn.close()
    records = []
    for row in rows:
        record = json.loads(row["data_json"])
        record["_db_id"] = row["id"]
        record["_created_at"] = row["created_at"]
        records.append(record)
    return records


def delete_village(db_id: int):
    """Delete a village record by its database id."""
    conn = get_connection()
    conn.execute("DELETE FROM villages WHERE id = ?", (db_id,))
    conn.commit()
    conn.close()


def delete_all_villages():
    """Delete all village records."""
    conn = get_connection()
    conn.execute("DELETE FROM villages")
    conn.commit()
    conn.close()


def get_village_count() -> int:
    """Return the number of stored villages."""
    conn = get_connection()
    count = conn.execute("SELECT COUNT(*) FROM villages").fetchone()[0]
    conn.close()
    return count


# Initialize DB on import
init_db()
