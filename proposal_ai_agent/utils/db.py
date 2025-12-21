# utils/db.py
import sqlite3
from datetime import datetime
from typing import Optional, Tuple, List
from pathlib import Path
from utils.paths import SQLITE_PATH

def get_conn():
    conn = sqlite3.connect(SQLITE_PATH)
    conn.execute("PRAGMA foreign_keys = ON;")
    return conn

def init_db():
    conn = get_conn()
    cur = conn.cursor()
    # 고객 테이블
    cur.execute("""
    CREATE TABLE IF NOT EXISTS clients(
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        name TEXT UNIQUE NOT NULL,
        capability TEXT,
        headcount TEXT,
        past_projects TEXT,
        created_at TEXT NOT NULL
    );
    """)
    # 프로젝트 테이블
    cur.execute("""
    CREATE TABLE IF NOT EXISTS projects(
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        client_id INTEGER NOT NULL,
        title TEXT NOT NULL,
        direction TEXT NOT NULL,
        status TEXT DEFAULT 'NEW',
        created_at TEXT NOT NULL,
        FOREIGN KEY(client_id) REFERENCES clients(id) ON DELETE CASCADE
    );
    """)
    # RFP 파일 테이블
    cur.execute("""
    CREATE TABLE IF NOT EXISTS rfp_files(
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        project_id INTEGER NOT NULL,
        filename TEXT NOT NULL,
        stored_path TEXT NOT NULL,
        uploaded_at TEXT NOT NULL,
        FOREIGN KEY(project_id) REFERENCES projects(id) ON DELETE CASCADE
    );
    """)
    conn.commit()
    conn.close()

def upsert_client(name: str,
                  capability: str = "",
                  headcount: str = "",
                  past_projects: str = "") -> int:
    conn = get_conn()
    cur = conn.cursor()
    now = datetime.utcnow().isoformat()

    cur.execute("SELECT id FROM clients WHERE name=?", (name,))
    row = cur.fetchone()
    if row:
        cid = row[0]
        if any([capability, headcount, past_projects]):
            cur.execute("""
                UPDATE clients SET
                    capability=COALESCE(NULLIF(?, ''), capability),
                    headcount=COALESCE(NULLIF(?, ''), headcount),
                    past_projects=COALESCE(NULLIF(?, ''), past_projects)
                WHERE id=?
            """, (capability, headcount, past_projects, cid))
    else:
        cur.execute("""
            INSERT INTO clients(name, capability, headcount, past_projects, created_at)
            VALUES (?, ?, ?, ?, ?)
        """, (name, capability, headcount, past_projects, now))
        cid = cur.lastrowid

    conn.commit()
    conn.close()
    return cid

def create_project(client_id: int, title: str, direction: str) -> int:
    conn = get_conn()
    cur = conn.cursor()
    now = datetime.utcnow().isoformat()
    cur.execute("""
        INSERT INTO projects(client_id, title, direction, created_at)
        VALUES (?, ?, ?, ?)
    """, (client_id, title, direction, now))
    pid = cur.lastrowid
    conn.commit()
    conn.close()
    return pid

def attach_rfp(project_id: int, filename: str, stored_path: Path):
    conn = get_conn()
    cur = conn.cursor()
    now = datetime.utcnow().isoformat()
    cur.execute("""
        INSERT INTO rfp_files(project_id, filename, stored_path, uploaded_at)
        VALUES (?, ?, ?, ?)
    """, (project_id, filename, str(stored_path), now))
    conn.commit()
    conn.close()

def fetch_client_names() -> List[str]:
    conn = get_conn()
    rows = conn.execute("SELECT name FROM clients ORDER BY name").fetchall()
    conn.close()
    return [r[0] for r in rows]

def fetch_client_info(client_id: int) -> Tuple[str, str, str]:
    conn = get_conn()
    row = conn.execute("""
        SELECT capability, headcount, past_projects
        FROM clients WHERE id=?
    """, (client_id,)).fetchone()
    conn.close()
    return row if row else ("", "", "")

def fetch_client_id_by_name(name: str) -> Optional[int]:
    conn = get_conn()
    row = conn.execute("SELECT id FROM clients WHERE name=?", (name,)).fetchone()
    conn.close()
    return row[0] if row else None

def list_projects(client_id: int):
    conn = get_conn()
    rows = conn.execute("""
        SELECT id, title, direction, created_at, status
        FROM projects WHERE client_id=? ORDER BY id DESC
    """, (client_id,)).fetchall()
    conn.close()
    return rows

def list_rfp_files(project_id: int):
    conn = get_conn()
    rows = conn.execute("""
        SELECT filename, stored_path, uploaded_at
        FROM rfp_files WHERE project_id=? ORDER BY id DESC
    """, (project_id,)).fetchall()
    conn.close()
    return rows
