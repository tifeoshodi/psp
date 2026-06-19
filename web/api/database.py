import sqlite3
import os
import tempfile
import uuid

def get_db_path():
    # If on Vercel, use /tmp. Otherwise use local directory.
    if os.environ.get("VERCEL"):
        return os.path.join(tempfile.gettempdir(), "project_scheduler.db")
    return os.path.join(os.path.dirname(__file__), "project_scheduler.db")

def init_db():
    path = get_db_path()
    conn = sqlite3.connect(path)
    # Enable foreign keys and WAL mode
    conn.execute("PRAGMA foreign_keys = ON")
    try:
        conn.execute("PRAGMA journal_mode = WAL")
    except sqlite3.OperationalError:
        pass # WAL mode not supported on some file systems
    
    cursor = conn.cursor()
    
    # Create projects table
    cursor.execute('''
    CREATE TABLE IF NOT EXISTS projects (
        id TEXT PRIMARY KEY,
        title TEXT NOT NULL,
        start_date TEXT,
        calendar_format TEXT CHECK(calendar_format IN ('5-day week', '6-day week', '7-day week')) DEFAULT '5-day week',
        logo_path TEXT,
        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
        updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
    )
    ''')
    
    # Create activities table
    cursor.execute('''
    CREATE TABLE IF NOT EXISTS activities (
        id TEXT PRIMARY KEY,
        project_id TEXT NOT NULL,
        task TEXT NOT NULL,
        action_needed TEXT,
        duration INTEGER NOT NULL DEFAULT 0,
        precursor TEXT,
        sequence INTEGER NOT NULL,
        resources TEXT,
        budget REAL NOT NULL DEFAULT 0.0,
        section TEXT CHECK(section IN ('Pre-Kickoff Activities', 'Post Kick-off Activities')),
        FOREIGN KEY (project_id) REFERENCES projects(id) ON DELETE CASCADE
    )
    ''')
    
    conn.commit()
    conn.close()

def get_db_connection():
    path = get_db_path()
    conn = sqlite3.connect(path)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON")
    return conn

# Initialize the database on module import
init_db()
