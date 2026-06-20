import sqlite3
import os
import tempfile
import uuid

# Check if Turso is configured
TURSO_URL = os.environ.get("TURSO_DATABASE_URL")
TURSO_TOKEN = os.environ.get("TURSO_AUTH_TOKEN")

# Enforce HTTPS for Vercel Serverless compatibility (WebSockets are often dropped)
if TURSO_URL and TURSO_URL.startswith("libsql://"):
    TURSO_URL = TURSO_URL.replace("libsql://", "https://", 1)

class DummyCursor:
    def __init__(self, result):
        self.result = result

    def fetchall(self):
        # Return a list of dictionaries representing rows
        return [dict(row) for row in self.result.rows]

    def fetchone(self):
        # Return a single dictionary or None
        if self.result.rows:
            return dict(self.result.rows[0])
        return None

class DBConnectionWrapper:
    def __init__(self, url, token):
        import libsql_client
        # Connect to Turso over HTTP
        self.client = libsql_client.create_client_sync(url=url, auth_token=token)

    def execute(self, sql, args=None):
        if args is None:
            args = []
        # Ensure args is a list as required by some libsql-client bindings
        if isinstance(args, tuple):
            args = list(args)
            
        # Execute query against Turso
        result = self.client.execute(sql, args)
        return DummyCursor(result)

    def commit(self):
        # Turso client automatically handles transactions/commits on simple execute calls
        pass

    def close(self):
        self.client.close()

def get_db_path():
    # Fallback to local SQLite file
    if os.environ.get("VERCEL"):
        return os.path.join(tempfile.gettempdir(), "project_scheduler.db")
    return os.path.join(os.path.dirname(__file__), "project_scheduler.db")

def get_db_connection():
    if TURSO_URL and TURSO_TOKEN:
        return DBConnectionWrapper(TURSO_URL, TURSO_TOKEN)
    else:
        path = get_db_path()
        conn = sqlite3.connect(path)
        conn.row_factory = sqlite3.Row
        conn.execute("PRAGMA foreign_keys = ON")
        return conn

def init_db():
    conn = get_db_connection()
    
    # Only set journal mode if using local SQLite
    if not (TURSO_URL and TURSO_TOKEN):
        try:
            conn.execute("PRAGMA journal_mode = WAL")
        except sqlite3.OperationalError:
            pass
            
    # Create projects table
    conn.execute('''
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
    conn.execute('''
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

# Initialize the database tables on module import
init_db()
