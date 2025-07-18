import sqlite3

DB_PATH = "matches.db"

def init_db():
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute('DROP TABLE IF EXISTS matches')
    c.execute('''CREATE TABLE matches (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        category TEXT,
        subject TEXT,
        sender TEXT,
        date TEXT,
        file_path TEXT,
        sheet TEXT,
        cell TEXT,
        cell_value TEXT
    )''')
    conn.commit()
    conn.close()
