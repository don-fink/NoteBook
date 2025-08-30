"""
db_version.py
Provides functions to get and set the SQLite database schema version using PRAGMA user_version.
"""
import sqlite3

def get_db_version(db_path):
    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()
    cursor.execute('PRAGMA user_version')
    version = cursor.fetchone()[0]
    conn.close()
    return version

def set_db_version(version, db_path):
    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()
    cursor.execute(f'PRAGMA user_version = {int(version)}')
    conn.commit()
    conn.close()
