"""
db_access.py
Provides functions to retrieve notebooks from the database.
"""
import sqlite3

def get_notebooks(db_path):
    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM notebooks")
    rows = cursor.fetchall()
    conn.close()
    return rows
