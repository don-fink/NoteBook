"""
db_sections.py
Provides functions to retrieve sections for a given notebook from the database.
"""
import sqlite3

def get_sections_by_notebook_id(notebook_id, db_path):
    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM sections WHERE notebook_id = ?", (notebook_id,))
    rows = cursor.fetchall()
    conn.close()
    return rows
