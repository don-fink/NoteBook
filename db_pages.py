"""
db_pages.py
Provides functions to retrieve pages for a given section from the database.
"""
import sqlite3

def get_pages_by_section_id(section_id, db_path):
    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM pages WHERE section_id = ?", (section_id,))
    rows = cursor.fetchall()
    conn.close()
    return rows
