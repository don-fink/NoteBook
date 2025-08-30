import sqlite3

def seed_database(path="notes.db"):
    conn = sqlite3.connect(path)
    conn.execute("PRAGMA foreign_keys = ON")
    cur = conn.cursor()

    # Create schema
    with open("schema.sql", "r") as f:
        cur.executescript(f.read())

    # Insert Notebooks
    notebooks = [("Work",), ("Personal",)]
    cur.executemany(
        "INSERT INTO notebooks (title, order_index) VALUES (?, ?)",
        [(nb[0], idx) for idx, nb in enumerate(notebooks)]
    )

    # Fetch notebook IDs
    cur.execute("SELECT id, title FROM notebooks")
    nb_map = {title: nid for nid, title in cur.fetchall()}

    # Insert Sections
    sections = [
        (nb_map["Work"], "Projects"),
        (nb_map["Work"], "Meeting Notes"),
        (nb_map["Personal"], "Recipes"),
        (nb_map["Personal"], "Travel Plans"),
    ]
    cur.executemany(
        "INSERT INTO sections (notebook_id, title, order_index) VALUES (?, ?, ?)",
        [(nbid, title, idx) for idx, (nbid, title) in enumerate(sections)]
    )

    # Fetch section IDs
    cur.execute("SELECT id, title FROM sections")
    sec_map = {title: sid for sid, title in cur.fetchall()}

    # Insert Pages
    pages = [
        (sec_map["Projects"], "Q3 Roadmap", "<h1>Q3 Roadmap</h1><p>Define milestones...</p>"),
        (sec_map["Meeting Notes"], "Team Sync 2025-08-20", "<p>Attendees: ...</p>"),
        (sec_map["Recipes"], "Chocolate Cake", "<h2>Ingredients</h2><ul><li>...</li></ul>"),
        (sec_map["Travel Plans"], "Lisbon Itinerary", "<p>Day 1: Alfama...</p>"),
    ]
    cur.executemany(
        "INSERT INTO pages (section_id, title, content_html, order_index) VALUES (?, ?, ?, ?)",
        [(sid, title, html, idx) for idx, (sid, title, html) in enumerate(pages)]
    )

    conn.commit()
    conn.close()
    print(f"Seeded database at {path}")

if __name__ == "__main__":
    seed_database()