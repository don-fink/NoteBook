PRAGMA foreign_keys = ON;

CREATE TABLE notebooks (
  id            INTEGER PRIMARY KEY AUTOINCREMENT,
  title         TEXT    NOT NULL,
  created_at    TEXT    NOT NULL DEFAULT (datetime('now')),
  modified_at   TEXT    NOT NULL DEFAULT (datetime('now')),
  order_index   INTEGER NOT NULL DEFAULT 0
);

CREATE TABLE sections (
  id            INTEGER PRIMARY KEY AUTOINCREMENT,
  notebook_id   INTEGER NOT NULL,
  title         TEXT    NOT NULL,
  color_hex     TEXT,
  created_at    TEXT    NOT NULL DEFAULT (datetime('now')),
  modified_at   TEXT    NOT NULL DEFAULT (datetime('now')),
  order_index   INTEGER NOT NULL DEFAULT 0,
  FOREIGN KEY (notebook_id) REFERENCES notebooks(id) ON DELETE CASCADE
);

CREATE TABLE pages (
  id            INTEGER PRIMARY KEY AUTOINCREMENT,
  section_id    INTEGER NOT NULL,
  title         TEXT    NOT NULL,
  content_html  TEXT    NOT NULL,
  created_at    TEXT    NOT NULL DEFAULT (datetime('now')),
  modified_at   TEXT    NOT NULL DEFAULT (datetime('now')),
  order_index   INTEGER NOT NULL DEFAULT 0,
  FOREIGN KEY (section_id) REFERENCES sections(id) ON DELETE CASCADE
);