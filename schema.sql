PRAGMA foreign_keys = ON;

CREATE TABLE notebooks (
  id            INTEGER PRIMARY KEY AUTOINCREMENT,
  title         TEXT    NOT NULL,
  created_at    TEXT    NOT NULL DEFAULT (datetime('now')),
  modified_at   TEXT    NOT NULL DEFAULT (datetime('now')),
  order_index   INTEGER NOT NULL DEFAULT 0,
  deleted_at    TEXT    NULL
);

CREATE TABLE sections (
  id            INTEGER PRIMARY KEY AUTOINCREMENT,
  notebook_id   INTEGER NOT NULL,
  title         TEXT    NOT NULL,
  color_hex     TEXT,
  created_at    TEXT    NOT NULL DEFAULT (datetime('now')),
  modified_at   TEXT    NOT NULL DEFAULT (datetime('now')),
  order_index   INTEGER NOT NULL DEFAULT 0,
  deleted_at    TEXT    NULL,
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
  parent_page_id INTEGER NULL,
  deleted_at    TEXT    NULL,
  FOREIGN KEY (section_id) REFERENCES sections(id) ON DELETE CASCADE
);

-- Helpful index for hierarchical lookups
CREATE INDEX IF NOT EXISTS idx_pages_parent ON pages(parent_page_id);

-- Index for soft-delete queries (filter active vs deleted items)
CREATE INDEX IF NOT EXISTS idx_notebooks_deleted ON notebooks(deleted_at);
CREATE INDEX IF NOT EXISTS idx_sections_deleted ON sections(deleted_at);
CREATE INDEX IF NOT EXISTS idx_pages_deleted ON pages(deleted_at);