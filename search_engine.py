"""Finding pages: which pages in a scope match a query, in left-panel order.

Pages match on their title and text (HTML stripped, so table text counts). Every word of the
query must appear in the page. A word can match exactly (case and accents ignored, and part
of a longer word counts, so "pump" finds "pumpkin"), or, unless exact matching is asked for,
nearly: a word of 4 or more letters may be off by one letter (two for 7 or more letters).
Exact matches are listed first, then near matches, each in left-panel order.

No Qt here, so this can be used and tested on its own.
"""

import re
import sqlite3
import unicodedata
from dataclasses import dataclass, field
from html.parser import HTMLParser

WORD_RE = re.compile(r"[^\W_]+(?:'[^\W_]+)*", re.UNICODE)

_BLOCK_TAGS = {"p", "div", "br", "li", "tr", "td", "th", "h1", "h2", "h3", "h4", "h5", "h6", "hr"}


class _TextExtractor(HTMLParser):
    def __init__(self):
        super().__init__(convert_charrefs=True)
        self.parts = []
        self._skip = 0

    def handle_starttag(self, tag, attrs):
        if tag in ("style", "script", "head", "title"):
            self._skip += 1
        elif tag in _BLOCK_TAGS:
            self.parts.append(" ")

    def handle_endtag(self, tag):
        if tag in ("style", "script", "head", "title"):
            self._skip = max(0, self._skip - 1)
        elif tag in _BLOCK_TAGS:
            self.parts.append(" ")

    def handle_data(self, data):
        if not self._skip:
            self.parts.append(data)


def html_to_text(html: str) -> str:
    """The visible text of a page's HTML."""
    if not html:
        return ""
    parser = _TextExtractor()
    try:
        parser.feed(html)
        parser.close()
    except Exception:
        return re.sub(r"<[^>]+>", " ", html)
    return "".join(parts for parts in parser.parts)


def fold(text: str) -> str:
    """Lowercase and drop accents, so "Purée" and "puree" compare equal."""
    decomposed = unicodedata.normalize("NFKD", text or "")
    return "".join(c for c in decomposed if not unicodedata.combining(c)).casefold()


def query_terms(query: str) -> list:
    return [fold(w) for w in WORD_RE.findall(query or "")]


def _allowed_typos(term: str) -> int:
    if len(term) < 4:
        return 0
    return 1 if len(term) <= 6 else 2


def _within(a: str, b: str, limit: int) -> bool:
    """True if a and b differ by at most limit single-letter edits."""
    if abs(len(a) - len(b)) > limit:
        return False
    prev = list(range(len(b) + 1))
    for i, ca in enumerate(a, 1):
        cur = [i] + [0] * len(b)
        best = cur[0]
        for j, cb in enumerate(b, 1):
            cur[j] = min(prev[j] + 1, cur[j - 1] + 1, prev[j - 1] + (ca != cb))
            best = min(best, cur[j])
        if best > limit:
            return False
        prev = cur
    return prev[-1] <= limit


@dataclass
class PageText:
    page_id: int
    section_id: int
    notebook_id: int
    title: str
    text: str
    _folded: str = field(default=None, repr=False)
    _words: dict = field(default=None, repr=False)

    def folded(self) -> str:
        if self._folded is None:
            self._folded = fold(self.title + "\n" + self.text)
        return self._folded

    def words(self) -> dict:
        """Folded word -> the word as written (first spelling seen)."""
        if self._words is None:
            self._words = {}
            for w in WORD_RE.findall(self.title + "\n" + self.text):
                self._words.setdefault(fold(w), w)
        return self._words


@dataclass
class Match:
    page_id: int
    section_id: int
    notebook_id: int
    title: str
    exact: bool
    highlight: list  # Strings to highlight on the page


def match_page(page: PageText, terms: list, exact_only: bool = False):
    """A Match if every term is found in the page, else None."""
    if not terms:
        return None
    hay = page.folded()
    highlight, near = [], False
    for term in terms:
        if term in hay:
            highlight.append(term)
            continue
        limit = 0 if exact_only else _allowed_typos(term)
        if not limit:
            return None
        found = [w for f, w in page.words().items() if _within(term, f, limit)]
        if not found:
            return None
        near = True
        highlight.extend(found)
    return Match(page.page_id, page.section_id, page.notebook_id, page.title, not near, highlight)


def search(pages: list, query: str, exact_only: bool = False) -> list:
    """Matches in order: exact matches first, then near matches, each in the given page order."""
    terms = query_terms(query)
    exact, near = [], []
    for page in pages:
        m = match_page(page, terms, exact_only)
        if m is not None:
            (exact if m.exact else near).append(m)
    return exact + near


# --- Reading pages from the database ---


def load_pages(db_path: str, scope: tuple = ("database", None), include_deleted: bool = False):
    """Pages in scope, in the order the left panel shows them.

    scope is ("database", None), ("binder", id), ("section", id) or ("page", id); a page
    scope includes its subpages."""
    con = sqlite3.connect(db_path)
    try:
        live = "" if include_deleted else " AND deleted_at IS NULL"
        binders = con.execute(f"SELECT id FROM notebooks WHERE 1=1{live} ORDER BY order_index, id")
        binder_ids = [r[0] for r in binders]
        kind, sid = scope
        if kind == "binder":
            binder_ids = [b for b in binder_ids if b == sid]
        sections = []
        for b in binder_ids:
            rows = con.execute(
                f"SELECT id FROM sections WHERE notebook_id = ?{live} ORDER BY order_index, id",
                (b,),
            )
            sections.extend((b, r[0]) for r in rows)
        if kind == "section":
            sections = [(b, s) for b, s in sections if s == sid]
        rows = con.execute(
            f"SELECT id, section_id, parent_page_id, title, content_html FROM pages WHERE 1=1{live} "
            "ORDER BY order_index, id"
        ).fetchall()
    finally:
        con.close()

    children = {}
    by_id = {}
    for pid, sec, parent, title, html in rows:
        by_id[pid] = (sec, title or "", html or "")
        key = (sec, int(parent) if parent not in (None, "") else None)
        children.setdefault(key, []).append(pid)

    out = []

    def walk(binder_id, section_id, parent_id, wanted):
        for pid in children.get((section_id, parent_id), []):
            take = wanted or (kind == "page" and pid == sid)
            if take:
                sec, title, html = by_id[pid]
                out.append(PageText(pid, section_id, binder_id, title, html_to_text(html)))
            walk(binder_id, section_id, pid, take)

    for binder_id, section_id in sections:
        walk(binder_id, section_id, None, kind != "page")
    return out
