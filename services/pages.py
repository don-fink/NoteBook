"""
services.pages

Stable import surface for page CRUD and ordering helpers. Currently re-exports
db_pages functions so UI code can depend on this module as we refactor.
"""

from db_pages import (  # noqa: F401
    create_page,
    delete_page,
    get_page_by_id,
    get_pages_by_section_id,
    set_pages_order,
    set_pages_parent_and_order,
    update_page_content,
    update_page_title,
)
