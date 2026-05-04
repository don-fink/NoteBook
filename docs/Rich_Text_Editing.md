# Rich Text Editing

The right panel is a full-featured rich-text editor. You can format text, create lists, embed images, and paste content from other applications with fine-grained control over what formatting is kept.

## Basic Formatting

Use the toolbar buttons or keyboard shortcuts:

| Format | Shortcut |
|--------|----------|
| Bold | Ctrl+B |
| Italic | Ctrl+I |
| Underline | Ctrl+U |
| Strikethrough | (toolbar button) |

Select the text you want to format first, then apply the shortcut or click the toolbar button. Clicking a formatting button when no text is selected turns that format on for subsequent typing.

## Font and Size

Use the font name and size dropdowns in the toolbar to change the typeface and point size. Select the text first, then choose from the dropdowns.

## Lists

To create a bulleted or numbered list:

1. Place your cursor at the start of a line.
2. Click the appropriate list button in the toolbar (bullet or numbered).
3. Type your item and press **Enter** to start the next item.

To create sub-levels:

- Press **Tab** to indent a list item (move it one level deeper)
- Press **Shift+Tab** to outdent a list item (move it one level up)
- Press **Ctrl+Up** / **Ctrl+Down** to reorder items within the same level

To change the list style (e.g. from disc bullets to numbered), use **Format → List Schemes**.

To end the list, press **Enter** twice at the end of the last item, or click the list button again to toggle it off.

## Paste Modes

NoteBook gives you control over what happens when you paste content copied from another application:

| Mode | What it does |
|------|-------------|
| **Rich** (default) | Pastes with all original formatting preserved |
| **Text-only** | Strips all formatting; pastes as plain text |
| **Match Style** | Keeps structure but applies the current document font |
| **Clean** | Removes most styling while keeping links and images |

Set your default paste mode under **Edit → Default Paste Mode**. Regardless of the default, you can always press **Ctrl+Shift+V** to paste as plain text.

**Tip:** If pasted content looks wrong — wrong font, unexpected colours, or odd spacing — try **Edit → Paste Clean Formatting** or **Ctrl+Shift+V** to paste it again as plain text.

## Images

To insert an image into a page:

1. Use **Notebook → Insert Attachment** and select an image file.
2. Once inserted, click the image to select it — resize handles will appear at the corners.
3. Drag a corner handle to resize the image.

Images are stored inside the database (or in a companion `.media` folder alongside it), so they move with your database file automatically.

## Undo and Redo

- **Ctrl+Z** — Undo the last edit
- **Ctrl+Y** — Redo the last undone edit

Undo history is per-page and is reset when you navigate to a different page.
