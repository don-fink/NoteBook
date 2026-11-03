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
2. To change its size, move the mouse over the image and drag one of the small square handles at its corners. An outline shows the new size, with its width and height in pixels, and the image changes when you let go. The image always keeps its shape, and it can't be dragged wider than the page, or than its table cell. **Ctrl+Z** puts it back.
3. You can also right-click the image and choose **Image Properties…**, **Fit to editor width**, or **Reset to original size**. With the caret next to an image, **F2** or **Ctrl+Shift+I** opens Image Properties.

Videos on a page appear as a picture from the video, and their corner handles work the same way. Dragging a handle doesn't open the video; clicking the picture still does.

Images are stored inside the database (or in a companion `.media` folder alongside it), so they move with your database file automatically.

## Links

To turn text into a link:

1. Select the text.
2. Right-click the selection and choose **Add Link…**.
3. Type or paste the link, then click **OK**. **OK** stays unavailable until the entry can be used as a link.

You can enter a web address (`example.com` becomes `https://example.com`), an email address, or a file path. To change a link, select its text and add the new link over it. Click a link to open it.

### Linking to a file

The easiest way to get a file's full address is **Copy as path** in File Explorer:

1. In File Explorer, right-click the file and choose **Copy as path**. On Windows 10, hold **Shift** while you right-click. You can also select the file and press **Ctrl+Shift+C** on Windows 11.
2. In NoteBook, select the text for the link, right-click it and choose **Add Link…**.
3. Press **Ctrl+V** to paste the address, then click **OK**.

The link points to the file where it is now and doesn't copy it into your notebook, so the link stops working if the file is moved or renamed. To keep a copy of the file with your notes, use **Notebook → Insert Attachment** instead.

### Linking to an email address

Clicking an email link opens whichever program Windows uses for email links. To change it, go to **Settings → Apps → Default apps** in Windows and set a default for the link type `MAILTO`.

## Undo and Redo

- **Ctrl+Z** — Undo the last edit
- **Ctrl+Y** — Redo the last undone edit

Undo history is per-page and is reset when you navigate to a different page.
