# Tables and Currency Columns

NoteBook supports rich tables with automatic currency formatting and column totals, making it useful for budgets, expense tracking, and any structured data you want to keep alongside your notes.

## Inserting a Table

1. Place your cursor in the page where you want the table.
2. Choose **Insert → Insert Table…**, or right-click and choose **Insert → Table…**.
3. Enter the number of rows and columns you want.

## Presets

A preset is a saved copy of a table or Planning Register that you can insert again on any page. It comes back exactly as you saved it, including any text and amounts in it, so you can keep pre-filled tables as well as empty layouts.

Tables and Planning Registers have separate preset lists. Presets are stored in the database, so they're included in your database backups. Each database has its own presets.

All preset commands are on the **Insert** menu, and the same commands are under **Insert** when you right-click a page:

- **Insert Table from Preset…** / **Insert Planning Register from Preset…** — pick a preset from the list and insert it at the cursor.
- **Save Table as Preset…** / **Save Planning Register as Preset…** — place the cursor in the table or register, choose the command and give the preset a name. If the name is already used, NoteBook asks before replacing that preset. When you right-click inside a table, the Save commands are on the menu too.
- **Manage Table Presets…** / **Manage Planning Register Presets…** — rename or delete saved presets.

Saving a table that sits inside a Planning Register saves just that table. To save the whole register, use **Save Planning Register as Preset…** from anywhere inside it.

## Navigating a Table

| Action | Key |
|--------|-----|
| Move to next cell | Tab |
| Move to previous cell | Shift+Tab |
| Add a new row at the end | Tab from the very last cell |

You can also click any cell to position your cursor there directly.

## Editing Cells

Click into any cell and type normally. Text in table cells supports the same rich text formatting as the rest of the page — bold, italic, font changes, and so on.

## Currency Columns

You can mark one or more columns to automatically format their values as currency and display a running total at the bottom.

**To mark a column as currency:**

1. Right-click anywhere in the table.
2. Choose **Mark Column(s) as Currency + Total**.
3. Select which columns should be treated as currency.

Once a column is marked:

- Numbers you type are automatically formatted as `$1,234.56` when you leave the cell.
- A **Total** row appears at the bottom of the column and updates automatically as you edit values.
- The Total row is calculated and read-only — you cannot type into it directly.

**To remove currency formatting:**

1. Right-click the table.
2. Choose **Remove Currency from Column** and select the column to unmark.

## Table Context Menu

Right-clicking anywhere inside a table gives you a context menu with options to:

- Insert or delete rows and columns
- Mark or unmark columns as currency
- Save the table, or the Planning Register it's in, as a preset
- Insert a table, Planning Register or preset (under **Insert**)

## Tips

- Press **Tab** in the last cell of a table to automatically insert a new row — no need to use the context menu for routine data entry.
- You can type a plain number like `1500` into a currency cell. The `$1,500.00` formatting is applied automatically when you move to another cell.
- If a Total is not updating, click away from the table and back — the calculation refreshes on focus changes.
- Tables work well for simple budgets. Put your income rows at the top, expenses below, and mark the amount column as currency for instant totals.
