# Excel Search Tool

A simple, client-side web application to load Excel files and search their contents by keywords or exact phrases with customizable search and output options.

---

## Features

- **Load Excel Files** (.xlsx, .xls, .xlsm) directly in your browser (no server needed).
- Select **search mode**:
  - 1 keyword
  - 2 keywords (with optional sequential match)
  - 3 keywords (with optional sequential match)
  - Exact phrase
- Select **one or more columns to search** within.
- Select **which columns to include in the output**.
- Highlight matched keywords in the results.
- Export search results as:
  - **HTML file** (styled table)
  - **Excel file** (.xlsx)
- Responsive UI with dropdown checkboxes for column selections.
- Formats Excel dates into human-readable `DD-MM-YY` format in the output.

---

## How to Use

1. **Open the HTML file** in a modern web browser (Chrome, Firefox, Edge).
2. Click **Choose File** and select an Excel file to load.
3. Select the **Search mode** from the dropdown:
   - If selecting 2 or 3 keywords mode, an option appears to enable matching keywords in sequence.
4. Enter the keyword(s) or exact phrase to search.
5. Select the **column to search** (from columns in the loaded Excel file).
6. Click the **output columns dropdown** and check the columns you want to include in the results.
7. Click the **Search** button.
8. Results appear in a table with matched keywords highlighted.
9. Use the **Download HTML** or **Download Excel** buttons to save the results.

---

## Supported Excel File Formats

- `.xlsx`
- `.xls`
- `.xlsm`

---

## Notes

- The tool reads only the **first worksheet** of the Excel file.
- The default selected search column is `"Content"` if present; otherwise, the first column.
- The default output includes the `"Content"` column checked.
- Search is **case-insensitive**.
- Sequential keyword matching means keywords must appear in order but can have other text in between.
- Duplicate rows are automatically filtered out.
- Date columns are automatically formatted from Excel's numeric date codes.

---

## Dependencies

- [SheetJS (xlsx) library](https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.16.9/xlsx.full.min.js) — included via CDN for parsing Excel files.

---

## Customization

- You can modify the styles in the `<style>` section to adjust appearance.
- The code supports multiple keywords and exact phrase search modes and can be extended for additional features.

---

## License

This project is open-source and free to use.
