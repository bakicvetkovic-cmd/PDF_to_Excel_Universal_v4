# User Guide — PDF to Excel Universal Converter v4

## 1. Overview

PDF to Excel Universal Converter is a Windows desktop application that extracts selectable text from PDF pages and places it into an Excel workbook. You define the output columns by moving red vertical boundaries and set the first row to process with the blue **Data Start (Y)** line.

The application supports:
- Moving and resizing column boundaries on the PDF preview.
- Adding, deleting, and renaming output columns.
- Saving and loading reusable JSON templates.
- Using shared column widths across all pages or independent widths for each page.
- Using one shared data-start position or independent Data Start (Y) positions per page.
- Exporting extracted rows to an `.xlsx` workbook with a header row, filters, and frozen header.

## 2. Requirements and installation

- Windows 10 or Windows 11.
- Python 3.10 or newer.
- Internet connection during first-time package installation.

If Python is not installed, download it from https://www.python.org/downloads/ and enable **Add Python to PATH** during installation.

Extract the ZIP archive, then double-click `INSTALL.bat`. It installs:
- PyMuPDF — reads and renders PDF pages and extracts text.
- openpyxl — creates Excel workbooks.

The batch file then starts the application. If installation fails, verify Python is installed and available using the `py` command.

## 3. Main window

### Toolbar
- **Open PDF** — select a PDF file.
- **◀ / ▶** — move to the previous or next page.
- **Add Column** — add an output column.
- **Delete Column** — delete the selected column in the column list.
- **Rename** — rename the selected output column.
- **Save Template** — save column names, boundaries, and data-start settings as JSON.
- **Load Template** — load a previously saved JSON template.
- **Separate Column Widths by Page** — create independent column boundaries for each PDF page.
- **Separate Data Start Y by Page** — allow the blue start line to be positioned independently on each page.
- **EXPORT TO EXCEL** — extract the rows and save an Excel workbook.

The right-hand panel lists output columns and their horizontal coordinates (**From** and **To**). Select a column in this list before deleting or renaming it.

## 4. Setting up the PDF

1. Click **Open PDF** and choose a PDF.
2. The PDF page appears in the preview. Red vertical lines indicate column boundaries. The blue horizontal line indicates where extraction begins.
3. Drag a red line left or right to adjust the boundary. The column list updates to show the coordinates.
4. Drag the blue line up or down to change the starting Y coordinate.
5. Use **◀ / ▶** to inspect other pages.

By default, column boundaries and the blue start line are shared across all pages. This works when pages have the same layout.

### Different column widths on different pages
If page layouts differ, click **Separate Column Widths by Page**. The current column layout is copied to each page. Then navigate page by page and adjust that page’s red boundaries independently. This separation is enabled for the current session; save a template to preserve the settings.

### Different data start positions on different pages
Click **Separate Data Start Y by Page** and adjust the blue line on each page individually. This is useful when headers or page breaks differ. Save a template to preserve the per-page values.

## 5. Managing output columns

- **Add Column** asks for a column name and appends a new column to the current layout.
- **Delete Column** removes the selected column from the list.
- **Rename** changes the selected column’s header.
- The initial column layout is a preset intended as a starting point for accounting-card style PDFs. Adjust or replace it for your own document.

Column boundaries are horizontal X coordinates on the PDF page. During extraction, each word is assigned to the column containing the horizontal center of that word.

## 6. Saving and reusing templates

Click **Save Template**, enter a name, and the application saves a JSON template in your user profile under `.pdf_excel_universal/templates`.

A template stores column names, X boundaries, the shared data-start Y coordinate, and—when enabled—per-page column widths and per-page Y positions.

To reuse a template, open a PDF with a compatible layout and click **Load Template**. Check the preview and adjust boundaries if the PDF differs.

## 7. Exporting to Excel

1. Confirm the column boundaries and Data Start line on all relevant pages.
2. Click **EXPORT TO EXCEL**.
3. Choose the output `.xlsx` file name and location.
4. The application creates a worksheet named **Data**, writes the column names in the first row, and appends extracted rows from the PDF pages.
5. The header row is frozen and an Excel filter is enabled.

For columns whose names contain **debit**, **credit**, **balance**, **amount**, or **turnover**, the converter attempts to parse values as numbers using European number formatting (periods as thousands separators and comma as decimal separator). Other values are exported as text.

## 8. Important limitations and troubleshooting

- **Scanned PDF:** This version does not include OCR. If you cannot select/copy text from the PDF in a PDF reader, the document is likely image-only and must be OCR-processed first.
- **PDF layout:** This is a coordinate-based extraction tool, not a general-purpose table-recognition engine. Complex tables, merged cells, multi-line rows, rotated text, and irregular layouts may need manual adjustment or may not extract perfectly.
- **Text order:** Words are grouped into rows by their vertical positions and then assigned to columns by their horizontal centers. Review the resulting Excel file.
- **Missing packages:** Run `INSTALL.bat` again, or install dependencies manually with `py -m pip install pymupdf openpyxl`.
- **No PDF text detected:** Confirm the PDF contains selectable text and is not a scan.
- **Unexpected column values:** Adjust the red boundaries so each word’s center falls inside the intended column.

Always review the exported workbook before using it for accounting, reporting, or other important work.
