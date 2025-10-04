#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
===========================================================
QUICK START: HOW TO RUN THIS SCRIPT
===========================================================

1. In Scribus, go to Script → Execute Script…
2. Browse to this file (e.g., update_titles.py) and select it.
3. When prompted, choose your CSV file.
4. The script will insert the CSV text into the first linked frame
   (named "TitleFrame" by default) and flow it across the chain.

===========================================================
DETAILED INSTRUCTIONS
===========================================================

1. Saving and Running
   - Save this file with a .py extension (e.g., update_titles.py).
   - Run it in Scribus via Script → Execute Script….
   - The script will prompt you to choose a CSV file.
   - It will then insert the CSV text into the first linked frame and flow it across the chain.

2. Naming the First Linked Frame
   - You must give the first text frame in your linked chain a name.
   - To do this:
       a. Select the first frame in Scribus.
       b. Go to Windows → Properties → X, Y, Z tab (or right-click → Properties).
       c. In the “Name” field, type exactly the name you set in the script (default: TitleFrame).
   - Only the first frame needs a name. Scribus will automatically flow the text into the rest of the linked frames.

3. CSV Requirements
   - Save your spreadsheet as CSV UTF-8 (Comma delimited) from Excel.
   - If your cells contain Alt+Enter line breaks, Excel will wrap them in quotes automatically.
   - This script handles embedded commas, quotes, and newlines correctly.
   - If your CSV has a header row, set SKIP_HEADER = True in the script.
   - If your CSV has no header row, leave SKIP_HEADER = False.

4. Fonts in Scribus
   - The font name in the script must match exactly the name Scribus uses internally.
   - To check:
       a. In Scribus, go to File → Preferences → Fonts.
       b. Look at the list of installed fonts.
       c. Copy the name exactly as it appears (including spaces, capitalization, and style).
   - Example: If Scribus lists it as "Comic Sans MS Regular", you must use exactly:
       FONT_NAME = "Comic Sans MS Regular"

5. Updating the Script
   - To change the font, update the FONT_NAME variable near the top of the script.
   - To change font size, update FONT_SIZE.
   - To change line spacing, update LINE_SPACING.
   - To change alignment, update JUSTIFY_MODE:
       0 = left, 1 = center, 2 = right, 3 = block (justified).
   - To skip a header row in your CSV, set SKIP_HEADER = True.

6. Debugging
   - A debug block is included but commented out.
   - To re-enable it, remove the # in front of the lines inside the "# ;;;;;;;;" markers.
   - This will show the first 3 parsed rows in message boxes for verification.

7. Success Message
   - After running, you’ll see a message like: “42 rows updated.”
   - This number refers to the number of Excel rows processed from your CSV.
   - If a single cell contains multiple lines (from Alt+Enter), it still counts as one row.
   - So the count matches your Excel rows, not the number of visible line breaks.

===========================================================
COMMON ERRORS & FIXES
===========================================================

• Error: Row 1 shows \ufeff at the start
  - Cause: Excel adds a BOM (Byte Order Mark) when saving as UTF-8.
  - Fix: This script already opens with encoding="utf-8-sig", which strips the BOM.

• Error: Quotes still appear around text
  - Cause: Excel wraps cells with commas/newlines in quotes.
  - Fix: The script strips wrapping quotes automatically.

• Error: Font not found
  - Cause: The font name in the script doesn’t match Scribus’ internal name.
  - Fix: Go to File → Preferences → Fonts and copy the name exactly as shown.

• Error: “Frame not found”
  - Cause: The first linked frame isn’t named correctly.
  - Fix: Select the first frame, open Properties → X, Y, Z tab, and set its name to match FRAME_NAME.

• Error: Text overflows
  - Cause: Not enough linked frames to hold all the text.
  - Fix: Add more linked frames, or reduce text size/line spacing.
  - The script will warn you with:
    “Text flows beyond linked frames starting at 'TitleFrame'.”
    → Add more linked frames or adjust formatting (font size/line spacing).

• Error: First row skipped
  - Cause: SKIP_HEADER is set to True.
  - Fix: If your CSV has no header, set SKIP_HEADER = False.

===========================================================
END OF INSTRUCTIONS
===========================================================
"""

import csv
import scribus

# Formatting
FRAME_NAME     = "TitleFrame"            # Must match the name of the FIRST linked frame in Scribus
FONT_NAME      = "Comic Sans MS Regular" # Must match Scribus font list exactly
FONT_SIZE      = 20                      # pt
JUSTIFY_MODE   = 1                       # 0=left, 1=center, 2=right, 3=block
LINE_SPACING   = 23                      # pt
LINE_MODE      = 0                       # 0=fixed, 1=automatic

# CSV parsing
COLUMN_INDEX   = 0                       # Use the first column
SKIP_HEADER    = False                   # Set True if your CSV has a header row
DELIMITER      = ","                     # Excel CSV default
QUOTECHAR      = '"'                     # Excel CSV default

def choose_csv_file():
    path = scribus.fileDialog("Select CSV File", "*.csv")
    if not path:
        scribus.messageBox("Cancelled", "No file selected.", icon=0)
        return None
    return path

def read_csv_column(path, column_index=COLUMN_INDEX, skip_header=SKIP_HEADER):
    values = []
    # utf-8-sig removes BOM if present
    with open(path, mode="r", encoding="utf-8-sig", newline="") as f:
        reader = csv.reader(f, delimiter=DELIMITER, quotechar=QUOTECHAR)
        for i, row in enumerate(reader):
            if skip_header and i == 0:
                continue
            if len(row) > column_index:
                cell = row[column_index]
                # Normalize Windows CRLF to LF
                cell = cell.replace("\r\n", "\n").replace("\r", "\n")
                # Strip wrapping quotes if they remain
                if cell.startswith('"') and cell.endswith('"'):
                    cell = cell[1:-1]
                values.append(cell)
            else:
                values.append("")
    return values

def apply_formatting(frame_name):
    length = scribus.getTextLength(frame_name)
    if length <= 0:
        return
    scribus.selectText(0, length, frame_name)
    scribus.setFont(FONT_NAME, frame_name)
    scribus.setFontSize(FONT_SIZE, frame_name)
    scribus.setTextAlignment(JUSTIFY_MODE, frame_name)
    scribus.setLineSpacingMode(LINE_MODE, frame_name)
    if LINE_MODE == 0:
        scribus.setLineSpacing(LINE_SPACING, frame_name)

def update_from_csv(path, frame_name=FRAME_NAME):
    if not scribus.haveDoc():
        scribus.messageBox("Error", "No document open.", icon=2)
        return
    if not scribus.objectExists(frame_name):
        scribus.messageBox("Error", f"Frame '{frame_name}' not found.", icon=2)
        return

    items = read_csv_column(path)
    if not items:
        scribus.messageBox("Error", "CSV has no data rows.", icon=2)
        return

    # ;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
    # 🔍 Debug: show first 3 parsed values in message boxes
    # for i, val in enumerate(items[:3]):
    #     scribus.messageBox("Debug", f"Row {i+1}:\n{repr(val)}", icon=0)
    # ;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;

    # Join rows with a single newline; embedded newlines inside cells are preserved
    text = "\n".join(items)

    scribus.selectObject(frame_name)
    scribus.deleteText(frame_name)
    scribus.insertText(text, 0, frame_name)

    apply_formatting(frame_name)

    if scribus.textOverflows(frame_name):
        scribus.messageBox(
            "Overflow Warning",
            f"Text flows beyond linked frames starting at '{frame_name}'.\n\n"
            "→ Add more linked frames or reduce font size/line spacing.",
            icon=1
        )
    else:
        row_count = len(items)
        scribus.messageBox("Success", f"{row_count} rows updated.", icon=0)

if __name__ == "__main__":
    csv_file = choose_csv_file()
    if csv_file:
        update_from_csv(csv_file, frame_name=FRAME_NAME)
