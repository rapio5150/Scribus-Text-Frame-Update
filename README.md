# Scribus CSV Text Frame Updater

A Python script for [Scribus](https://www.scribus.net) that updates a linked text frame chain with content from a CSV file.  
Designed for publishing pipelines where titles, captions, or other text need to be refreshed directly from Excel/CSV data.

---

## ✨ Features
- Reads a specified column from a CSV file (UTF‑8, Excel‑compatible).
- Handles embedded commas, quotes, and Alt+Enter line breaks correctly.
- Inserts text into the first named frame and flows across linked frames.
- Applies consistent font, size, alignment, and line spacing automatically.
- Provides clear success messages (e.g. **“40 rows updated.”**).
- Warns if text overflows beyond the linked frames.
- Includes optional debug markers to preview parsed rows.

---

## 📋 Requirements
- Scribus 1.5+ (tested with 1.7.0svn).
- Python 3 (bundled with Scribus).
- A Scribus document with a **linked text frame chain**.
- The first frame in the chain must be named (default: `TitleFrame`).

---

## 🚀 Usage
1. Save the script as `update_titles.py`.
2. In Scribus, go to **Script → Execute Script…** and select the file.
3. Choose your CSV file when prompted.
4. The script will:
   - Clear the first frame,
   - Insert the CSV text,
   - Flow it across linked frames,
   - Apply formatting,
   - Show a confirmation message.

---

## ⚙️ Configuration
Edit the variables at the top of the script to match your needs:

```python
FRAME_NAME   = "TitleFrame"            # First linked frame name
FONT_NAME    = "Comic Sans MS Regular" # Must match Scribus font list
FONT_SIZE    = 20                      # pt
JUSTIFY_MODE = 1                       # 0=left, 1=center, 2=right, 3=block
LINE_SPACING = 23                      # pt
SKIP_HEADER  = False                   # True if CSV has a header row
