# Excel Productivity Suite — User Guide

**Product Version:** 1.0.0  
**Host Application:** Microsoft Excel 2016 / 2019 / 2021 / Office 365 (Windows Desktop)  

---

## 1. Introduction

**Excel Productivity Suite** is a high-performance Windows desktop add-in designed to streamline everyday worksheet tasks, automate data cleaning, accelerate text transformations, and simplify workbook management.

### Key Highlights
- **Integrated Office Ribbon:** Beautifully organized tabs (`Excel Productivity` and `Workbook & Sheets`) placed naturally alongside standard Excel tabs.
- **Quick Tool Launcher (SuperSearch):** Press `Ctrl + Shift + P` or click **Tool Launcher** to search, filter, and launch any tool instantly without menu diving.
- **In-Memory Performance:** All range operations process data in managed 2D memory arrays, executing operations on tens of thousands of cells in milliseconds.
- **Live Before-and-After Previews:** High-impact text and range dialogs feature real-time previews showing sample rows before committing changes.
- **100% Local & Secure:** Zero cloud dependencies, zero data telemetry. All processing stays entirely on your local machine.

---

## 2. Ribbon Tour

The suite installs two primary Ribbon tabs:

### 2.1 Tab: `Excel Productivity`
Contains daily data manipulation and formatting utilities:
- **Productivity Suite Group:**
  - `Tool Launcher`: Open the searchable tool command palette.
  - `About Suite`: Inspect product version, architecture, and diagnostics.
- **Text Operations Group:**
  - `Extract Text`: Extract substrings before, after, or between delimiters into adjacent columns or overwrite.
  - `Split Cells`: Split cell contents across multiple columns or down into multiple rows.
  - `Remove Characters`: Strip numbers, letters, symbols, non-printable codes, or custom characters.
  - `Trim Spaces`: Collapse excess spaces, remove leading/trailing spaces, or clean non-breaking spaces.
  - `Change Case`: Convert to UPPERCASE, lowercase, Proper Case, or Sentence case.
  - `Add Text`: Prepend prefixes or append suffixes with duplicate prevention.
- **Range & Cells Group:**
  - `Fill Blank Cells`: Populate empty cells using values from above, below, left, right, or constants.
- **Data Cleaning Group:**
  - `Remove Duplicates`: Identify and clear duplicate rows in selected data.
  - `Formulas to Values`: Replace live calculated formulas with static values in-place instantly.

### 2.2 Tab: `Workbook & Sheets`
Contains workbook-wide management utilities:
- **Worksheet Tools Group:**
  - `Sheet Visibility`: Bulk manager to show, hide, or very-hide worksheets. Includes "Unhide All Hidden Sheets" and "Hide Inactive Sheets".
  - `Rename Sheets`: Bulk rename worksheets using prefix, suffix, find-and-replace, and sequential numbering rules.

---

## 3. Tool Walkthroughs

### 3.1 Quick Tool Launcher
- **How to Open:** Click **Tool Launcher** on the Ribbon or press `Ctrl + Shift + P`.
- **How to Use:**
  1. Type any keyword (e.g., `trim`, `split`, `case`, `blanks`).
  2. Use `Up` and `Down` arrow keys to highlight a tool.
  3. Press `Enter` or double-click to launch the utility immediately.
  4. Use the category pills at the top (`Text`, `Range & Cells`, `Data Cleaning`, `Worksheets`) to filter by category.

---

### 3.2 Extract Text
- **When to Use:** When you have compound data in single cells (e.g., `Smith, John`, `Order #1042 [US]`, `name@domain.com`).
- **Steps:**
  1. Select the range of cells you wish to extract from.
  2. Click **Extract Text** on the Ribbon.
  3. Choose extraction mode: `Before Delimiter`, `After Delimiter`, or `Between Delimiters`.
  4. Enter your delimiter(s).
  5. Inspect the **Live Preview** table.
  6. Click **Apply & Extract**.

---

### 3.3 Split Cells
- **When to Use:** When a single cell contains multiple values separated by commas, spaces, semicolons, or line breaks that you want to expand across columns or into multiple rows.
- **Steps:**
  1. Select the column of cells to split.
  2. Click **Split Cells**.
  3. Select your split direction (`Split to Columns` or `Split to Rows`).
  4. Choose your delimiter (Comma, Space, Semicolon, Newline, or Custom).
  5. Check **Trim whitespace from tokens** to ensure clean output.
  6. Review the **Live Preview** and click **Split Cells**.

---

### 3.4 Fill Blank Cells
- **When to Use:** When flattening tabular data or preparing pivot table output for analysis where category headers only appear on the first row of each group.
- **Steps:**
  1. Select the range containing blank cells.
  2. Click **Fill Blank Cells**.
  3. Select **Fill Down (Value from above)** or choose another direction.
  4. Click **Fill Blank Cells**. All empty cells are filled in milliseconds.

---

### 3.5 Remove Duplicates
- **When to Use:** When cleaning messy data exports with duplicate rows.
- **Steps:**
  1. Select the data table or range.
  2. Click **Remove Duplicates**.
  3. Check **My data has headers** if your table has a header row (preserves headers).
  4. Select which specific columns to evaluate in the **Columns to Compare** checklist, or click **Select All**.
  5. Choose whether to **Keep First occurrence** or **Keep Last occurrence**.
  6. Review the live duplicate row count in the **Preview Box**.
  7. Click **Clear Duplicates**. A confirmation prompt ensures intentional execution.

---

### 3.6 Compare Ranges / Datasets (Phase 2)
- **When to Use:** When reconciling two sheets, comparing budget vs. actuals, or auditing changes between two versions of a data table.
- **Steps:**
  1. Select the primary source range in Excel.
  2. Click **Compare Ranges** on the Ribbon.
  3. Enter the target range address (e.g. `B1:B20` or `Sheet2!A1:D10`).
  4. Choose your comparison mode:
     - `Cell-by-cell (Side-by-side)`: Directly compares cells at identical relative coordinates.
     - `Values in Source Not in Target`: Finds values missing from the target range.
     - `Common Values`: Finds intersecting matching records.
  5. Inspect the discrepancy table showing cell coordinates, source and target values, and discrepancy types.

---

### 3.7 Select Special Cells (Phase 2)
- **When to Use:** When you need to quickly locate, audit, or jump to specific cells in a large table.
- **Steps:**
  1. Select the range to inspect.
  2. Click **Select Special** on the Ribbon.
  3. Choose your filter criteria (`Blank Cells`, `Error Cells`, `Formula Cells`, `Numeric`, `Text`, `Maximum`, or `Minimum`).
  4. The dialog displays all matched cells with their addresses and values.
  5. Click **Select First Matching** to jump directly to the target cell in Excel.

---

### 3.8 Batch Rename Worksheets
- **When to Use:** When organizing monthly or quarterly reports (e.g. renaming `Sheet1`, `Sheet2` to `Report_2026_01`, `Report_2026_02`).
- **Steps:**
  1. Click **Rename Sheets** under the **Workbook & Sheets** tab.
  2. Enter a prefix (e.g. `2026_Q3_`) or suffix.
  3. Check **Append sequential numbers** if you want automatic numbering.
  4. Review the **Proposed Sheet Names Preview** to ensure all names conform to Excel rules.
  5. Click **Rename All Sheets**.

---

## 4. Performance & Safety Tips

1. **Large Range Optimization:** For ranges exceeding 50,000 cells, the suite automatically suppresses Excel screen updating and calculation during processing and restores them immediately upon completion.
2. **Formula Cells:** Destructive text tools skip cells containing formulas by default, safeguarding your calculation models against unintended text overwrites.
3. **No Excel Freezing:** All dialogs run on the STA thread with Excel window ownership, ensuring dialogs stay properly centered and never trap keyboard focus.
