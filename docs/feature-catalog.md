# Feature Catalog: Excel Productivity Suite

**Product Identity:** Excel Productivity Suite  
**Version:** 1.0.0-dev  
**Host Application:** Microsoft Excel 2016+ (64-bit / 32-bit Windows Desktop)  

---

## 1. Feature Lifecycle & Development Phases

To ensure stability, code quality, and maintainability, features are categorized into phased milestones:
- **Phase 1: Core Productivity Utilities (Active Implementation Target)**
- **Phase 2: Advanced Data & Range Operations**
- **Phase 3: Workbook & Multi-Sheet Architecture**
- **Phase 4: Advanced Auditing, Reporting & Analytics**

---

## 2. Phase 1 — Core Utilities Catalog (Detailed Specifications)

### 2.1 Searchable Tool Launcher (SuperSearch)
- **ID:** `CMD_LAUNCHER`
- **Category:** Navigation & Discovery
- **Description:** Quick-access modal launcher providing instantaneous search, category browsing, and execution of all suite tools.
- **Inputs:** Keyboard search query string, category filter selector.
- **Behavior:**
  - Queries centralized `ICommandRegistry`.
  - Filters by Name, Keywords, and Category with fuzzy matching.
  - Displays keyboard shortcut, category pill, and brief description.
  - Tracks frequently used and favorited tools.
- **Edge Cases:** Empty query displays categorized tree; invalid search displays helpful suggestion.
- **Status:** **Implemented**

---

### 2.2 Text Extraction Before / After / Between Delimiters
- **ID:** `CMD_TEXT_EXTRACT`
- **Category:** Text Operations
- **Description:** Extracts substrings relative to specified delimiters (e.g., extract text before comma, between parentheses, or after hyphens).
- **Options:**
  - Mode: `Before Delimiter`, `After Delimiter`, `Between Delimiters`.
  - Delimiters: Custom string, Comma, Space, Hyphen, Slash, Semicolon, Tab.
  - Occurrence: First occurrence, Last occurrence, or N-th occurrence.
  - Case Sensitivity: Toggle for delimiter matching.
  - Overwrite original or write to adjacent offset column.
- **Data Safety:** Skips blank and numeric-only cells unless requested; preserves cell formulas or converts to text on request; live preview grid.
- **Performance:** Reads entire selected range into 2D array, extracts strings in managed memory, writes back in single COM call.
- **Status:** **Implemented**

---

### 2.3 Split Text into Columns or Rows
- **ID:** `CMD_TEXT_SPLIT`
- **Category:** Text Operations / Merge & Split
- **Description:** Splits delimited cell values into multiple adjacent columns or multiple new rows.
- **Options:**
  - Delimiters: Comma, Semicolon, Space, Newline (`\n`, `\r\n`), Custom text.
  - Split Direction: Across Columns or Down into Rows.
  - Blank Handling: Ignore consecutive delimiters or retain empty tokens.
  - Trimming: Automatically trim leading/trailing whitespace of tokens.
- **Data Safety:** Checks for existing data in target columns/rows; prompts user before overwriting non-empty cells.
- **Status:** **Implemented**

---

### 2.4 Remove Selected Characters
- **ID:** `CMD_TEXT_REMOVE_CHARS`
- **Category:** Text Operations
- **Description:** Removes characters from cells based on character class or custom specifications.
- **Options:**
  - Remove Numeric characters (`0-9`).
  - Remove Alphabetic characters (`a-z`, `A-Z`).
  - Remove Non-alphanumeric / Symbols (`!@#$%^&*...`).
  - Remove Non-printable characters (control codes `0x00`-`0x1F`, `0x7F`).
  - Remove Specific Custom Characters / String list.
- **Data Safety:** Live before-and-after preview of first 10 selected cells; options to preserve formulas.
- **Status:** **Implemented**

---

### 2.5 Trim and Normalize Whitespace
- **ID:** `CMD_TEXT_TRIM_SPACES`
- **Category:** Text Operations
- **Description:** Cleans messy text strings with irregular spacing.
- **Options:**
  - Leading and trailing spaces only (equivalent to standard clean trim).
  - Excess internal spaces (collapses multiple consecutive spaces into a single space).
  - All spaces (removes every space character).
  - Non-breaking spaces (converts ASCII `160` / Unicode `\u00A0` to standard space).
- **Data Safety:** In-memory batch array processing; live preview.
- **Status:** **Implemented**

---

### 2.6 Change Text Case
- **ID:** `CMD_TEXT_CHANGE_CASE`
- **Category:** Text Operations
- **Description:** Transforms capitalization of text cells across selected ranges.
- **Options:**
  - `UPPERCASE`
  - `lowercase`
  - `Proper Case` (Capitalize Each Word)
  - `Sentence case` (Capitalize first letter of each sentence)
  - `tOGGLE cASE` (Invert existing capitalization)
- **Data Safety:** Skips formulas and pure numbers; respects locale-specific casing rules; live preview.
- **Status:** **Implemented**

---

### 2.7 Add Prefix and Suffix
- **ID:** `CMD_TEXT_ADD_AFFIX`
- **Category:** Text Operations
- **Description:** Appends or prepends specified text to all cells in the selected range.
- **Options:**
  - Prefix text string.
  - Suffix text string.
  - Condition: Skip blank cells, skip cells that already contain the prefix/suffix.
- **Data Safety:** Batch 2D array processing; live preview.
- **Status:** **Implemented**

---

### 2.8 Insert Text at Specified Position
- **ID:** `CMD_TEXT_INSERT_AT_POS`
- **Category:** Text Operations
- **Description:** Inserts custom text at a specified zero- or 1-based index or relative to an anchor character.
- **Options:**
  - Position from beginning (e.g. at 3rd character).
  - Position from end (e.g. 2 characters before end).
  - Before / after first uppercase letter, first number, or specific character.
- **Data Safety:** Handles strings shorter than the target position gracefully; live preview.
- **Status:** **Implemented**

---

### 2.9 Find and Replace within Selected Ranges
- **ID:** `CMD_RANGE_FIND_REPLACE`
- **Category:** Data Cleaning / Range Operations
- **Description:** High-precision search and replace operating strictly on the user's selected range or active worksheet.
- **Options:**
  - Match case (case-sensitive).
  - Match entire cell content only.
  - Regular expression (Regex) search and replacement.
  - Search in Values or Formulas.
- **Status:** **Implemented**

---

### 2.10 Remove Duplicate Values
- **ID:** `CMD_DATA_REMOVE_DUPLICATES`
- **Category:** Data Cleaning
- **Description:** Identifies and removes duplicate rows or values in the selected range with granular criteria.
- **Options:**
  - Evaluate entire rows or designated key columns.
  - Keep First occurrence or Keep Last occurrence.
  - Action: Delete entire rows, Clear cell contents only, or Highlight duplicates.
- **Data Safety:** Pre-calculation check; confirmation prompt showing total duplicates detected prior to deletion.
- **Status:** **Implemented**

---

### 2.11 Fill Blank Cells
- **ID:** `CMD_RANGE_FILL_BLANKS`
- **Category:** Range & Cells
- **Description:** Fills empty cells within a range using intelligent fill strategies.
- **Options:**
  - Fill with Value from Above (Fill Down).
  - Fill with Value from Below (Fill Up).
  - Fill with Value from Left / Right.
  - Fill with Fixed Value (Constant string or number).
  - Fill with Linear Interpolation (Numeric average of bounding cells).
- **Status:** **Implemented**

---

### 2.12 Convert Formulas to Values
- **ID:** `CMD_FORMULA_TO_VALUE`
- **Category:** Formulas & Data Cleaning
- **Description:** Replaces cell formulas with their calculated results in-place across the selected range or entire worksheet.
- **Options:**
  - Selected range only.
  - Entire active sheet.
  - All sheets in workbook.
  - Handle error cells (`#N/A`, `#VALUE!`, `#REF!`): Option to retain errors, replace with empty string, or replace with custom placeholder.
- **Performance:** Direct array assignment (`range.Value2 = range.Value2`).
- **Status:** **Implemented**

---

### 2.13 Batch Cell Formatting
- **ID:** `CMD_BATCH_CELL_FORMAT`
- **Category:** Range & Cells / Formatting
- **Description:** Applies standardized formatting rules (number formatting, dates, text alignment, borders, number scales).
- **Options:**
  - Currency formatting with customizable symbol.
  - Date & Time standard ISO formatting (`YYYY-MM-DD`).
  - Percentage formatting with precision control.
  - Thousands separator (`#,##0.00`).
  - Clear all formats / Reset to General.
- **Status:** **Implemented**

---

### 2.14 Manage Worksheet Visibility
- **ID:** `CMD_SHEET_VISIBILITY`
- **Category:** Worksheet Management
- **Description:** Interactive manager to batch toggle worksheet visibility (Visible, Hidden, VeryHidden).
- **Options:**
  - Unhide all hidden sheets in one click.
  - Hide all inactive sheets (keep only current sheet visible).
  - Multi-select checkbox list of all sheets with status badges.
- **Data Safety:** Enforces Excel rule requiring at least one worksheet to remain visible.
- **Status:** **Implemented**

---

### 2.15 Rename Worksheets in Bulk
- **ID:** `CMD_SHEET_BATCH_RENAME`
- **Category:** Worksheet Management
- **Description:** Renames multiple worksheets simultaneously using flexible naming templates.
- **Options:**
  - Prefix / Suffix insertion.
  - Find and replace within sheet names.
  - Sequential numbering (e.g., `Report_01`, `Report_02`...).
  - Live preview list showing original and proposed new sheet names.
- **Data Safety:** Validates sheet names against Excel naming rules (max 31 characters, no invalid characters `\ / ? * : [ ]`, no duplicate names).
- **Status:** **Implemented**

---

### 2.16 Compare Ranges / Datasets (Phase 2)
- **ID:** `CMD_RANGE_COMPARE`
- **Category:** Range & Cells
- **Description:** Performs high-performance memory comparison between two worksheet ranges or datasets, identifying cell discrepancies, value mismatches, and missing entries.
- **Options:**
  - Mode: `Cell-by-cell (Side-by-side)`, `Values in Source Not in Target`, `Common / Matching Values`.
  - Criteria: `Match case`, `Trim whitespace`, `Ignore blank cells`.
- **Output:** Interactive discrepancy table displaying cell position, source value, target value, and discrepancy classification.
- **Status:** **Implemented (Phase 2)**

---

### 2.17 Select Special Cells (Phase 2)
- **ID:** `CMD_SELECT_SPECIAL`
- **Category:** Range & Cells
- **Description:** Locates, filters, and inspects cells matching specific attributes within the active selection.
- **Criteria:**
  - `Blank Cells`, `Non-Blank / Constant Cells`, `Error Cells (#N/A, #VALUE!, etc.)`.
  - `Formula Cells`, `Numeric Cells`, `Text / Label Cells`, `Maximum Numeric Value`, `Minimum Numeric Value`.
- **Interactivity:** Displays tabular list of matched cells with coordinates and values; allows one-click navigation/selection in Excel.
- **Status:** **Implemented (Phase 2)**

---

## 3. Phase 2 & 3 Catalog Overview (Roadmap)

| Feature | Category | Target Phase | Status |
| :--- | :--- | :--- | :--- |
| **Compare Ranges & Datasets** | Range & Cells | Phase 2 | **Implemented** |
| **Select Special Cells** | Range & Cells | Phase 2 | **Implemented** |
| **Merge / Combine Worksheets** | Consolidation | Phase 3 | Backlog |
| **Split Data by Key Column** | Worksheet | Phase 3 | Backlog |
| **Swap Two Ranges** | Range & Cells | Phase 3 | Backlog |
| **Transpose Table Dimensions** | Range & Cells | Phase 3 | Backlog |
| **Exact Formula Copy** | Formulas | Phase 3 | Backlog |
| **Convert Formula References (Abs/Rel)** | Formulas | Phase 3 | Backlog |
| **Table of Contents (Index Sheet)** | Worksheet | Phase 3 | Backlog |
| **Export Range to CSV / JSON / Image** | Import & Export | Phase 3 | Backlog |
| **Sort Worksheets in Bulk** | Worksheet | Phase 3 | Backlog |
