"""
Pandas Workbench — CSV automation GUI
────────────────────────────────────────
• Load multiple CSV files, each exposed as a named DataFrame
• Write pandas scripts in a syntax-highlighted editor with line numbers
• Inline autocomplete (type to trigger, Ctrl+Space, arrows/Tab/Enter to accept, Esc to dismiss)
• Snippets panel with descriptions — click to preview, double-click to insert
• Execute scripts; result DataFrame previewed in a table
• Save / load / manage reusable scripts
• Error output + execution time in status bar

Requirements:  sudo pacman -S python-pandas  (tkinter: pacman -S tk)
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog
import os
import re
import time
import traceback
import io
import sys
from contextlib import redirect_stdout
from typing import Any

try:
    import pandas as pd
    PANDAS_OK = True
except ImportError:
    pd = None  # type: ignore[assignment]
    PANDAS_OK = False

# ── Paths ──────────────────────────────────────────────────────────────────
APP_DIR     = os.path.join(os.path.expanduser("~"), ".pandas_workbench")
SCRIPTS_DIR = os.path.join(APP_DIR, "scripts")
os.makedirs(SCRIPTS_DIR, exist_ok=True)

# ── Colour palette ─────────────────────────────────────────────────────────
C = {
    "bg":       "#0d0f17",
    "panel":    "#13151f",
    "border":   "#1e2130",
    "surface":  "#1a1d2e",
    "surface2": "#222538",
    "text":     "#e2e4f0",
    "muted":    "#5a5f7a",
    "accent":   "#7c6af5",
    "accent2":  "#50e3c2",
    "err":      "#f5546a",
    "ok":       "#50e3a0",
    "kw":       "#c792ea",
    "str_col":  "#c3e88d",
    "num":      "#f78c6c",
    "comment":  "#546e7a",
    "builtin":  "#82aaff",
    "sel":      "#2d3155",
    "ac_bg":    "#1e2235",
    "ac_sel":   "#7c6af5",
}

_PLAT       = sys.platform
FONT_MONO   = ("JetBrains Mono", 11) if _PLAT != "darwin" else ("Menlo", 12)
FONT_MONO_SM= (FONT_MONO[0], 10)
FONT_UI     = ("Segoe UI", 10) if _PLAT == "win32" else ("SF Pro Display", 10) if _PLAT == "darwin" else ("Ubuntu", 10)
FONT_UI_B   = (FONT_UI[0], 10, "bold")
FONT_SMALL  = (FONT_UI[0], 9)

# ── Syntax highlighting ────────────────────────────────────────────────────
KEYWORDS = {
    "import","from","as","def","class","return","if","elif","else","for",
    "while","in","not","and","or","True","False","None","with","try",
    "except","finally","raise","pass","lambda","yield","del","global","is",
}
BUILTINS = {
    "print","len","range","list","dict","set","tuple","str","int","float",
    "bool","type","isinstance","enumerate","zip","map","filter","sorted",
    "pd","df","result","read_csv","concat","merge","groupby","agg","apply",
    "DataFrame","Series","to_csv","reset_index","fillna","dropna","rename",
    "drop","copy","sort_values","value_counts","pivot_table","melt","explode",
    "astype","strip","upper","lower","contains","startswith","endswith","np",
}

_KW_PAT  = re.compile(r'\b(' + '|'.join(re.escape(k) for k in KEYWORDS) + r')\b')
_BT_PAT  = re.compile(r'\b(' + '|'.join(re.escape(k) for k in BUILTINS) + r')\b')
_STR_PAT = re.compile(r'(\"\"\"[\s\S]*?\"\"\"|\'\'\'[\s\S]*?\'\'\'|"[^"\\]*(?:\\.[^"\\]*)*"|\'[^\'\\]*(?:\\.[^\'\\]*)*\')')
_NUM_PAT = re.compile(r'\b(\d+\.?\d*)\b')
_CMT_PAT = re.compile(r'(#.*)')


def highlight(tw: tk.Text) -> None:
    for tag in ("kw","bt","str","num","cmt"):
        tw.tag_remove(tag, "1.0", "end")
    lines = tw.get("1.0", "end-1c").split("\n")
    def mark(pat, tag):
        for ln, line in enumerate(lines, 1):
            for m in pat.finditer(line):
                tw.tag_add(tag, f"{ln}.{m.start()}", f"{ln}.{m.end()}")
    mark(_CMT_PAT, "cmt")
    mark(_STR_PAT, "str")
    mark(_NUM_PAT, "num")
    mark(_KW_PAT,  "kw")
    mark(_BT_PAT,  "bt")


def apply_tags(tw: tk.Text) -> None:
    tw.tag_config("kw",  foreground=C["kw"])
    tw.tag_config("bt",  foreground=C["builtin"])
    tw.tag_config("str", foreground=C["str_col"])
    tw.tag_config("num", foreground=C["num"])
    tw.tag_config("cmt", foreground=C["comment"])


# ── Snippets  (name, description, code) ───────────────────────────────────
SNIPPETS: list[tuple[str, str, str]] = [

    # ── Combining files ───────────────────────────────────────────────────
    (
        "Merge on key column",
        "JOIN two files on a shared ID column — like a SQL LEFT JOIN.\n"
        "Change 'id' to your shared column name.\n"
        "how= options: left, right, inner, outer",
        "result = pd.merge(df1, df2, on='id', how='left')\n"
    ),
    (
        "Merge on different column names",
        "JOIN when the key column has different names in each file.\n"
        "e.g. file1 has 'user_id', file2 has 'id'.",
        "result = pd.merge(df1, df2, left_on='user_id', right_on='id', how='left')\n"
    ),
    (
        "Stack rows (concat)",
        "Append rows from multiple files into one table.\n"
        "Files should have the same columns.\n"
        "ignore_index=True resets the row numbers.",
        "result = pd.concat([df1, df2], ignore_index=True)\n"
    ),
    (
        "Stack many files",
        "Combine more than 2 files vertically.\n"
        "Add as many DataFrames to the list as needed.",
        "result = pd.concat([df1, df2, df3], ignore_index=True)\n"
    ),

    # ── Filtering rows ────────────────────────────────────────────────────
    (
        "Filter: exact match",
        "Keep only rows where a column equals a specific value.\n"
        "Change 'column' and 'value' to match your data.",
        "result = df1[df1['column'] == 'value']\n"
    ),
    (
        "Filter: contains text",
        "Keep rows where a column contains a substring.\n"
        "na=False safely ignores blank cells.\n"
        "case=False makes it case-insensitive.",
        "result = df1[df1['column'].str.contains('text', na=False, case=False)]\n"
    ),
    (
        "Filter: number range",
        "Keep rows where a numeric column is within a range.",
        "result = df1[(df1['amount'] >= 100) & (df1['amount'] <= 500)]\n"
    ),
    (
        "Filter: multiple values (isin)",
        "Keep rows where a column matches any value in a list.\n"
        "Like SQL: WHERE status IN ('active', 'pending')",
        "result = df1[df1['status'].isin(['active', 'pending'])]\n"
    ),
    (
        "Filter: exclude values",
        "Exclude rows that match a value — opposite of isin.\n"
        "The ~ means NOT.",
        "result = df1[~df1['status'].isin(['deleted', 'archived'])]\n"
    ),
    (
        "Filter: non-empty rows",
        "Drop rows where a specific column is blank or null.",
        "result = df1[df1['column'].notna() & (df1['column'] != '')]\n"
    ),
    (
        "Filter: top N rows",
        "Keep only the first N rows (useful for previewing large files).",
        "result = df1.head(100)\n"
    ),
    (
        "Filter: random sample",
        "Take a random sample of N rows.\n"
        "Useful for testing on a subset of large data.",
        "result = df1.sample(n=100, random_state=42)\n"
    ),

    # ── Cleaning data ─────────────────────────────────────────────────────
    (
        "Drop duplicate rows",
        "Remove completely identical rows.\n"
        "To deduplicate on specific columns, use subset=['col1','col2'].",
        "result = df1.drop_duplicates()\n"
        "# result = df1.drop_duplicates(subset=['email'])  # dedupe by email\n"
    ),
    (
        "Fill missing values",
        "Replace blank/NaN cells with a default.\n"
        "You can specify a different default per column.",
        "result = df1.fillna({'column': 'unknown', 'amount': 0})\n"
    ),
    (
        "Drop rows with any blanks",
        "Remove any row that has at least one empty cell.",
        "result = df1.dropna()\n"
    ),
    (
        "Drop rows where column is blank",
        "Remove rows only when a specific column is empty.",
        "result = df1.dropna(subset=['column'])\n"
    ),
    (
        "Strip whitespace from all text",
        "Trim leading/trailing spaces from every string cell.\n"
        "Useful when CSVs were exported from spreadsheets.",
        "result = df1.apply(lambda col: col.str.strip() if col.dtype == 'object' else col)\n"
    ),
    (
        "Standardise text case",
        "Convert a text column to lowercase (or upper/title).",
        "result = df1.copy()\n"
        "result['column'] = result['column'].str.lower()\n"
        "# .str.upper()  or  .str.title()  also available\n"
    ),
    (
        "Fix column data types",
        "Convert columns to the correct type after loading.\n"
        "Useful when numbers or dates were loaded as text.\n"
        "errors='coerce' turns unparseable values into NaN.",
        "result = df1.copy()\n"
        "result['amount'] = pd.to_numeric(result['amount'], errors='coerce')\n"
        "result['date']   = pd.to_datetime(result['date'],  errors='coerce')\n"
    ),
    (
        "Replace values in column",
        "Swap specific values in a column — like find-and-replace.\n"
        "Add as many old→new pairs as needed.",
        "result = df1.copy()\n"
        "result['column'] = result['column'].replace({\n"
        "    'old_value': 'new_value',\n"
        "    'Y': 'Yes',\n"
        "    'N': 'No',\n"
        "})\n"
    ),
    (
        "Remove special characters",
        "Strip non-alphanumeric characters from a text column.\n"
        "Useful for cleaning phone numbers, codes etc.",
        "result = df1.copy()\n"
        "result['column'] = result['column'].str.replace(r'[^\\w\\s]', '', regex=True)\n"
    ),

    # ── Reshaping columns ─────────────────────────────────────────────────
    (
        "Select columns",
        "Keep only the columns you want in the output.\n"
        "Any column not listed is dropped.",
        "result = df1[['col1', 'col2', 'col3']]\n"
    ),
    (
        "Drop columns",
        "Remove specific columns you don't need.",
        "result = df1.drop(columns=['unwanted_col1', 'unwanted_col2'])\n"
    ),
    (
        "Rename columns",
        "Give columns new names. Add as many pairs as needed.",
        "result = df1.rename(columns={\n"
        "    'old_name_1': 'new_name_1',\n"
        "    'old_name_2': 'new_name_2',\n"
        "})\n"
    ),
    (
        "Reorder columns",
        "Set a specific column order in the output.\n"
        "Any columns not listed are dropped.",
        "result = df1[['col_a', 'col_b', 'col_c']]\n"
    ),
    (
        "Add calculated column",
        "Create a new column derived from existing ones.\n"
        "Supports math, string joins, conditionals etc.",
        "result = df1.copy()\n"
        "result['full_name'] = result['first_name'] + ' ' + result['last_name']\n"
        "result['total']     = result['qty'] * result['unit_price']\n"
    ),
    (
        "Conditional column (if/else)",
        "Set a column value based on a condition.\n"
        "Like Excel IF() — needs numpy (usually installed with pandas).",
        "import numpy as np\n"
        "result = df1.copy()\n"
        "result['label'] = np.where(result['score'] >= 50, 'pass', 'fail')\n"
    ),
    (
        "Multiple conditions (nested if)",
        "Assign different labels based on multiple conditions.\n"
        "Like nested IF() in Excel.",
        "import numpy as np\n"
        "result = df1.copy()\n"
        "result['grade'] = np.select(\n"
        "    condlist=[\n"
        "        result['score'] >= 90,\n"
        "        result['score'] >= 70,\n"
        "        result['score'] >= 50,\n"
        "    ],\n"
        "    choicelist=['A', 'B', 'C'],\n"
        "    default='F'\n"
        ")\n"
    ),
    (
        "Split column by delimiter",
        "Split one column into two or more.\n"
        "e.g. 'John Smith' → first and last name columns.",
        "result = df1.copy()\n"
        "result[['first', 'last']] = result['full_name'].str.split(' ', n=1, expand=True)\n"
    ),
    (
        "Extract substring / regex",
        "Pull a part of a string using a pattern.\n"
        "e.g. extract a year from a mixed text column.",
        "result = df1.copy()\n"
        "result['year'] = result['description'].str.extract(r'(\\d{4})')\n"
    ),

    # ── Sorting & ranking ─────────────────────────────────────────────────
    (
        "Sort by column",
        "Sort rows by one column, ascending or descending.",
        "result = df1.sort_values('column', ascending=True)\n"
        "# ascending=False  →  largest/latest first\n"
    ),
    (
        "Sort by multiple columns",
        "Sort by a primary column, then a secondary one.\n"
        "Each can have its own direction.",
        "result = df1.sort_values(['dept', 'salary'], ascending=[True, False])\n"
    ),
    (
        "Add rank column",
        "Add a numeric rank column based on a score.\n"
        "method='dense' means no gaps in rank numbers.",
        "result = df1.copy()\n"
        "result['rank'] = result['score'].rank(ascending=False, method='dense').astype(int)\n"
        "result = result.sort_values('rank')\n"
    ),
    (
        "Top N per group",
        "Keep the top N rows within each group.\n"
        "e.g. top 3 sales per region.",
        "result = (\n"
        "    df1.sort_values('amount', ascending=False)\n"
        "       .groupby('category')\n"
        "       .head(3)\n"
        "       .reset_index(drop=True)\n"
        ")\n"
    ),

    # ── Aggregation & summarising ─────────────────────────────────────────
    (
        "Group & sum",
        "Group rows by a category and sum a numeric column.\n"
        "Like a pivot table in Excel.",
        "result = df1.groupby('category').agg({'amount': 'sum'}).reset_index()\n"
    ),
    (
        "Group & multiple aggregations",
        "Compute count, sum, mean etc. for different columns at once.",
        "result = df1.groupby('category').agg(\n"
        "    total   = ('amount', 'sum'),\n"
        "    count   = ('amount', 'count'),\n"
        "    average = ('amount', 'mean'),\n"
        "    maximum = ('amount', 'max'),\n"
        ").reset_index()\n"
    ),
    (
        "Value counts",
        "Count how many times each value appears in a column.\n"
        "Great for checking data quality or distributions.",
        "result = df1['column'].value_counts().reset_index()\n"
        "result.columns = ['value', 'count']\n"
    ),
    (
        "Pivot table",
        "Reshape data: rows become one axis, columns another.\n"
        "aggfunc can be sum, mean, count, max, min.",
        "result = df1.pivot_table(\n"
        "    index='row_col',\n"
        "    columns='col_col',\n"
        "    values='value_col',\n"
        "    aggfunc='sum',\n"
        "    fill_value=0,\n"
        ").reset_index()\n"
    ),
    (
        "Unpivot / melt (wide → long)",
        "Reverse a pivot: turn column headers into row values.\n"
        "Useful when months or dates are spread across columns.",
        "result = df1.melt(\n"
        "    id_vars=['id', 'name'],   # columns to keep as-is\n"
        "    var_name='month',         # new column for old headers\n"
        "    value_name='amount',      # new column for values\n"
        ")\n"
    ),
    (
        "Running total (cumulative sum)",
        "Add a column showing the cumulative sum over rows.\n"
        "Sort by date first for a meaningful running total.",
        "result = df1.sort_values('date').copy()\n"
        "result['running_total'] = result['amount'].cumsum()\n"
    ),
    (
        "Percentage of total",
        "Add a column showing each row's % share of the total.",
        "result = df1.copy()\n"
        "result['pct'] = (result['amount'] / result['amount'].sum() * 100).round(2)\n"
    ),
    (
        "Percentage within group",
        "Each row's share of its group total (not overall total).",
        "result = df1.copy()\n"
        "result['group_total'] = result.groupby('category')['amount'].transform('sum')\n"
        "result['pct_of_group'] = (result['amount'] / result['group_total'] * 100).round(2)\n"
    ),
    (
        "Cross-tabulation",
        "Count combinations of two categorical columns.\n"
        "Like a frequency matrix.",
        "result = pd.crosstab(df1['col_a'], df1['col_b']).reset_index()\n"
    ),

    # ── Date & time ───────────────────────────────────────────────────────
    (
        "Parse dates & extract parts",
        "Convert a text date column and extract year/month/day.\n"
        "errors='coerce' turns unparseable dates into NaN.",
        "result = df1.copy()\n"
        "result['date']  = pd.to_datetime(result['date_col'], errors='coerce')\n"
        "result['year']  = result['date'].dt.year\n"
        "result['month'] = result['date'].dt.month\n"
        "result['day']   = result['date'].dt.day\n"
    ),
    (
        "Filter by date range",
        "Keep only rows within a specific date range.",
        "result = df1.copy()\n"
        "result['date'] = pd.to_datetime(result['date_col'], errors='coerce')\n"
        "result = result[\n"
        "    (result['date'] >= '2024-01-01') &\n"
        "    (result['date'] <= '2024-12-31')\n"
        "]\n"
    ),
    (
        "Days between two dates",
        "Calculate the number of days between two date columns.",
        "result = df1.copy()\n"
        "result['start']     = pd.to_datetime(result['start_date'])\n"
        "result['end']       = pd.to_datetime(result['end_date'])\n"
        "result['days_diff'] = (result['end'] - result['start']).dt.days\n"
    ),
    (
        "Group by month/year",
        "Aggregate data by month and year.\n"
        "Useful for monthly trend reports.",
        "result = df1.copy()\n"
        "result['date']       = pd.to_datetime(result['date_col'])\n"
        "result['year_month'] = result['date'].dt.to_period('M').astype(str)\n"
        "result = result.groupby('year_month').agg({'amount': 'sum'}).reset_index()\n"
    ),

    # ── Lookup & matching ─────────────────────────────────────────────────
    (
        "VLOOKUP (add column from lookup table)",
        "Like VLOOKUP in Excel — attach extra info from a second file\n"
        "based on a key column.",
        "# df2 is the lookup table (e.g. product_id → product_name)\n"
        "result = df1.merge(df2[['id', 'label']], on='id', how='left')\n"
    ),
    (
        "Flag rows present in both files",
        "Mark which rows in file1 also appear in file2.\n"
        "Adds a True/False column.",
        "result = df1.copy()\n"
        "result['in_file2'] = result['id'].isin(df2['id'])\n"
    ),
    (
        "Anti-join (rows only in file1)",
        "Find records in file1 with NO match in file2.\n"
        "Useful for spotting missing or unmatched records.",
        "merged = df1.merge(df2[['id']], on='id', how='left', indicator=True)\n"
        "result = merged[merged['_merge'] == 'left_only'].drop(columns=['_merge'])\n"
    ),
    (
        "Inner join (rows in both files)",
        "Keep only rows that exist in BOTH files (matched records only).",
        "result = pd.merge(df1, df2, on='id', how='inner')\n"
    ),
    (
        "Full outer join (all rows)",
        "Keep ALL rows from both files.\n"
        "Unmatched cells will be NaN.",
        "result = pd.merge(df1, df2, on='id', how='outer')\n"
    ),
    (
        "Fuzzy match (approximate names)",
        "Find the closest match in file2 for each row in file1.\n"
        "Useful when names/strings aren't spelled exactly the same.\n"
        "Requires: pip install thefuzz  or  pacman -S python-thefuzz",
        "from thefuzz import process\n\n"
        "def best_match(val, choices, threshold=80):\n"
        "    match, score = process.extractOne(str(val), choices)\n"
        "    return match if score >= threshold else None\n\n"
        "choices = df2['name'].tolist()\n"
        "result  = df1.copy()\n"
        "result['matched_name'] = result['name'].apply(lambda v: best_match(v, choices))\n"
    ),

    # ── Output formatting ─────────────────────────────────────────────────
    (
        "Round numeric columns",
        "Round all float columns to 2 decimal places.\n"
        "Or target one specific column.",
        "result = df1.round(2)\n"
        "# result['amount'] = result['amount'].round(2)  # single column\n"
    ),
    (
        "Remove suffix columns after merge",
        "After merging, drop the redundant _x/_y suffix columns.",
        "result = df1.merge(df2, on='id', how='left', suffixes=('', '_drop'))\n"
        "result = result[[c for c in result.columns if not c.endswith('_drop')]]\n"
    ),
    (
        "Flatten pivot column names",
        "After pivot_table, column names can become tuples — flatten them.",
        "result = df1.pivot_table(index='a', columns='b', values='c', aggfunc='sum')\n"
        "result.columns = ['_'.join(str(s) for s in col).strip() for col in result.columns]\n"
        "result = result.reset_index()\n"
    ),
    (
        "Reindex / ensure all columns exist",
        "Guarantee a fixed set of columns exist, adding blanks for missing ones.\n"
        "Useful when stacking files with slightly different schemas.",
        "expected_cols = ['id', 'name', 'email', 'amount', 'status']\n"
        "result = df1.reindex(columns=expected_cols)\n"
    ),
    (
        "Export multiple sheets to Excel",
        "Write multiple DataFrames to separate sheets in one Excel file.\n"
        "Requires: pip install openpyxl  or  pacman -S python-openpyxl",
        "with pd.ExcelWriter('output.xlsx', engine='openpyxl') as writer:\n"
        "    df1.to_excel(writer, sheet_name='Sheet1', index=False)\n"
        "    df2.to_excel(writer, sheet_name='Sheet2', index=False)\n"
        "result = df1  # set result so the preview still works\n"
    ),
]


# ══════════════════════════════════════════════════════════════════════════
#  Main application
# ══════════════════════════════════════════════════════════════════════════

class Workbench(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Pandas Workbench")
        self.geometry("1440x860")
        self.minsize(1100, 700)
        self.configure(bg=C["bg"])

        self.loaded_files: dict[str, str] = {}
        self.current_script_path: str | None = None
        self._hl_after: str | None = None
        self._result_df: Any = None  # holds pd.DataFrame when set

        # Inline autocomplete state
        self._ac_active   = False
        self._ac_col_mode = False
        self._ac_matches: list[str] = []
        self._ac_index   = 0
        self._ac_frame:   tk.Frame   | None = None
        self._ac_listbox: tk.Listbox | None = None

        self._build_styles()
        self._build_layout()

        if not PANDAS_OK:
            messagebox.showerror("pandas not found",
                "Install pandas:\n\n  sudo pacman -S python-pandas")

    # ── Styles ─────────────────────────────────────────────────────────────

    def _build_styles(self):
        s = ttk.Style(self)
        s.theme_use("clam")
        s.configure(".",              background=C["bg"],      foreground=C["text"],  font=FONT_UI)
        s.configure("TFrame",         background=C["bg"])
        s.configure("TLabel",         background=C["bg"],      foreground=C["text"])
        s.configure("TScrollbar",     background=C["surface"], troughcolor=C["bg"],
                    bordercolor=C["bg"], arrowcolor=C["muted"])
        s.configure("TButton",        background=C["surface2"],foreground=C["text"],
                    borderwidth=0, focusthickness=0, padding=(10, 6))
        s.map("TButton",              background=[("active", C["accent"])])
        s.configure("Accent.TButton", background=C["accent"],  foreground="#fff",
                    borderwidth=0, padding=(12, 7))
        s.map("Accent.TButton",       background=[("active", "#9580ff")])
        s.configure("Danger.TButton", background="#3a1a22",    foreground=C["err"],
                    borderwidth=0, padding=(8, 5))
        s.map("Danger.TButton",       background=[("active", "#5a2030")])
        s.configure("Treeview",       background=C["surface"], foreground=C["text"],
                    fieldbackground=C["surface"], rowheight=26, borderwidth=0, font=FONT_MONO_SM)
        s.configure("Treeview.Heading", background=C["surface2"], foreground=C["accent2"],
                    font=FONT_UI_B, relief="flat")
        s.map("Treeview",             background=[("selected", C["sel"])],
                                      foreground=[("selected", C["text"])])
        s.configure("TNotebook",      background=C["bg"], borderwidth=0, tabmargins=0)
        s.configure("TNotebook.Tab",  background=C["surface"], foreground=C["muted"],
                    padding=(14, 6), borderwidth=0)
        s.map("TNotebook.Tab",        background=[("selected", C["panel"])],
                                      foreground=[("selected", C["text"])])

    # ── Layout ─────────────────────────────────────────────────────────────

    def _build_layout(self):
        topbar = tk.Frame(self, bg=C["panel"], height=48)
        topbar.pack(fill="x", side="top")
        topbar.pack_propagate(False)
        tk.Label(topbar, text="⬡  PANDAS WORKBENCH", bg=C["panel"],
                 fg=C["accent"], font=(FONT_UI[0], 12, "bold"), padx=20).pack(side="left", pady=10)
        tk.Label(topbar, text="CSV automation with live scripting",
                 bg=C["panel"], fg=C["muted"], font=FONT_SMALL).pack(side="left")

        main = tk.PanedWindow(self, orient="horizontal", bg=C["border"],
                              sashwidth=4, sashrelief="flat", bd=0)
        main.pack(fill="both", expand=True)
        main.add(self._build_left(main),   minsize=220, width=260)
        main.add(self._build_center(main), minsize=400, width=680)
        main.add(self._build_right(main),  minsize=300, width=480)

        self.statusbar = tk.Label(self, text="Ready.", bg=C["surface"],
                                  fg=C["muted"], font=FONT_SMALL,
                                  anchor="w", padx=16, pady=5)
        self.statusbar.pack(fill="x", side="bottom")

    # ── Left panel ─────────────────────────────────────────────────────────

    def _build_left(self, parent):
        f = tk.Frame(parent, bg=C["panel"])

        self._slabel(f, "FILES").pack(fill="x", padx=14, pady=(14, 4))
        br = tk.Frame(f, bg=C["panel"])
        br.pack(fill="x", padx=10, pady=(0, 8))
        ttk.Button(br, text="+ Load CSV", style="Accent.TButton",
                   command=self._load_csv).pack(side="left", padx=(0, 6))
        ttk.Button(br, text="✕ Remove",   style="Danger.TButton",
                   command=self._remove_file).pack(side="left")

        lf = tk.Frame(f, bg=C["surface"])
        lf.pack(fill="both", expand=True, padx=10, pady=(0, 10))
        self.file_listbox = tk.Listbox(lf, bg=C["surface"], fg=C["text"],
                                        selectbackground=C["sel"], selectforeground=C["text"],
                                        font=FONT_MONO_SM, borderwidth=0, highlightthickness=0,
                                        activestyle="none", relief="flat")
        sb = ttk.Scrollbar(lf, command=self.file_listbox.yview)
        self.file_listbox.configure(yscrollcommand=sb.set)
        self.file_listbox.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")
        self.file_listbox.bind("<Double-1>",        self._rename_file)
        self.file_listbox.bind("<<ListboxSelect>>", self._on_file_select)

        self._slabel(f, "FILE INFO").pack(fill="x", padx=14, pady=(6, 4))
        self.file_info = tk.Label(f, text="Select a file to inspect.",
                                   bg=C["panel"], fg=C["muted"], font=FONT_SMALL,
                                   justify="left", wraplength=230, anchor="nw", padx=14)
        self.file_info.pack(fill="x", pady=(0, 10))

        self._slabel(f, "SNIPPETS").pack(fill="x", padx=14, pady=(6, 2))
        tk.Label(f, text="Click = description  |  Double-click = insert",
                 bg=C["panel"], fg=C["muted"], font=(FONT_UI[0], 8)).pack(padx=14, anchor="w")

        # Search box
        search_frame = tk.Frame(f, bg=C["surface2"],
                                highlightthickness=1, highlightbackground=C["border"])
        search_frame.pack(fill="x", padx=10, pady=(4, 0))
        tk.Label(search_frame, text="⌕", bg=C["surface2"], fg=C["muted"],
                 font=(FONT_UI[0], 10)).pack(side="left", padx=(6, 2))
        self.snip_search_var = tk.StringVar()
        self.snip_search_var.trace_add("write", self._filter_snippets)
        tk.Entry(search_frame, textvariable=self.snip_search_var,
                 bg=C["surface2"], fg=C["text"], insertbackground=C["text"],
                 relief="flat", font=FONT_SMALL, highlightthickness=0,
                 ).pack(side="left", fill="x", expand=True, ipady=4, padx=(0, 6))

        snf = tk.Frame(f, bg=C["surface"])
        snf.pack(fill="both", expand=True, padx=10, pady=(2, 4))
        self.snip_list = tk.Listbox(snf, bg=C["surface"], fg=C["accent2"],
                                     selectbackground=C["sel"], selectforeground=C["text"],
                                     font=FONT_SMALL, borderwidth=0, highlightthickness=0,
                                     activestyle="none", relief="flat")
        ssb = ttk.Scrollbar(snf, command=self.snip_list.yview)
        self.snip_list.configure(yscrollcommand=ssb.set)
        # _snip_index_map maps listbox position → SNIPPETS index
        self._snip_index_map: list[int] = list(range(len(SNIPPETS)))
        for name, _, _ in SNIPPETS:
            self.snip_list.insert("end", f"  {name}")
        self.snip_list.pack(side="left", fill="both", expand=True)
        ssb.pack(side="right", fill="y")
        self.snip_list.bind("<Double-1>",        self._insert_snippet)
        self.snip_list.bind("<<ListboxSelect>>", self._show_snippet_desc)

        self._slabel(f, "SNIPPET INFO").pack(fill="x", padx=14, pady=(4, 2))
        self.snip_desc = tk.Label(f, text="Select a snippet to see what it does.",
                                   bg=C["panel"], fg=C["muted"], font=(FONT_UI[0], 8),
                                   justify="left", wraplength=230, anchor="nw", padx=14)
        self.snip_desc.pack(fill="x", pady=(0, 10))
        return f

    # ── Center panel ───────────────────────────────────────────────────────

    def _build_center(self, parent):
        f = tk.Frame(parent, bg=C["bg"])

        # Toolbar
        tb = tk.Frame(f, bg=C["panel"], height=42)
        tb.pack(fill="x")
        tb.pack_propagate(False)
        ttk.Button(tb, text="▶  Run (Ctrl+Enter)", style="Accent.TButton",
                   command=self._run_script).pack(side="left", padx=(10, 4), pady=6)
        ttk.Button(tb, text="💾 Save", command=self._save_script).pack(side="left", padx=4, pady=6)
        ttk.Button(tb, text="📂 Open", command=self._open_script).pack(side="left", padx=4, pady=6)
        ttk.Button(tb, text="🗑 Clear", command=self._clear_editor).pack(side="left", padx=4, pady=6)
        self.script_name_var = tk.StringVar(value="untitled")
        tk.Label(tb, text="Name:", bg=C["panel"], fg=C["muted"],
                 font=FONT_SMALL).pack(side="left", padx=(16, 4))
        tk.Entry(tb, textvariable=self.script_name_var,
                 bg=C["surface2"], fg=C["text"], insertbackground=C["text"],
                 relief="flat", font=FONT_SMALL, width=18,
                 highlightthickness=1, highlightcolor=C["accent"],
                 highlightbackground=C["border"]).pack(side="left", ipady=4)

        # Editor area — container uses place() so we can overlay the AC popup
        self._ec = tk.Frame(f, bg=C["bg"])
        self._ec.pack(fill="both", expand=True)

        ef = tk.Frame(self._ec, bg=C["bg"])
        ef.place(relx=0, rely=0, relwidth=1, relheight=1)

        self.line_nums = tk.Text(ef, width=4, bg=C["surface"], fg=C["muted"],
                                  font=FONT_MONO, state="disabled", relief="flat",
                                  bd=0, highlightthickness=0, padx=6)
        self.line_nums.pack(side="left", fill="y")

        evs = ttk.Scrollbar(ef, orient="vertical")
        ehs = ttk.Scrollbar(ef, orient="horizontal")

        self.editor = tk.Text(
            ef, bg=C["bg"], fg=C["text"],
            insertbackground=C["accent2"], selectbackground=C["sel"],
            font=FONT_MONO, relief="flat", bd=0, highlightthickness=0,
            undo=True, wrap="none", tabs="4c",
            yscrollcommand=self._sync_scroll,
            xscrollcommand=ehs.set,
            padx=12, pady=8, spacing1=2, spacing3=2,
        )
        evs.configure(command=self._on_editor_scroll)
        ehs.configure(command=self.editor.xview)
        evs.pack(side="right", fill="y")
        ehs.pack(side="bottom", fill="x")
        self.editor.pack(side="left", fill="both", expand=True)
        apply_tags(self.editor)

        # Bindings
        self.editor.bind("<KeyRelease>",     self._on_key_release)
        self.editor.bind("<Tab>",            self._on_tab)
        self.editor.bind("<Return>",         self._on_return)
        self.editor.bind("<Escape>",         lambda _: self._hide_ac())
        self.editor.bind("<Up>",             self._on_up)
        self.editor.bind("<Down>",           self._on_down)
        self.editor.bind("<Control-Return>", lambda _: self._run_script())
        self.editor.bind("<Control-s>",      lambda _: self._save_script())
        self.editor.bind("<Control-z>",      lambda _: self.editor.edit_undo())
        self.editor.bind("<Control-space>",  self._trigger_ac)
        self.editor.bind("<Button-1>",       lambda _: self._hide_ac())

        self.editor.insert("1.0",
            "# ── IDE IMPORTS (auto-updated as you load files) ──────────────\n"
            "# import pandas as pd\n"
            "# ───────────────────────────────────────────────────────────────\n\n"
            "# Assign your final output to: result\n\n"
            "result = df1\n"
        )
        self._schedule_highlight()
        self._update_line_numbers()

        # Output box
        of = tk.Frame(f, bg=C["panel"], height=110)
        of.pack(fill="x")
        of.pack_propagate(False)
        tk.Label(of, text="OUTPUT", bg=C["panel"], fg=C["muted"],
                 font=(FONT_UI[0], 8, "bold"), padx=12).pack(anchor="w", pady=(6, 0))
        self.output_box = tk.Text(of, bg=C["panel"], fg=C["ok"],
                                   font=FONT_MONO_SM, relief="flat", bd=0,
                                   highlightthickness=0, state="disabled", height=4, padx=12)
        self.output_box.pack(fill="both", expand=True)
        self.output_box.tag_config("err",  foreground=C["err"])
        self.output_box.tag_config("ok",   foreground=C["ok"])
        self.output_box.tag_config("info", foreground=C["accent2"])
        return f

    # ── Right panel ─────────────────────────────────────────────────────────

    def _build_right(self, parent):
        f = tk.Frame(parent, bg=C["panel"])
        nb = ttk.Notebook(f)
        nb.pack(fill="both", expand=True)

        # Result tab
        rt = tk.Frame(nb, bg=C["panel"])
        nb.add(rt, text="  Result  ")
        self.result_info = tk.Label(rt, text="No result yet.",
                                     bg=C["panel"], fg=C["muted"],
                                     font=FONT_SMALL, anchor="w", padx=12, pady=6)
        self.result_info.pack(fill="x")
        tf = tk.Frame(rt, bg=C["panel"])
        tf.pack(fill="both", expand=True, padx=8, pady=(0, 8))
        self.result_tree = ttk.Treeview(tf, show="headings", selectmode="browse")
        vsb = ttk.Scrollbar(tf, orient="vertical",   command=self.result_tree.yview)
        hsb = ttk.Scrollbar(tf, orient="horizontal", command=self.result_tree.xview)
        self.result_tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        self.result_tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        tf.rowconfigure(0, weight=1)
        tf.columnconfigure(0, weight=1)
        ttk.Button(rt, text="💾  Export Result as CSV", style="Accent.TButton",
                   command=self._export_result).pack(side="left", padx=8, pady=(0, 8))

        # Saved scripts tab
        st = tk.Frame(nb, bg=C["panel"])
        nb.add(st, text="  Saved Scripts  ")
        self._slabel(st, "SAVED SCRIPTS").pack(fill="x", padx=14, pady=(12, 6))
        sbr = tk.Frame(st, bg=C["panel"])
        sbr.pack(fill="x", padx=10, pady=(0, 8))
        ttk.Button(sbr, text="🔄 Refresh", command=self._refresh_scripts).pack(side="left", padx=(0, 6))
        ttk.Button(sbr, text="🗑 Delete",  style="Danger.TButton",
                   command=self._delete_script).pack(side="left")
        scf = tk.Frame(st, bg=C["surface"])
        scf.pack(fill="both", expand=True, padx=10, pady=(0, 10))
        self.scripts_list = tk.Listbox(scf, bg=C["surface"], fg=C["text"],
                                        selectbackground=C["sel"], selectforeground=C["text"],
                                        font=FONT_MONO_SM, borderwidth=0, highlightthickness=0,
                                        activestyle="none", relief="flat")
        scsb = ttk.Scrollbar(scf, command=self.scripts_list.yview)
        self.scripts_list.configure(yscrollcommand=scsb.set)
        self.scripts_list.pack(side="left", fill="both", expand=True)
        scsb.pack(side="right", fill="y")
        self.scripts_list.bind("<Double-1>", self._load_saved_script)
        tk.Label(st, text="Double-click to load into editor",
                 bg=C["panel"], fg=C["muted"], font=(FONT_UI[0], 8)).pack(padx=14, anchor="w")
        self._refresh_scripts()
        return f

    # ── Small helpers ───────────────────────────────────────────────────────

    def _slabel(self, parent, text):
        row = tk.Frame(parent, bg=C["panel"])
        tk.Label(row, text=text, bg=C["panel"], fg=C["accent"],
                 font=(FONT_UI[0], 8, "bold")).pack(side="left")
        return row

    def _set_status(self, msg, color=None):
        self.statusbar.configure(text=msg, fg=color or C["muted"])

    def _write_output(self, text, tag="ok"):
        self.output_box.configure(state="normal")
        self.output_box.delete("1.0", "end")
        self.output_box.insert("end", text, tag)
        self.output_box.configure(state="disabled")

    def _update_ide_header(self) -> None:
        """Rewrite the comment block at the top of the editor with current file imports."""
        MARKER_START = "# ── IDE IMPORTS"
        MARKER_END   = "# ───────────────────────────────────────────────────────────────"

        content = self.editor.get("1.0", "end-1c")

        # Find the existing header block and remove it
        start_idx = content.find(MARKER_START)
        end_idx   = content.find(MARKER_END)
        if start_idx != -1 and end_idx != -1:
            # end_idx points to start of last marker line — skip to end of that line
            end_of_block = end_idx + len(MARKER_END)
            # also consume the trailing newline if present
            if end_of_block < len(content) and content[end_of_block] == "\n":
                end_of_block += 1
            self.editor.delete("1.0", f"1.0 + {end_of_block}c")

        # Build new header
        lines = ["# ── IDE IMPORTS (auto-updated as you load files) ──────────────"]
        lines.append("# import pandas as pd")
        for name, path in self.loaded_files.items():
            lines.append(f"# {name} = pd.read_csv(r\"{path}\")")
        lines.append("# ───────────────────────────────────────────────────────────────")
        header = "\n".join(lines) + "\n"

        self.editor.insert("1.0", header)
        self._schedule_highlight()
        self._update_line_numbers()

    # ── File management ────────────────────────────────────────────────────

    def _load_csv(self):
        paths = filedialog.askopenfilenames(
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")])
        if not paths:
            return
        for path in paths:
            base = os.path.splitext(os.path.basename(path))[0]
            name = re.sub(r'\W+', '_', base).lower()
            orig = name; i = 2
            while name in self.loaded_files:
                name = f"{orig}_{i}"; i += 1
            self.loaded_files[name] = path
            self.file_listbox.insert("end", f"  {name}  ←  {os.path.basename(path)}")
        self._update_ide_header()
        self._set_status(f"{len(paths)} file(s) loaded.", C["ok"])

    def _remove_file(self):
        sel = self.file_listbox.curselection()
        if not sel:
            return
        idx  = sel[0]
        name = self.file_listbox.get(idx).strip().split("  ←  ")[0].strip()
        del self.loaded_files[name]
        self.file_listbox.delete(idx)
        self.file_info.configure(text="")
        self._update_ide_header()

    def _rename_file(self, _=None):
        sel = self.file_listbox.curselection()
        if not sel:
            return
        idx  = sel[0]
        old  = self.file_listbox.get(idx).strip().split("  ←  ")[0].strip()
        path = self.loaded_files[old]
        new  = simpledialog.askstring("Rename", f"New variable name for '{old}':",
                                      initialvalue=old, parent=self)
        if not new:
            return
        new = re.sub(r'\W+', '_', new).lower()
        if new in self.loaded_files and new != old:
            messagebox.showerror("Name taken", f"'{new}' is already used."); return
        del self.loaded_files[old]
        self.loaded_files[new] = path
        self.file_listbox.delete(idx)
        self.file_listbox.insert(idx, f"  {new}  ←  {os.path.basename(path)}")
        self.file_listbox.selection_set(idx)
        self._update_ide_header()

    def _on_file_select(self, _=None):
        sel = self.file_listbox.curselection()
        if not sel or not PANDAS_OK:
            return
        assert pd is not None
        name = self.file_listbox.get(sel[0]).strip().split("  ←  ")[0].strip()
        path = self.loaded_files.get(name)
        if not path:
            return
        try:
            df = pd.read_csv(path, nrows=1)
            n_rows = sum(1 for _ in open(path)) - 1
            cols_preview = ", ".join(df.columns[:6])
            if len(df.columns) > 6:
                cols_preview += " …"
            self.file_info.configure(
                text=f"var: {name}\nrows: {n_rows:,}\ncols: {len(df.columns)}\n{cols_preview}",
                fg=C["text"])
        except Exception as e:
            self.file_info.configure(text=str(e), fg=C["err"])

    # ── Editor ─────────────────────────────────────────────────────────────

    def _on_key_release(self, event=None):
        self._update_line_numbers()
        self._schedule_highlight()
        if event and event.keysym not in {
            "Up","Down","Left","Right","Return","Tab","Escape",
            "Home","End","Prior","Next",
            "Control_L","Control_R","Alt_L","Alt_R","Shift_L","Shift_R",
            "F1","F2","F3","F4","F5","F6","F7","F8","F9","F10","F11","F12",
        }:
            self._trigger_ac()

    def _schedule_highlight(self):
        if self._hl_after:
            self.after_cancel(self._hl_after)
        self._hl_after = self.after(250, lambda: (
            highlight(self.editor), setattr(self, '_hl_after', None)))

    def _on_tab(self, _):
        if self._ac_active:
            self._apply_ac(); return "break"
        self.editor.insert("insert", "    ")
        return "break"

    def _on_return(self, _):
        if self._ac_active:
            self._apply_ac(); return "break"
        line   = self.editor.get("insert linestart", "insert")
        indent = len(line) - len(line.lstrip())
        if line.rstrip().endswith(":"):
            indent += 4
        self.editor.insert("insert", "\n" + " " * indent)
        self._update_line_numbers()
        return "break"

    def _on_up(self, _):
        if self._ac_active:
            self._ac_index = max(0, self._ac_index - 1)
            self._update_ac_sel(); return "break"

    def _on_down(self, _):
        if self._ac_active:
            self._ac_index = min(len(self._ac_matches) - 1, self._ac_index + 1)
            self._update_ac_sel(); return "break"

    def _sync_scroll(self, *args):
        self.line_nums.yview_moveto(args[0])

    def _on_editor_scroll(self, *args):
        self.editor.yview(*args)
        self.line_nums.yview(*args)

    def _update_line_numbers(self):
        n    = self.editor.get("1.0", "end-1c").count("\n") + 1
        nums = "\n".join(str(i) for i in range(1, n + 1))
        self.line_nums.configure(state="normal")
        self.line_nums.delete("1.0", "end")
        self.line_nums.insert("1.0", nums)
        self.line_nums.configure(state="disabled")

    def _clear_editor(self):
        if messagebox.askyesno("Clear", "Clear the editor?"):
            self.editor.delete("1.0", "end")
            self._update_line_numbers()

    def _filter_snippets(self, *_) -> None:
        query = self.snip_search_var.get().lower().strip()
        self.snip_list.delete(0, "end")
        self._snip_index_map = []
        for i, (name, desc, _code) in enumerate(SNIPPETS):
            if query in name.lower() or query in desc.lower():
                self.snip_list.insert("end", f"  {name}")
                self._snip_index_map.append(i)
        self.snip_desc.configure(text="Select a snippet to see what it does.", fg=C["muted"])

    def _insert_snippet(self, _=None):
        sel = self.snip_list.curselection()
        if not sel:
            return
        real_idx = self._snip_index_map[sel[0]]
        _, _, code = SNIPPETS[real_idx]
        self.editor.insert("insert", code)
        self._schedule_highlight()
        self._update_line_numbers()
        self.editor.focus_set()

    def _show_snippet_desc(self, _=None):
        sel = self.snip_list.curselection()
        if not sel:
            return
        real_idx = self._snip_index_map[sel[0]]
        _, desc, _ = SNIPPETS[real_idx]
        self.snip_desc.configure(text=desc, fg=C["text"])

    # ── Inline autocomplete ────────────────────────────────────────────────

    def _get_prefix(self) -> str:
        line = self.editor.get("insert linestart", "insert")
        m    = re.search(r'[\w.]+$', line)
        return m.group() if m else ""

    def _get_col_context(self) -> list[str] | None:
        """
        If the cursor is inside  varname['<cursor>  or  varname["<cursor>
        return the CSV column headers for that varname, else None.
        """
        if not PANDAS_OK:
            return None
        line = self.editor.get("insert linestart", "insert")
        m = re.search(r'(\w+)\[[\'"]([^\'"]*?)$', line)
        if not m:
            return None
        varname = m.group(1)
        # resolve varname → file path
        path: str | None = None
        if varname in self.loaded_files:
            path = self.loaded_files[varname]
        else:
            # handle df1/df2/... aliases
            alias = re.match(r'^df(\d+)$', varname)
            if alias:
                idx = int(alias.group(1)) - 1
                keys = list(self.loaded_files.keys())
                if 0 <= idx < len(keys):
                    path = self.loaded_files[keys[idx]]
        if path is None:
            return None
        try:
            assert pd is not None
            cols = list(pd.read_csv(path, nrows=0).columns)
            return cols
        except Exception:
            return None

    def _get_candidates(self) -> list[str]:
        seen: set[str] = set()
        result: list[str] = []
        for w in (
            sorted(KEYWORDS) +
            sorted(BUILTINS) +
            sorted(self.loaded_files.keys()) +
            [f"df{i+1}" for i in range(len(self.loaded_files))]
        ):
            if w not in seen:
                seen.add(w); result.append(w)
        return result

    def _trigger_ac(self, event=None):
        # ── Column context: inside varname['  ────────────────────────────
        col_candidates = self._get_col_context()
        if col_candidates is not None:
            line    = self.editor.get("insert linestart", "insert")
            m       = re.search(r'\[[\'"]([^\'"]*?)$', line)
            prefix  = m.group(1) if m else ""
            matches = [c for c in col_candidates if c.lower().startswith(prefix.lower())] or col_candidates
            self._ac_matches  = matches
            self._ac_index    = 0
            self._ac_active   = True
            self._ac_col_mode = True
            self._render_ac()
            return
        self._ac_col_mode = False

        # ── Normal keyword/variable context ───────────────────────────────
        prefix  = self._get_prefix()
        if not prefix:
            self._hide_ac(); return
        matches = [c for c in self._get_candidates() if c.startswith(prefix) and c != prefix]
        if not matches:
            self._hide_ac(); return
        self._ac_matches = matches
        self._ac_index   = 0
        self._ac_active  = True
        self._render_ac()
        if event and getattr(event, "keysym", None) == "space":
            return "break"

    def _render_ac(self):
        if self._ac_frame and self._ac_frame.winfo_exists():
            self._ac_frame.destroy()

        try:
            bbox = self.editor.bbox("insert")
        except Exception:
            return
        if not bbox:
            return

        x, y, _, h = bbox
        # offset for line-number gutter width
        gutter_w = self.line_nums.winfo_width()
        ax = gutter_w + x
        ay = y + h + 2

        n        = min(8, len(self._ac_matches))
        row_h    = 20
        widget_h = n * row_h + 6

        self._ac_frame = tk.Frame(self._ec, bg=C["ac_bg"],
                                   highlightthickness=1,
                                   highlightbackground=C["accent"])
        self._ac_frame.place(x=ax, y=ay, width=200, height=widget_h)
        self._ac_frame.lift()

        self._ac_listbox = tk.Listbox(
            self._ac_frame,
            bg=C["ac_bg"], fg=C["text"],
            selectbackground=C["ac_sel"], selectforeground="#fff",
            font=FONT_MONO_SM, borderwidth=0, highlightthickness=0,
            activestyle="none", relief="flat", height=n,
        )
        self._ac_listbox.pack(fill="both", expand=True, padx=2, pady=2)
        for m in self._ac_matches:
            self._ac_listbox.insert("end", m)
        self._update_ac_sel()
        # Clicking applies immediately
        self._ac_listbox.bind("<Button-1>", lambda _: self.after(10, self._apply_ac))

    def _update_ac_sel(self):
        if not self._ac_listbox:
            return
        self._ac_listbox.selection_clear(0, "end")
        self._ac_listbox.selection_set(self._ac_index)
        self._ac_listbox.see(self._ac_index)

    def _apply_ac(self):
        if not self._ac_active or not self._ac_matches:
            return
        word = self._ac_matches[self._ac_index]
        if self._ac_col_mode:
            # Replace everything after the opening quote up to cursor
            line = self.editor.get("insert linestart", "insert")
            m    = re.search(r'\[[\'"]([^\'"]*?)$', line)
            if m:
                self.editor.delete(f"insert - {len(m.group(1))}c", "insert")
            self.editor.insert("insert", word)
        else:
            prefix = self._get_prefix()
            if prefix:
                self.editor.delete(f"insert - {len(prefix)}c", "insert")
            self.editor.insert("insert", word)
        self._hide_ac()
        self._schedule_highlight()

    def _hide_ac(self, *_):
        self._ac_active  = False
        self._ac_matches = []
        if self._ac_frame and self._ac_frame.winfo_exists():
            self._ac_frame.destroy()
        self._ac_frame   = None
        self._ac_listbox = None

    # ── Script execution ───────────────────────────────────────────────────

    def _run_script(self):
        if not PANDAS_OK:
            self._write_output("pandas not installed.\nsudo pacman -S python-pandas", "err"); return
        assert pd is not None
        code = self.editor.get("1.0", "end-1c").strip()
        if not code:
            return

        namespace: dict = {"pd": pd}
        for name, path in self.loaded_files.items():
            try:
                namespace[name] = pd.read_csv(path)
            except Exception as e:
                self._write_output(f"Error loading '{name}': {e}", "err"); return
        for i, name in enumerate(self.loaded_files.keys(), 1):
            namespace[f"df{i}"] = namespace[name]

        buf = io.StringIO()
        t0  = time.perf_counter()
        try:
            with redirect_stdout(buf):
                exec(code, namespace)   # noqa: S102
            elapsed = time.perf_counter() - t0
        except Exception:
            elapsed = time.perf_counter() - t0
            self._write_output(traceback.format_exc(), "err")
            self._set_status(f"Error — {elapsed*1000:.1f} ms", C["err"]); return

        result  = namespace.get("result")
        printed = buf.getvalue()

        if result is None:
            self._write_output(
                "Script ran OK but no 'result' variable was set.\n"
                "Assign your output to: result", "info")
        elif not isinstance(result, pd.DataFrame):
            self._write_output(f"'result' is {type(result).__name__}, expected a DataFrame.", "info")
        else:
            rows, cols = result.shape
            self._populate_result(result)
            self.result_info.configure(
                text=f"  {rows:,} rows × {cols} columns   ·   {elapsed*1000:.1f} ms",
                fg=C["accent2"])
            out = printed if printed else f"✓  {elapsed*1000:.1f} ms  —  {rows:,} rows × {cols} columns"
            self._write_output(out, "ok")
            self._set_status(f"✓  {rows:,} × {cols}  in {elapsed*1000:.1f} ms", C["ok"])

    def _populate_result(self, df: Any, max_rows: int = 500):
        tree = self.result_tree
        tree.delete(*tree.get_children())
        cols = list(df.columns.astype(str))
        tree["columns"] = cols
        for c in cols:
            tree.heading(c, text=c)
            tree.column(c, width=100, minwidth=60, stretch=True)
        for _, row in df.head(max_rows).iterrows():
            tree.insert("", "end", values=[str(v) for v in row])
        self._result_df = df

    # ── Export ─────────────────────────────────────────────────────────────

    def _export_result(self):
        if self._result_df is None:
            messagebox.showinfo("No result", "Run a script first."); return
        path = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("CSV", "*.csv")],
            title="Export result as CSV")
        if path:
            self._result_df.to_csv(path, index=False)
            self._set_status(f"Exported → {path}", C["ok"])
            messagebox.showinfo("Saved", f"CSV saved:\n{path}")

    # ── Script library ─────────────────────────────────────────────────────

    def _save_script(self):
        name = re.sub(r'[^\w\-]', '_', self.script_name_var.get().strip() or "untitled")
        path = os.path.join(SCRIPTS_DIR, f"{name}.py")
        with open(path, "w", encoding="utf-8") as fh:
            fh.write(self.editor.get("1.0", "end-1c"))
        self.current_script_path = path
        self._refresh_scripts()
        self._set_status(f"Saved: {name}.py", C["ok"])

    def _open_script(self):
        path = filedialog.askopenfilename(
            initialdir=SCRIPTS_DIR,
            filetypes=[("Python scripts", "*.py"), ("All files", "*.*")])
        if path:
            self._load_script_from_path(path)

    def _load_script_from_path(self, path: str):
        with open(path, "r", encoding="utf-8") as fh:
            code = fh.read()
        self.editor.delete("1.0", "end")
        self.editor.insert("1.0", code)
        self.script_name_var.set(os.path.splitext(os.path.basename(path))[0])
        self.current_script_path = path
        self._schedule_highlight()
        self._update_line_numbers()
        self._set_status(f"Loaded: {os.path.basename(path)}", C["accent2"])

    def _refresh_scripts(self):
        self.scripts_list.delete(0, "end")
        for fn in sorted(os.listdir(SCRIPTS_DIR)):
            if fn.endswith(".py"):
                self.scripts_list.insert("end", f"  {fn}")

    def _load_saved_script(self, _=None):
        sel = self.scripts_list.curselection()
        if not sel:
            return
        fn   = self.scripts_list.get(sel[0]).strip()
        path = os.path.join(SCRIPTS_DIR, fn)
        self._load_script_from_path(path)

    def _delete_script(self):
        sel = self.scripts_list.curselection()
        if not sel:
            return
        fn = self.scripts_list.get(sel[0]).strip()
        if messagebox.askyesno("Delete", f"Delete '{fn}'?"):
            os.remove(os.path.join(SCRIPTS_DIR, fn))
            self._refresh_scripts()


# ══════════════════════════════════════════════════════════════════════════
if __name__ == "__main__":
    app = Workbench()
    app.mainloop()
