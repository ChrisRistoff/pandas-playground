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
        "Imagine two spreadsheets that share a common column — like a customer ID.\n"
        "This sticks them side-by-side, matching rows by that shared column.\n\n"
        "→ Change 'id' to whatever your shared column is called.\n"
        "→ df1 is your main file, df2 is the file you're pulling extra info from.\n"
        "→ Rows in df1 with no match in df2 are still kept (blank for the missing columns).",
        "# Change 'id' to the column name that both files share\n"
        "# Example: if both files have an 'order_id' column, write on='order_id'\n"
        "result = pd.merge(df1, df2, on='id', how='left')\n"
    ),
    (
        "Merge on different column names",
        "Same idea as a regular merge, but the shared column has a different name in each file.\n"
        "For example, one file calls it 'user_id' and the other just calls it 'id'.\n\n"
        "→ left_on is the column name in your first file.\n"
        "→ right_on is the column name in your second file.",
        "# left_on  = the column name in df1 (your first file)\n"
        "# right_on = the column name in df2 (your second file)\n"
        "# Change these to match your actual column names\n"
        "result = pd.merge(df1, df2, left_on='user_id', right_on='id', how='left')\n"
    ),
    (
        "Stack rows (concat)",
        "Takes two files and stacks all their rows into one big file, one on top of the other.\n"
        "Think of it like copy-pasting the rows from one spreadsheet below the other.\n\n"
        "→ Both files should have the same column names for this to make sense.\n"
        "→ The row numbers will be reset automatically.",
        "# Stacks df1 and df2 on top of each other into one file\n"
        "# Both files should have the same columns\n"
        "result = pd.concat([df1, df2], ignore_index=True)\n"
    ),
    (
        "Stack many files",
        "Same as stacking two files, but you can add as many as you like.\n"
        "Useful if you have monthly exports that all need to be combined into one.\n\n"
        "→ Just add more file names inside the square brackets, separated by commas.",
        "# Add as many files as you need inside the brackets\n"
        "# Example: [january, february, march] if those are your file names\n"
        "result = pd.concat([df1, df2, df3], ignore_index=True)\n"
    ),

    # ── Filtering rows ────────────────────────────────────────────────────
    (
        "Filter: exact match",
        "Keeps only the rows where a column contains a specific value — like filtering in Excel.\n"
        "Everything else is removed from the result.\n\n"
        "→ Replace 'column' with the name of the column you want to filter.\n"
        "→ Replace 'value' with the exact text or number you're looking for.",
        "# Replace 'column' with your column name, e.g. 'status' or 'country'\n"
        "# Replace 'value' with what you want to keep, e.g. 'active' or 'UK'\n"
        "result = df1[df1['column'] == 'value']\n"
    ),
    (
        "Filter: contains text",
        "Keeps rows where a column contains a word or phrase anywhere inside it.\n"
        "Like using Ctrl+F search but as a filter on your data.\n\n"
        "→ Replace 'column' with the column to search in.\n"
        "→ Replace 'text' with the word or phrase to search for.\n"
        "→ It ignores uppercase/lowercase differences automatically.",
        "# Replace 'column' with the column to search in, e.g. 'description'\n"
        "# Replace 'text' with what you're searching for, e.g. 'refund'\n"
        "result = df1[df1['column'].str.contains('text', na=False, case=False)]\n"
    ),
    (
        "Filter: number range",
        "Keeps only rows where a number falls between two values.\n"
        "For example, keep all orders between £100 and £500.\n\n"
        "→ Replace 'amount' with your numeric column name.\n"
        "→ Change 100 and 500 to your minimum and maximum values.",
        "# Replace 'amount' with your numeric column, e.g. 'price' or 'age'\n"
        "# Change 100 and 500 to your min and max values\n"
        "result = df1[(df1['amount'] >= 100) & (df1['amount'] <= 500)]\n"
    ),
    (
        "Filter: multiple values (isin)",
        "Keeps rows that match any one of several values — like ticking multiple boxes in a filter.\n"
        "Much easier than writing separate filters for each value.\n\n"
        "→ Replace 'status' with your column name.\n"
        "→ Replace the values in the list with the ones you want to keep.",
        "# Replace 'status' with your column name\n"
        "# List all the values you want to keep inside the square brackets\n"
        "result = df1[df1['status'].isin(['active', 'pending'])]\n"
    ),
    (
        "Filter: exclude values",
        "The opposite of the one above — removes rows that match certain values and keeps everything else.\n"
        "Useful for getting rid of rows you don't want, like deleted or test records.\n\n"
        "→ Replace 'status' with your column name.\n"
        "→ List the values you want to throw away.",
        "# Replace 'status' with your column name\n"
        "# List the values you want to REMOVE inside the brackets\n"
        "result = df1[~df1['status'].isin(['deleted', 'archived'])]\n"
    ),
    (
        "Filter: non-empty rows",
        "Removes any row where a particular column has been left blank.\n"
        "Great for cleaning up exports that have missing entries.\n\n"
        "→ Replace 'column' with the column you want to check for blanks.",
        "# Replace 'column' with the column you want to check, e.g. 'email'\n"
        "# Rows where that column is blank will be removed\n"
        "result = df1[df1['column'].notna() & (df1['column'] != '')]\n"
    ),
    (
        "Filter: top N rows",
        "Keeps only the first N rows of your file.\n"
        "Handy when your file is huge and you just want a quick look at the top.\n\n"
        "→ Change 100 to however many rows you want.",
        "# Change 100 to however many rows you want to keep\n"
        "result = df1.head(100)\n"
    ),
    (
        "Filter: random sample",
        "Picks a random selection of rows from your file.\n"
        "Useful for spot-checking a large dataset without looking at all of it.\n\n"
        "→ Change 100 to how many random rows you want.\n"
        "→ The random_state=42 just makes sure you get the same random rows each time you run it.",
        "# Change 100 to how many random rows you want\n"
        "# random_state=42 means you'll get the same sample each time — remove it for a different sample each run\n"
        "result = df1.sample(n=100, random_state=42)\n"
    ),

    # ── Cleaning data ─────────────────────────────────────────────────────
    (
        "Drop duplicate rows",
        "Finds and removes rows that are completely identical — keeping only one copy of each.\n"
        "If two rows have the exact same data in every column, one of them gets removed.\n\n"
        "→ If you only want to check for duplicates in one column (like email), use the second line.",
        "# Removes rows that are 100% identical across all columns\n"
        "result = df1.drop_duplicates()\n"
        "\n"
        "# To remove duplicates based on just one column (keeps the first occurrence):\n"
        "# result = df1.drop_duplicates(subset=['email'])\n"
    ),
    (
        "Fill missing values",
        "Finds empty/blank cells and fills them in with something you choose.\n"
        "You can set a different fill value for different columns.\n\n"
        "→ Replace 'column' with your column name.\n"
        "→ Replace 'unknown' or 0 with whatever you want the blank cells to say.",
        "# Replace 'column' with your column name\n"
        "# Replace 'unknown' with the text to use for blank text cells\n"
        "# Replace 0 with the number to use for blank number cells\n"
        "result = df1.fillna({'column': 'unknown', 'amount': 0})\n"
    ),
    (
        "Drop rows with any blanks",
        "Removes every row that has at least one empty cell anywhere in it.\n"
        "After this, every remaining row will be fully filled in.\n\n"
        "→ Warning: this can remove a lot of rows if your data has many blanks.",
        "# Removes any row that has even one empty cell\n"
        "result = df1.dropna()\n"
    ),
    (
        "Drop rows where column is blank",
        "Like the one above, but only removes rows where one specific column is empty.\n"
        "Other columns can still have blanks — only the column you name matters.\n\n"
        "→ Replace 'column' with the column that must not be blank.",
        "# Replace 'column' with the column that must have a value\n"
        "# e.g. 'email' — removes any row where email is missing\n"
        "result = df1.dropna(subset=['column'])\n"
    ),
    (
        "Strip whitespace from all text",
        "Removes invisible spaces from the start and end of text in every cell.\n"
        "These hidden spaces are a common cause of 'Why won't these rows match?!' problems.\n\n"
        "→ No changes needed — this works on the whole file automatically.",
        "# Removes leading/trailing spaces from every text cell in the file\n"
        "# This fixes a very common invisible data quality problem\n"
        "result = df1.apply(lambda col: col.str.strip() if col.dtype == 'object' else col)\n"
    ),
    (
        "Standardise text case",
        "Makes all the text in a column the same case — all lowercase, all UPPERCASE, or Title Case.\n"
        "Useful when your data has mixed capitalisation like 'london', 'London', 'LONDON'.\n\n"
        "→ Replace 'column' with your column name.\n"
        "→ Choose .str.lower(), .str.upper(), or .str.title() depending on what you want.",
        "# Replace 'column' with the column you want to change\n"
        "result = df1.copy()\n"
        "result['column'] = result['column'].str.lower()   # all lowercase\n"
        "# result['column'] = result['column'].str.upper()  # ALL UPPERCASE\n"
        "# result['column'] = result['column'].str.title()  # Title Case\n"
    ),
    (
        "Fix column data types",
        "Sometimes when a CSV loads, numbers come in as text and dates come in as plain text.\n"
        "This fixes them so the computer understands what type of data they are.\n\n"
        "→ Replace 'amount' with your number column name.\n"
        "→ Replace 'date' with your date column name.\n"
        "→ Any values that can't be converted will be left blank rather than causing an error.",
        "# Replace 'amount' with the column that should be a number\n"
        "# Replace 'date' with the column that should be a date\n"
        "result = df1.copy()\n"
        "result['amount'] = pd.to_numeric(result['amount'], errors='coerce')\n"
        "result['date']   = pd.to_datetime(result['date'],  errors='coerce')\n"
    ),
    (
        "Replace values in column",
        "Swaps specific values in a column with new ones — like a find-and-replace in Word.\n"
        "You can do as many swaps as you like in one go.\n\n"
        "→ Replace 'column' with your column name.\n"
        "→ Add as many 'old value': 'new value' pairs as you need.",
        "# Replace 'column' with your column name\n"
        "# Add as many swaps as you need: 'what it says now': 'what you want it to say'\n"
        "result = df1.copy()\n"
        "result['column'] = result['column'].replace({\n"
        "    'old_value': 'new_value',\n"
        "    'Y': 'Yes',\n"
        "    'N': 'No',\n"
        "})\n"
    ),
    (
        "Remove special characters",
        "Strips out punctuation and symbols from a text column, leaving only letters, numbers and spaces.\n"
        "Really useful for cleaning up phone numbers, product codes, or messy exported data.\n\n"
        "→ Replace 'column' with the column you want to clean.",
        "# Replace 'column' with the column you want to clean\n"
        "# This removes anything that isn't a letter, number, or space\n"
        "result = df1.copy()\n"
        "result['column'] = result['column'].str.replace(r'[^\\w\\s]', '', regex=True)\n"
    ),

    # ── Reshaping columns ─────────────────────────────────────────────────
    (
        "Select columns",
        "Keeps only the columns you name and throws away the rest.\n"
        "Great for trimming down a file with dozens of columns to just the ones you care about.\n\n"
        "→ Replace col1, col2, col3 with your actual column names.\n"
        "→ Add or remove column names from the list as needed.",
        "# List the columns you want to keep — everything else will be removed\n"
        "# Replace col1, col2, col3 with your actual column names\n"
        "result = df1[['col1', 'col2', 'col3']]\n"
    ),
    (
        "Drop columns",
        "Removes specific columns you don't want, keeping everything else.\n"
        "The opposite of selecting — useful when there are only one or two columns to get rid of.\n\n"
        "→ Replace the column names with the ones you want to delete.",
        "# List the columns you want to DELETE — everything else is kept\n"
        "result = df1.drop(columns=['unwanted_col1', 'unwanted_col2'])\n"
    ),
    (
        "Rename columns",
        "Gives your columns new names — like double-clicking a column header in Excel to rename it.\n\n"
        "→ On the left of the colon: the current column name.\n"
        "→ On the right of the colon: what you want to rename it to.\n"
        "→ Add as many pairs as you need.",
        "# Left side: the current column name  |  Right side: the new name you want\n"
        "result = df1.rename(columns={\n"
        "    'old_name_1': 'new_name_1',\n"
        "    'old_name_2': 'new_name_2',\n"
        "})\n"
    ),
    (
        "Reorder columns",
        "Changes the order that columns appear in — left to right.\n"
        "Any columns you don't include will be dropped.\n\n"
        "→ List your column names in the order you want them to appear.",
        "# List column names in the order you want them to appear left to right\n"
        "# Any column not mentioned here will be removed from the result\n"
        "result = df1[['col_a', 'col_b', 'col_c']]\n"
    ),
    (
        "Add calculated column",
        "Creates a brand new column by doing a calculation using other columns.\n"
        "Like adding a formula column in Excel.\n\n"
        "→ Change 'full_name', 'first_name', 'last_name' to your actual column names.\n"
        "→ You can combine text with + ' ' + or do maths like * and +.",
        "# Creates a new column called 'full_name' by joining two text columns\n"
        "# Creates a new column called 'total' by multiplying qty by unit_price\n"
        "# Change the column names and formula to suit your data\n"
        "result = df1.copy()\n"
        "result['full_name'] = result['first_name'] + ' ' + result['last_name']\n"
        "result['total']     = result['qty'] * result['unit_price']\n"
    ),
    (
        "Conditional column (if/else)",
        "Adds a new column whose value depends on a condition — just like =IF() in Excel.\n"
        "If the condition is true, it gets one value. If not, it gets another.\n\n"
        "→ Replace 'score' with your column name.\n"
        "→ Change >= 50 to your condition.\n"
        "→ Change 'pass' and 'fail' to whatever labels you want.",
        "# Adds a new column 'label' based on whether a condition is true or false\n"
        "# Change 'score' to your column, >= 50 to your condition, and 'pass'/'fail' to your labels\n"
        "import numpy as np\n"
        "result = df1.copy()\n"
        "result['label'] = np.where(result['score'] >= 50, 'pass', 'fail')\n"
    ),
    (
        "Multiple conditions (nested if)",
        "Like the conditional column above, but with more than two outcomes.\n"
        "Like a nested =IF(IF(IF())) in Excel, but much easier to read.\n\n"
        "→ Add or remove conditions in the condlist.\n"
        "→ Each condition matches up with a label in choicelist (top to bottom).\n"
        "→ 'default' is what gets used if none of the conditions match.",
        "# Each condition in condlist matches the label at the same position in choicelist\n"
        "# The first condition that is true wins — order matters!\n"
        "# Change 'score', the numbers, and the labels to suit your data\n"
        "import numpy as np\n"
        "result = df1.copy()\n"
        "result['grade'] = np.select(\n"
        "    condlist=[\n"
        "        result['score'] >= 90,\n"
        "        result['score'] >= 70,\n"
        "        result['score'] >= 50,\n"
        "    ],\n"
        "    choicelist=['A', 'B', 'C'],\n"
        "    default='F'   # used when none of the conditions above are true\n"
        ")\n"
    ),
    (
        "Split column by delimiter",
        "Splits one column into two separate columns at a dividing character.\n"
        "For example, splitting 'John Smith' into a 'first' and 'last' column.\n\n"
        "→ Replace 'full_name' with the column you want to split.\n"
        "→ Change ' ' to whatever character separates the two parts (a space, comma, dash, etc.).\n"
        "→ Replace 'first' and 'last' with your new column names.",
        "# Replace 'full_name' with the column to split\n"
        "# Change ' ' to the character that divides the two parts — e.g. ',' or '-'\n"
        "# Replace 'first' and 'last' with what you want the two new columns to be called\n"
        "result = df1.copy()\n"
        "result[['first', 'last']] = result['full_name'].str.split(' ', n=1, expand=True)\n"
    ),
    (
        "Extract substring / regex",
        "Pulls out a specific piece of text from inside a column using a pattern.\n"
        "For example, extracting a 4-digit year from a column that contains mixed text like 'Report 2024 Final'.\n\n"
        "→ Replace 'description' with your column name.\n"
        "→ The pattern \\d{4} means 'find 4 digits in a row' — change it for other patterns.",
        "# Pulls out the first 4-digit number found in each cell\n"
        "# Replace 'description' with your column name\n"
        "# The pattern r'(\\d{4})' finds 4 digits — change it if you need a different pattern\n"
        "result = df1.copy()\n"
        "result['year'] = result['description'].str.extract(r'(\\d{4})')\n"
    ),

    # ── Sorting & ranking ─────────────────────────────────────────────────
    (
        "Sort by column",
        "Sorts all the rows by a column — either smallest to largest, or largest to smallest.\n"
        "Like clicking a column header to sort in Excel.\n\n"
        "→ Replace 'column' with the column you want to sort by.\n"
        "→ ascending=True = A to Z or lowest to highest.\n"
        "→ ascending=False = Z to A or highest to lowest.",
        "# Replace 'column' with the column to sort by\n"
        "# ascending=True = A→Z or lowest first  |  ascending=False = Z→A or highest first\n"
        "result = df1.sort_values('column', ascending=True)\n"
    ),
    (
        "Sort by multiple columns",
        "Sorts by a first column, then uses a second column to break any ties.\n"
        "For example: sort by department (A-Z), then by salary (highest first).\n\n"
        "→ Replace 'dept' and 'salary' with your column names.\n"
        "→ Set True or False for each column's sort direction.",
        "# First sorts by 'dept' A→Z, then within each dept sorts by 'salary' highest first\n"
        "# Replace column names and True/False to match what you need\n"
        "result = df1.sort_values(['dept', 'salary'], ascending=[True, False])\n"
    ),
    (
        "Add rank column",
        "Adds a numbered rank column — 1st, 2nd, 3rd etc. — based on a score.\n"
        "The highest score gets rank 1 by default.\n\n"
        "→ Replace 'score' with the column you want to rank by.\n"
        "→ The result will be sorted so rank 1 appears at the top.",
        "# Replace 'score' with the column to rank by — highest value gets rank 1\n"
        "# The result is sorted so rank 1 appears at the top\n"
        "result = df1.copy()\n"
        "result['rank'] = result['score'].rank(ascending=False, method='dense').astype(int)\n"
        "result = result.sort_values('rank')\n"
    ),
    (
        "Top N per group",
        "Finds the top N rows within each group — not just the overall top.\n"
        "For example: the top 3 sales in each region, or the top 5 products per category.\n\n"
        "→ Replace 'amount' with the column to rank by.\n"
        "→ Replace 'category' with the column that defines the groups.\n"
        "→ Change .head(3) to however many top results you want per group.",
        "# Replace 'amount' with the column to find the top values in\n"
        "# Replace 'category' with the column that defines your groups (e.g. 'region', 'department')\n"
        "# Change head(3) to the number of top results you want per group\n"
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
        "Groups your rows by a category and adds up a number for each group.\n"
        "Like a pivot table in Excel — great for totals by region, department, product, etc.\n\n"
        "→ Replace 'category' with the column you want to group by.\n"
        "→ Replace 'amount' with the column you want to add up.",
        "# Replace 'category' with what you want to group by (e.g. 'region', 'product')\n"
        "# Replace 'amount' with the number column you want totalled\n"
        "result = df1.groupby('category').agg({'amount': 'sum'}).reset_index()\n"
    ),
    (
        "Group & multiple aggregations",
        "Groups your rows and calculates several things at once — total, count, average, max.\n"
        "All in one go instead of running separate calculations.\n\n"
        "→ Replace 'category' with your grouping column.\n"
        "→ Replace 'amount' with the number column you're summarising.\n"
        "→ You can rename total, count, average, maximum to whatever you like.",
        "# Replace 'category' with the column to group by\n"
        "# Replace 'amount' with the number column to summarise\n"
        "# The names on the left (total, count, etc.) become your new column names\n"
        "result = df1.groupby('category').agg(\n"
        "    total   = ('amount', 'sum'),\n"
        "    count   = ('amount', 'count'),\n"
        "    average = ('amount', 'mean'),\n"
        "    maximum = ('amount', 'max'),\n"
        ").reset_index()\n"
    ),
    (
        "Value counts",
        "Counts how many times each unique value appears in a column.\n"
        "Great for answering 'how many orders per country?' or 'how many customers per status?'\n\n"
        "→ Replace 'column' with the column you want to count.",
        "# Replace 'column' with the column you want to count values in\n"
        "# The result will have two columns: the value, and how many times it appears\n"
        "result = df1['column'].value_counts().reset_index()\n"
        "result.columns = ['value', 'count']\n"
    ),
    (
        "Pivot table",
        "Reorganises your data so one column becomes the row labels and another becomes the column headers.\n"
        "Exactly like creating a pivot table in Excel.\n\n"
        "→ Replace 'row_col' with what you want as row labels (left side).\n"
        "→ Replace 'col_col' with what you want spread across the top as column headers.\n"
        "→ Replace 'value_col' with the numbers to fill the table with.",
        "# row_col   = what appears as rows (left side of the table)\n"
        "# col_col   = what spreads across the top as column headers\n"
        "# value_col = the numbers that fill the table\n"
        "# aggfunc   = what to do with the numbers: 'sum', 'mean', 'count', 'max', 'min'\n"
        "result = df1.pivot_table(\n"
        "    index='row_col',\n"
        "    columns='col_col',\n"
        "    values='value_col',\n"
        "    aggfunc='sum',\n"
        "    fill_value=0,   # use 0 instead of blank for missing combinations\n"
        ").reset_index()\n"
    ),
    (
        "Unpivot / melt (wide → long)",
        "The reverse of a pivot table — turns column headers into rows.\n"
        "Useful when your data has months or dates spread across many columns and you want them in one column instead.\n\n"
        "→ Replace 'id' and 'name' with the columns that should stay as they are.\n"
        "→ 'month' will become the name of the new column holding the old header names.\n"
        "→ 'amount' will become the name of the new column holding the values.",
        "# id_vars  = columns that should stay as they are (not unpivoted)\n"
        "# var_name = what to call the new column that holds the old header names\n"
        "# value_name = what to call the new column that holds the values\n"
        "result = df1.melt(\n"
        "    id_vars=['id', 'name'],\n"
        "    var_name='month',\n"
        "    value_name='amount',\n"
        ")\n"
    ),
    (
        "Running total (cumulative sum)",
        "Adds a column that shows a running total — each row shows the total so far up to that point.\n"
        "Like a bank statement balance that grows with each transaction.\n\n"
        "→ Replace 'date' with your date column so rows are in the right order first.\n"
        "→ Replace 'amount' with the number column you want to accumulate.",
        "# Sorts by date first so the running total goes in the right order\n"
        "# Replace 'date' with your date column and 'amount' with the number to accumulate\n"
        "result = df1.sort_values('date').copy()\n"
        "result['running_total'] = result['amount'].cumsum()\n"
    ),
    (
        "Percentage of total",
        "Adds a column showing what percentage of the grand total each row represents.\n"
        "For example, each product's share of total sales.\n\n"
        "→ Replace 'amount' with the number column you want to calculate percentages from.",
        "# Replace 'amount' with your number column\n"
        "# Each row will show its value as a % of the overall total\n"
        "result = df1.copy()\n"
        "result['pct'] = (result['amount'] / result['amount'].sum() * 100).round(2)\n"
    ),
    (
        "Percentage within group",
        "Like percentage of total, but calculated within each group separately.\n"
        "For example, each product's share of sales within its own category — not the whole dataset.\n\n"
        "→ Replace 'category' with your grouping column.\n"
        "→ Replace 'amount' with your number column.",
        "# Replace 'category' with the column that defines your groups\n"
        "# Replace 'amount' with your number column\n"
        "# Each row will show its value as a % of its group's total\n"
        "result = df1.copy()\n"
        "result['group_total'] = result.groupby('category')['amount'].transform('sum')\n"
        "result['pct_of_group'] = (result['amount'] / result['group_total'] * 100).round(2)\n"
    ),
    (
        "Cross-tabulation",
        "Counts how many times each combination of two columns appears.\n"
        "Like asking 'how many customers in each country have each status?'\n\n"
        "→ Replace 'col_a' with your first category column (becomes the rows).\n"
        "→ Replace 'col_b' with your second category column (becomes the column headers).",
        "# Replace 'col_a' with the column for rows, 'col_b' with the column for headers\n"
        "# The numbers in the table show how many rows match each combination\n"
        "result = pd.crosstab(df1['col_a'], df1['col_b']).reset_index()\n"
    ),

    # ── Date & time ───────────────────────────────────────────────────────
    (
        "Parse dates & extract parts",
        "Converts a column of date text into real dates, then pulls out the year, month, and day as separate columns.\n"
        "Useful when your CSV has dates stored as plain text like '2024-03-15'.\n\n"
        "→ Replace 'date_col' with your date column name.\n"
        "→ Any dates that can't be read will be left blank rather than causing an error.",
        "# Replace 'date_col' with the name of your date column\n"
        "# This creates separate year, month, and day columns from your date\n"
        "result = df1.copy()\n"
        "result['date']  = pd.to_datetime(result['date_col'], errors='coerce')\n"
        "result['year']  = result['date'].dt.year\n"
        "result['month'] = result['date'].dt.month\n"
        "result['day']   = result['date'].dt.day\n"
    ),
    (
        "Filter by date range",
        "Keeps only rows that fall within a specific date range.\n"
        "Like filtering a spreadsheet to show only this year's data.\n\n"
        "→ Replace 'date_col' with your date column name.\n"
        "→ Change '2024-01-01' and '2024-12-31' to your start and end dates.",
        "# Replace 'date_col' with your date column name\n"
        "# Change the start and end dates to your desired range (format: YYYY-MM-DD)\n"
        "result = df1.copy()\n"
        "result['date'] = pd.to_datetime(result['date_col'], errors='coerce')\n"
        "result = result[\n"
        "    (result['date'] >= '2024-01-01') &\n"
        "    (result['date'] <= '2024-12-31')\n"
        "]\n"
    ),
    (
        "Days between two dates",
        "Calculates the number of days between two date columns and puts it in a new column.\n"
        "For example, the number of days between an order date and a delivery date.\n\n"
        "→ Replace 'start_date' and 'end_date' with your two date column names.",
        "# Replace 'start_date' and 'end_date' with your two date column names\n"
        "# A new column 'days_diff' will show the number of days between them\n"
        "result = df1.copy()\n"
        "result['start']     = pd.to_datetime(result['start_date'])\n"
        "result['end']       = pd.to_datetime(result['end_date'])\n"
        "result['days_diff'] = (result['end'] - result['start']).dt.days\n"
    ),
    (
        "Group by month/year",
        "Groups your data by month and year and totals a number column for each month.\n"
        "Perfect for creating a monthly summary report from detailed transaction data.\n\n"
        "→ Replace 'date_col' with your date column.\n"
        "→ Replace 'amount' with the number column to total each month.",
        "# Replace 'date_col' with your date column name\n"
        "# Replace 'amount' with the number column to sum up each month\n"
        "# The result will have one row per month with the total for that month\n"
        "result = df1.copy()\n"
        "result['date']       = pd.to_datetime(result['date_col'])\n"
        "result['year_month'] = result['date'].dt.to_period('M').astype(str)\n"
        "result = result.groupby('year_month').agg({'amount': 'sum'}).reset_index()\n"
    ),

    # ── Lookup & matching ─────────────────────────────────────────────────
    (
        "VLOOKUP (add column from lookup table)",
        "Pulls extra information from a second file and adds it to your first file — exactly like VLOOKUP in Excel.\n"
        "For example, you have a sales file with product IDs, and a second file with product names. This joins them.\n\n"
        "→ Replace 'id' with the shared column that links the two files.\n"
        "→ Replace 'label' with the column from df2 that you want to bring across.",
        "# df2 is your lookup table — the file with the extra info you want to bring in\n"
        "# Replace 'id' with the column that both files share\n"
        "# Replace 'label' with the column from df2 you want added to df1\n"
        "result = df1.merge(df2[['id', 'label']], on='id', how='left')\n"
    ),
    (
        "Flag rows present in both files",
        "Adds a True/False column to show which rows in your first file also exist in your second file.\n"
        "Useful for quickly spotting which records appear in both lists.\n\n"
        "→ Replace 'id' with the column to compare between the two files.",
        "# Adds a column 'in_file2' that is True if the row's id also appears in df2\n"
        "# Replace 'id' with the column to match on\n"
        "result = df1.copy()\n"
        "result['in_file2'] = result['id'].isin(df2['id'])\n"
    ),
    (
        "Anti-join (rows only in file1)",
        "Finds rows that exist in your first file but have NO matching row in your second file.\n"
        "Great for finding missing records — like customers who placed an order but never got an invoice.\n\n"
        "→ Replace 'id' with the column to match between the two files.",
        "# Finds rows in df1 that have NO match in df2\n"
        "# Replace 'id' with the column to match on\n"
        "merged = df1.merge(df2[['id']], on='id', how='left', indicator=True)\n"
        "result = merged[merged['_merge'] == 'left_only'].drop(columns=['_merge'])\n"
    ),
    (
        "Inner join (rows in both files)",
        "Keeps only the rows that have a matching entry in BOTH files — everything else is removed.\n"
        "If a row only exists in one file, it won't appear in the result.\n\n"
        "→ Replace 'id' with the shared column between your two files.",
        "# Keeps only rows that exist in BOTH files\n"
        "# Rows with no match in the other file are removed entirely\n"
        "# Replace 'id' with the column both files share\n"
        "result = pd.merge(df1, df2, on='id', how='inner')\n"
    ),
    (
        "Full outer join (all rows)",
        "Combines both files and keeps every row from both — even ones with no match.\n"
        "Where there's no match, the missing columns will just be blank.\n\n"
        "→ Replace 'id' with the shared column between your two files.",
        "# Keeps ALL rows from both files\n"
        "# Where a row has no match in the other file, the missing columns will be blank\n"
        "# Replace 'id' with the column both files share\n"
        "result = pd.merge(df1, df2, on='id', how='outer')\n"
    ),
    (
        "Fuzzy match (approximate names)",
        "Matches rows between two files even when the text isn't spelled exactly the same.\n"
        "For example, 'John Smith' and 'Jon Smith' would still be matched.\n\n"
        "→ You need to install an extra library first: sudo pacman -S python-thefuzz\n"
        "→ Replace 'name' with the column you want to match on.\n"
        "→ The threshold=80 means 80% similar — lower it to match more loosely, raise it to be stricter.",
        "# FIRST: install the fuzzy matching library\n"
        "# Run this in your terminal: sudo pacman -S python-thefuzz\n"
        "#\n"
        "# Replace 'name' with the column you want to match across both files\n"
        "# threshold=80 means 80% similar — lower = more matches, higher = stricter\n"
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
        "Rounds all decimal numbers in your file to 2 decimal places.\n"
        "Tidies up results that end up with many decimal places after calculations.\n\n"
        "→ Change the 2 to however many decimal places you want.\n"
        "→ Use the commented-out line instead if you only want to round one specific column.",
        "# Rounds every number in the file to 2 decimal places\n"
        "result = df1.round(2)\n"
        "\n"
        "# To round just one column instead, use this:\n"
        "# result['amount'] = result['amount'].round(2)\n"
    ),
    (
        "Remove suffix columns after merge",
        "After joining two files that share column names, pandas adds _x and _y suffixes to avoid confusion.\n"
        "This removes the duplicate columns you don't need.\n\n"
        "→ Replace 'id' with your shared column. This keeps df1's version and drops df2's duplicates.",
        "# Joins the files but marks duplicate columns from df2 with '_drop'\n"
        "# Then removes all columns ending in '_drop'\n"
        "# Replace 'id' with your shared column name\n"
        "result = df1.merge(df2, on='id', how='left', suffixes=('', '_drop'))\n"
        "result = result[[c for c in result.columns if not c.endswith('_drop')]]\n"
    ),
    (
        "Flatten pivot column names",
        "After creating a pivot table, the column names sometimes turn into awkward combinations like ('sales', 'sum').\n"
        "This converts them into simple flat names like 'sales_sum'.\n\n"
        "→ No changes needed — just run this after your pivot_table step.",
        "# Run this after creating a pivot table if your column names look like tuples\n"
        "# It joins them into simple names with an underscore, e.g. ('sales', 'sum') → 'sales_sum'\n"
        "result = df1.pivot_table(index='a', columns='b', values='c', aggfunc='sum')\n"
        "result.columns = ['_'.join(str(s) for s in col).strip() for col in result.columns]\n"
        "result = result.reset_index()\n"
    ),
    (
        "Reindex / ensure all columns exist",
        "Makes sure your output always has a specific set of columns — even if some are missing from the source file.\n"
        "Any missing columns are added automatically with blank values.\n"
        "Useful when combining files that don't always have the exact same columns.\n\n"
        "→ Replace the column names in the list with the ones you always want in your output.",
        "# List every column you want to guarantee exists in the output\n"
        "# If a column is missing from the source file, it will be added with blank values\n"
        "expected_cols = ['id', 'name', 'email', 'amount', 'status']\n"
        "result = df1.reindex(columns=expected_cols)\n"
    ),
    (
        "Export multiple sheets to Excel",
        "Saves multiple files as separate sheets inside a single Excel workbook.\n"
        "Instead of having separate CSV files, everything ends up neatly in one .xlsx file.\n\n"
        "→ You need to install openpyxl first: sudo pacman -S python-openpyxl\n"
        "→ Change 'output.xlsx' to whatever you want the file to be called.\n"
        "→ Change 'Sheet1' and 'Sheet2' to whatever you want the tab names to be.",
        "# FIRST: install the Excel library\n"
        "# Run in terminal: sudo pacman -S python-openpyxl\n"
        "#\n"
        "# Change 'output.xlsx' to your desired filename\n"
        "# Change 'Sheet1'/'Sheet2' to whatever you want the tabs to be called\n"
        "with pd.ExcelWriter('output.xlsx', engine='openpyxl') as writer:\n"
        "    df1.to_excel(writer, sheet_name='Sheet1', index=False)\n"
        "    df2.to_excel(writer, sheet_name='Sheet2', index=False)\n"
        "result = df1  # keeps the preview working in this app\n"
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
