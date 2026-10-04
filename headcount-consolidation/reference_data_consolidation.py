# ============================================================
# REFERENCE DATA / HEADCOUNT / APPIAN CONSOLIDATION
# FULL CONSOLIDATED SCRIPT
#
# Fixes in this version:
#   1. FRC reconciliation now uses ALL Appian rows (previously rows
#      without a Position ID were dropped before the reconciliation).
#   2. apply_data_types only strips a trailing ".0" and only converts
#      a column to numbers/dates when no value would be lost.
#   3. Rows with a blank Position Number are kept (previously all but
#      one were silently dropped by drop_duplicates).
#   4. FRC reconciliation no longer uppercases FRC Code in the
#      exported Reference file.
#   5. Pre/Post comparison is text-based, so 1234 vs "1234" is not
#      reported as a change.
#   6. HIRING_FLG can optionally be reset each run (RESET_HIRING_FLG).
#   7. Removed positions can optionally be dropped from the Reference
#      output (DROP_REMOVED_POSITIONS).
#   8. ID columns are formatted as "0" instead of "#,##0".
#   9. Columns are converted to object dtype before cell-by-cell
#      updates (avoids pandas 2.x warnings / pandas 3 errors).
#  10. Job Summary exclusions can be maintained on the Config sheet.
#
# Comments_RD rules (applied in this order):
#   a. Vacant last month, occupied this month      -> "Onboarded"
#   b. STATUS_APPROVAL = AWAITING COO-BE-LEAD APPROVAL
#      or AWAITING APPROVAL FROM OPCO              -> "AWAITING FRC APPROVAL"
#   c. STATUS_APPROVAL = OPCO APPROVED and
#      Comments_RD blank                           -> "RAISE IJP"
#
# Input files: leave the C:\YOUR_PATH placeholders in section 1 and a
# Windows "Open" box asks for each file when the script runs.
#
# Leaver rule:
#   STATUS = VACANT                                -> Leaver cleared
# Filled position rule:
#   Vacant last month, occupied this month         -> Committed Offer,
#   Candidate Name, Joiner flag, Joiner Month and Comments_RD blanked
# Hiring flag rule (HIRING_FLG and HIRING AGAINST PHYSICAL_RD):
#   Vacant last month with flag = YES, filled
#   this month                                     -> flag removed
# New position rule:
#   Position Number not in last month's Reference  -> NEW_POSITION_FLAG = YES
#   (existing flag values are never overwritten)
# LWD rule:
#   Current month Employee ID = Employee ID (MR_Leavers_CM)
#                                                  -> LWD = Termination Date
#   (only Employee Class = Employee; exited employees / Vacant
#    positions are not updated)
#
# Outputs (all in the OUTPUT folder): Reference, Exception Report,
# MR Hiring Update, FRC Reconciliation and Run_Log_<timestamp>.txt
# ============================================================

import os
import re
import sys

import numpy as np
import pandas as pd

from datetime import datetime
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.utils import get_column_letter


# ============================================================
# 1. INPUT FILES / OUTPUT FOLDER / OPTIONS
# ============================================================

reference_file = r"C:\YOUR_PATH\Reference_Data_Hierarchy.xlsx"
hc_file = r"C:\YOUR_PATH\HC_Current_Month.xlsx"
appian_file = r"C:\YOUR_PATH\Appian.xlsx"
# Leavers file - its Termination Date updates LWD (section 9G).
# Optional: if this file is not found, the LWD update is skipped.
leavers_file = r"C:\YOUR_PATH\MR_Leavers_CM.xlsx"
# Outputs go to an OUTPUT folder next to this script.
# To use a different folder, replace this line with e.g.
#   output_folder = r"C:\YOUR_PATH\OUTPUT"
output_folder = os.path.join(os.path.dirname(os.path.abspath(__file__)), "OUTPUT")

# True  -> HIRING_FLG is cleared and re-derived every run
# False -> existing YES values are kept (original behaviour)
RESET_HIRING_FLG = False

# True  -> positions no longer in HC are removed from the Reference output
# False -> they stay in the Reference output (original behaviour);
#          they are always listed in the Exception Report
DROP_REMOVED_POSITIONS = False


# ------------------------------------------------------------
# FILE PICKER
#
# If a path above is still the C:\YOUR_PATH placeholder, a Windows
# "Open" box pops up to choose that file - no need to type paths.
# ------------------------------------------------------------

def pick_file(title):
    try:
        import tkinter as tk
        from tkinter import filedialog

        root = tk.Tk()
        root.withdraw()
        root.attributes("-topmost", True)  # show the box in front of Jupyter
        path = filedialog.askopenfilename(
            title=title,
            filetypes=[("Excel files", "*.xlsx *.xlsm"), ("All files", "*.*")],
        )
        root.destroy()
    except Exception as err:
        raise RuntimeError(
            f"Could not open a file picker ({err}). "
            "Type the file paths in section 1 instead."
        )
    if not path:
        raise RuntimeError(f"No file chosen for: {title}")
    return path


if "YOUR_PATH" in reference_file:
    reference_file = pick_file("Select the REFERENCE DATA HIERARCHY file")
if "YOUR_PATH" in hc_file:
    hc_file = pick_file("Select the HC CURRENT MONTH file")
if "YOUR_PATH" in appian_file:
    appian_file = pick_file("Select the APPIAN file")

os.makedirs(output_folder, exist_ok=True)

timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")

output_reference_file = os.path.join(
    output_folder, f"Reference_Data_Hierarchy_CM_{timestamp}.xlsx"
)
exception_file = os.path.join(
    output_folder, f"Exception_Report_{timestamp}.xlsx"
)
hiring_update_file = os.path.join(
    output_folder, f"MR_Hiring_Update_{timestamp}.xlsx"
)
frc_reconciliation_file = os.path.join(
    output_folder, f"FRC_Reconciliation_Report_{timestamp}.xlsx"
)
run_log_file = os.path.join(
    output_folder, f"Run_Log_{timestamp}.txt"
)


# ------------------------------------------------------------
# RUN LOG
#
# Everything printed (counts, warnings, errors) is shown on screen
# AND written to Run_Log_<timestamp>.txt in the output folder.
# ------------------------------------------------------------

class TeeToLog:
    def __init__(self, stream, log_path):
        self.stream = stream
        self.log_path = log_path

    def write(self, text):
        self.stream.write(text)
        # Open/append/close each time so the log is complete
        # even if the run stops with an error
        with open(self.log_path, "a", encoding="utf-8") as f:
            f.write(text)

    def flush(self):
        self.stream.flush()

    def __getattr__(self, name):
        return getattr(self.stream, name)


def original_stream(stream):
    # Unwrap a log set up by an earlier run (e.g. re-running a notebook cell)
    while hasattr(stream, "log_path"):
        stream = stream.stream
    return stream


sys.stdout = TeeToLog(original_stream(sys.stdout), run_log_file)
sys.stderr = TeeToLog(original_stream(sys.stderr), run_log_file)

print(f"Run started: {datetime.now():%d-%b-%Y %H:%M:%S}")
print(f"Reference file : {reference_file}")
print(f"HC file        : {hc_file}")
print(f"Appian file    : {appian_file}")


# ============================================================
# 2. COLUMN DEFINITIONS
# ============================================================

KEY_COL = "Position Number"
EMPLOYEE_NAME_COL = "Employee Name"
OLD_EMP_RD_COL = "OLD_EMP_RD"
COMMENTS_RD_COL = "Comments_RD"
ONBOARDED_COMMENT = "Onboarded"
NEW_POSITION_FLAG_COL = "NEW_POSITION_FLAG"
# Other spellings of the same field (case, spaces, "_" and "-" are
# ignored when matching), so the existing column is always reused.
NEW_POSITION_FLAG_CANDIDATES = [
    NEW_POSITION_FLAG_COL, "NEW_POSITION_FLG", "NEW_POS_FLAG", "NEW_POS_FLG",
    "NEW_POSITION", "NEW_POSITION_YN",
]
JOB_SUMMARY_COL = "Job Summary"
EXPECTED_JOB_SUMMARY = "Financial insight and advisory support specialist"

FRC_COL = "FRC Code"
POSITION_ID_COL = "Position ID"

NUMERIC_COLUMNS = [
    "Position Number",
    "Position ID",
    "Existing Position ID",
    "Employee ID",
    "GCB Level",
    "Position Career Band",
    "Functional Manager Position Level Position Number",
    "Functional Manager Job Level Employee ID",
]

DATE_COLUMNS = [
    "EXIT DATE",
    "Exit Date",
    "JOINER MONTH",
    "Joiners Month",
    "LWD",
    "LWD_Pre",
    "LWD_Post",
    "Termination Date",
    "Terminate Date",
]

DEFAULT_JOB_SUMMARY_EXCLUSIONS = [
    "Operations Leadership",
    "Business Finance - Modelling",
    "Bus Finance Asst - Modelling",
    "Risk Analytics and Modelling",
]


# ============================================================
# 3. HELPER FUNCTIONS
# ============================================================

def clean_columns(df):
    df = df.copy()
    df.columns = [str(c).replace("\xa0", " ").strip() for c in df.columns]
    return df


def clean_compare_value(value):
    if pd.isna(value):
        return ""
    if isinstance(value, (pd.Timestamp, datetime)):
        return value.strftime("%Y-%m-%d")
    if isinstance(value, float) and value.is_integer():
        return int(value)
    if isinstance(value, str):
        return value.strip()
    return value


def values_differ(old_val, new_val):
    # Compare as text so 1234 (number) and "1234" (text) are equal
    return str(clean_compare_value(old_val)) != str(clean_compare_value(new_val))


def clean_position_id_series(series):
    def clean_one(value):
        if pd.isna(value):
            return np.nan
        value = str(value).strip()
        if value == "" or value.lower() in ["nan", "none", "<na>"]:
            return np.nan
        # Remove trailing .0 from Excel numeric IDs
        if re.fullmatch(r"\d+\.0", value):
            value = value[:-2]
        return value

    return series.apply(clean_one).astype(object)


def clean_frc_series(series):
    return (
        series.fillna("")
        .astype(str)
        .str.replace("\xa0", " ", regex=False)
        .str.strip()
        .str.upper()
    )


def clean_text_series(series, lower=False):
    s = series.fillna("").astype(str).str.strip()
    return s.str.lower() if lower else s.str.upper()


def normalise_text(value):
    if pd.isna(value):
        return ""
    return str(value).replace("\xa0", " ").strip().lower()


def is_blank_series(series):
    s = series.astype(str).str.strip().str.lower()
    return series.isna() | s.isin(["", "nan", "none", "nat", "<na>"])


def find_column(df, name):
    # Case-insensitive column match, e.g. "comments_rd" -> "Comments_RD"
    for col in df.columns:
        if str(col).strip().lower() == name.strip().lower():
            return col
    return None


def find_column_loose(df, candidates):
    # Like find_column, but also ignores spaces, "_" and "-",
    # e.g. "New Position Flag" matches "NEW_POSITION_FLAG"
    def key(name):
        return re.sub(r"[^a-z0-9]", "", str(name).lower())
    for name in candidates:
        for col in df.columns:
            if key(col) == key(name):
                return col
    return None


def is_vacant(value):
    return str(clean_compare_value(value)).strip().lower() == "vacant"


def split_blank_keys(df, key_col):
    # Rows with no key cannot be indexed/de-duplicated safely,
    # so they are set aside and added back afterwards.
    has_key = df[key_col].notna()
    return df[has_key].copy(), df[~has_key].copy()


def drop_duplicate_keys(df, key_col, label):
    dupes = df[key_col].duplicated(keep="first")
    if dupes.any():
        print(
            f"WARNING: {dupes.sum()} duplicate {key_col} rows in {label} "
            "- only the first occurrence is used."
        )
    return df[~dupes].copy()


def apply_data_types(df):
    df = df.copy()

    for col in NUMERIC_COLUMNS:
        if col not in df.columns:
            continue

        cleaned = (
            df[col]
            .astype(str)
            .str.strip()
            .str.replace(r"\.0$", "", regex=True)
        )
        numeric = pd.to_numeric(cleaned, errors="coerce")
        blank = is_blank_series(df[col])

        # Only convert when nothing would be lost
        # (e.g. an alphanumeric ID keeps the column as text)
        lost = numeric.isna() & ~blank
        if numeric.notna().any() and not lost.any():
            df[col] = numeric.astype("Int64")
        elif lost.any():
            # Mixed column: numeric IDs become numbers,
            # alphanumeric IDs are kept exactly as they are
            numeric_ok = numeric.notna() & (numeric % 1 == 0)
            df[col] = df[col].astype(object)
            df.loc[numeric_ok, col] = numeric[numeric_ok].astype("int64")
            df.loc[blank, col] = None
            print(
                f"Note: '{col}' has {lost.sum()} non-numeric value(s) "
                "- those are kept as text."
            )

    for col in DATE_COLUMNS:
        if col not in df.columns:
            continue

        try:
            dates = pd.to_datetime(df[col], errors="coerce", format="mixed")
        except (TypeError, ValueError):
            dates = pd.to_datetime(df[col], errors="coerce")

        blank = is_blank_series(df[col])
        lost = dates.isna() & ~blank
        if not lost.any():
            df[col] = dates
        else:
            print(
                f"Note: '{col}' kept as text - {lost.sum()} "
                "value(s) could not be read as dates."
            )

    return df


def format_excel(file_path):
    if not os.path.exists(file_path):
        return

    wb = load_workbook(file_path)

    header_fill = PatternFill(fill_type="solid", fgColor="DB0011")
    header_font = Font(color="FFFFFF", bold=True)

    for ws in wb.worksheets:
        if ws.max_row < 1:
            continue

        # Header formatting
        for cell in ws[1]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center", vertical="center")

        ws.freeze_panes = "A2"

        if ws.max_column > 0:
            ws.auto_filter.ref = ws.dimensions

        # Auto width
        for col_idx in range(1, ws.max_column + 1):
            max_length = 0
            for row_idx in range(1, min(ws.max_row, 5000) + 1):
                value = ws.cell(row=row_idx, column=col_idx).value
                if value is not None:
                    max_length = max(max_length, len(str(value)))
            width = min(max(max_length + 2, 10), 50)
            ws.column_dimensions[get_column_letter(col_idx)].width = width

        # Numeric/date formatting
        header_map = {
            ws.cell(1, col).value: col for col in range(1, ws.max_column + 1)
        }

        for header in NUMERIC_COLUMNS:
            if header in header_map:
                col_idx = header_map[header]
                for row_idx in range(2, ws.max_row + 1):
                    cell = ws.cell(row=row_idx, column=col_idx)
                    if isinstance(cell.value, (int, float)):
                        # IDs: no thousands separator
                        cell.number_format = "0"

        for header in DATE_COLUMNS:
            if header in header_map:
                col_idx = header_map[header]
                for row_idx in range(2, ws.max_row + 1):
                    cell = ws.cell(row=row_idx, column=col_idx)
                    if isinstance(cell.value, (datetime, pd.Timestamp)):
                        cell.number_format = "dd-mmm-yyyy"

    wb.save(file_path)


# ============================================================
# 4. CONFIG HELPERS
# ============================================================

def get_config_list(config_df, candidate_columns):
    for col in candidate_columns:
        if col in config_df.columns:
            values = []
            for value in config_df[col]:
                if pd.notna(value):
                    value = str(value).strip()
                    if value:
                        values.append(value)
            if values:
                return values
    return []


def get_export_column_order(config_df):
    return get_config_list(
        config_df, ["Fields", "Field", "Column", "Column Name"]
    )


def apply_column_sequence(df, requested_order):
    if not requested_order:
        return df
    first = [c for c in requested_order if c in df.columns]
    remaining = [c for c in df.columns if c not in first]
    return df[first + remaining]


def get_appian_config(appian_config_df):
    config = clean_columns(appian_config_df)
    field_col = config.columns[0]

    fields = []
    primary_key = None

    for _, row in config.iterrows():
        field = row.get(field_col)
        if pd.isna(field):
            continue
        field = str(field).strip()
        if not field:
            continue

        fields.append(field)

        # Second column contains Primary_Key marker
        if len(config.columns) > 1:
            marker = row.get(config.columns[1])
            if pd.notna(marker) and str(marker).strip().lower() == "primary_key":
                primary_key = field

    if primary_key is None:
        primary_key = "Position ID"

    return primary_key, fields


def get_job_summary_exclusions(config_df):
    # Add a "Job Summary Exclusions" column on the Config sheet to
    # maintain this list without editing the script.
    exclusions = get_config_list(
        config_df,
        ["Job Summary Exclusions", "Job Summary Exclusion", "JOB_SUMMARY_EXCLUSIONS"],
    ) or DEFAULT_JOB_SUMMARY_EXCLUSIONS

    return {normalise_text(x) for x in exclusions}


def get_hiring_update_columns(config_df):
    # If not available, all columns are used later.
    return get_config_list(
        config_df, ["Hiring Update", "MR Hiring Update", "Hiring_Update"]
    )


# ============================================================
# 5. READ INPUT FILES
# ============================================================

reference_xls = pd.ExcelFile(reference_file)
reference_sheet = reference_xls.sheet_names[0]

for candidate in ["Reference_Data", "Reference Data", "Reference_Data_PM"]:
    if candidate in reference_xls.sheet_names:
        reference_sheet = candidate
        break

ref_df = pd.read_excel(reference_file, sheet_name=reference_sheet)
config_df = pd.read_excel(reference_file, sheet_name="Config")
appian_config_df = pd.read_excel(reference_file, sheet_name="APPIAN_CONFIG")
new_migrations_df = pd.read_excel(reference_file, sheet_name="New_Migrations")

hc_df = pd.read_excel(hc_file)
appian_df = pd.read_excel(appian_file)

ref_df = clean_columns(ref_df)
config_df = clean_columns(config_df)
appian_config_df = clean_columns(appian_config_df)
new_migrations_df = clean_columns(new_migrations_df)
hc_df = clean_columns(hc_df)
appian_df = clean_columns(appian_df)


# ============================================================
# 6. CLEAN PRIMARY KEYS
# ============================================================

ref_df[KEY_COL] = clean_position_id_series(ref_df[KEY_COL])
hc_df[KEY_COL] = clean_position_id_series(hc_df[KEY_COL])

hc_blank_keys = hc_df[KEY_COL].isna().sum()
if hc_blank_keys:
    print(
        f"WARNING: {hc_blank_keys} HC row(s) have a blank {KEY_COL} "
        "and cannot be matched."
    )


# ============================================================
# 7. IDENTIFY REMOVED / COMMON / NEW POSITIONS
# ============================================================

ref_keys = set(ref_df[KEY_COL].dropna())
hc_keys = set(hc_df[KEY_COL].dropna())

removed_keys = ref_keys - hc_keys
new_keys = hc_keys - ref_keys
common_keys = ref_keys & hc_keys

positions_removed_df = ref_df[ref_df[KEY_COL].isin(removed_keys)].copy()


# ============================================================
# 8. NEW POSITIONS ADDED FROM CM
# ============================================================

# NEW_POSITION_FLAG: "YES" for positions that did not exist last month.
# The EXISTING flag column in the Reference is used (no new column is
# ever created). Flags already present (on existing rows, or a value
# already on the new row) are kept, never overwritten.
updated_ref_df = ref_df.copy()
new_flag_col = find_column_loose(updated_ref_df, NEW_POSITION_FLAG_CANDIDATES)

new_positions_df = hc_df[hc_df[KEY_COL].isin(new_keys)].copy()
new_positions_df = drop_duplicate_keys(new_positions_df, KEY_COL, "HC new positions")

if new_flag_col is None:
    print(f"WARNING: no '{NEW_POSITION_FLAG_COL}' column in the Reference - "
          "new positions not flagged (no column added).")
else:
    print(f"New position flag column: '{new_flag_col}'")
    # If HC also carries the flag (any spelling), line it up with the Reference column
    hc_flag_col = find_column_loose(new_positions_df, NEW_POSITION_FLAG_CANDIDATES)
    if hc_flag_col is not None and hc_flag_col != new_flag_col:
        new_positions_df = new_positions_df.drop(columns=[new_flag_col], errors="ignore")
        new_positions_df = new_positions_df.rename(columns={hc_flag_col: new_flag_col})
    if new_flag_col not in new_positions_df.columns:
        new_positions_df[new_flag_col] = ""
    new_positions_df[new_flag_col] = new_positions_df[new_flag_col].astype(object)
    flag_blank = is_blank_series(new_positions_df[new_flag_col])
    new_positions_df.loc[flag_blank, new_flag_col] = "YES"
    print(f"New positions this month: {len(new_positions_df)} "
          f"({int(flag_blank.sum())} flagged YES, {int((~flag_blank).sum())} kept existing flag)")

for col in updated_ref_df.columns:
    if col not in new_positions_df.columns:
        new_positions_df[col] = ""

new_positions_df = new_positions_df[updated_ref_df.columns]

updated_ref_df = pd.concat([updated_ref_df, new_positions_df], ignore_index=True)

# Object dtype so cells can take any value during the updates below
updated_ref_df = updated_ref_df.astype(object)

# Rows with a blank Position Number are set aside and added back
# after the HC and Appian updates (sections 9 and 9A.2)
updated_ref_df, ref_blank_key_rows = split_blank_keys(updated_ref_df, KEY_COL)

if len(ref_blank_key_rows):
    print(
        f"Note: {len(ref_blank_key_rows)} Reference row(s) have a blank "
        f"{KEY_COL} - kept as-is, not updated from HC/Appian."
    )


# ============================================================
# 9. UPDATE CHANGED FIELDS FROM HEADCOUNT AND AUDIT
# ============================================================

other_changes = []
onboarded_count = 0
filled_keys = set()  # vacant last month, filled this month

# Use the existing comments column if present (any case), else create it
comments_col = find_column(updated_ref_df, COMMENTS_RD_COL)
if comments_col is None:
    comments_col = COMMENTS_RD_COL
    updated_ref_df[comments_col] = ""
    print(f"Note: '{COMMENTS_RD_COL}' not found in Reference - column added.")

updated_ref_df = drop_duplicate_keys(updated_ref_df, KEY_COL, "Reference")
hc_unique_df = drop_duplicate_keys(hc_df.dropna(subset=[KEY_COL]), KEY_COL, "HC")

updated_ref_df = updated_ref_df.set_index(KEY_COL)
hc_indexed = hc_unique_df.set_index(KEY_COL)

common_columns = [
    c for c in updated_ref_df.columns
    if c in hc_indexed.columns and c != new_flag_col  # existing flags are never overwritten
]

for pos in common_keys:
    if pos not in updated_ref_df.index or pos not in hc_indexed.index:
        continue

    old_emp_rd = ""
    if OLD_EMP_RD_COL in updated_ref_df.columns:
        old_emp_rd = clean_compare_value(updated_ref_df.at[pos, OLD_EMP_RD_COL])

    change_record = {
        KEY_COL: clean_compare_value(pos),
        OLD_EMP_RD_COL: old_emp_rd,
    }
    has_change = False

    # Last month's employee, read before HC overwrites it
    prev_employee = ""
    if EMPLOYEE_NAME_COL in updated_ref_df.columns:
        prev_employee = updated_ref_df.at[pos, EMPLOYEE_NAME_COL]

    for col in common_columns:
        old_val = updated_ref_df.at[pos, col]
        new_val = hc_indexed.at[pos, col]

        if values_differ(old_val, new_val):
            change_record[f"{col}_Pre"] = clean_compare_value(old_val)
            change_record[f"{col}_Post"] = clean_compare_value(new_val)
            updated_ref_df.at[pos, col] = clean_compare_value(new_val)
            has_change = True

    # --------------------------------------------------------
    # OLD_EMP_RD
    #
    # If current employee is VACANT: retain previous OLD_EMP_RD.
    # Otherwise: update OLD_EMP_RD with current Employee Name.
    # --------------------------------------------------------
    if (
        OLD_EMP_RD_COL in updated_ref_df.columns
        and EMPLOYEE_NAME_COL in hc_indexed.columns
    ):
        current_employee = clean_compare_value(
            hc_indexed.at[pos, EMPLOYEE_NAME_COL]
        )
        if str(current_employee).strip().lower() != "vacant":
            updated_ref_df.at[pos, OLD_EMP_RD_COL] = current_employee

    # --------------------------------------------------------
    # Comments_RD
    #
    # Vacant last month and occupied this month -> "Onboarded"
    # --------------------------------------------------------
    if EMPLOYEE_NAME_COL in hc_indexed.columns:
        current_employee = hc_indexed.at[pos, EMPLOYEE_NAME_COL]

        if (
            is_vacant(prev_employee)
            and not is_vacant(current_employee)
            and clean_compare_value(current_employee) != ""
        ):
            change_record[f"{comments_col}_Pre"] = clean_compare_value(
                updated_ref_df.at[pos, comments_col]
            )
            change_record[f"{comments_col}_Post"] = ONBOARDED_COMMENT
            updated_ref_df.at[pos, comments_col] = ONBOARDED_COMMENT
            onboarded_count += 1
            filled_keys.add(pos)
            has_change = True

    if has_change:
        other_changes.append(change_record)

updated_ref_df = updated_ref_df.reset_index()

other_changes_df = pd.DataFrame(other_changes)

print(f"Positions marked '{ONBOARDED_COMMENT}' (vacant -> occupied): {onboarded_count}")


# ============================================================
# 9A.0 UPDATE APPIAN POSITION ID FROM NEW_MIGRATIONS
#
# ONLY IF POSITION ID IS BLANK
# ============================================================

new_migrations_df[FRC_COL] = clean_frc_series(new_migrations_df[FRC_COL])
appian_df[FRC_COL] = clean_frc_series(appian_df[FRC_COL])

appian_df[POSITION_ID_COL] = clean_position_id_series(appian_df[POSITION_ID_COL])
new_migrations_df[POSITION_ID_COL] = clean_position_id_series(
    new_migrations_df[POSITION_ID_COL]
)

# Lookup table: where duplicate FRC Codes exist in New_Migrations,
# the first Position ID is used.
migration_lookup = (
    new_migrations_df
    .dropna(subset=[POSITION_ID_COL])
    .drop_duplicates(subset=[FRC_COL], keep="first")
    .set_index(FRC_COL)[POSITION_ID_COL]
)

blank_position = appian_df[POSITION_ID_COL].isna()
before_blank = blank_position.sum()

appian_df.loc[blank_position, POSITION_ID_COL] = (
    appian_df.loc[blank_position, FRC_COL].map(migration_lookup)
)

after_blank = appian_df[POSITION_ID_COL].isna().sum()

print(f"Position ID blanks before update : {before_blank}")
print(f"Position ID blanks after update  : {after_blank}")
print(f"Rows updated from New_Migrations : {before_blank - after_blank}")

matches = appian_df[FRC_COL].isin(new_migrations_df[FRC_COL]).sum()
print(f"Matching FRC Codes found: {matches}")


# ============================================================
# 9A.1 APPIAN DERIVED LOOKUP KEY
#
# If Position ID is blank AND Role Type contains Replacement,
# use Existing Position ID
# ============================================================

APPIAN_KEY_COL = "Position ID"
APPIAN_EXISTING_POS_COL = "Existing Position ID"
ROLE_TYPE_COL = "Role Type"
APPIAN_LOOKUP_KEY_COL = "APPIAN_LOOKUP_KEY"

appian_df[APPIAN_KEY_COL] = clean_position_id_series(appian_df[APPIAN_KEY_COL])
appian_df[APPIAN_EXISTING_POS_COL] = clean_position_id_series(
    appian_df[APPIAN_EXISTING_POS_COL]
)

role_type_clean = clean_text_series(appian_df[ROLE_TYPE_COL], lower=True)

appian_df[APPIAN_LOOKUP_KEY_COL] = appian_df[APPIAN_KEY_COL]

replacement_mask = (
    appian_df[APPIAN_KEY_COL].isna()
    & role_type_clean.str.contains("replacement", na=False)
)
appian_df.loc[replacement_mask, APPIAN_LOOKUP_KEY_COL] = (
    appian_df.loc[replacement_mask, APPIAN_EXISTING_POS_COL]
)


# ============================================================
# 9A.2 UPDATE REFERENCE FROM APPIAN
#
# A separate lookup table is used so that appian_df keeps every
# row for the FRC reconciliation in section 12A.
# ============================================================

_, appian_fields = get_appian_config(appian_config_df)

appian_lookup_df = drop_duplicate_keys(
    appian_df.dropna(subset=[APPIAN_LOOKUP_KEY_COL]),
    APPIAN_LOOKUP_KEY_COL,
    "Appian",
)

updated_ref_df = updated_ref_df.set_index(KEY_COL)
appian_indexed = appian_lookup_df.set_index(APPIAN_LOOKUP_KEY_COL)

appian_common_keys = (
    set(updated_ref_df.index.dropna()) & set(appian_indexed.index.dropna())
)

print(f"Appian Position IDs matched to Reference: {len(appian_common_keys)}")

appian_update_cols = [
    c for c in appian_fields
    if c in updated_ref_df.columns and c in appian_indexed.columns
]

for pos in appian_common_keys:
    for col in appian_update_cols:
        old_val = updated_ref_df.at[pos, col]
        new_val = appian_indexed.at[pos, col]

        if values_differ(old_val, new_val):
            updated_ref_df.at[pos, col] = clean_compare_value(new_val)

updated_ref_df = updated_ref_df.reset_index()

# Add back rows with a blank Position Number
updated_ref_df = pd.concat(
    [updated_ref_df, ref_blank_key_rows], ignore_index=True
)

if DROP_REMOVED_POSITIONS:
    updated_ref_df = updated_ref_df[
        ~updated_ref_df[KEY_COL].isin(removed_keys)
    ].reset_index(drop=True)
    print(f"Removed positions dropped from Reference: {len(removed_keys)}")


# ============================================================
# 9B. DERIVE ROLE MOVEMENT
# ============================================================

ROLE_MOVEMENT_COL = "Role Movement"

if FRC_COL in updated_ref_df.columns and ROLE_MOVEMENT_COL in updated_ref_df.columns:
    frc = clean_frc_series(updated_ref_df[FRC_COL])
    role_movement = updated_ref_df[ROLE_MOVEMENT_COL].fillna("").astype(str).str.strip()

    blank_role = role_movement.eq("")

    updated_ref_df.loc[
        blank_role & frc.str.startswith("NEW/BTO"), ROLE_MOVEMENT_COL
    ] = "NEW BTO ROLE"

    updated_ref_df.loc[
        blank_role & frc.str.startswith("NEW/MIGRATIONS"), ROLE_MOVEMENT_COL
    ] = "NEW MIGRATION ROLE"
else:
    print("Role Movement logic skipped because required columns are missing.")


# ============================================================
# 9C. FRC ENDING OLD -> OPCO STATUS COMPLETED
# ============================================================

OPCO_STATUS_COL = "OPCO Status"

if FRC_COL in updated_ref_df.columns and OPCO_STATUS_COL in updated_ref_df.columns:
    frc_clean = clean_frc_series(updated_ref_df[FRC_COL])
    updated_ref_df.loc[frc_clean.str.endswith("_OLD"), OPCO_STATUS_COL] = "COMPLETED"
else:
    print(
        "FRC OLD to OPCO Status logic skipped "
        "because required columns are missing."
    )


# ============================================================
# 9D. DERIVE STATUS_APPROVAL
# ============================================================

SUB_FUNCTION_HEAD_COL = "Sub Function Head Status"
COO_STATUS_COL = "COO Status"
STATUS_APPROVAL_COL = "STATUS_APPROVAL"

if all(
    c in updated_ref_df.columns
    for c in [SUB_FUNCTION_HEAD_COL, COO_STATUS_COL, OPCO_STATUS_COL, STATUS_APPROVAL_COL]
):
    sub_function = clean_text_series(updated_ref_df[SUB_FUNCTION_HEAD_COL], lower=True)
    coo = clean_text_series(updated_ref_df[COO_STATUS_COL], lower=True)
    opco = clean_text_series(updated_ref_df[OPCO_STATUS_COL], lower=True)

    updated_ref_df[STATUS_APPROVAL_COL] = np.select(
        [
            sub_function.isin(["not started", "in progress"]),
            coo.eq("not started"),
            coo.eq("in progress"),
            coo.eq("completed") & opco.eq("completed"),
            coo.eq("completed") & opco.isin(["", "in progress", "not started"]),
        ],
        [
            "AWAITING SUB FUNCTION HEAD APPROVAL",
            "AWAITING COO-BE-LEAD APPROVAL",
            "AWAITING COO-BE-LEAD APPROVAL",
            "OPCO APPROVED",
            "AWAITING APPROVAL FROM OPCO",
        ],
        default="",
    )
else:
    print("STATUS_APPROVAL logic skipped because required columns are missing.")


# ============================================================
# 9D.1 Comments_RD FROM STATUS_APPROVAL
#
# STATUS_APPROVAL = AWAITING COO-BE-LEAD APPROVAL
#                or AWAITING APPROVAL FROM OPCO
#   -> Comments_RD = AWAITING FRC APPROVAL
# ============================================================

AWAITING_FRC_STATUSES = [
    "AWAITING COO-BE-LEAD APPROVAL",
    "AWAITING APPROVAL FROM OPCO",
]
AWAITING_FRC_COMMENT = "AWAITING FRC APPROVAL"

if STATUS_APPROVAL_COL in updated_ref_df.columns:
    awaiting_frc = clean_text_series(updated_ref_df[STATUS_APPROVAL_COL]).isin(
        AWAITING_FRC_STATUSES
    )
    updated_ref_df.loc[awaiting_frc, comments_col] = AWAITING_FRC_COMMENT
    print(
        f"Positions marked '{AWAITING_FRC_COMMENT}' in {comments_col}: "
        f"{awaiting_frc.sum()}"
    )

    # --------------------------------------------------------
    # STATUS_APPROVAL = OPCO APPROVED and Comments_RD blank
    #   -> Comments_RD = RAISE IJP
    # Existing comments are left as they are.
    # --------------------------------------------------------
    opco_approved = clean_text_series(updated_ref_df[STATUS_APPROVAL_COL]).eq(
        "OPCO APPROVED"
    )
    comments_blank = is_blank_series(updated_ref_df[comments_col])
    raise_ijp = opco_approved & comments_blank

    updated_ref_df.loc[raise_ijp, comments_col] = "RAISE IJP"
    print(f"Positions marked 'RAISE IJP' in {comments_col}: {raise_ijp.sum()}")
else:
    print(f"{comments_col} from STATUS_APPROVAL skipped - column missing.")


# ============================================================
# 9D.2 FILLED POSITIONS -> OFFER / CANDIDATE / JOINER FIELDS BLANKED
#
# Vacant in the PM file AND occupied in the current month (same check
# as "Onboarded"):
#   Committed Offer, Candidate Name, Joiner flag, Joiner Month and
#   Comments_RD are made blank.
# Runs after the Comments_RD rules (so they end up blank) and before
# HIRING_FLG is derived from the joiner flag (9E).
# Column names are matched ignoring case, spaces, "_" and "-".
# ============================================================

FILLED_CLEAR_FIELDS = {
    "Committed Offer": ["COMMITTED OFFER", "COMMITTED_OFFER_FLAG", "COMMITTED OFFER FLG"],
    "Candidate Name": ["CANDIDATE NAME", "CANDIDATE"],
    "Joiner flag": ["JOINER FLAG", "JOINER_FLG", "JOINERS FLAG", "JOINERS_FLG", "JOINERS", "JOINER"],
    "Joiner Month": ["JOINER MONTH", "JOINERS MONTH", "JOINING MONTH"],
    "Comments_RD": [comments_col],
}

filled_cleared = []
filled_mask = updated_ref_df[KEY_COL].isin(filled_keys)

for label, candidates in FILLED_CLEAR_FIELDS.items():
    col = find_column_loose(updated_ref_df, candidates)
    if col is None:
        print(f"Note: '{label}' column not found - not cleared for filled positions.")
        continue

    had_value = filled_mask & ~is_blank_series(updated_ref_df[col])
    for idx in updated_ref_df.index[had_value]:
        filled_cleared.append({
            KEY_COL: updated_ref_df.at[idx, KEY_COL],
            EMPLOYEE_NAME_COL: (updated_ref_df.at[idx, EMPLOYEE_NAME_COL]
                                if EMPLOYEE_NAME_COL in updated_ref_df.columns else ""),
            "Field": col,
            "Pre": clean_compare_value(updated_ref_df.at[idx, col]),
            "Post": "",
        })
    updated_ref_df.loc[filled_mask, col] = None
    print(f"'{col}' blanked for filled positions: {int(had_value.sum())} value(s) cleared")

filled_cleared_df = pd.DataFrame(
    filled_cleared, columns=[KEY_COL, EMPLOYEE_NAME_COL, "Field", "Pre", "Post"]
)


# ============================================================
# 9E. UPDATE HIRING_FLG
#
# YES if:
#   HIRING AGAINST PHYSICAL_RD = YES
#   OR JOINERS = YES
# ============================================================

HIRING_FLG_COL = "HIRING_FLG"
HIRING_AGAINST_COL = "HIRING AGAINST PHYSICAL_RD"
JOINERS_COL = "JOINERS"

# ------------------------------------------------------------
# Vacant in the PM file AND flag = YES in the PM file AND filled
# in the current month (same check as "Onboarded") -> flag removed.
# HIRING AGAINST PHYSICAL_RD is cleared here, before HIRING_FLG is
# derived from it; HIRING_FLG itself is cleared in 9E.1.
# ------------------------------------------------------------
flags_removed = []


def remove_flag_for_filled(col):
    pm_col = find_column(ref_df, col)
    cur_col = find_column(updated_ref_df, col)
    if pm_col is None or cur_col is None:
        print(f"{col} removal for filled positions skipped - column missing.")
        return

    pm_yes = set(ref_df.loc[clean_text_series(ref_df[pm_col]).eq("YES"), KEY_COL].dropna())
    mask = updated_ref_df[KEY_COL].isin(filled_keys & pm_yes)

    names = (
        updated_ref_df.loc[mask, EMPLOYEE_NAME_COL]
        if EMPLOYEE_NAME_COL in updated_ref_df.columns
        else pd.Series("", index=updated_ref_df.index[mask])
    )
    for pos, name in zip(updated_ref_df.loc[mask, KEY_COL], names):
        flags_removed.append({
            KEY_COL: pos, EMPLOYEE_NAME_COL: name,
            "Field": cur_col, "Pre": "YES", "Post": "",
        })

    updated_ref_df.loc[mask, cur_col] = None
    print(f"{cur_col} removed (vacant last month, filled this month): {int(mask.sum())}")


remove_flag_for_filled(HIRING_AGAINST_COL)

if HIRING_FLG_COL in updated_ref_df.columns:
    blank = pd.Series("", index=updated_ref_df.index)

    hiring_against = (
        clean_text_series(updated_ref_df[HIRING_AGAINST_COL])
        if HIRING_AGAINST_COL in updated_ref_df.columns
        else blank
    )
    joiners = (
        clean_text_series(updated_ref_df[JOINERS_COL])
        if JOINERS_COL in updated_ref_df.columns
        else blank
    )

    if RESET_HIRING_FLG:
        updated_ref_df[HIRING_FLG_COL] = ""

    updated_ref_df.loc[
        hiring_against.eq("YES") | joiners.eq("YES"), HIRING_FLG_COL
    ] = "YES"
else:
    print(f"{HIRING_FLG_COL} not found. Skipping HIRING_FLG update.")


# ============================================================
# 9E.1 FILLED POSITIONS -> HIRING_FLG REMOVED
#
# Same rule as HIRING AGAINST PHYSICAL_RD above. Runs after 9E so
# the flag is not set back to YES.
# ============================================================

remove_flag_for_filled(HIRING_FLG_COL)

hiring_flag_removed_df = pd.DataFrame(
    flags_removed, columns=[KEY_COL, EMPLOYEE_NAME_COL, "Field", "Pre", "Post"]
)


# ============================================================
# 9F. STATUS = VACANT -> LEAVER CLEARED
#
# Column names are matched ignoring case (e.g. "Status", "STATUS").
# ============================================================

STATUS_COL = "Status"
LEAVER_COL = "Leaver"

status_col = find_column(updated_ref_df, STATUS_COL)
leaver_col = find_column(updated_ref_df, LEAVER_COL)

if status_col is not None and leaver_col is not None:
    vacant_status = clean_text_series(updated_ref_df[status_col]).eq("VACANT")
    updated_ref_df.loc[vacant_status, leaver_col] = None
    print(f"{leaver_col} cleared for STATUS = VACANT: {vacant_status.sum()}")
else:
    missing = [n for n, c in [(STATUS_COL, status_col), (LEAVER_COL, leaver_col)] if c is None]
    print(f"Leaver clearing skipped - column(s) missing: {', '.join(missing)}")


# ============================================================
# 9G. LWD FROM MR_LEAVERS_CM (Termination Date)
#
# The CURRENT MONTH Employee ID (after the HC update) is compared
# with Employee ID in the leavers file; on a match, LWD is set to the
# leaver's Termination Date (latest date if listed more than once).
#
# Employees who have already exited (position Vacant / Employee ID
# blank in the current month) are not matched, so no Termination
# Date is added for them.
#
# Only records with Employee Class = "Employee" are compared (taken
# from the leavers file; if it has no Employee Class column, the
# Reference Data's Employee Class is used instead).
# ============================================================

LWD_COL = "LWD"
TERMINATION_DATE_CANDIDATES = [
    "Termination Date", "Terminate Date", "Terminated Date", "Date of Termination",
]
EMPLOYEE_ID_CANDIDATES = ["Employee ID", "EmployeeID", "Emp ID", "Employee Number"]
EMPLOYEE_CLASS_COL = "Employee Class"
EMPLOYEE_CLASS_VALUE = "EMPLOYEE"

lwd_updates_df = pd.DataFrame(columns=[KEY_COL, "Employee ID", f"{LWD_COL}_Pre", f"{LWD_COL}_Post"])
leavers_unmatched_df = pd.DataFrame()


def first_column(df, candidates):
    for name in candidates:
        col = find_column(df, name)
        if col is not None:
            return col
    return None


if not os.path.exists(leavers_file):
    print(f"LWD update skipped - leavers file not found: {leavers_file}")
else:
    leavers_df = clean_columns(pd.read_excel(leavers_file))

    term_col = first_column(leavers_df, TERMINATION_DATE_CANDIDATES)
    lv_emp_col = first_column(leavers_df, EMPLOYEE_ID_CANDIDATES)
    cur_emp_col = first_column(updated_ref_df, EMPLOYEE_ID_CANDIDATES)

    if term_col is None:
        print("LWD update skipped - no 'Termination Date' column in the leavers file.")
    elif lv_emp_col is None:
        print("LWD update skipped - no 'Employee ID' column in the leavers file.")
    elif cur_emp_col is None:
        print("LWD update skipped - no 'Employee ID' column in the Reference Data Hierarchy.")
    else:
        lwd_col = find_column(updated_ref_df, LWD_COL)
        if lwd_col is None:
            lwd_col = LWD_COL
            updated_ref_df[lwd_col] = None
            print(f"Note: '{LWD_COL}' not found in Reference - column added.")

        # Only Employee Class = Employee
        lv_class_col = find_column(leavers_df, EMPLOYEE_CLASS_COL)
        ref_class_col = find_column(updated_ref_df, EMPLOYEE_CLASS_COL)
        if lv_class_col is not None:
            is_employee = clean_text_series(leavers_df[lv_class_col]).eq(EMPLOYEE_CLASS_VALUE)
            print(f"Leavers with Employee Class = Employee: {int(is_employee.sum())} "
                  f"of {len(leavers_df)} (others ignored for LWD)")
            leavers_df = leavers_df[is_employee].copy()
        elif ref_class_col is None:
            print(f"Note: no '{EMPLOYEE_CLASS_COL}' column in the leavers file or the "
                  "Reference - all leavers compared.")

        # Leaver Employee ID -> latest Termination Date
        leavers_df["_EMP"] = clean_position_id_series(leavers_df[lv_emp_col])
        leavers_df["_TERM_DATE"] = pd.to_datetime(leavers_df[term_col], errors="coerce")
        leaver_dates = (
            leavers_df.dropna(subset=["_EMP", "_TERM_DATE"])
            .sort_values("_TERM_DATE")
            .drop_duplicates("_EMP", keep="last")
            .set_index("_EMP")["_TERM_DATE"]
        )

        # Current month Employee ID of each Reference row. Blank if Vacant,
        # and ignored for positions not in this month's HC (removed positions
        # still carry last month's Employee ID).
        ref_emp = clean_position_id_series(updated_ref_df[cur_emp_col])
        ref_emp[~updated_ref_df[KEY_COL].isin(hc_keys)] = np.nan
        if lv_class_col is None and ref_class_col is not None:
            ref_is_employee = clean_text_series(updated_ref_df[ref_class_col]).eq(EMPLOYEE_CLASS_VALUE)
            ref_emp[~ref_is_employee] = np.nan
            print(f"Reference rows with Employee Class = Employee: {int(ref_is_employee.sum())}")

        hit = ref_emp.isin(leaver_dates.index)
        new_lwd = ref_emp[hit].map(leaver_dates)
        old_lwd = updated_ref_df.loc[hit, lwd_col]
        changed = [values_differ(o, n) for o, n in zip(old_lwd, new_lwd)]

        lwd_updates_df = pd.DataFrame({
            KEY_COL: updated_ref_df.loc[hit, KEY_COL][changed].values,
            "Employee ID": ref_emp[hit][changed].values,
            f"{LWD_COL}_Pre": [clean_compare_value(v) for v in old_lwd[changed]],
            f"{LWD_COL}_Post": new_lwd[changed].values,
        })
        updated_ref_df.loc[hit, lwd_col] = new_lwd

        not_matched = (
            ~leavers_df["_EMP"].isin(set(ref_emp.dropna()))
            | leavers_df["_TERM_DATE"].isna()
        )
        leavers_unmatched_df = leavers_df[not_matched].drop(columns=["_EMP", "_TERM_DATE"])

        print(f"Leavers in file                  : {len(leavers_df)}")
        print(f"Reference rows matched on Emp ID : {int(hit.sum())}")
        print(f"LWD values changed               : {len(lwd_updates_df)}")
        if len(leavers_unmatched_df):
            print(f"Note: {len(leavers_unmatched_df)} leaver row(s) not in the current month "
                  "(already exited) or without a Termination Date - see 'Leavers not matched' sheet.")


# ============================================================
# 10. COLUMN SEQUENCING FROM CONFIG
# ============================================================

final_column_order = get_export_column_order(config_df)
updated_ref_df = apply_column_sequence(updated_ref_df, final_column_order)

positions_removed_df = positions_removed_df.reindex(
    columns=updated_ref_df.columns, fill_value=""
)
new_positions_df = new_positions_df.reindex(
    columns=updated_ref_df.columns, fill_value=""
)

if not other_changes_df.empty:
    ordered_cols = [
        c for c in [KEY_COL, OLD_EMP_RD_COL] if c in other_changes_df.columns
    ]
    remaining_cols = [c for c in other_changes_df.columns if c not in ordered_cols]
    other_changes_df = other_changes_df[ordered_cols + remaining_cols]
else:
    other_changes_df = pd.DataFrame(columns=[KEY_COL, OLD_EMP_RD_COL])


# ============================================================
# 11. JOB SUMMARY EXCEPTIONS
# ============================================================

job_summary_exclusions = get_job_summary_exclusions(config_df)
expected_job_value = normalise_text(EXPECTED_JOB_SUMMARY)

if JOB_SUMMARY_COL in updated_ref_df.columns:
    job_summary_clean = updated_ref_df[JOB_SUMMARY_COL].apply(normalise_text)
    job_summary_exceptions_df = updated_ref_df[
        (job_summary_clean != expected_job_value)
        & (~job_summary_clean.isin(job_summary_exclusions))
    ].copy()
else:
    job_summary_exceptions_df = pd.DataFrame()


# ============================================================
# 12. DATA TYPE HANDLING
# ============================================================

updated_ref_df = apply_data_types(updated_ref_df)
positions_removed_df = apply_data_types(positions_removed_df)
new_positions_df = apply_data_types(new_positions_df)
other_changes_df = apply_data_types(other_changes_df)
job_summary_exceptions_df = apply_data_types(job_summary_exceptions_df)


# ============================================================
# 12A. FRC RECONCILIATION REPORT
#
# APPIAN FRC vs REFERENCE FRC
# Uses cleaned copies so the exported Reference FRC is unchanged.
# Uses every Appian row, including those without a Position ID.
# ============================================================

if FRC_COL in updated_ref_df.columns and FRC_COL in appian_df.columns:
    ref_frc = clean_frc_series(updated_ref_df[FRC_COL])
    appian_frc = clean_frc_series(appian_df[FRC_COL])

    appian_frc_set = set(appian_frc) - {""}
    ref_frc_set = set(ref_frc) - {""}

    appian_not_in_ref = appian_df[
        appian_frc.isin(appian_frc_set - ref_frc_set)
    ].copy()

    ref_not_in_appian = updated_ref_df[
        ref_frc.isin(ref_frc_set - appian_frc_set)
    ].copy()

    control_summary_df = pd.DataFrame({
        "Metric": [
            "Total FRC in Appian",
            "Total FRC in Reference",
            "Matching FRC",
            "Appian FRC Not in Reference",
            "Reference FRC Not in Appian",
        ],
        "Count": [
            len(appian_frc_set),
            len(ref_frc_set),
            len(appian_frc_set & ref_frc_set),
            len(appian_frc_set - ref_frc_set),
            len(ref_frc_set - appian_frc_set),
        ],
    })

    with pd.ExcelWriter(frc_reconciliation_file, engine="openpyxl") as writer:
        control_summary_df.to_excel(writer, index=False, sheet_name="Control Summary")
        appian_not_in_ref.to_excel(writer, index=False, sheet_name="Appian_FRC_Not_in_Ref")
        ref_not_in_appian.to_excel(writer, index=False, sheet_name="Ref_FRC_Not_in_Appian")

    format_excel(frc_reconciliation_file)
else:
    print("FRC reconciliation skipped - FRC Code column missing.")


# ============================================================
# 12B. CREATE MR HIRING UPDATE FILE
# ============================================================

hiring_update_cols = get_hiring_update_columns(config_df)

if HIRING_FLG_COL in updated_ref_df.columns:
    # If no specific config list found, use all reference columns.
    if hiring_update_cols:
        valid_hiring_cols = [c for c in hiring_update_cols if c in updated_ref_df.columns]
    else:
        valid_hiring_cols = list(updated_ref_df.columns)

    if HIRING_FLG_COL not in valid_hiring_cols:
        valid_hiring_cols.append(HIRING_FLG_COL)

    hiring_update_df = updated_ref_df[
        clean_text_series(updated_ref_df[HIRING_FLG_COL]).eq("YES")
    ][valid_hiring_cols].copy()

    hiring_update_df = apply_data_types(hiring_update_df)

    with pd.ExcelWriter(hiring_update_file, engine="openpyxl") as writer:
        hiring_update_df.to_excel(writer, index=False, sheet_name="MR_Hiring_Update")

    format_excel(hiring_update_file)
else:
    print("MR Hiring Update file skipped.")


# ============================================================
# 13. EXPORT OUTPUT FILES
#
# Config goes only to Reference output.
# New_Migrations is also retained.
# ============================================================

with pd.ExcelWriter(output_reference_file, engine="openpyxl") as writer:
    updated_ref_df.to_excel(writer, index=False, sheet_name="Reference_Data_CM")
    config_df.to_excel(writer, index=False, sheet_name="Config")
    appian_config_df.to_excel(writer, index=False, sheet_name="APPIAN_CONFIG")
    new_migrations_df.to_excel(writer, index=False, sheet_name="New_Migrations")

# ------------------------------------------------------------
# EXCEPTION REPORT
# ------------------------------------------------------------

with pd.ExcelWriter(exception_file, engine="openpyxl") as writer:
    positions_removed_df.to_excel(writer, index=False, sheet_name="positions removed")
    new_positions_df.to_excel(writer, index=False, sheet_name="new positions added in CM")
    other_changes_df.to_excel(writer, index=False, sheet_name="Other changes")
    job_summary_exceptions_df.to_excel(writer, index=False, sheet_name="Job Summary Exceptions")
    apply_data_types(lwd_updates_df).to_excel(writer, index=False, sheet_name="LWD updates")
    apply_data_types(leavers_unmatched_df).to_excel(writer, index=False, sheet_name="Leavers not matched")
    apply_data_types(hiring_flag_removed_df).to_excel(writer, index=False, sheet_name="Hiring flags removed")
    filled_cleared_df.to_excel(writer, index=False, sheet_name="Filled - fields cleared")


# ============================================================
# 14. FORMAT OUTPUT FILES
# ============================================================

format_excel(output_reference_file)
format_excel(exception_file)


# ============================================================
# 15. COMPLETION MESSAGE
# ============================================================

print("")
print("Completed successfully.")
print(f"Reference output created: {output_reference_file}")
print(f"Exception report created: {exception_file}")

if os.path.exists(hiring_update_file):
    print(f"MR Hiring Update file created: {hiring_update_file}")

if os.path.exists(frc_reconciliation_file):
    print(f"FRC Reconciliation report created: {frc_reconciliation_file}")

print(f"Run log saved: {run_log_file}")

# Stop writing to the log file
sys.stdout = original_stream(sys.stdout)
sys.stderr = original_stream(sys.stderr)
