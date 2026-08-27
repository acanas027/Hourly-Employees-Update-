"""
Employee Hours Roster Updater
=============================
Merges a weekly "Employee Hours" export into the cumulative historical roster.

  1. Sums the new week's hours into the running totals
  2. Appends new hires with their information
  3. Overwrites attribute fields that changed (status, department, supervisor, etc.)
  4. Emits an updated .xlsx that becomes next week's historical file

Matching key: Employee Full Name + Hire Date (strict).
"""

import csv
import io
import re
from datetime import datetime

import pandas as pd
import streamlit as st
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

# ----------------------------------------------------------------------------
# Schema
# ----------------------------------------------------------------------------

# Canonical column order of the historical file (Book1), columns A-V.
CANONICAL_COLUMNS = [
    "Employment Status",                        # A
    "Employee Full Name",                       # B
    "Employee Pay Rule",                        # C
    "Temp Agency Code",                         # D
    "Temp Hourly Rates",                        # E
    "Hire Date",                                # F
    "Rehire Date",                              # G
    "Profit Center",                            # H
    "Shift",                                    # I
    "Department",                               # J
    "Reports To",                               # K
    "Regular Hours",                            # L
    "Overtime Hours",                           # M
    "Productive Hours",                         # N
    "Holiday Credit HOLCR",                     # O
    "Absence - Unplanned Hours - Excused",      # P
    "Absence - Planned Hours - Excused",        # Q
    "Absence - PTO Unplanned Hours - Excused",  # R
    "Leave",                                    # S
    "Actual Hours",                             # T
    "Job",                                      # U
    "Absence  - Unplanned Hours - Unexcused",   # V  (note: double space, as in source)
]

# The weekly export renames three absence columns. Mapped positionally.
COLUMN_ALIASES = {
    "Absence Unplanned Hours": "Absence - Unplanned Hours - Excused",
    "Planned Absenteeism Hours": "Absence - Planned Hours - Excused",
    "Extended Absenteeism Hours": "Absence - PTO Unplanned Hours - Excused",
}

# Columns that accumulate week over week.
HOURS_COLUMNS = [
    "Regular Hours",
    "Overtime Hours",
    "Productive Hours",
    "Holiday Credit HOLCR",
    "Absence - Unplanned Hours - Excused",
    "Absence - Planned Hours - Excused",
    "Absence - PTO Unplanned Hours - Excused",
    "Leave",
    "Actual Hours",
    "Absence  - Unplanned Hours - Unexcused",
]

# Columns overwritten from the new file when they change.
ATTRIBUTE_COLUMNS = [
    "Employment Status",
    "Employee Pay Rule",
    "Temp Agency Code",
    "Temp Hourly Rates",
    "Rehire Date",
    "Profit Center",
    "Shift",
    "Department",
    "Reports To",
    "Job",
]

# Columns forming the match key. Never overwritten.
KEY_COLUMNS = ["Employee Full Name", "Hire Date"]

DATE_COLUMNS = ["Hire Date", "Rehire Date"]

# Helper key in column W: =CONCATENATE(A,B,C,H,I,J,K,U)
HELPER_KEY_SOURCE_COLS = ["A", "B", "C", "H", "I", "J", "K", "U"]
HELPER_KEY_COL_INDEX = len(CANONICAL_COLUMNS) + 1  # column W


# ----------------------------------------------------------------------------
# Loading
# ----------------------------------------------------------------------------

def _normalize_header(name) -> str:
    """Trim and collapse whitespace, preserving the intentional double space."""
    if name is None:
        return ""
    return str(name).strip()


def _find_header_row(raw: pd.DataFrame, sentinel: str = "Employment Status") -> int:
    """Locate the header row; weekly exports carry metadata lines above it."""
    for idx in range(min(40, len(raw))):
        row = raw.iloc[idx].astype(str).str.strip()
        if (row == sentinel).any():
            return idx
    raise ValueError(
        f"Could not find a header row containing '{sentinel}' in the first 40 rows."
    )


def _extract_metadata(raw: pd.DataFrame, header_row: int) -> dict:
    """Pull the label/value metadata lines that sit above the header."""
    meta = {}
    for idx in range(header_row):
        label = str(raw.iloc[idx, 0]).strip().rstrip(":")
        if not label or label.lower() == "nan":
            continue
        value = ""
        if raw.shape[1] > 1:
            value = str(raw.iloc[idx, 1]).strip()
            if value.lower() == "nan":
                value = ""
        meta[label] = value
    return meta


def load_table(uploaded_file) -> tuple[pd.DataFrame, dict]:
    """Read an uploaded .csv/.xlsx into a canonical DataFrame plus its metadata."""
    name = uploaded_file.name.lower()
    uploaded_file.seek(0)

    if name.endswith(".csv"):
        # The export's metadata lines have 2 fields while the header has 22,
        # which defeats pandas' tokenizer. Read rows manually and pad.
        text = uploaded_file.read()
        if isinstance(text, bytes):
            text = text.decode("utf-8-sig", errors="replace")
        rows = list(csv.reader(io.StringIO(text)))
        if not rows:
            raise ValueError("The file appears to be empty.")
        width = max(len(r) for r in rows)
        rows = [r + [""] * (width - len(r)) for r in rows]
        raw = pd.DataFrame(rows).replace("", pd.NA)
    else:
        raw = pd.read_excel(uploaded_file, header=None, dtype=object)

    header_row = _find_header_row(raw)
    metadata = _extract_metadata(raw, header_row)

    headers = [_normalize_header(h) for h in raw.iloc[header_row].tolist()]
    df = raw.iloc[header_row + 1:].copy()
    df.columns = headers
    df = df.reset_index(drop=True)

    # Drop unnamed/helper trailing columns (the CONCATENATE and COUNTIF columns).
    df = df.loc[:, [c for c in df.columns if c != ""]]
    df = df.rename(columns=COLUMN_ALIASES)

    missing = [c for c in CANONICAL_COLUMNS if c not in df.columns]
    if missing:
        raise ValueError(
            f"'{uploaded_file.name}' is missing expected column(s): {', '.join(missing)}"
        )

    df = df[CANONICAL_COLUMNS].copy()

    # Drop rows with no name and no hire date - trailing blanks from the export.
    blank = df["Employee Full Name"].isna() & df["Hire Date"].isna()
    df = df[~blank].reset_index(drop=True)

    for col in HOURS_COLUMNS:
        df[col] = pd.to_numeric(df[col], errors="coerce")
    for col in DATE_COLUMNS:
        df[col] = pd.to_datetime(df[col], errors="coerce")

    return df, metadata


# ----------------------------------------------------------------------------
# Keys
# ----------------------------------------------------------------------------

def build_key(df: pd.DataFrame) -> pd.Series:
    """Name + Hire Date, case- and whitespace-insensitive."""
    name = (
        df["Employee Full Name"].fillna("").astype(str)
        .str.strip().str.upper().str.replace(r"\s+", " ", regex=True)
    )
    hire = df["Hire Date"].dt.strftime("%Y-%m-%d").fillna("")
    return name + " | " + hire


def collapse_duplicates(df: pd.DataFrame, keys: pd.Series) -> tuple[pd.DataFrame, pd.DataFrame]:
    """Sum hours across duplicate keys; keep the last non-null attribute values."""
    df = df.copy()
    df["__key"] = keys
    dupe_mask = df["__key"].duplicated(keep=False)
    dupes = df[dupe_mask].drop(columns="__key").copy()
    if not dupe_mask.any():
        return df.drop(columns="__key"), dupes

    agg = {}
    for col in df.columns:
        if col == "__key":
            continue
        agg[col] = "sum" if col in HOURS_COLUMNS else "last"

    collapsed = (
        df.groupby("__key", sort=False, dropna=False)
        .agg(agg)
        .reset_index(drop=True)
    )
    # groupby.sum turns all-NaN groups into 0; restore genuine blanks.
    for col in HOURS_COLUMNS:
        all_nan = df.groupby("__key", sort=False, dropna=False)[col].apply(
            lambda s: s.isna().all()
        ).reset_index(drop=True)
        collapsed.loc[all_nan.values, col] = pd.NA
        collapsed[col] = pd.to_numeric(collapsed[col], errors="coerce")

    return collapsed, dupes


# ----------------------------------------------------------------------------
# Comparison helpers
# ----------------------------------------------------------------------------

def values_differ(old, new) -> bool:
    """True when a field genuinely changed, ignoring blank/format noise."""
    old_blank = pd.isna(old) or str(old).strip() == ""
    new_blank = pd.isna(new) or str(new).strip() == ""
    if old_blank and new_blank:
        return False
    if old_blank != new_blank:
        return True

    if isinstance(old, (pd.Timestamp, datetime)) or isinstance(new, (pd.Timestamp, datetime)):
        o = pd.to_datetime(old, errors="coerce")
        n = pd.to_datetime(new, errors="coerce")
        if pd.notna(o) and pd.notna(n):
            return o.normalize() != n.normalize()

    try:
        return abs(float(old) - float(new)) > 1e-9
    except (TypeError, ValueError):
        pass

    return str(old).strip() != str(new).strip()


def display_value(v) -> str:
    if pd.isna(v) or str(v).strip() == "":
        return "(blank)"
    if isinstance(v, (pd.Timestamp, datetime)):
        return v.strftime("%-m/%-d/%Y")
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return str(v)


# ----------------------------------------------------------------------------
# Merge
# ----------------------------------------------------------------------------

def normalized_name(value) -> str:
    """Employee name stripped of case and whitespace noise, for rehire linking."""
    if pd.isna(value):
        return ""
    return re.sub(r"\s+", " ", str(value).strip()).upper()


def link_rehires(hist, new, dropped_keys, new_hire_keys):
    """Reconnect employees whose Hire Date was rewritten between exports.

    A person absent from this week's export whose name reappears under a new
    hire date is a rehire, not a departure plus a new hire. Their cumulative
    hours must carry forward.

    Pairs 1:1 only. When a name has several candidates on either side the
    pairing is genuinely ambiguous, so it is flagged for manual review rather
    than guessed at - misassigning hours between two people who share a name
    is worse than leaving it for a human.
    """
    dropped_by_name, new_by_name = {}, {}

    for key in dropped_keys:
        name = normalized_name(hist.loc[key, "Employee Full Name"])
        if name and name != "TEMP-":
            dropped_by_name.setdefault(name, []).append(key)

    for key in new_hire_keys:
        name = normalized_name(new.loc[key, "Employee Full Name"])
        if name and name != "TEMP-":
            new_by_name.setdefault(name, []).append(key)

    pairs, ambiguous = [], []
    for name in set(dropped_by_name) & set(new_by_name):
        old_keys, fresh_keys = dropped_by_name[name], new_by_name[name]
        if len(old_keys) == 1 and len(fresh_keys) == 1:
            pairs.append((old_keys[0], fresh_keys[0]))
        else:
            ambiguous.append({
                "Employee Full Name": new.loc[fresh_keys[0], "Employee Full Name"],
                "Rows in historical": len(old_keys),
                "Rows in export": len(fresh_keys),
                "Cumulative hours at risk": round(
                    float(hist.loc[old_keys, "Actual Hours"].sum(skipna=True)), 2
                ),
            })
    return pairs, ambiguous


def sync_duplicate_rows(df: pd.DataFrame, fields: list, temp_only: bool = False):
    """Make every row for one person agree on person-level fields.

    Someone with several stints has several rows. Facts about the person -
    chiefly Employment Status - should not disagree between them, so the most
    recent row wins and the older rows are brought into line. No row is removed
    and no hours move; each row keeps the hours it earned.

    Stint-level fields (Shift, Department, Reports To, Job) are deliberately
    not synced by default: an old row's shift describes where that person
    actually worked while earning those hours.
    """
    if not fields:
        return df, pd.DataFrame(), pd.DataFrame()

    df = df.copy().reset_index(drop=True)
    names = df["Employee Full Name"].apply(normalized_name)

    # Rank rows within a person: newest hire date wins, then rehire date,
    # then file order. Missing dates sort oldest.
    order = pd.DataFrame({
        "hire": df["Hire Date"],
        "rehire": df["Rehire Date"],
        "pos": range(len(df)),
    })

    sync_log, groups_seen = [], []

    for name, idx in names.groupby(names).groups.items():
        if not name or name == "TEMP-" or len(idx) < 2:
            continue
        if temp_only and not name.startswith("TEMP-"):
            continue

        block = order.loc[idx].sort_values(
            ["hire", "rehire", "pos"], na_position="first"
        )
        latest_i = block.index[-1]
        older = list(block.index[:-1])

        changed_here = 0
        for col in fields:
            authority = df.at[latest_i, col]
            for i in older:
                if values_differ(df.at[i, col], authority):
                    sync_log.append({
                        "Employee Full Name": df.at[i, "Employee Full Name"],
                        "Row hire date": display_value(df.at[i, "Hire Date"]),
                        "Field": col,
                        "Was": display_value(df.at[i, col]),
                        "Now": display_value(authority),
                        "Matched to hire date": display_value(df.at[latest_i, "Hire Date"]),
                        "TEMP": "Yes" if name.startswith("TEMP-") else "No",
                    })
                    df.at[i, col] = authority
                    changed_here += 1

        groups_seen.append({
            "Employee Full Name": df.at[latest_i, "Employee Full Name"],
            "Rows": len(idx),
            "Latest hire date": display_value(df.at[latest_i, "Hire Date"]),
            "Fields changed": changed_here,
            "TEMP": "Yes" if name.startswith("TEMP-") else "No",
        })

    return df, pd.DataFrame(sync_log), pd.DataFrame(groups_seen)


def newest_key_per_person(hist, new, temp_only: bool = False) -> set:
    """Keys that are their person's most recent row.

    Used to keep the weekly attribute overwrite off older rows for any field
    the duplicate sync owns. Without this the two stages fight every week: the
    export reports an old stint as Terminated, the sync reports the person as
    Active, and the change log fills with a flip that nets to nothing.

    Rows the sync ignores - blank names, and non-TEMP names when the sync is
    limited to TEMPs - are all treated as newest so their normal update path
    is untouched.
    """
    frames = [
        hist[["Employee Full Name", "Hire Date", "Rehire Date"]],
        new.loc[[k for k in new.index if k not in hist.index],
                ["Employee Full Name", "Hire Date", "Rehire Date"]],
    ]
    cand = pd.concat(frames)
    cand = cand[~cand.index.duplicated(keep="first")]
    cand = cand.assign(
        __name=cand["Employee Full Name"].apply(normalized_name),
        __pos=range(len(cand)),
    )

    synced = cand["__name"].ne("") & cand["__name"].ne("TEMP-")
    if temp_only:
        synced &= cand["__name"].str.startswith("TEMP-")

    # Anything the sync won't touch keeps its normal update path.
    newest = set(cand.index[~synced])

    considered = cand[synced].sort_values(
        ["__name", "Hire Date", "Rehire Date", "__pos"], na_position="first"
    )
    newest |= set(considered.groupby("__name", sort=False).tail(1).index)
    return newest


def merge_rosters(hist: pd.DataFrame, new: pd.DataFrame, drop_missing: bool,
                  link_rehire: bool = True, sync_fields: list = None,
                  sync_temp_only: bool = False) -> dict:
    hist_keys = build_key(hist)
    new_keys = build_key(new)

    hist, hist_dupes = collapse_duplicates(hist, hist_keys)
    new, new_dupes = collapse_duplicates(new, new_keys)

    hist_keys = build_key(hist)
    new_keys = build_key(new)

    hist = hist.set_index(hist_keys)
    new = new.set_index(new_keys)

    matched = [k for k in hist.index if k in new.index]
    new_hire_keys = [k for k in new.index if k not in hist.index]
    dropped_keys = [k for k in hist.index if k not in new.index]

    # Fields the duplicate sync owns must not be written onto older rows.
    sync_owned = set(sync_fields or [])
    newest_keys = (
        newest_key_per_person(hist, new, temp_only=sync_temp_only)
        if sync_owned else set(hist.index) | set(new.index)
    )

    result = hist.copy()
    hours_log, field_log, deferred_log = [], [], []

    for key in matched:
        h_row, n_row = hist.loc[key], new.loc[key]

        added = {}
        for col in HOURS_COLUMNS:
            old_v, new_v = h_row[col], n_row[col]
            if pd.isna(old_v) and pd.isna(new_v):
                continue
            total = (0 if pd.isna(old_v) else float(old_v)) + \
                    (0 if pd.isna(new_v) else float(new_v))
            result.at[key, col] = total
            if pd.notna(new_v) and float(new_v) != 0:
                added[col] = float(new_v)

        if added:
            hours_log.append({
                "Employee Full Name": h_row["Employee Full Name"],
                "Hire Date": display_value(h_row["Hire Date"]),
                **{c: round(v, 2) for c, v in added.items()},
                "Hours Added (Actual)": round(added.get("Actual Hours", 0), 2),
            })

        for col in ATTRIBUTE_COLUMNS:
            old_v, new_v = h_row[col], n_row[col]
            if not values_differ(old_v, new_v):
                continue

            if col in sync_owned and key not in newest_keys:
                # An older stint. The sync sets this field from the person's
                # newest row, so writing the export's value here would only be
                # undone a moment later.
                deferred_log.append({
                    "Employee Full Name": h_row["Employee Full Name"],
                    "Hire Date": display_value(h_row["Hire Date"]),
                    "Field": col,
                    "Export says": display_value(new_v),
                    "Left as": display_value(old_v),
                })
                continue

            result.at[key, col] = new_v
            field_log.append({
                "Employee Full Name": h_row["Employee Full Name"],
                "Hire Date": display_value(h_row["Hire Date"]),
                "Field": col,
                "Was": display_value(old_v),
                "Now": display_value(new_v),
            })

    new_hires = new.loc[new_hire_keys].copy() if new_hire_keys else new.iloc[0:0].copy()
    dropped = hist.loc[dropped_keys].copy() if dropped_keys else hist.iloc[0:0].copy()

    # --- Rehire reconciliation ---------------------------------------------
    # Someone whose Hire Date was rewritten looks like a departure plus a new
    # hire. Relink them so their cumulative hours survive.
    rehire_log, ambiguous = [], []
    rehire_rows = {}

    if link_rehire:
        pairs, ambiguous = link_rehires(hist, new, dropped_keys, new_hire_keys)

        for old_key, fresh_key in pairs:
            h_row, n_row = hist.loc[old_key], new.loc[fresh_key]
            merged_row = n_row.copy()  # adopt current attributes and hire date

            for col in HOURS_COLUMNS:
                old_v, new_v = h_row[col], n_row[col]
                if pd.isna(old_v) and pd.isna(new_v):
                    merged_row[col] = pd.NA
                else:
                    merged_row[col] = (0 if pd.isna(old_v) else float(old_v)) + \
                                      (0 if pd.isna(new_v) else float(new_v))

            rehire_rows[fresh_key] = merged_row
            rehire_log.append({
                "Employee Full Name": n_row["Employee Full Name"],
                "Hire Date was": display_value(h_row["Hire Date"]),
                "Hire Date now": display_value(n_row["Hire Date"]),
                "Status was": display_value(h_row["Employment Status"]),
                "Status now": display_value(n_row["Employment Status"]),
                "Hours carried forward": round(
                    0 if pd.isna(h_row["Actual Hours"]) else float(h_row["Actual Hours"]), 2
                ),
                "Hours added this week": round(
                    0 if pd.isna(n_row["Actual Hours"]) else float(n_row["Actual Hours"]), 2
                ),
            })

        paired_old = {o for o, _ in pairs}
        paired_new = {f for _, f in pairs}

        # A relinked person is no longer a departure or a new hire.
        dropped_keys = [k for k in dropped_keys if k not in paired_old]
        new_hire_keys = [k for k in new_hire_keys if k not in paired_new]
        new_hires = new.loc[new_hire_keys].copy() if new_hire_keys else new.iloc[0:0].copy()
        dropped = hist.loc[dropped_keys].copy() if dropped_keys else hist.iloc[0:0].copy()

        # The superseded historical row always goes, toggle or not.
        if paired_old:
            result = result.drop(index=list(paired_old))

    if drop_missing and dropped_keys:
        result = result.drop(index=dropped_keys)

    if rehire_rows:
        result = pd.concat([result, pd.DataFrame(rehire_rows).T])

    if new_hire_keys:
        result = pd.concat([result, new_hires])

    result = result.reset_index(drop=True)
    for col in HOURS_COLUMNS:
        result[col] = pd.to_numeric(result[col], errors="coerce")
    for col in DATE_COLUMNS:
        result[col] = pd.to_datetime(result[col], errors="coerce")

    rows_before_sync = len(result)
    result, sync_log, sync_groups = sync_duplicate_rows(
        result, sync_fields or [], temp_only=sync_temp_only
    )
    assert len(result) == rows_before_sync, "sync must never add or remove rows"

    return {
        "result": result,
        "hours_log": pd.DataFrame(hours_log),
        "field_log": pd.DataFrame(field_log),
        "rehire_log": pd.DataFrame(rehire_log),
        "sync_log": sync_log,
        "deferred_log": pd.DataFrame(deferred_log),
        "sync_groups": sync_groups,
        "ambiguous": pd.DataFrame(ambiguous),
        "new_hires": new_hires.reset_index(drop=True),
        "dropped": dropped.reset_index(drop=True),
        "hist_dupes": hist_dupes,
        "new_dupes": new_dupes,
        "matched_count": len(matched),
    }


# ----------------------------------------------------------------------------
# Excel output
# ----------------------------------------------------------------------------

def to_excel_bytes(df: pd.DataFrame, sheet_name: str = "Employee Hours Test") -> bytes:
    wb = Workbook()
    ws = wb.active
    ws.title = sheet_name

    header_font = Font(name="Arial", bold=True)
    body_font = Font(name="Arial")

    for col_idx, header in enumerate(CANONICAL_COLUMNS, start=1):
        cell = ws.cell(row=1, column=col_idx, value=header)
        cell.font = header_font

    for r, (_, row) in enumerate(df.iterrows(), start=2):
        for c, col in enumerate(CANONICAL_COLUMNS, start=1):
            value = row[col]
            if pd.isna(value):
                value = None
            elif isinstance(value, pd.Timestamp):
                value = value.to_pydatetime()
            cell = ws.cell(row=r, column=c, value=value)
            cell.font = body_font
            if col in DATE_COLUMNS:
                cell.number_format = "m/d/yyyy"
            elif col in HOURS_COLUMNS:
                cell.number_format = "#,##0.00"

        refs = ",".join(f"{letter}{r}" for letter in HELPER_KEY_SOURCE_COLS)
        helper = ws.cell(row=r, column=HELPER_KEY_COL_INDEX, value=f"=CONCATENATE({refs})")
        helper.font = body_font

    widths = {"A": 18, "B": 28, "C": 14, "D": 16, "E": 14, "F": 12, "G": 12,
              "H": 13, "I": 20, "J": 12, "K": 22, "U": 28}
    for letter, width in widths.items():
        ws.column_dimensions[letter].width = width
    for i in range(12, 23):
        ws.column_dimensions[get_column_letter(i)].width = 15
    ws.column_dimensions[get_column_letter(HELPER_KEY_COL_INDEX)].width = 40

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(CANONICAL_COLUMNS))}1"

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


# ----------------------------------------------------------------------------
# UI
# ----------------------------------------------------------------------------

st.set_page_config(page_title="Employee Hours Roster Updater", page_icon="📋", layout="wide")

st.title("Employee Hours Roster Updater")
st.caption(
    "Roll a weekly hours export into the cumulative roster. "
    "Employees are matched on **Employee Full Name + Hire Date**."
)

with st.sidebar:
    st.header("Options")
    link_rehire = st.checkbox(
        "Link rehires and carry their hours forward",
        value=True,
        help=(
            "When someone's Hire Date is rewritten between exports, treat it as a "
            "rehire rather than a departure plus a new hire, so their cumulative "
            "hours survive. Inactive-to-active needs no linking - the hire date "
            "doesn't move, so those hours already carry."
        ),
    )
    drop_missing = st.checkbox(
        "Drop employees missing from the weekly export",
        value=True,
        help=(
            "On: anyone absent from this week's export is removed, along with their "
            "cumulative hours. Off: they stay in the file untouched. Linked rehires "
            "are exempt either way."
        ),
    )
    sort_output = st.checkbox(
        "Sort output by Employee Full Name",
        value=False,
        help="Off keeps the historical row order and appends new hires at the bottom.",
    )
    st.divider()
    st.subheader("Duplicate rows")
    st.caption(
        "When one person has several rows, the newest row wins and older rows "
        "are brought into line. No row is removed and no hours move."
    )
    sync_fields = st.multiselect(
        "Fields to sync onto older rows",
        options=ATTRIBUTE_COLUMNS,
        default=["Employment Status"],
        help=(
            "Employment Status describes the person, so it shouldn't differ "
            "between their rows. Shift, Department, Reports To and Job describe "
            "a particular stint - syncing those rewrites where someone worked "
            "while earning the hours on that row."
        ),
    )
    sync_temp_only = st.checkbox(
        "Only sync TEMP- names",
        value=False,
        help=(
            "Matching is on name alone. Repeat TEMP stints are reliably the same "
            "person; two regular employees can share a name and are not."
        ),
    )
    st.divider()
    st.caption(
        "Hours columns accumulate. Status, pay rule, agency, rate, rehire date, "
        "profit center, shift, department, supervisor, and job are overwritten "
        "when they change. Name and hire date are never touched."
    )

col_a, col_b = st.columns(2)
with col_a:
    hist_file = st.file_uploader(
        "Historical file (last week's roster)", type=["xlsx", "xlsm", "csv"], key="hist"
    )
with col_b:
    new_file = st.file_uploader(
        "This week's Employee Hours export", type=["csv", "xlsx", "xlsm"], key="new"
    )

if not (hist_file and new_file):
    st.info("Upload both files to continue.")
    st.stop()

try:
    hist_df, hist_meta = load_table(hist_file)
except Exception as exc:
    st.error(f"Could not read the historical file: {exc}")
    st.stop()

try:
    new_df, new_meta = load_table(new_file)
except Exception as exc:
    st.error(f"Could not read the weekly export: {exc}")
    st.stop()

period = new_meta.get("Time Period", "")
if period:
    st.success(f"Weekly export period: **{period}**")

merged = merge_rosters(
    hist_df, new_df,
    drop_missing=drop_missing,
    link_rehire=link_rehire,
    sync_fields=sync_fields,
    sync_temp_only=sync_temp_only,
)
out_df = merged["result"]

if sort_output:
    out_df = out_df.sort_values(
        "Employee Full Name", key=lambda s: s.astype(str).str.upper()
    ).reset_index(drop=True)

# --- Summary -----------------------------------------------------------------
st.subheader("Summary")
m1, m2, m3, m4, m5, m6 = st.columns(6)
m1.metric("Rows in", f"{len(hist_df):,}")
m2.metric("Matched & updated", f"{merged['matched_count']:,}")
m3.metric("Rehires linked", f"{len(merged['rehire_log']):,}")
m4.metric("New hires added", f"{len(merged['new_hires']):,}")
m5.metric(
    "Dropped" if drop_missing else "Missing (kept)",
    f"{len(merged['dropped']):,}",
)
m6.metric("Rows out", f"{len(out_df):,}")

hours_added = new_df["Actual Hours"].sum(skipna=True)
carried = merged["rehire_log"]["Hours carried forward"].sum() if len(merged["rehire_log"]) else 0
st.caption(
    f"{hours_added:,.2f} actual hours added this week across "
    f"{int(new_df['Actual Hours'].notna().sum()):,} employees. "
    f"Field changes applied: {len(merged['field_log']):,}. "
    f"Hours preserved through rehire links: {carried:,.2f}."
)

if sync_fields:
    st.caption(
        "Synced across duplicate rows: " + ", ".join(sync_fields) +
        ". The newest row is the authority for these, so the weekly overwrite "
        "leaves them alone on older rows."
    )

# --- Warnings ----------------------------------------------------------------
if len(merged["ambiguous"]):
    st.error(
        f"{len(merged['ambiguous'])} name(s) could be a rehire but have several "
        "possible matches, so they were left alone rather than guessed at. "
        "See the **Needs review** tab - these hours will be lost if you don't "
        "resolve them by hand."
    )

if drop_missing and len(merged["dropped"]):
    lost = merged["dropped"]["Actual Hours"].sum(skipna=True)
    st.warning(
        f"{len(merged['dropped'])} employee(s) will be removed, taking "
        f"{lost:,.2f} cumulative actual hours with them. Review the **Dropped** "
        "tab before downloading."
    )

if len(merged["new_dupes"]):
    st.warning(
        f"{len(merged['new_dupes'])} row(s) in the weekly export share a "
        "name + hire date. Their hours were summed together into one row."
    )

if len(merged["hist_dupes"]):
    st.warning(
        f"{len(merged['hist_dupes'])} row(s) in the historical file share a "
        "name + hire date. Their hours were summed together into one row."
    )

# --- Change log --------------------------------------------------------------
st.subheader("Change log")
tabs = st.tabs([
    f"Field changes ({len(merged['field_log'])})",
    f"Rehires linked ({len(merged['rehire_log'])})",
    f"New hires ({len(merged['new_hires'])})",
    f"Dropped ({len(merged['dropped'])})",
    f"Duplicate sync ({len(merged['sync_log'])})",
    f"Held back ({len(merged['deferred_log'])})",
    f"Needs review ({len(merged['ambiguous'])})",
    f"Hours added ({len(merged['hours_log'])})",
    "Duplicates",
    "Preview",
])

with tabs[0]:
    if len(merged["field_log"]):
        log = merged["field_log"]
        fields = st.multiselect(
            "Filter by field", sorted(log["Field"].unique()), key="field_filter"
        )
        if fields:
            log = log[log["Field"].isin(fields)]
        st.dataframe(log, use_container_width=True, hide_index=True)
        st.caption("Counts by field: " + ", ".join(
            f"{f} ({n})" for f, n in merged["field_log"]["Field"].value_counts().items()
        ))
    else:
        st.info("No attribute changes this week.")

with tabs[1]:
    if len(merged["rehire_log"]):
        st.dataframe(
            merged["rehire_log"].sort_values("Hours carried forward", ascending=False),
            use_container_width=True, hide_index=True,
        )
        st.caption(
            "These employees came back under a rewritten hire date. Their cumulative "
            "hours were carried onto the new row and the superseded row removed."
        )
    else:
        st.info("No rehires detected this week.")

with tabs[2]:
    if len(merged["new_hires"]):
        st.dataframe(
            merged["new_hires"][
                ["Employee Full Name", "Employment Status", "Hire Date", "Shift",
                 "Department", "Reports To", "Job", "Temp Agency Code", "Actual Hours"]
            ],
            use_container_width=True, hide_index=True,
        )
    else:
        st.info("No new hires this week.")

with tabs[3]:
    if len(merged["dropped"]):
        st.dataframe(
            merged["dropped"][
                ["Employee Full Name", "Employment Status", "Hire Date", "Shift",
                 "Department", "Reports To", "Job", "Actual Hours"]
            ],
            use_container_width=True, hide_index=True,
        )
        st.caption(
            "These exist in the historical file but not in this week's export. "
            "'Actual Hours' is their cumulative total to date."
        )
    else:
        st.info("Every historical employee appeared in this week's export.")

with tabs[4]:
    if len(merged["sync_log"]):
        log = merged["sync_log"]
        non_temp = log[log["TEMP"] == "No"]
        if len(non_temp):
            st.warning(
                f"{len(non_temp)} of these are regular employees matched on name "
                "alone. Confirm they are the same person and not two people who "
                "share a name."
            )
        st.dataframe(log, use_container_width=True, hide_index=True)
        if len(merged["sync_groups"]):
            st.caption("People with more than one row:")
            st.dataframe(
                merged["sync_groups"].sort_values("Rows", ascending=False),
                use_container_width=True, hide_index=True,
            )
    elif len(merged["sync_groups"]):
        st.info(
            f"{len(merged['sync_groups'])} people have multiple rows, but their "
            "selected fields already agree. Nothing to change."
        )
    else:
        st.info("No selected fields to sync.")

with tabs[5]:
    if len(merged["deferred_log"]):
        st.dataframe(merged["deferred_log"], use_container_width=True, hide_index=True)
        st.caption(
            "These are older stints. The export reports the status that stint "
            "ended on, but the person's newest row is the authority, so the value "
            "was left alone instead of being written and immediately re-synced. "
            "Nothing here needs action - it's shown so the hold-back isn't silent."
        )
    else:
        st.info("Nothing was held back this week.")

with tabs[6]:
    if len(merged["ambiguous"]):
        st.dataframe(merged["ambiguous"], use_container_width=True, hide_index=True)
        st.caption(
            "This name appears on both sides but with more than one candidate row, "
            "so pairing them automatically could hand one person's hours to another. "
            "Resolve these in the historical file before re-running."
        )
    else:
        st.info("Nothing needs manual review.")

with tabs[7]:
    if len(merged["hours_log"]):
        st.dataframe(
            merged["hours_log"].sort_values("Hours Added (Actual)", ascending=False),
            use_container_width=True, hide_index=True,
        )
    else:
        st.info("No hours were added this week.")

with tabs[8]:
    if len(merged["new_dupes"]):
        st.markdown("**Weekly export**")
        st.dataframe(
            merged["new_dupes"][
                ["Employee Full Name", "Hire Date", "Employment Status",
                 "Temp Agency Code", "Shift", "Department", "Actual Hours"]
            ].sort_values(["Employee Full Name", "Hire Date"]),
            use_container_width=True, hide_index=True,
        )
    if len(merged["hist_dupes"]):
        st.markdown("**Historical file**")
        st.dataframe(
            merged["hist_dupes"][
                ["Employee Full Name", "Hire Date", "Employment Status",
                 "Temp Agency Code", "Shift", "Department", "Actual Hours"]
            ].sort_values(["Employee Full Name", "Hire Date"]),
            use_container_width=True, hide_index=True,
        )
    if not len(merged["new_dupes"]) and not len(merged["hist_dupes"]):
        st.info("No duplicate name + hire date combinations found.")

with tabs[9]:
    st.dataframe(out_df.head(200), use_container_width=True, hide_index=True)
    st.caption(f"Showing the first 200 of {len(out_df):,} rows.")

# --- Download ----------------------------------------------------------------
st.subheader("Download")

label = re.sub(r"[^0-9A-Za-z]+", "_", period).strip("_") if period else \
    datetime.now().strftime("%Y_%m_%d")
filename = f"Employee_Hours_Historical_{label}.xlsx"

st.download_button(
    "Download updated historical file",
    data=to_excel_bytes(out_df),
    file_name=filename,
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    type="primary",
)
st.caption(f"`{filename}` — upload this as the historical file next week.")
