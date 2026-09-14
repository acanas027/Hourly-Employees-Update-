"""
Employee Hours Roster Updater
=============================
Merges a weekly "Employee Hours" export into the cumulative historical roster.

  1. Sums the new week's hours into the running totals
  2. Appends new hires with their information
  3. Overwrites attribute fields that changed (status, department, supervisor, etc.)
  4. Persists the cumulative historical roster in Google Sheets
  5. Archives each weekly input so prior weeks can be safely replaced/replayed

Matching key: Employee Full Name + Hire Date (strict).
"""

import csv
import hashlib
import io
import re
import time
from datetime import datetime, timezone

import gspread
import pandas as pd
import streamlit as st
from openpyxl import Workbook
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


def header_fingerprint(name) -> str:
    """Case- and whitespace-insensitive form of a column name.

    Column V is 'Absence  - Unplanned Hours - Unexcused' with a double space,
    a quirk of the source export. Google Sheets collapses runs of whitespace on
    some paste paths, so an exact-match lookup breaks on a header that is
    visually identical. Matching on the fingerprint tolerates that in either
    direction without letting Excused and Unexcused collide.
    """
    return re.sub(r"\s+", " ", str(name or "").strip()).upper()


def resolve_headers(headers: list) -> list:
    """Map raw sheet/file headers onto canonical names where they match."""
    canonical_by_fp = {header_fingerprint(c): c for c in CANONICAL_COLUMNS}
    alias_by_fp = {header_fingerprint(k): v for k, v in COLUMN_ALIASES.items()}

    resolved = []
    for raw in headers:
        fp = header_fingerprint(raw)
        resolved.append(alias_by_fp.get(fp) or canonical_by_fp.get(fp) or _normalize_header(raw))
    return resolved

def parse_hours_series(series: pd.Series) -> pd.Series:
    """
    Convert formatted hour values from Google Sheets, CSV, and Excel into
    actual numeric values.

    Examples:
        "3,267.66"  -> 3267.66
        "$1,250.00" -> 1250.00
        "(40.00)"   -> -40.00
        ""          -> NaN
        None        -> NaN
    """
    text = series.astype("string").str.strip()

    # Detect accounting-style negative numbers before removing parentheses.
    negative_mask = text.str.match(r"^\(.*\)$", na=False)

    text = (
        text
        .str.replace(",", "", regex=False)
        .str.replace("$", "", regex=False)
        .str.replace("(", "", regex=False)
        .str.replace(")", "", regex=False)
        .str.strip()
    )

    result = pd.to_numeric(text, errors="coerce")

    # Restore negatives represented with parentheses.
    result.loc[negative_mask] = -result.loc[negative_mask].abs()

    # to_numeric on a string column returns a pandas *nullable* dtype (Int64 /
    # Float64). merge_rosters assigns a whole frame of hours into .loc[...] on a
    # string-indexed DataFrame, and pandas takes an integer-coercion path there
    # that indexes positionally, raising KeyError. Plain float64 avoids it and
    # still keeps blanks as NaN.
    return result.astype("float64")


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

    headers = resolve_headers(raw.iloc[header_row].tolist())
    df = raw.iloc[header_row + 1:].copy()
    df.columns = headers
    df = df.reset_index(drop=True)

    # Drop unnamed/helper trailing columns (the CONCATENATE and COUNTIF columns).
    df = df.loc[:, [c for c in df.columns if c != ""]]
    df = df.loc[:, ~df.columns.duplicated()]

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
        df[col] = parse_hours_series(df[col])
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


def is_day_month_swap(a, b) -> bool:
    """True when two dates are the same numbers with day and month exchanged.

    5/3/2018 vs 3/5/2018. Google Sheets produces these in bulk when the
    spreadsheet locale is not United States, and the damage is invisible: the
    dates still look plausible and the rehire logic quietly absorbs them.
    """
    a, b = pd.to_datetime(a, errors="coerce"), pd.to_datetime(b, errors="coerce")
    if pd.isna(a) or pd.isna(b) or a == b:
        return False
    return a.year == b.year and a.day == b.month and a.month == b.day


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

    # Attribute columns are written cell by cell below. If pandas inferred a
    # numeric dtype for one (Department read as int64, say) then assigning the
    # export's string value raises. Object dtype accepts either.
    for col in ATTRIBUTE_COLUMNS:
        if result[col].dtype != object:
            result[col] = result[col].astype(object)

    hours_log, field_log, deferred_log = [], [], []

    if matched:
        # --- Hours, vectorized ---------------------------------------------
        # Elementwise add over aligned indexes. A per-row Python loop here cost
        # seconds on every rerun, which reads as the app freezing.
        h_hours = hist.loc[matched, HOURS_COLUMNS]
        n_hours = new.loc[matched, HOURS_COLUMNS].apply(pd.to_numeric, errors="coerce")
        h_hours = h_hours.apply(pd.to_numeric, errors="coerce")

        both_blank = h_hours.isna() & n_hours.isna()
        totals = (h_hours.fillna(0) + n_hours.fillna(0)).mask(both_blank)
        result.loc[matched, HOURS_COLUMNS] = totals

        contributed = n_hours.fillna(0).ne(0)
        rows_with_hours = contributed.any(axis=1)
        for key in n_hours.index[rows_with_hours]:
            added = {c: round(float(n_hours.at[key, c]), 2)
                     for c in HOURS_COLUMNS if contributed.at[key, c]}
            hours_log.append({
                "Employee Full Name": hist.at[key, "Employee Full Name"],
                "Hire Date": display_value(hist.at[key, "Hire Date"]),
                **added,
                "Hours Added (Actual)": round(
                    float(n_hours.at[key, "Actual Hours"] or 0), 2
                ),
            })

        # --- Attributes ------------------------------------------------------
        # values_differ is careful but slow. A plain string comparison flags a
        # superset of real changes, so it cheaply narrows the candidates and
        # values_differ only runs on those.
        h_attr = hist.loc[matched, ATTRIBUTE_COLUMNS]
        n_attr = new.loc[matched, ATTRIBUTE_COLUMNS]
        candidates = h_attr.astype(str).values != n_attr.astype(str).values

        for r_i, key in enumerate(matched):
            for c_i, col in enumerate(ATTRIBUTE_COLUMNS):
                if not candidates[r_i, c_i]:
                    continue
                old_v, new_v = h_attr.iat[r_i, c_i], n_attr.iat[r_i, c_i]
                if not values_differ(old_v, new_v):
                    continue

                if col in sync_owned and key not in newest_keys:
                    # An older stint. The sync sets this field from the person's
                    # newest row, so writing the export's value here would only
                    # be undone a moment later.
                    deferred_log.append({
                        "Employee Full Name": hist.at[key, "Employee Full Name"],
                        "Hire Date": display_value(hist.at[key, "Hire Date"]),
                        "Field": col,
                        "Export says": display_value(new_v),
                        "Left as": display_value(old_v),
                    })
                    continue

                result.at[key, col] = new_v
                field_log.append({
                    "Employee Full Name": hist.at[key, "Employee Full Name"],
                    "Hire Date": display_value(hist.at[key, "Hire Date"]),
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
    date_swaps = 0

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

            if is_day_month_swap(h_row["Hire Date"], n_row["Hire Date"]):
                date_swaps += 1

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
        "date_swaps": date_swaps,
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


def sheet_to_excel_bytes(ws) -> bytes:
    """
    Export the worksheet exactly as it exists in Google Sheets.

    No pandas conversion.
    No date parsing.
    No numeric conversion.
    No column filtering.

    Every visible cell from Google Sheets is written directly into Excel.
    """
    wb = Workbook()
    excel_ws = wb.active
    excel_ws.title = ws.title

    values = api_call(ws.get_all_values)

    for r, row in enumerate(values, start=1):
        for c, value in enumerate(row, start=1):
            excel_ws.cell(row=r, column=c, value=value)

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


# ----------------------------------------------------------------------------
# Google Sheets persistence
# ----------------------------------------------------------------------------
# Two tabs, nothing else:
#   Historical   the cumulative roster - the source of truth
#   Backup       a copy of Historical taken just before each save
#
# There is no weekly archive and no update log. Those existed so a past week
# could be corrected by replaying every week in order, which is not something
# this workflow does. Backup gives one level of undo instead: if a week is ever
# saved twice, copy Backup over Historical and the doubled hours are gone.

HISTORICAL_SHEET = "Historical"
BACKUP_SHEET = "Backup"


def api_call(fn, *args, **kwargs):
    """Run a Google Sheets call, waiting out quota errors.

    Google allows 60 reads and 60 writes per minute per user. The quota is per
    minute, so waiting clears it.
    """
    delay = 10
    for attempt in range(6):
        try:
            return fn(*args, **kwargs)
        except Exception as exc:
            text = str(exc)
            quota = "429" in text or "RESOURCE_EXHAUSTED" in text or "RATE_LIMIT" in text
            if not quota or attempt == 5:
                raise
            time.sleep(delay)
            delay = min(delay * 2, 70)


def file_sha256(file_bytes: bytes) -> str:
    return hashlib.sha256(file_bytes).hexdigest()


def _sheet_cell_value(value, column=None):
    """Convert pandas/numpy values into JSON-safe Google Sheets cell values."""
    if pd.isna(value):
        return ""
    if column in DATE_COLUMNS:
        dt = pd.to_datetime(value, errors="coerce")
        return "" if pd.isna(dt) else dt.strftime("%Y-%m-%d")
    if isinstance(value, (pd.Timestamp, datetime)):
        return value.strftime("%Y-%m-%d %H:%M:%S")
    if hasattr(value, "item") and not isinstance(value, (str, bytes)):
        try:
            value = value.item()
        except (ValueError, AttributeError):
            pass
    return value


def dataframe_to_sheet_rows(df: pd.DataFrame) -> list:
    rows = []
    for _, row in df.iterrows():
        rows.append([_sheet_cell_value(row.get(col, ""), col) for col in CANONICAL_COLUMNS])
    return rows


def roster_from_sheet_values(values: list, source_name: str) -> pd.DataFrame:
    """Read a Google Sheet table into the same canonical dataframe as load_table."""
    if not values:
        return pd.DataFrame(columns=CANONICAL_COLUMNS)

    headers = resolve_headers(values[0])

    missing = [c for c in CANONICAL_COLUMNS if c not in headers]
    if missing:
        raise ValueError(
            f"Google Sheet tab '{source_name}' is missing expected column(s): "
            + ", ".join(missing)
        )

    rows = values[1:]
    width = len(headers)
    padded = [r + [""] * (width - len(r)) for r in rows]
    df = pd.DataFrame(padded, columns=headers)
    df = df.loc[:, ~df.columns.duplicated()]
    df = df[CANONICAL_COLUMNS].copy()
    df = df.replace("", pd.NA)

    blank = df["Employee Full Name"].isna() & df["Hire Date"].isna()
    df = df[~blank].reset_index(drop=True)

    for col in HOURS_COLUMNS:
        df[col] = parse_hours_series(df[col])
    for col in DATE_COLUMNS:
        df[col] = pd.to_datetime(df[col], errors="coerce")

    return df


@st.cache_resource(show_spinner=False)
def get_google_spreadsheet():
    """Authenticate with a service account stored in Streamlit secrets."""
    try:
        credentials = dict(st.secrets["gcp_service_account"])
        sheet_id = str(st.secrets["GOOGLE_SHEET_ID"]).strip()
    except Exception as exc:
        raise RuntimeError(
            "Google Sheets secrets are not configured. Add GOOGLE_SHEET_ID and "
            "[gcp_service_account] to Streamlit secrets."
        ) from exc

    # TOML often stores literal \\n sequences; Google expects real newlines.
    if "private_key" in credentials:
        credentials["private_key"] = credentials["private_key"].replace("\\n", "\n")

    client = gspread.service_account_from_dict(credentials)
    return client.open_by_key(sheet_id)


@st.cache_resource(show_spinner=False)
def _cache_state() -> dict:
    """One mutable dict shared by every session in this server process.

    The cache version has to live beside the cache itself. session_state resets
    on every page load while cache_data persists, so a returning visitor would
    otherwise start back at version 0 and hit a stale read.
    """
    return {"version": 0}


def sheet_version() -> int:
    return _cache_state()["version"]


def bump_sheet_cache():
    """Invalidate cached tab reads after a write."""
    _cache_state()["version"] += 1
    try:
        fetch_tab_values.clear()
    except Exception:
        pass


@st.cache_data(ttl=60, show_spinner=False)
def fetch_tab_values(title: str, version: int) -> list:
    """Read one tab. `version` is part of the cache key, not used in the body."""
    book = get_google_spreadsheet()
    try:
        return api_call(book.worksheet(title).get_all_values)
    except gspread.WorksheetNotFound:
        return []


def get_or_create_worksheet(book, title: str, rows: int, cols: int, known=None):
    if known is not None and title in known:
        return known[title]
    try:
        return api_call(book.worksheet, title)
    except gspread.WorksheetNotFound:
        return book.add_worksheet(title=title, rows=rows, cols=cols)


def write_roster_worksheet(ws, df: pd.DataFrame):
    """Overwrite a tab with the roster. Header first, then every row."""
    values = [CANONICAL_COLUMNS] + dataframe_to_sheet_rows(df)
    needed_rows = max(1000, len(values) + 50)
    needed_cols = max(26, len(CANONICAL_COLUMNS))
    if ws.row_count < needed_rows or ws.col_count < needed_cols:
        api_call(ws.resize, rows=max(ws.row_count, needed_rows),
                 cols=max(ws.col_count, needed_cols))
    api_call(ws.clear)
    api_call(ws.update, values, "A1", raw=True)
    api_call(ws.freeze, rows=1)


def count_people(log, was, now):
    """Distinct people whose Employment Status moved from `was` to `now`."""
    names = set()
    for frame in log:
        if not len(frame):
            continue
        if not {"Field", "Was", "Now"} <= set(frame.columns):
            continue
        hit = frame[
            (frame["Field"] == "Employment Status")
            & (frame["Was"] == was)
            & (frame["Now"] == now)
        ]
        names |= {normalized_name(v) for v in hit["Employee Full Name"]}
    names.discard("")
    return len(names)


# ----------------------------------------------------------------------------
# Cached wrappers
# ----------------------------------------------------------------------------
# Streamlit reruns the whole script on every interaction. Without these the app
# re-reads the sheet and re-runs the merge on each click.


@st.cache_data(show_spinner=False, max_entries=4)
def load_table_cached(file_bytes: bytes, filename: str):
    buf = io.BytesIO(file_bytes)
    buf.name = filename
    return load_table(buf)


@st.cache_data(show_spinner="Merging...", max_entries=8)
def merge_cached(hist_df, new_df):
    return merge_rosters(
        hist_df, new_df,
        drop_missing=DROP_MISSING,
        link_rehire=LINK_REHIRE,
        sync_fields=list(SYNC_FIELDS),
        sync_temp_only=SYNC_TEMP_ONLY,
    )


@st.cache_data(show_spinner="Building workbook...", max_entries=2)
def build_excel_cached(df):
    return to_excel_bytes(df)


# ----------------------------------------------------------------------------
# UI
# ----------------------------------------------------------------------------

st.set_page_config(page_title="Employee Hours Roster Updater", page_icon="", layout="wide")

st.title("Employee Hours Roster Updater")
st.caption(
    "Upload this week's Employee Hours export. Hours are added to the running "
    "totals and any changed details are updated. Employees are matched on "
    "**Employee Full Name + Hire Date**."
)

# Fixed behaviour.
LINK_REHIRE = True
DROP_MISSING = True
SORT_OUTPUT = True
SYNC_FIELDS = list(ATTRIBUTE_COLUMNS)   # every attribute field
SYNC_TEMP_ONLY = False                  # applies to everyone, not just TEMPs

with st.sidebar:
    st.header("How this runs")
    st.markdown(
        "**Storage**  \n"
        "The `Historical` tab in Google Sheets is the source of truth. `Backup` "
        "holds the copy taken just before the last save."
    )
    st.markdown(
        "**Matching**  \n"
        "Employee Full Name + Hire Date."
    )
    st.markdown(
        "**Hours**  \n"
        "Every hours column accumulates week over week."
    )
    st.markdown(
        "**Rehires**  \n"
        "A rewritten hire date is treated as a rehire, not a departure plus a "
        "new hire, so cumulative hours carry forward."
    )
    st.markdown(
        "**Missing from the export**  \n"
        "Removed from the roster, along with their cumulative hours."
    )
    st.markdown(
        "**Duplicate rows**  \n"
        "When one person has several rows, the newest row is the authority for "
        "every attribute field and older rows are brought into line. No row is "
        "removed and no hours move."
    )
    st.markdown(
        "**Never changed**  \n"
        "Employee Full Name and Hire Date. Output is sorted by name."
    )
    st.markdown(
        "**If a week is saved twice**  \n"
        "Copy the `Backup` tab over `Historical` in Google Sheets. That undoes "
        "the last save."
    )

# --- Connect ------------------------------------------------------------------
try:
    book = get_google_spreadsheet()
    known = {w.title: w for w in api_call(book.worksheets)}
    historical_ws = get_or_create_worksheet(book, HISTORICAL_SHEET, 5000, 26, known)
    backup_ws = get_or_create_worksheet(book, BACKUP_SHEET, 5000, 26, known)
    historical_values = fetch_tab_values(HISTORICAL_SHEET, sheet_version())
    historical_df = (
        roster_from_sheet_values(historical_values, HISTORICAL_SHEET)
        if historical_values else pd.DataFrame(columns=CANONICAL_COLUMNS)
    )
except Exception as exc:
    st.error(f"Could not connect to Google Sheets: {exc}")
    st.stop()

if historical_df.empty:
    st.error(
        f"The **{HISTORICAL_SHEET}** tab is empty. Import your current roster "
        "into it with the same A-V headers, then reload."
    )
    st.stop()

st.success(f"Connected — **{len(historical_df):,} rows** in the roster.")

with st.expander("Download the current roster"):
    st.download_button(
        "Download roster as Excel",
        data=sheet_to_excel_bytes(historical_ws),
        file_name="Employee_Hours_Historical_Current.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

st.divider()

# --- Upload -------------------------------------------------------------------
new_file = st.file_uploader(
    "This week's Employee Hours export",
    type=["csv", "xlsx", "xlsm"],
    key="new_week_file",
)

if not new_file:
    st.info("Upload this week's export when you're ready.")
    st.stop()

file_bytes = new_file.getvalue()
file_hash = file_sha256(file_bytes)

try:
    new_df, new_meta = load_table_cached(file_bytes, new_file.name)
except Exception as exc:
    st.error(f"Could not read the export: {exc}")
    st.stop()

period = re.sub(r"\s+", " ", str(new_meta.get("Time Period", "")).strip())
if period:
    st.caption(f"Export period: {period}")

merged = merge_cached(historical_df, new_df)
out_df = merged["result"]
if SORT_OUTPUT:
    out_df = out_df.sort_values(
        "Employee Full Name", key=lambda s: s.astype(str).str.upper()
    ).reset_index(drop=True)

# --- What this week does ------------------------------------------------------
terminated = count_people([merged["field_log"], merged["sync_log"]], "Active", "Terminated")
reactivated = count_people([merged["field_log"], merged["sync_log"]], "Terminated", "Active")
hours_added = pd.to_numeric(new_df["Actual Hours"], errors="coerce").sum(skipna=True)

swaps = merged.get("date_swaps", 0)
if swaps >= 5:
    st.error(
        f"**Stop - {swaps} hire dates look day/month swapped.** They were counted "
        "as rehires, but they are almost certainly the same people with corrupted "
        "dates (5/3/2018 stored as 3/5/2018). This happens when the Google Sheet "
        "locale is not United States. Fix it under File > Settings > Locale and "
        "re-import the Historical tab before saving."
    )

st.subheader("This week")
c1, c2, c3 = st.columns(3)
c1.metric("New hires added", f"{len(merged['new_hires']):,}")
c2.metric("Active to Terminated", f"{terminated:,}")
c3.metric("Hours to add", f"{hours_added:,.2f}")

if reactivated:
    st.caption(f"{reactivated:,} went the other way, Terminated to Active.")

if len(merged["dropped"]):
    lost = merged["dropped"]["Actual Hours"].sum(skipna=True)
    names = ", ".join(merged["dropped"]["Employee Full Name"].astype(str).head(5))
    st.warning(
        f"{len(merged['dropped'])} removed for not appearing in the export, "
        f"taking {lost:,.2f} cumulative hours: {names}"
        + (" ..." if len(merged["dropped"]) > 5 else "")
    )

if len(merged["ambiguous"]):
    st.error(
        f"{len(merged['ambiguous'])} name(s) have several possible rehire matches "
        "and were left alone: "
        + ", ".join(merged["ambiguous"]["Employee Full Name"].astype(str))
    )

st.caption(f"Roster goes from {len(historical_df):,} to {len(out_df):,} rows.")

# --- Save ---------------------------------------------------------------------
st.divider()

saved = st.session_state.get("saved_hashes", set())

if file_hash in saved:
    st.success("Saved. Nothing further to do.")
else:
    st.warning("Nothing has been written yet. Review the numbers above, then save.")
    if st.button("Save to Google Sheet", type="primary"):
        try:
            with st.spinner("Backing up, then saving..."):
                # Backup first: this is the only undo there is.
                write_roster_worksheet(backup_ws, historical_df)
                write_roster_worksheet(historical_ws, out_df)
                bump_sheet_cache()
            saved.add(file_hash)
            st.session_state["saved_hashes"] = saved
            st.success(f"Saved. The roster now has {len(out_df):,} rows.")
        except Exception as exc:
            st.error(
                f"Not saved: {exc}\n\nThe sheet may be partly written. Reload the "
                "page and check the roster row count before trying again."
            )
            st.stop()

if file_hash in st.session_state.get("saved_hashes", set()):
    st.download_button(
        "Download the updated roster",
        data=build_excel_cached(out_df),
        file_name="Employee_Hours_Historical_Updated.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
