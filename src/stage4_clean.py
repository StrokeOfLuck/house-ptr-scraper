# Auto-extracted from the user's V8.2 cleanup notebook.
# Resolver rules are preserved; Colab mounting/display-only QA was removed.

from pathlib import Path
import re

import numpy as np
import pandas as pd

from config import (
    V81_TRANSACTIONS,
    V82_TRANSACTIONS,
    ensure_folders,
)

ensure_folders()

INPUT_CSV = V81_TRANSACTIONS
OUTPUT_CSV = V82_TRANSACTIONS

# Optional external securities reference table. The original notebook's
# normal production setting is None.
STOCK_REFERENCE_CSV = None

print("Input: ", INPUT_CSV)
print("Output:", OUTPUT_CSV)

# ============================================================================
# STAGE 4 — LOAD + SCHEMA CHECK
# ============================================================================
# Refuse to run if the V8.1 source file or required raw-evidence columns are missing.

if not INPUT_CSV.exists():
    raise FileNotFoundError(
        f"Could not find V8.1 CSV:\n{INPUT_CSV}"
    )

df = pd.read_csv(
    INPUT_CSV,
    low_memory=False
)

print(f"Loaded {len(df):,} V8.1 transactions")
print(f"Columns: {len(df.columns):,}")

required = {
    "asset",
    "ticker",
    "asset_type",
    "asset_raw",
    "asset_lookup_context",
    "review_reason",
    "review_level",
    "needs_review",
    "geometry_quality_score",
}

missing = sorted(required - set(df.columns))

if missing:
    raise ValueError(
        "V8.1 CSV is missing required columns: "
        + ", ".join(missing)
    )

# ============================================================================
# STAGE 4 — TICKER RESOLUTION RULES
# ============================================================================
# Resolve stock tickers from source-faithful asset evidence before [ST], preserving uncertain candidates and avoiding destructive semantic blacklists.

PAREN_RE = re.compile(r"\(([^()]*)\)")

COMPACT_SYMBOL_RE = re.compile(
    r"^[A-Za-z][A-Za-z0-9.$-]{0,7}$"
)

STRUCTURAL_ACCEPT_MAX_LEN = 6


# ----------------------------------------------------------------------
# Convert missing values to blank text and normalize repeated
# whitespace without changing the underlying evidence columns.
# ----------------------------------------------------------------------
def clean_text(value):
    if pd.isna(value):
        return ""

    return re.sub(
        r"\s+",
        " ",
        str(value)
    ).strip()


# ----------------------------------------------------------------------
# Normalize a candidate ticker to uppercase for comparison while
# leaving the original raw source text untouched elsewhere.
# ----------------------------------------------------------------------
def normalize_symbol(value):
    value = clean_text(value)

    if not value:
        return ""

    return value.upper()


# ----------------------------------------------------------------------
# Keep only source text before the first [ST] asset-type marker so
# later Subholding, Description, or Comments parentheticals cannot
# become the ticker.
# ----------------------------------------------------------------------
def asset_section_before_st(value):
    text = clean_text(value)

    if not text:
        return ""

    return re.split(
        r"\[ST\]",
        text,
        maxsplit=1,
        flags=re.I
    )[0].strip()


# ----------------------------------------------------------------------
# Collect compact symbol-shaped parentheticals from the asset section
# in their original order.
# ----------------------------------------------------------------------
def symbol_like_parentheticals(value):
    text = asset_section_before_st(value)

    output = []

    for raw in PAREN_RE.findall(text):
        candidate = clean_text(raw)

        if COMPACT_SYMBOL_RE.fullmatch(candidate):
            output.append(candidate)

    return output


# ----------------------------------------------------------------------
# Search the preserved V8.1 asset evidence in priority order and
# return the last compact parenthetical before [ST] as the raw ticker
# candidate.
# ----------------------------------------------------------------------
def extract_ticker_candidate(row):
    for column in [
        "asset_lookup_context",
        "asset_raw",
        "asset",
    ]:
        candidates = symbol_like_parentheticals(
            row.get(column, "")
        )

        if candidates:
            return normalize_symbol(
                candidates[-1]
            )

    return ""


# ----------------------------------------------------------------------
# Optionally load a separate securities reference table for
# validation. V8.2 does not require this table and never deletes a
# candidate merely because the reference lacks it.
# ----------------------------------------------------------------------
def load_reference_symbols(path):
    if path is None:
        return set()

    path = Path(path)

    if not path.exists():
        print(
            "Reference table not found; continuing without it:",
            path
        )
        return set()

    reference = pd.read_csv(
        path,
        low_memory=False
    )

    symbol_column = None

    for candidate in [
        "symbol",
        "ticker",
        "stock_symbol",
    ]:
        if candidate in reference.columns:
            symbol_column = candidate
            break

    if symbol_column is None:
        raise ValueError(
            "Reference CSV does not have a symbol/ticker column."
        )

    symbols = {
        normalize_symbol(value)
        for value in reference[symbol_column]
        if clean_text(value)
    }

    print(
        f"Loaded {len(symbols):,} reference symbols "
        f"from {path.name}"
    )

    return symbols


REFERENCE_SYMBOLS = load_reference_symbols(
    STOCK_REFERENCE_CSV
)


# ----------------------------------------------------------------------
# Resolve one row conservatively: preserve source candidates, accept
# strong structural candidates, retain V8.1 corroboration, and mark
# longer uncertain candidates for review instead of destroying them.
# ----------------------------------------------------------------------
def resolve_ticker(row):
    asset_type = normalize_symbol(
        row.get("asset_type", "")
    )

    old_ticker = normalize_symbol(
        row.get("ticker", "")
    )

    if asset_type != "ST":
        return {
            "ticker_candidate_raw": "",
            "ticker_resolved": old_ticker,
            "ticker_parse_status": "not_applicable",
            "ticker_validation_source": "",
        }

    candidate = extract_ticker_candidate(
        row
    )

    if not candidate:
        if old_ticker:
            return {
                "ticker_candidate_raw": "",
                "ticker_resolved": old_ticker,
                "ticker_parse_status": "legacy_only_review",
                "ticker_validation_source": "v8_1_only",
            }

        return {
            "ticker_candidate_raw": "",
            "ticker_resolved": "",
            "ticker_parse_status": "missing_candidate",
            "ticker_validation_source": "",
        }

    if (
        REFERENCE_SYMBOLS
        and candidate in REFERENCE_SYMBOLS
    ):
        return {
            "ticker_candidate_raw": candidate,
            "ticker_resolved": candidate,
            "ticker_parse_status": "validated",
            "ticker_validation_source": "reference_table",
        }

    if old_ticker == candidate:
        return {
            "ticker_candidate_raw": candidate,
            "ticker_resolved": candidate,
            "ticker_parse_status": "accepted",
            "ticker_validation_source": "house_ptr_structure+v8_1",
        }

    # Do not blacklist words such as FUND, CORP, ETN, ADS, or REIT.
    # If the source structure says a compact parenthetical is the asset
    # symbol, preserve it.
    if len(candidate) <= STRUCTURAL_ACCEPT_MAX_LEN:
        return {
            "ticker_candidate_raw": candidate,
            "ticker_resolved": candidate,
            "ticker_parse_status": "accepted",
            "ticker_validation_source": "house_ptr_structure",
        }

    # Longer source candidates are preserved for review.
    if old_ticker == candidate:
        return {
            "ticker_candidate_raw": candidate,
            "ticker_resolved": candidate,
            "ticker_parse_status": "accepted_long_existing",
            "ticker_validation_source": "house_ptr_structure+v8_1",
        }

    return {
        "ticker_candidate_raw": candidate,
        "ticker_resolved": "",
        "ticker_parse_status": "ambiguous_preserved",
        "ticker_validation_source": "house_ptr_structure",
    }


# ----------------------------------------------------------------------
# Clean only the presentation asset field by removing the exact ticker
# parenthetical that V8.2 accepted; raw V8.1 evidence remains intact.
# ----------------------------------------------------------------------
def remove_resolved_ticker_from_asset(asset, ticker):
    asset = clean_text(asset)
    ticker = normalize_symbol(ticker)

    if not asset or not ticker:
        return asset

    pattern = re.compile(
        r"\(\s*"
        + re.escape(ticker)
        + r"\s*\)",
        re.I
    )

    cleaned = pattern.sub(
        " ",
        asset,
        count=1
    )

    cleaned = re.sub(
        r"\s+",
        " ",
        cleaned
    ).strip()

    return cleaned.strip(" -")


# ----------------------------------------------------------------------
# Turn the semicolon-delimited review-reason string back into a list
# so ticker-related reasons can be updated without disturbing
# unrelated QA flags.
# ----------------------------------------------------------------------
def split_review_reasons(value):
    text = clean_text(value)

    if not text:
        return []

    return [
        item.strip()
        for item in text.split(";")
        if item.strip()
    ]


LOW_REVIEW_REASONS = {
    "nonstandard_exact",
    "adjacent transaction has same core signature",
}


# ----------------------------------------------------------------------
# Recalculate review reason, level, and boolean status after ticker
# resolution so successfully recovered tickers no longer carry stale
# missing-ticker warnings.
# ----------------------------------------------------------------------
def rebuild_review_fields(row):
    reasons = split_review_reasons(
        row.get("review_reason", "")
    )

    resolved_ticker = normalize_symbol(
        row.get("ticker", "")
    )

    status = clean_text(
        row.get(
            "ticker_parse_status",
            ""
        )
    )

    if resolved_ticker:
        reasons = [
            reason
            for reason in reasons
            if reason != "ST asset missing ticker"
        ]

    if (
        normalize_symbol(
            row.get("asset_type", "")
        ) == "ST"
        and not resolved_ticker
    ):
        if status == "ambiguous_preserved":
            new_reason = "ticker candidate ambiguous"
        else:
            new_reason = "ST asset missing ticker"

        if new_reason not in reasons:
            reasons.append(
                new_reason
            )

    if not reasons:
        review_level = ""
    elif all(
        reason in LOW_REVIEW_REASONS
        for reason in reasons
    ):
        review_level = "low"
    else:
        review_level = "high"

    return {
        "review_reason": "; ".join(reasons),
        "review_level": review_level,
        "needs_review": bool(reasons),
    }

# ============================================================================
# STAGE 4 — APPLY REVERSIBLE CLEANUP
# ============================================================================
# Create a new output DataFrame, copy V8.1 fields into audit columns, apply ticker resolution, and update only ticker-related QA.

out_df = df.copy()

# Preserve the V8.1 values for auditing.
out_df.insert(
    out_df.columns.get_loc("asset") + 1,
    "asset_v8_1",
    out_df["asset"]
)

out_df.insert(
    out_df.columns.get_loc("ticker") + 1,
    "ticker_v8_1",
    out_df["ticker"]
)

out_df["review_reason_v8_1"] = out_df[
    "review_reason"
]

out_df["review_level_v8_1"] = out_df[
    "review_level"
]

out_df["needs_review_v8_1"] = out_df[
    "needs_review"
]

out_df["geometry_quality_score_v8_1"] = out_df[
    "geometry_quality_score"
]

# Recover explicit exact amounts in cached V8.1 rows without re-extracting PDFs.
# Keep all other QA reasons and preserve the original review fields above.
exact_text = out_df["amount_raw"].fillna("").astype(str).str.strip()
recover_exact = (
    out_df["amount_status"].eq("missing_range_bound")
    & exact_text.str.fullmatch(r"\$[\d,]+\.\d{2}")
    & out_df["continuation_raw"].fillna("").astype(str).str.strip().eq("")
    & out_df["amount_min"].isna()
    & out_df["amount_max"].isna()
    & out_df["amount_exact"].isna()
)
out_df.loc[recover_exact, "amount_exact"] = pd.to_numeric(
    exact_text[recover_exact].str.replace(r"[$,]", "", regex=True)
)
out_df.loc[recover_exact, "amount_status"] = "nonstandard_exact"
for index in out_df.index[recover_exact]:
    reasons = split_review_reasons(out_df.at[index, "review_reason"])
    if "missing_range_bound" in reasons:
        reasons = [r for r in reasons if r != "missing_range_bound"]
        if "nonstandard_exact" not in reasons:
            reasons.append("nonstandard_exact")
        out_df.at[index, "review_reason"] = "; ".join(reasons)
        # Replace the old 20-point penalty with the exact-amount 5-point penalty.
        out_df.at[index, "geometry_quality_score"] = min(
            100, out_df.at[index, "geometry_quality_score"] + 15
        )

resolved = out_df.apply(
    resolve_ticker,
    axis=1,
    result_type="expand"
)

for column in resolved.columns:
    out_df[column] = resolved[column]

out_df["ticker"] = out_df[
    "ticker_resolved"
]

out_df.drop(
    columns=["ticker_resolved"],
    inplace=True
)

out_df["asset"] = out_df.apply(
    lambda row: remove_resolved_ticker_from_asset(
        row["asset_v8_1"],
        row["ticker"]
    ),
    axis=1
)

out_df["ticker_changed"] = (
    out_df["ticker"]
        .fillna("")
        .astype(str)
        .str.upper()
    !=
    out_df["ticker_v8_1"]
        .fillna("")
        .astype(str)
        .str.upper()
)

out_df["asset_changed_by_ticker_resolver"] = (
    out_df["asset"]
        .fillna("")
        .astype(str)
    !=
    out_df["asset_v8_1"]
        .fillna("")
        .astype(str)
)

review_updates = out_df.apply(
    rebuild_review_fields,
    axis=1,
    result_type="expand"
)

for column in [
    "review_reason",
    "review_level",
    "needs_review",
]:
    out_df[column] = review_updates[
        column
    ]

# If V8.2 fixed an old missing-ticker penalty, restore those 5 points.
old_missing_ticker = (
    out_df["review_reason_v8_1"]
        .fillna("")
        .astype(str)
        .str.contains(
            r"(?:^|;\s*)ST asset missing ticker(?:;|$)",
            regex=True
        )
)

fixed_missing_ticker = (
    old_missing_ticker
    & out_df["ticker"]
        .fillna("")
        .astype(str)
        .str.strip()
        .ne("")
)

out_df.loc[
    fixed_missing_ticker,
    "geometry_quality_score"
] = (
    pd.to_numeric(
        out_df.loc[
            fixed_missing_ticker,
            "geometry_quality_score"
        ],
        errors="coerce"
    )
    .fillna(100)
    .add(5)
    .clip(upper=100)
)

print(
    f"V8.2 rows: {len(out_df):,}"
)

# ============================================================================
# STAGE 4 — SAVE FINAL V8.2 DATASET WITH EXPLICIT ANALYSIS COLUMN NAMES
# ============================================================================
#
# The V8.2 resolver works internally with the columns named "asset" and
# "ticker". Those are the CURRENT CLEANED values.
#
# Before saving, rename them so the final CSV makes the version and purpose
# obvious:
#
#   asset  -> asset_v8_2_cleaned
#   ticker -> ticker_v8_2_cleaned
#
# The preserved V8.1 audit columns remain:
#
#   asset_v8_1
#   ticker_v8_1
#
# Therefore, for charts, SNA, enrichment, websites, and other analysis:
#
#   USE asset_v8_2_cleaned
#   USE ticker_v8_2_cleaned

OUTPUT_CSV.parent.mkdir(
    parents=True,
    exist_ok=True
)

# Make a separate save DataFrame so the in-memory QA logic above remains
# unchanged and the rename affects only the final V8.2 file.
save_df = out_df.rename(
    columns={
        "asset": "asset_v8_2_cleaned",
        "ticker": "ticker_v8_2_cleaned",
    }
).copy()

# Put the current V8.2 analysis columns directly beside their V8.1 audit
# counterparts. This makes spreadsheet inspection much easier.
preferred_front_columns = [
    "filing_id",
    "politician",
    "member_status",
    "state_district",
    "source_pdf",
    "source_year",
    "transaction_number_in_filing",
    "owner",
    "asset_v8_2_cleaned",
    "asset_v8_1",
    "ticker_v8_2_cleaned",
    "ticker_v8_1",
    "asset_type",
]

preferred_front_columns = [
    column
    for column in preferred_front_columns
    if column in save_df.columns
]

remaining_columns = [
    column
    for column in save_df.columns
    if column not in preferred_front_columns
]

save_df = save_df[
    preferred_front_columns
    + remaining_columns
]

save_df.to_csv(
    OUTPUT_CSV,
    index=False
)

print("Saved V8.2 transactions:")
print(OUTPUT_CSV)

print()
print(f"{len(save_df):,} transaction rows")

print()
print("Analysis columns:")
print("  USE: asset_v8_2_cleaned")
print("  USE: ticker_v8_2_cleaned")
print("  AUDIT ONLY: asset_v8_1")
print("  AUDIT ONLY: ticker_v8_1")
