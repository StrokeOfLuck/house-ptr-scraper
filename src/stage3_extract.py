# Auto-extracted from the user's V8.1 production notebook.
# Parser logic is intentionally preserved; only storage/runtime configuration
# was adapted from Colab to a normal Python process.

import fitz
import pandas as pd
import xlsxwriter

from config import (
    ROOT as CONFIG_ROOT,
    PDF_ROOT as CONFIG_PDF_ROOT,
    YEARS as CONFIG_YEARS,
    DATA_ROOT,
    STATUS_ROOT,
    V81_TRANSACTIONS,
    V81_FALLBACK,
    V81_EXCEL,
    V81_CHECKPOINT,
    MAX_NEW_PDFS_THIS_RUN as CONFIG_MAX_NEW_PDFS_THIS_RUN,
    ensure_folders,
)

ensure_folders()

# ============================================================================
# STAGE 3 — CONFIGURATION + SCHEMAS
# ============================================================================
# Centralize every production path, year, save/reset switch, validation regex, and output-column definition.

from pathlib import Path
import re
import math
import warnings
import os
from datetime import datetime

ROOT = CONFIG_ROOT
PDF_ROOT = CONFIG_PDF_ROOT

YEARS = CONFIG_YEARS

CSV_DIR = DATA_ROOT
STATUS_DIR = STATUS_ROOT

OUTPUT_CSV = V81_TRANSACTIONS

FALLBACK_CSV = V81_FALLBACK

OUTPUT_XLSX = V81_EXCEL

CHECKPOINT_CSV = V81_CHECKPOINT

CSV_DIR.mkdir(
    parents=True,
    exist_ok=True
)

STATUS_DIR.mkdir(
    parents=True,
    exist_ok=True
)

# ---------------------------------------------------------
# Resume / batch controls
# ---------------------------------------------------------

# Save after this many NEW PDFs.
SAVE_EVERY = 10

# None = process every remaining PDF. Overridable via the
# HOUSE_PTR_MAX_NEW_PDFS env var (see config.py) -- the GitHub Action uses
# this as a safety chunk size for incremental runs.
MAX_NEW_PDFS_THIS_RUN = CONFIG_MAX_NEW_PDFS_THIS_RUN

# Leave FALSE for normal use and resume.
# TRUE intentionally deletes V8 output/checkpoint files and starts over.
RESET_V8 = False

VALID_OWNER_CODES = {
    "",
    "JT",
    "SP",
    "DC",
    "SELF",
}

VALID_TXN_CODES = {
    "P",
    "S",
    "E",
}

DATE_RE = re.compile(
    r"\d{1,2}/\d{1,2}/\d{4}"
)

MONEY_RE = re.compile(
    r"\$[\d,]+(?:\.\d{1,2})?"
)

ASSET_TYPE_RE = re.compile(
    r"\[([A-Za-z]{1,4})\]"
)

# Parenthetical candidate only. V8 validates candidates before accepting them
# as tickers and removes only the accepted ticker from the asset name.
TICKER_CANDIDATE_RE = re.compile(
    r"\(([A-Za-z][A-Za-z0-9.$\-]{0,11})\)"
)

TXN_RE = re.compile(
    r"(?<!\w)(P|S|E)(?:\s*\((partial|full)\))?(?!\w)",
    re.I
)

HOUSE_PTR_BASE = (
    "https://disclosures-clerk.house.gov/"
    "public_disc/ptr-pdfs"
)


# ----------------------------------------------------------------------
# Build the official House Clerk URL for one source PTR PDF so every
# extracted transaction can link back to the original disclosure.
# ----------------------------------------------------------------------
def make_original_pdf_url(
    year,
    filing_id
):
    return (
        f"{HOUSE_PTR_BASE}/"
        f"{year}/"
        f"{filing_id}.pdf"
    )


# Original link remains directly beside needs_review.
TRANSACTION_COLUMNS = [
    "filing_id",
    "politician",
    "member_status",
    "state_district",
    "source_pdf",
    "source_year",
    "transaction_number_in_filing",
    "owner",
    "asset",
    "ticker",
    "asset_type",
    "transaction_type",
    "transaction_date",
    "notification_date",
    "amount_min",
    "amount_max",
    "amount_exact",
    "amount_category",
    "amount_status",
    "filing_status",
    "subholding",
    "location",
    "description",
    "comments",
    "page",
    "last_page",
    "page_parse_method",
    "page_recovery_used",
    "filing_page_completeness_issue",
    "suspected_missed_pages",
    "page_continuation_used",
    "duplicate_geometry_rows_merged",
    "possible_adjacent_same_signature",
    "row_shaded",
    "geometry_quality_score",
    "unicode_repair_applied",
    "unicode_repair_count",
    "needs_review",
    "original_pdf_url",
    "review_level",
    "review_reason",
    "owner_raw",
    "asset_raw",
    "asset_lookup_context",
    "transaction_type_raw",
    "transaction_date_raw",
    "notification_date_raw",
    "amount_raw",
    "continuation_raw",
    "detail_raw",
]


FALLBACK_COLUMNS = [
    "filing_id",
    "politician",
    "member_status",
    "state_district",
    "source_pdf",
    "source_year",
    "needs_review",
    "original_pdf_url",
    "review_level",
    "review_reason",
    "pages_with_words",
    "pages_with_geometry",
    "table_pages",
    "recovered_pages",
    "suspected_missed_pages",
    "geometry_error",
]


CHECKPOINT_COLUMNS = [
    "source_key",
    "filing_id",
    "source_pdf",
    "source_year",
    "parser_status",
    "transaction_rows",
    "pages_with_words",
    "pages_with_geometry",
    "table_pages",
    "recovered_pages",
    "suspected_missed_pages",
    "geometry_error",
    "processed_at",
]


if RESET_V8:
    print("RESET_V8=True — deleting existing V8 outputs.")

    for path in [
        OUTPUT_CSV,
        FALLBACK_CSV,
        OUTPUT_XLSX,
        CHECKPOINT_CSV,
    ]:
        if path.exists():
            path.unlink()
            print("Deleted:", path)

print("Years:", ", ".join(YEARS))
print("Transactions:", OUTPUT_CSV)
print("Fallback:", FALLBACK_CSV)
print("Checkpoint:", CHECKPOINT_CSV)

# ============================================================================
# STAGE 3 — SOURCE PDF INDEX
# ============================================================================
# Enumerate the original verified PDFs before parsing so progress and checkpoint state refer to stable source paths.

pdf_items = []

missing_year_folders = []

for year in YEARS:
    year_dir = PDF_ROOT / year

    if not year_dir.exists():
        missing_year_folders.append(
            year
        )
        continue

    for pdf_path in sorted(
        year_dir.glob("*.pdf")
    ):
        source_key = (
            f"{year}/"
            f"{pdf_path.name}"
        )

        pdf_items.append(
            (
                source_key,
                pdf_path
            )
        )

pdf_items = sorted(
    pdf_items,
    key=lambda item: (
        int(
            item[1].parent.name
        ),
        item[1].name
    )
)

PDF_INDEX = {
    source_key: pdf_path
    for source_key, pdf_path
    in pdf_items
}

print(
    "Total source PDFs discovered:",
    len(
        pdf_items
    )
)

for year in YEARS:
    count = sum(
        1
        for _, path in pdf_items
        if path.parent.name == year
    )

    print(
        year,
        "PDFs:",
        count
    )

if missing_year_folders:
    print(
        "WARNING — missing year folders:",
        missing_year_folders
    )

# ============================================================================
# STAGE 3 — AMOUNT PARSER
# ============================================================================
# Interpret disclosure amount text without borrowing unrelated numbers from detail prose.

STANDARD_AMOUNT_RANGES = {
    (1001, 15000),
    (15001, 50000),
    (50001, 100000),
    (100001, 250000),
    (250001, 500000),
    (500001, 1000000),
    (1000001, 5000000),
    (5000001, 25000000),
    (25000001, 50000000),
}

DETAIL_LABEL_RE = re.compile(
    r"\b(?:"
    r"FILING\s+STATUS"
    r"|SUBHOLDING\s+OF"
    r"|LOCATION"
    r"|DESCRIPTION"
    r"|COMMENTS?"
    r")\s*:",
    re.I
)

CAP_GAINS_HEADER_MONEY_RE = re.compile(
    r"\$200\s*\?",
    re.I
)

OPEN_WITH_MONEY_RE = re.compile(
    r"\b(?:over|more\s+than|greater\s+than)"
    r"\s*(\$[\d,]+(?:\.\d{1,2})?)",
    re.I
)

OPEN_TRAILING_RE = re.compile(
    r"\b(?:over|more\s+than|greater\s+than)\s*$",
    re.I
)


# ----------------------------------------------------------------------
# Convert one House money string such as '$15,001' into a numeric
# value while preserving blanks when no usable value exists.
# ----------------------------------------------------------------------
def money_number(text):
    if not text:
        return None

    cleaned = (
        str(text)
        .replace("$", "")
        .replace(",", "")
    )

    try:
        return float(cleaned)
    except Exception:
        return None


# ----------------------------------------------------------------------
# Take only the safe beginning of continuation text that can
# legitimately complete a split amount, stopping before detail labels
# such as Description or Comments.
# ----------------------------------------------------------------------
def amount_continuation_prefix(continuation_raw):
    """
    Keep only continuation text before the first PTR detail label.
    This prevents DESCRIPTION / COMMENTS text from changing the amount.
    """
    text = clean_space(
        continuation_raw
    )

    if not text:
        return ""

    prefix = DETAIL_LABEL_RE.split(
        text,
        maxsplit=1
    )[0]

    # Ignore "Cap. Gains > $200?" header leakage.
    prefix = CAP_GAINS_HEADER_MONEY_RE.sub(
        " ",
        prefix
    )

    return clean_space(
        prefix
    )


# ----------------------------------------------------------------------
# Interpret the House Amount column conservatively as a standard
# range, open-ended range, exact nonstandard amount, or review case
# without borrowing unrelated dollar values from later prose.
# ----------------------------------------------------------------------
def parse_amount_text(amount_raw, continuation_raw=""):
    """
    Parse PTR amount fields without borrowing dollar values or wording
    from descriptions/comments.
    """
    amount_text = clean_space(
        amount_raw
    )

    continuation_amount = amount_continuation_prefix(
        continuation_raw
    )

    combined = clean_space(
        amount_text
        + " "
        + continuation_amount
    )

    values = [
        money_number(value)
        for value in MONEY_RE.findall(
            combined
        )
    ]

    values = [
        value
        for value in values
        if value is not None
    ]

    result = {
        "amount_min": None,
        "amount_max": None,
        "amount_exact": None,
        "amount_category": "",
        "amount_status": "",
    }

    # A complete recognized PTR range wins first.
    # This blocks unrelated phrases such as "rolled over".
    if len(values) >= 2:
        pair = (
            int(round(values[0])),
            int(round(values[1]))
        )

        if pair in STANDARD_AMOUNT_RANGES:
            result["amount_min"] = pair[0]
            result["amount_max"] = pair[1]
            result["amount_category"] = (
                f"${pair[0]:,} - ${pair[1]:,}"
            )
            result["amount_status"] = "valid_range"
            return result

    # True open-ended wording directly tied to money.
    floor_value = None

    direct_open = OPEN_WITH_MONEY_RE.search(
        combined
    )

    if direct_open:
        floor_value = money_number(
            direct_open.group(1)
        )

    # Handle a split row such as:
    # amount_raw       = "Spouse/DC Over"
    # continuation_raw = "... $1,000,000 Filing Status: New"
    elif OPEN_TRAILING_RE.search(
        amount_text
    ):
        continuation_values = [
            money_number(value)
            for value in MONEY_RE.findall(
                continuation_amount
            )
        ]

        continuation_values = [
            value
            for value in continuation_values
            if value is not None
        ]

        if continuation_values:
            floor_value = continuation_values[0]

    if floor_value is not None:
        if float(
            floor_value
        ).is_integer():
            floor_value = int(
                floor_value
            )

            result["amount_min"] = floor_value + 1
            result["amount_category"] = (
                f"Over ${floor_value:,}"
            )
        else:
            result["amount_min"] = floor_value
            result["amount_category"] = (
                f"Over ${floor_value:,.2f}"
            )

        result["amount_status"] = "valid_open_ended"
        return result

    if len(values) >= 2:
        result["amount_status"] = "nonstandard_pair"
        return result

    if len(values) == 1:
        value = values[0]

        if (
            not float(value).is_integer()
            or value < 1001
            # Explicit cents distinguish exact whole-dollar disclosures
            # from a truncated standard range bound.
            or re.fullmatch(r"\$[\d,]+\.\d{2}", combined) is not None
        ):
            result["amount_exact"] = value
            result["amount_status"] = "nonstandard_exact"
        else:
            result["amount_status"] = "missing_range_bound"

        return result

    result["amount_status"] = "missing_amount"
    return result


# ============================================================================
# STAGE 3 — GEOMETRY + TEXT REPAIR
# ============================================================================
# Define the physical-layout helpers, deterministic House Unicode repair, relaxed table recovery, and independent date-anchor reconstruction used throughout the parser.

HOUSE_BROKEN_UNICODE_START = 0x0283
HOUSE_BROKEN_UNICODE_END = 0x029C
HOUSE_BROKEN_UNICODE_OFFSET = 0x0222

# Standard House PTR table X positions on a 612-point page.
# These are only used until / unless a real table on the same PDF gives us
# exact column geometry.
HOUSE_TEMPLATE_WIDTH = 612.0
HOUSE_TEMPLATE_X = [
    22.0,
    62.25,
    101.25,
    258.0,
    322.5,
    378.0,
    442.5,
    522.0,
    576.0,
]

COLUMN_NAMES = [
    "id",
    "owner",
    "asset",
    "transaction_type",
    "transaction_date",
    "notification_date",
    "amount",
    "cap_gains",
]

DETAIL_LABEL_LINE_RE = re.compile(
    r"(?:"
    r"FILING\s+STATUS"
    r"|SUBHOLDING\s+OF"
    r"|LOCATION"
    r"|DESCRIPTION"
    r"|COMMENTS?"
    r")\s*:",
    re.I
)

SECTION_BOUNDARY_RE = re.compile(
    r"(?:"
    r"For the complete list of asset type abbreviations"
    r"|Asset Class Details"
    r"|Investment Vehicle Details"
    r"|Initial Public Offerings"
    r"|Certification and Signature"
    r"|Filing ID\s*#"
    r")",
    re.I
)


# ----------------------------------------------------------------------
# Count characters affected by the known House small-cap text-layer
# encoding bug so the repair can be measured and audited.
# ----------------------------------------------------------------------
def count_house_encoding_chars(value):
    return sum(
        1
        for char in str(value or "")
        if HOUSE_BROKEN_UNICODE_START <= ord(char) <= HOUSE_BROKEN_UNICODE_END
    )


# ----------------------------------------------------------------------
# Return True when a text value still contains a character from the
# known broken House Unicode range.
# ----------------------------------------------------------------------
def contains_house_encoding_chars(value):
    return count_house_encoding_chars(value) > 0


# ----------------------------------------------------------------------
# Repair the deterministic House PDF font/ToUnicode mapping error by
# shifting the affected IPA-range characters back to their intended
# letters.
# ----------------------------------------------------------------------
def normalize_house_pdf_text(value):
    output = []

    for char in str(value or ""):
        code = ord(char)

        if HOUSE_BROKEN_UNICODE_START <= code <= HOUSE_BROKEN_UNICODE_END:
            output.append(chr(code - HOUSE_BROKEN_UNICODE_OFFSET))
        else:
            output.append(char)

    return "".join(output)


# ----------------------------------------------------------------------
# Collapse repeated whitespace without applying the House Unicode
# repair; this is useful when we need to count how many corrupt
# characters existed in the original text layer.
# ----------------------------------------------------------------------
def clean_space_raw(value):
    return re.sub(r"\s+", " ", str(value or "")).strip()


# ----------------------------------------------------------------------
# Apply the House Unicode repair and then normalize whitespace for
# ordinary parser comparisons and output fields.
# ----------------------------------------------------------------------
def clean_space(value):
    value = normalize_house_pdf_text(value)
    return re.sub(r"\s+", " ", str(value or "")).strip()


# ----------------------------------------------------------------------
# Read embedded PDF text from one physical rectangle without Unicode
# repair so raw text-layer diagnostics remain possible.
# ----------------------------------------------------------------------
def clip_text_raw(page, rect):
    return clean_space_raw(page.get_text("text", clip=rect))


# ----------------------------------------------------------------------
# Read embedded PDF text from one physical rectangle, repair House
# Unicode, and normalize whitespace.
# ----------------------------------------------------------------------
def clip_text(page, rect):
    return clean_space(page.get_text("text", clip=rect))


# ----------------------------------------------------------------------
# Call PyMuPDF table detection defensively and return an empty list
# instead of crashing the entire archive when one page has unusual
# geometry.
# ----------------------------------------------------------------------
def safe_find_tables(page, **kwargs):
    try:
        return page.find_tables(**kwargs).tables
    except Exception:
        return []


# ----------------------------------------------------------------------
# Convert a recognized House table header into the true X-coordinate
# boundaries for ID, owner, asset, dates, amount, and related columns.
# ----------------------------------------------------------------------
def header_column_boxes(table):
    if not table.rows:
        return None

    header_cells = table.rows[0].cells

    if len(header_cells) < 8:
        return None

    boxes = {}

    for name, cell in zip(COLUMN_NAMES, header_cells[:8]):
        if cell is None:
            return None

        boxes[name] = (
            float(cell[0]),
            float(cell[2])
        )

    return boxes


# ----------------------------------------------------------------------
# Provide a scaled House PTR column template when a page contains
# transactions but PyMuPDF did not successfully recognize its repeated
# table header.
# ----------------------------------------------------------------------
def default_house_column_boxes(page):
    scale = float(page.rect.width) / HOUSE_TEMPLATE_WIDTH
    xs = [x * scale for x in HOUSE_TEMPLATE_X]

    return {
        name: (xs[i], xs[i + 1])
        for i, name in enumerate(COLUMN_NAMES)
    }


# ----------------------------------------------------------------------
# Return the leftmost and rightmost X coordinates covered by the known
# PTR columns.
# ----------------------------------------------------------------------
def column_span(column_boxes):
    return (
        min(x0 for x0, _ in column_boxes.values()),
        max(x1 for _, x1 in column_boxes.values()),
    )


# ----------------------------------------------------------------------
# Normalize detected table header names into one searchable string
# used to decide whether a table is the PTR transaction table.
# ----------------------------------------------------------------------
def table_header_joined(table):
    try:
        names = [clean_space(name).lower() for name in table.header.names]
    except Exception:
        names = []

    return " ".join(names)


# ----------------------------------------------------------------------
# Recognize a PTR transaction table by requiring core header concepts
# such as Asset, Transaction, and Amount.
# ----------------------------------------------------------------------
def is_ptr_header_table(table):
    joined = table_header_joined(table)

    return (
        "asset" in joined
        and "transaction" in joined
        and "amount" in joined
    )


# ----------------------------------------------------------------------
# Find the normal House PTR transaction table on a page using PyMuPDF
# table detection and header validation.
# ----------------------------------------------------------------------
def find_ptr_table(page):
    for table in safe_find_tables(page):
        if is_ptr_header_table(table):
            return table

    return None


# ----------------------------------------------------------------------
# Measure how much a candidate table overlaps the expected PTR column
# region; this supports relaxed continuation-page recovery.
# ----------------------------------------------------------------------
def bbox_x_overlap_ratio(bbox, column_boxes):
    left, right = column_span(column_boxes)
    a = max(float(bbox[0]), left)
    b = min(float(bbox[2]), right)

    overlap = max(0.0, b - a)
    denom = max(1.0, right - left)

    return overlap / denom


# ----------------------------------------------------------------------
# Select plausible PTR table geometry even when PyMuPDF failed to
# label the repeated header correctly, using position and row
# structure instead.
# ----------------------------------------------------------------------
def find_relaxed_ptr_table(page, column_boxes):
    """
    Secondary table selector for pages where PyMuPDF finds table geometry but
    fails to label the repeated PTR header correctly.
    """
    candidates = []

    for table in safe_find_tables(page):
        if not table.rows:
            continue

        first_cells = table.rows[0].cells

        if len(first_cells) < 6:
            continue

        overlap = bbox_x_overlap_ratio(table.bbox, column_boxes)

        if overlap < 0.80:
            continue

        score = (
            overlap,
            len(table.rows),
            float(table.bbox[3]) - float(table.bbox[1]),
        )

        candidates.append((score, table))

    if not candidates:
        return None

    candidates.sort(key=lambda item: item[0], reverse=True)
    return candidates[0][1]


# ----------------------------------------------------------------------
# Clip one physical row across the known X-coordinate column
# boundaries and return the text belonging to each PTR field.
# ----------------------------------------------------------------------
def extract_row_columns(page, row_bbox, column_boxes):
    y0 = float(row_bbox[1])
    y1 = float(row_bbox[3])

    output = {}

    for name, (x0, x1) in column_boxes.items():
        output[name] = clip_text(
            page,
            fitz.Rect(x0, y0, x1, y1)
        )

    return output


# ----------------------------------------------------------------------
# Detect whether a physical row looks like the repeated PTR header so
# it is not mistaken for a transaction.
# ----------------------------------------------------------------------
def is_header_like_text(text):
    t = clean_space(text).lower()

    score = sum(
        token in t
        for token in [
            "owner",
            "asset",
            "transaction",
            "notification",
            "amount",
        ]
    )

    return score >= 3


# ----------------------------------------------------------------------
# Turn PyMuPDF table rows into parser-ready physical row objects
# containing full text, core-column text, bounding boxes, and parse
# method.
# ----------------------------------------------------------------------
def table_row_items(page, table, column_boxes, method):
    rows = list(table.rows)

    if not rows:
        return []

    # Some relaxed tables include the header as the first row; others begin
    # directly with transaction content.
    first_bbox = tuple(float(x) for x in rows[0].bbox)
    first_text = clip_text(page, fitz.Rect(*first_bbox))

    if is_header_like_text(first_text):
        rows = rows[1:]

    items = []

    for row in rows:
        row_bbox = tuple(float(x) for x in row.bbox)
        raw_full_text = clip_text_raw(page, fitz.Rect(*row_bbox))
        full_text = clean_space(raw_full_text)

        if not full_text:
            continue

        items.append({
            "bbox": row_bbox,
            "core_bbox": row_bbox,
            "raw_full_text": raw_full_text,
            "full_text": full_text,
            "core": extract_row_columns(page, row_bbox, column_boxes),
            "method": method,
        })

    return items


# ----------------------------------------------------------------------
# Return embedded PDF words that physically fall inside a requested
# horizontal column region.
# ----------------------------------------------------------------------
def words_in_x_range(page, x0, x1):
    return page.get_text(
        "words",
        clip=fitz.Rect(x0, 0, x1, page.rect.height)
    )


# ----------------------------------------------------------------------
# Find date-shaped words inside one date column and retain their
# coordinates for independent transaction-row detection.
# ----------------------------------------------------------------------
def date_words_in_column(page, x0, x1):
    output = []

    for word in words_in_x_range(page, x0, x1):
        text = clean_space(word[4])

        if not DATE_RE.fullmatch(text):
            continue

        output.append({
            "text": text,
            "x0": float(word[0]),
            "y0": float(word[1]),
            "x1": float(word[2]),
            "y1": float(word[3]),
            "yc": (float(word[1]) + float(word[3])) / 2.0,
        })

    return output


# ----------------------------------------------------------------------
# Pair Transaction Date and Notification Date words that sit on the
# same physical row, giving an independent count of likely
# transactions on the page.
# ----------------------------------------------------------------------
def paired_date_anchors(page, column_boxes, y_tolerance=3.5):
    tx_words = date_words_in_column(
        page,
        *column_boxes["transaction_date"]
    )

    notice_words = date_words_in_column(
        page,
        *column_boxes["notification_date"]
    )

    used_notice = set()
    pairs = []

    for tx in tx_words:
        options = []

        for idx, notice in enumerate(notice_words):
            if idx in used_notice:
                continue

            distance = abs(tx["yc"] - notice["yc"])

            if distance <= y_tolerance:
                options.append((distance, idx, notice))

        if not options:
            continue

        _, idx, notice = min(options, key=lambda item: item[0])
        used_notice.add(idx)

        pairs.append({
            "transaction": tx,
            "notification": notice,
            "yc": (tx["yc"] + notice["yc"]) / 2.0,
        })

    pairs.sort(key=lambda item: item["yc"])
    return pairs


# ----------------------------------------------------------------------
# Group PyMuPDF words back into visual text lines so labels and
# section boundaries can be located by Y position.
# ----------------------------------------------------------------------
def grouped_page_lines(page):
    groups = {}

    for word in page.get_text("words"):
        key = (int(word[5]), int(word[6]))
        groups.setdefault(key, []).append(word)

    output = []

    for words in groups.values():
        words = sorted(words, key=lambda w: float(w[0]))
        text = clean_space(" ".join(str(w[4]) for w in words))

        output.append({
            "text": text,
            "y0": min(float(w[1]) for w in words),
            "y1": max(float(w[3]) for w in words),
        })

    output.sort(key=lambda line: line["y0"])
    return output


# ----------------------------------------------------------------------
# Find where Filing Status, Subholding, Location, Description, or
# Comments begins beneath one transaction row.
# ----------------------------------------------------------------------
def first_detail_label_y(page, y0, y1):
    for line in grouped_page_lines(page):
        if line["y0"] < y0 or line["y0"] >= y1:
            continue

        if DETAIL_LABEL_LINE_RE.search(line["text"]):
            return line["y0"]

    return None


# ----------------------------------------------------------------------
# Find the lower boundary of the transaction section so date-anchor
# recovery does not absorb later sections such as Asset Class Details.
# ----------------------------------------------------------------------
def transaction_section_bottom(page, after_y):
    candidates = []

    for line in grouped_page_lines(page):
        if line["y0"] <= after_y:
            continue

        if SECTION_BOUNDARY_RE.search(line["text"]):
            candidates.append(line["y0"])

    if candidates:
        return min(candidates) - 1.0

    return float(page.rect.height) - 2.0


# ----------------------------------------------------------------------
# Estimate the bottom of the repeated House transaction header when
# normal table detection is unavailable.
# ----------------------------------------------------------------------
def estimated_header_bottom(page, first_anchor_y):
    relevant = []

    for word in page.get_text("words"):
        if float(word[1]) >= first_anchor_y:
            continue

        text = clean_space(word[4]).lower()

        if text in {
            "id",
            "owner",
            "asset",
            "transaction",
            "type",
            "date",
            "notification",
            "amount",
            "cap.",
            "gains",
            "$200?",
        }:
            relevant.append(float(word[3]))

    if relevant:
        return max(relevant) + 3.0

    return max(0.0, first_anchor_y - 55.0)


# ----------------------------------------------------------------------
# Reconstruct transaction groups from paired date anchors and known
# House column geometry when ordinary table recognition misses part or
# all of a valid transaction page.
# ----------------------------------------------------------------------
def date_anchor_items(page, column_boxes):
    """
    Reconstruct physical transaction groups from paired date anchors while
    retaining the known House column X geometry.
    """
    anchors = paired_date_anchors(page, column_boxes)

    if not anchors:
        return [], None

    left, right = column_span(column_boxes)
    items = []

    first_core_y0 = min(
        anchors[0]["transaction"]["y0"],
        anchors[0]["notification"]["y0"]
    ) - 1.5

    header_bottom = estimated_header_bottom(page, first_core_y0)

    prefix_bbox = None

    if first_core_y0 > header_bottom + 2:
        prefix_bbox = (
            left,
            header_bottom,
            right,
            first_core_y0 - 1.0,
        )

    for idx, anchor in enumerate(anchors):
        tx = anchor["transaction"]
        notice = anchor["notification"]

        core_y0 = min(tx["y0"], notice["y0"]) - 1.5

        if idx + 1 < len(anchors):
            next_anchor = anchors[idx + 1]
            group_y1 = min(
                next_anchor["transaction"]["y0"],
                next_anchor["notification"]["y0"]
            ) - 1.5
        else:
            group_y1 = transaction_section_bottom(page, core_y0)

        detail_y = first_detail_label_y(page, core_y0, group_y1)

        if detail_y is not None:
            core_y1 = max(core_y0 + 8.0, detail_y - 0.5)
        else:
            core_y1 = min(
                group_y1,
                max(tx["y1"], notice["y1"]) + 24.0
            )

        core_bbox = (
            left,
            core_y0,
            right,
            core_y1,
        )

        full_bbox = (
            left,
            core_y0,
            right,
            group_y1,
        )

        raw_full_text = clip_text_raw(page, fitz.Rect(*full_bbox))
        full_text = clean_space(raw_full_text)

        items.append({
            "bbox": full_bbox,
            "core_bbox": core_bbox,
            "raw_full_text": raw_full_text,
            "full_text": full_text,
            "core": extract_row_columns(page, core_bbox, column_boxes),
            "method": "date_anchor_recovery",
        })

    return items, prefix_bbox


# ----------------------------------------------------------------------
# Compare two physical PDF rectangles with a small tolerance; V8.1
# uses this rather than text identity when deciding whether geometry
# is truly duplicated.
# ----------------------------------------------------------------------
def bbox_close(a, b, tolerance=1.25):
    if a is None or b is None:
        return False

    if len(a) != 4 or len(b) != 4:
        return False

    return all(
        abs(float(x) - float(y)) <= tolerance
        for x, y in zip(a, b)
    )


# ----------------------------------------------------------------------
# Inspect PDF drawing fills to record whether a transaction row
# appears shaded; this is a diagnostic feature, not the primary row
# detector.
# ----------------------------------------------------------------------
def row_is_shaded(page, row_bbox):
    row_rect = fitz.Rect(*row_bbox)
    substantial_fills = 0

    for drawing in page.get_drawings():
        fill = drawing.get("fill")
        rect = drawing.get("rect")

        if fill is None or rect is None:
            continue

        intersection = fitz.Rect(rect) & row_rect

        if (
            not intersection.is_empty
            and intersection.width > 20
            and intersection.height > 5
        ):
            substantial_fills += 1

    return substantial_fills >= 3


# ============================================================================
# STAGE 3 — REGRESSION TESTS: AMOUNTS
# ============================================================================
# Stop immediately if known amount edge cases no longer produce the expected classifications.

# V8 amount-parser sanity tests.
_test = parse_amount_text(
    "$1,001 - $15,000",
    (
        "Filing Status: New Comments: "
        "This portfolio is independently managed, "
        "over which I have no authority."
    )
)
assert _test["amount_status"] == "valid_range"
assert _test["amount_min"] == 1001
assert _test["amount_max"] == 15000

_test = parse_amount_text(
    "$250,001 - $500,000 3 (rolled over)",
    ""
)
assert _test["amount_status"] == "valid_range"

_test = parse_amount_text(
    "Spouse/DC Over $1,000,000",
    ""
)
assert _test["amount_status"] == "valid_open_ended"
assert _test["amount_min"] == 1000001

_test = parse_amount_text(
    "Spouse/DC Over",
    "Type Date Gains > $200? [GS] $1,000,000 Filing Status: New"
)
assert _test["amount_status"] == "valid_open_ended"
assert _test["amount_min"] == 1000001

_test = parse_amount_text(
    "$318.74",
    "Filing Status: New Description: Capital Call of $318.74"
)
assert _test["amount_status"] == "nonstandard_exact"
assert _test["amount_exact"] == 318.74

print("V8 amount-parser sanity tests: PASS")


# ============================================================================
# STAGE 3 — LOGICAL TRANSACTION HELPERS
# ============================================================================
# Turn physical row fragments into logical transactions, detail fields, conservative signatures, and V8.1 ticker candidates.

# ----------------------------------------------------------------------
# Normalize transaction text for conservative comparison by cleaning
# whitespace, case, and known presentation artifacts.
# ----------------------------------------------------------------------
def normalize_text(value):
    value = clean_space(value).lower()

    # Remove common checkbox/text-layer garbage for duplicate comparison.
    value = re.sub(
        r"\b(?:gfedc|gfedcb|gfedcba)\b",
        " ",
        value
    )

    return clean_space(value)


# ----------------------------------------------------------------------
# Extract the transaction type and the two dates from the physical
# core columns of one candidate row.
# ----------------------------------------------------------------------
def extract_core_values(core):
    type_match = TXN_RE.search(
        core.get("transaction_type", "")
    )

    transaction_dates = DATE_RE.findall(
        core.get("transaction_date", "")
    )

    notification_dates = DATE_RE.findall(
        core.get("notification_date", "")
    )

    transaction_type = ""

    if type_match:
        transaction_type = type_match.group(1).upper()

        if type_match.group(2):
            transaction_type += (
                " ("
                + type_match.group(2).lower()
                + ")"
            )

    return {
        "transaction_type": transaction_type,
        "transaction_date": (
            transaction_dates[0]
            if transaction_dates
            else ""
        ),
        "notification_date": (
            notification_dates[0]
            if notification_dates
            else ""
        ),
    }


# ----------------------------------------------------------------------
# Require a recognizable transaction type plus both transaction and
# notification dates before treating a physical row as a new
# transaction.
# ----------------------------------------------------------------------
def looks_like_new_transaction(core):
    values = extract_core_values(core)

    return bool(
        values["transaction_type"]
        and values["transaction_date"]
        and values["notification_date"]
    )


# ----------------------------------------------------------------------
# Build a compact comparison signature from owner, asset, transaction
# type, dates, and amount fields; identical signatures are preserved
# unless physical coordinates also prove duplication.
# ----------------------------------------------------------------------
def core_signature(core):
    values = extract_core_values(core)

    asset_text = re.split(
        r"FILING\s+STATUS:",
        core.get("asset", ""),
        flags=re.I
    )[0]

    owner_text = re.split(
        r"FILING\s+STATUS:",
        core.get("owner", ""),
        flags=re.I
    )[0]

    amount_values = MONEY_RE.findall(
        core.get("amount", "")
    )

    return (
        normalize_text(owner_text),
        normalize_text(asset_text),
        values["transaction_type"].lower(),
        values["transaction_date"],
        values["notification_date"],
        tuple(amount_values[:2]),
    )


# ----------------------------------------------------------------------
# Recognize rows or text fragments that look like continuation/detail
# material belonging to the previous transaction.
# ----------------------------------------------------------------------
def continuation_like(text):
    return bool(
        re.search(
            r"FILING\s+STATUS"
            r"|SUBHOLDING\s+OF"
            r"|LOCATION:"
            r"|DESCRIPTION:"
            r"|COMMENTS?:"
            r"|\$[\d,]+(?:\.\d{1,2})?"
            r"|\[[A-Za-z]{1,4}\]"
            r"|\([A-Za-z][A-Za-z0-9.$\-]{0,11}\)",
            str(text),
            re.I
        )
    )


# ----------------------------------------------------------------------
# Split Filing Status, Subholding, Location, Description, and Comments
# into separate fields using every known detail label as a boundary.
# ----------------------------------------------------------------------
def parse_details(text):
    text = clean_space(
        text
    )

    result = {
        "filing_status": "",
        "subholding": "",
        "location": "",
        "description": "",
        "comments": "",
    }

    boundary = (
        r"(?:"
        r"FILING\s+STATUS"
        r"|SUBHOLDING\s+OF"
        r"|LOCATION"
        r"|DESCRIPTION"
        r"|COMMENTS?"
        r")"
    )

    patterns = {
        "filing_status": (
            r"FILING\s+STATUS:\s*(.*?)"
            rf"(?=\s+{boundary}\s*:|$)"
        ),
        "subholding": (
            r"SUBHOLDING\s+OF:\s*(.*?)"
            rf"(?=\s+{boundary}\s*:|$)"
        ),
        "location": (
            r"LOCATION:\s*(.*?)"
            rf"(?=\s+{boundary}\s*:|$)"
        ),
        "description": (
            r"DESCRIPTION:\s*(.*?)"
            rf"(?=\s+{boundary}\s*:|$)"
        ),
        "comments": (
            r"COMMENTS?:\s*(.*?)"
            rf"(?=\s+{boundary}\s*:|$)"
        ),
    }

    for field, pattern in patterns.items():
        match = re.search(
            pattern,
            text,
            re.I
        )

        if match:
            result[field] = clean_space(
                match.group(1)
            )

    return result


# ----------------------------------------------------------------------
# Check whether two detail blocks contain conflicting nonblank values;
# this supports conservative handling of apparently similar
# transactions.
# ----------------------------------------------------------------------
def detail_conflict(existing_text, candidate_text):
    existing = parse_details(existing_text)
    candidate = parse_details(candidate_text)

    for field in [
        "description",
        "subholding",
        "location",
        "comments",
    ]:
        a = normalize_text(existing.get(field, ""))
        b = normalize_text(candidate.get(field, ""))

        if a and b and a != b:
            return True

    return False


NON_TICKER_PAREN_TERMS = {
    # Clear corporate / prose descriptors.
    "PLC",
    "LLC",
    "LLP",
    "LP",
    "INC",
    "CORP",
    "CORPORATION",
    "TRUST",
    "FUND",
    "CLASS",
    "COMMON",
    "ORDINARY",

    # Clear geographic parentheticals.
    "BELGIUM",
    "CANADA",
    "CHINA",
    "FRANCE",
    "GERMANY",
    "IRELAND",
    "ISRAEL",
    "JAPAN",
    "MEXICO",
    "NETHERLANDS",
    "SWITZERLAND",
    "TAIWAN",
    "UK",
    "USA",
    "US",
}

# Do NOT blacklist ETN / ETF / ADR / ADS / REIT here.
# Those strings can themselves be legitimate market symbols.


# ----------------------------------------------------------------------
# Apply the V8.1 ticker plausibility rule to a parenthetical
# candidate. V8.2 later performs a safer post-processing pass without
# rereading PDFs.
# ----------------------------------------------------------------------
def plausible_ticker_candidate(candidate):
    candidate = clean_space(candidate).upper()

    if not candidate:
        return False

    if candidate in NON_TICKER_PAREN_TERMS:
        return False

    # Ordinary US / OTC symbols are compact. Preferred symbols with a
    # punctuation marker are allowed a little more room.
    max_len = 8 if re.search(r"[.$-]", candidate) else 6

    if len(candidate) > max_len:
        return False

    if not re.fullmatch(r"[A-Z][A-Z0-9.$-]*", candidate):
        return False

    return True


# ----------------------------------------------------------------------
# Select a ticker candidate from the extracted asset context using
# V8.1 rules. The raw context is retained so V8.2 can correct edge
# cases reversibly.
# ----------------------------------------------------------------------
def extract_ticker_from_context(text, asset_type=""):
    text = clean_space(text)

    if not text:
        return ""

    candidates = []

    for match in TICKER_CANDIDATE_RE.finditer(text):
        candidate = match.group(1).upper()

        if plausible_ticker_candidate(candidate):
            candidates.append((match.start(), candidate))

    if not candidates:
        return ""

    # The true ticker is normally the last accepted compact parenthetical
    # before the [ST] asset-type code.
    return candidates[-1][1]


# ----------------------------------------------------------------------
# Remove only the parenthetical that V8.1 actually accepted as the
# ticker from the cleaned presentation asset field.
# ----------------------------------------------------------------------
def remove_accepted_ticker(asset_text, ticker):
    if not ticker:
        return clean_space(asset_text)

    pattern = re.compile(
        r"\(" + re.escape(ticker) + r"\)",
        re.I
    )

    return clean_space(
        pattern.sub(" ", asset_text, count=1)
    )


# ----------------------------------------------------------------------
# Merge two candidates only when their core signatures match AND they
# occupy essentially the same physical coordinates on the same page;
# identical text at different positions remains separate transactions.
# ----------------------------------------------------------------------
def should_merge_duplicate(
    current,
    candidate_signature,
    candidate_core_bbox,
    page_number
):
    """
    V8 rule: textual identity is NEVER enough to merge transactions.

    Merge only when the same signature is encountered at essentially the
    same physical PDF coordinates on the same page.
    """
    if current is None:
        return False

    if candidate_signature != current["signature"]:
        return False

    for prior in current.get("physical_core_bboxes", []):
        if prior.get("page") != page_number:
            continue

        if bbox_close(
            prior.get("bbox"),
            candidate_core_bbox
        ):
            return True

    return False


# ----------------------------------------------------------------------
# Convert a House MM/DD/YYYY date string into ISO YYYY-MM-DD while
# returning blank for malformed values rather than inventing a
# correction.
# ----------------------------------------------------------------------
def to_iso_date(value):
    if not value:
        return ""

    try:
        return datetime.strptime(
            value,
            "%m/%d/%Y"
        ).strftime("%Y-%m-%d")
    except Exception:
        return value

# V8 ticker sanity tests.
assert extract_ticker_from_context("Apple Inc. (AAPL) [ST]", "ST") == "AAPL"
assert extract_ticker_from_context("Wells Fargo preferred (WFC$V) [ST]", "ST") == "WFC$V"

# Regression tests from the real V8 snapshot:
# these abbreviations are also legitimate stock symbols.
assert extract_ticker_from_context(
    "Eaton Corporation, PLC Ordinary Shares (ETN) [ST]",
    "ST"
) == "ETN"

assert extract_ticker_from_context(
    "Alliance Data Systems Corporation (AdS) [ST]",
    "ST"
) == "ADS"

# Clear geographic prose still must not become a ticker.
assert extract_ticker_from_context(
    "Example issuer (Belgium) [ST]",
    "ST"
) == ""

assert "(AAPL)" not in remove_accepted_ticker(
    "Apple Inc. (AAPL)",
    "AAPL"
)

_test_current = {
    "signature": ("", "example", "p", "01/01/2024", "01/02/2024", ("$1,001", "$15,000")),
    "physical_core_bboxes": [
        {"page": 1, "bbox": (22.0, 100.0, 576.0, 140.0)}
    ],
}

assert should_merge_duplicate(
    _test_current,
    _test_current["signature"],
    (22.2, 100.2, 575.9, 140.1),
    1
)

# Same text/signature but a different physical row MUST be preserved.
assert not should_merge_duplicate(
    _test_current,
    _test_current["signature"],
    (22.0, 150.0, 576.0, 190.0),
    1
)

print("V8 physical-duplicate + ticker sanity tests: PASS")

# V8 detail-parser sanity test.
_test_details = parse_details(
    (
        "Filing Status: New "
        "Subholding Of: Example IRA "
        "Description: Example description. "
        "Comments: This is a separate comment."
    )
)

assert _test_details["filing_status"] == "New"
assert _test_details["subholding"] == "Example IRA"
assert _test_details["description"] == "Example description."
assert _test_details["comments"] == "This is a separate comment."

print("V8 detail-parser sanity test: PASS")


# ============================================================================
# STAGE 3 — ONE-PDF PARSER
# ============================================================================
# This function is the heart of V8.1. It performs normal table parsing, recovery parsing, logical transaction assembly, cleanup, validation, and filing diagnostics.

# ----------------------------------------------------------------------
# Parse one born-digital House PTR PDF end to end: detect/recover
# transaction pages, reconstruct logical transactions, clean fields,
# validate them, and return both transaction rows and filing
# diagnostics.
# ----------------------------------------------------------------------
def parse_pdf_geometry_v8(pdf_path):
    pdf_path = Path(pdf_path)
    doc = fitz.open(pdf_path)

    first_page_text = (
        doc[0].get_text("text")
        if len(doc)
        else ""
    )

    # ----------------------------------------------------------------------
    # Nested helper for one PDF: extract a named first-page metadata field
    # such as member name or state/district.
    # ----------------------------------------------------------------------
    def metadata(pattern):
        match = re.search(pattern, first_page_text, re.I)
        return clean_space(match.group(1)) if match else ""

    filing_meta = {
        "filing_id": pdf_path.stem,
        "politician": metadata(r"Name:\s*([^\n]+)"),
        "member_status": metadata(r"Status:\s*([^\n]+)"),
        "state_district": metadata(r"State/District:\s*([^\n]+)"),
        "source_pdf": pdf_path.name,
        "source_year": pdf_path.parent.name,
        "original_pdf_url": make_original_pdf_url(
            pdf_path.parent.name,
            pdf_path.stem
        ),
    }

    records = []
    pages_with_geometry = 0
    pages_with_words = 0
    table_pages = 0
    recovered_pages = []
    suspected_missed_pages = []

    current = None
    cached_column_boxes = None

    # ----------------------------------------------------------------------
    # Nested helper for one PDF: attach a continuation/detail region to
    # the currently open transaction while tracking page boundaries and
    # recovery use.
    # ----------------------------------------------------------------------
    def append_continuation_to_current(
        page,
        page_number,
        bbox,
        full_text,
        raw_full_text,
        recovery_used=False
    ):
        nonlocal current

        if current is None:
            return

        if not continuation_like(full_text):
            return

        current["continuation_text_parts"].append(full_text)
        current["unicode_repair_count"] += count_house_encoding_chars(raw_full_text)

        if recovery_used:
            current["page_recovery_used"] = True

        if (
            current["segments"]
            and current["segments"][-1]["page"] == page_number
        ):
            current["segments"][-1]["y1"] = max(
                current["segments"][-1]["y1"],
                float(bbox[3])
            )
        else:
            current["segments"].append({
                "page": page_number,
                "y0": float(bbox[1]),
                "y1": float(bbox[3]),
                "continuation": True,
            })
            current["page_continuation_used"] = True

        current["last_page"] = page_number

    for page_number, page in enumerate(doc, start=1):
        page_has_words = bool(page.get_text("words"))

        if page_has_words:
            pages_with_words += 1

        primary_table = find_ptr_table(page)
        primary_boxes = (
            header_column_boxes(primary_table)
            if primary_table is not None
            else None
        )

        if primary_boxes is not None:
            cached_column_boxes = primary_boxes

        column_boxes = (
            primary_boxes
            or cached_column_boxes
            or default_house_column_boxes(page)
        )

        anchors = paired_date_anchors(page, column_boxes)
        anchor_count = len(anchors)

        chosen_items = []
        prefix_bbox = None
        page_method = ""

        # --------------------------------------------------------
        # 1. Normal PTR table.
        # --------------------------------------------------------
        if primary_table is not None and primary_boxes is not None:
            normal_items = table_row_items(
                page,
                primary_table,
                primary_boxes,
                "table"
            )

            normal_new_count = sum(
                looks_like_new_transaction(item["core"])
                for item in normal_items
            )

            # If the independently observed paired dates prove that the
            # table parser found fewer transaction rows, recover the page
            # from the date anchors instead of silently dropping rows.
            if anchor_count > normal_new_count:
                recovered, prefix_bbox = date_anchor_items(
                    page,
                    column_boxes
                )

                if recovered:
                    chosen_items = recovered
                    page_method = "date_anchor_recovery"
                else:
                    chosen_items = normal_items
                    page_method = "table"
            else:
                chosen_items = normal_items
                page_method = "table"

            table_pages += 1

        # --------------------------------------------------------
        # 2. Relaxed table geometry when header labeling failed.
        # --------------------------------------------------------
        else:
            relaxed_table = find_relaxed_ptr_table(
                page,
                column_boxes
            )

            if relaxed_table is not None:
                relaxed_items = table_row_items(
                    page,
                    relaxed_table,
                    column_boxes,
                    "relaxed_table"
                )

                relaxed_new_count = sum(
                    looks_like_new_transaction(item["core"])
                    for item in relaxed_items
                )

                if anchor_count > relaxed_new_count:
                    recovered, prefix_bbox = date_anchor_items(
                        page,
                        column_boxes
                    )

                    if recovered:
                        chosen_items = recovered
                        page_method = "date_anchor_recovery"
                    else:
                        chosen_items = relaxed_items
                        page_method = "relaxed_table"
                else:
                    chosen_items = relaxed_items
                    page_method = "relaxed_table"

                table_pages += 1

            # ----------------------------------------------------
            # 3. No usable table object, but paired dates prove the
            #    page contains transaction rows.
            # ----------------------------------------------------
            elif anchor_count:
                chosen_items, prefix_bbox = date_anchor_items(
                    page,
                    column_boxes
                )

                if chosen_items:
                    page_method = "date_anchor_recovery"

        if chosen_items:
            pages_with_geometry += 1

        if page_method in {
            "relaxed_table",
            "date_anchor_recovery",
        }:
            recovered_pages.append(page_number)

        # A date-anchor recovery page may begin with the continuation of the
        # previous page's last transaction. Attach that prefix before the
        # first new transaction.
        if (
            page_method == "date_anchor_recovery"
            and prefix_bbox is not None
            and current is not None
        ):
            prefix_raw = clip_text_raw(
                page,
                fitz.Rect(*prefix_bbox)
            )
            prefix_text = clean_space(prefix_raw)

            append_continuation_to_current(
                page,
                page_number,
                prefix_bbox,
                prefix_text,
                prefix_raw,
                recovery_used=True
            )

        parsed_new_on_page = 0

        for item in chosen_items:
            row_bbox = tuple(float(x) for x in item["bbox"])
            core_bbox = tuple(float(x) for x in item["core_bbox"])
            raw_full_text = item["raw_full_text"]
            full_text = item["full_text"]
            core = item["core"]
            method = item["method"]

            if not full_text:
                continue

            row_unicode_repair_count = count_house_encoding_chars(raw_full_text)
            is_new = looks_like_new_transaction(core)
            signature = core_signature(core) if is_new else None

            if is_new:
                parsed_new_on_page += 1

            # ---------------------------------------------------
            # Possible repeated core signature.
            # V8 merges only if it is physically the same row.
            # ---------------------------------------------------
            if (
                is_new
                and current is not None
                and signature == current["signature"]
            ):
                merge_it = should_merge_duplicate(
                    current,
                    signature,
                    core_bbox,
                    page_number
                )

                if merge_it:
                    current["duplicate_geometry_rows_merged"] += 1
                    current["detail_text_parts"].append(full_text)
                    current["unicode_repair_count"] += row_unicode_repair_count
                    current["physical_core_bboxes"].append({
                        "page": page_number,
                        "bbox": core_bbox,
                    })

                    if method != "table":
                        current["page_recovery_used"] = True

                    current["last_page"] = page_number
                    continue

                # Same signature at a different physical location = preserve.
                current["possible_adjacent_same_signature"] = True

            # ---------------------------------------------------
            # True new transaction.
            # ---------------------------------------------------
            if is_new:
                if current is not None:
                    records.append(current)

                core_values = extract_core_values(core)

                current = {
                    **filing_meta,
                    "signature": signature,
                    "page": page_number,
                    "first_page": page_number,
                    "last_page": page_number,
                    "page_parse_method": method,
                    "page_recovery_used": method != "table",
                    "owner_raw": core.get("owner", ""),
                    "asset_raw": core.get("asset", ""),
                    "transaction_type_raw": core.get("transaction_type", ""),
                    "transaction_date_raw": core.get("transaction_date", ""),
                    "notification_date_raw": core.get("notification_date", ""),
                    "amount_raw": core.get("amount", ""),
                    "transaction_type": core_values["transaction_type"],
                    "transaction_date_raw_value": core_values["transaction_date"],
                    "notification_date_raw_value": core_values["notification_date"],
                    "detail_text_parts": [full_text],
                    "continuation_text_parts": [],
                    "page_continuation_used": False,
                    "duplicate_geometry_rows_merged": 0,
                    "possible_adjacent_same_signature": False,
                    "unicode_repair_count": row_unicode_repair_count,
                    "row_shaded": row_is_shaded(page, core_bbox),
                    "physical_core_bboxes": [{
                        "page": page_number,
                        "bbox": core_bbox,
                    }],
                    "segments": [{
                        "page": page_number,
                        "y0": row_bbox[1],
                        "y1": row_bbox[3],
                        "continuation": False,
                    }],
                }

                continue

            # ---------------------------------------------------
            # Detail / continuation row from a real table object.
            # Date-anchor recovery already groups details with the core row.
            # ---------------------------------------------------
            if (
                current is not None
                and continuation_like(full_text)
            ):
                append_continuation_to_current(
                    page,
                    page_number,
                    row_bbox,
                    full_text,
                    raw_full_text,
                    recovery_used=(method != "table")
                )

        # Independent completeness test.
        if anchor_count > parsed_new_on_page:
            suspected_missed_pages.append(page_number)

    if current is not None:
        records.append(current)

    doc.close()

    recovered_pages = sorted(set(recovered_pages))
    suspected_missed_pages = sorted(set(suspected_missed_pages))
    suspected_missed_pages_text = ",".join(str(x) for x in suspected_missed_pages)

    cleaned_rows = []

    for transaction_number, record in enumerate(records, start=1):
        continuation_raw = clean_space(
            " ".join(record["continuation_text_parts"])
        )

        detail_raw = clean_space(
            " ".join(record["detail_text_parts"])
            + " "
            + continuation_raw
        )

        continuation_before_status = re.split(
            r"FILING\s+STATUS:",
            continuation_raw,
            flags=re.I
        )[0]

        asset_source = re.split(
            r"FILING\s+STATUS:",
            record["asset_raw"],
            flags=re.I
        )[0]

        asset_source = clean_space(asset_source)

        asset_lookup_context = clean_space(
            asset_source + " " + continuation_before_status
        )

        owner = ""
        owner_match = re.search(
            r"\b(JT|SP|DC|SELF)\b",
            record["owner_raw"],
            re.I
        )

        if owner_match:
            owner = owner_match.group(1).upper()

        asset_type_match = ASSET_TYPE_RE.search(asset_lookup_context)
        asset_type = asset_type_match.group(1).upper() if asset_type_match else ""

        ticker = extract_ticker_from_context(
            asset_lookup_context,
            asset_type
        )

        asset = ASSET_TYPE_RE.sub(" ", asset_source)
        asset = remove_accepted_ticker(asset, ticker)
        asset = clean_space(asset).strip(" -")

        leaked_owner = re.search(
            r"\b(JT|SP|DC|SELF)\s*$",
            asset,
            re.I
        )

        if leaked_owner:
            owner = leaked_owner.group(1).upper()
            asset = re.sub(
                r"\b(JT|SP|DC|SELF)\s*$",
                "",
                asset,
                flags=re.I
            ).strip()

        details = parse_details(detail_raw)

        if not asset and details["description"]:
            asset = details["description"]

        amount_info = parse_amount_text(
            record["amount_raw"],
            continuation_raw
        )

        transaction_date = to_iso_date(record["transaction_date_raw_value"])
        notification_date = to_iso_date(record["notification_date_raw_value"])

        reasons = []

        if owner.upper() not in VALID_OWNER_CODES:
            reasons.append("unexpected owner code")

        if not asset:
            reasons.append("missing asset")

        base_type = (
            record["transaction_type"][:1].upper()
            if record["transaction_type"]
            else ""
        )

        if base_type not in VALID_TXN_CODES:
            reasons.append("unexpected transaction type")

        td = None
        nd = None

        try:
            td = datetime.strptime(transaction_date, "%Y-%m-%d")
        except Exception:
            reasons.append("invalid/missing transaction date")

        try:
            nd = datetime.strptime(notification_date, "%Y-%m-%d")
        except Exception:
            reasons.append("invalid/missing notification date")

        try:
            source_year_int = int(record["source_year"])
        except Exception:
            source_year_int = None

        if td and (
            td.year < 2012
            or (
                source_year_int is not None
                and td.year > source_year_int + 1
            )
        ):
            reasons.append("implausible transaction date year")

        if nd and (
            nd.year < 2012
            or (
                source_year_int is not None
                and nd.year > source_year_int + 1
            )
        ):
            reasons.append("implausible notification date year")

        if td and nd and nd < td:
            reasons.append("notification date before transaction date")

        if amount_info["amount_status"] not in {
            "valid_range",
            "valid_open_ended",
        }:
            reasons.append(amount_info["amount_status"])

        if asset_type == "ST" and not ticker:
            reasons.append("ST asset missing ticker")

        if re.search(r"\b(JT|SP|DC|SELF)\s*$", asset, re.I):
            reasons.append("owner code remains in asset")

        post_repair_fields = [
            asset,
            ticker,
            asset_type,
            details["filing_status"],
            details["subholding"],
            details["location"],
            details["description"],
            details["comments"],
            continuation_raw,
            detail_raw,
        ]

        if any(contains_house_encoding_chars(value) for value in post_repair_fields):
            reasons.append("unrepaired House font encoding")

        if record["possible_adjacent_same_signature"]:
            reasons.append("adjacent transaction has same core signature")

        if suspected_missed_pages:
            reasons.append(
                "possible missed transaction page(s): "
                + suspected_missed_pages_text
            )

        score = 100

        penalties = {
            "unexpected owner code": 15,
            "missing asset": 30,
            "unexpected transaction type": 30,
            "invalid/missing transaction date": 25,
            "invalid/missing notification date": 25,
            "notification date before transaction date": 25,
            "implausible transaction date year": 30,
            "implausible notification date year": 30,
            "unrepaired House font encoding": 30,
            "nonstandard_pair": 25,
            "nonstandard_exact": 5,
            "missing_range_bound": 20,
            "missing_amount": 25,
            "ST asset missing ticker": 5,
            "owner code remains in asset": 15,
            "adjacent transaction has same core signature": 5,
        }

        for reason in reasons:
            if reason.startswith("possible missed transaction page(s):"):
                score -= 35
            else:
                score -= penalties.get(reason, 10)

        if record["page_continuation_used"]:
            score -= 2

        if record["duplicate_geometry_rows_merged"]:
            score -= min(3, record["duplicate_geometry_rows_merged"])

        score = max(0, min(100, score))

        low_review_reasons = {
            "nonstandard_exact",
            "adjacent transaction has same core signature",
        }

        if not reasons:
            review_level = ""
        elif all(reason in low_review_reasons for reason in reasons):
            review_level = "low"
        else:
            review_level = "high"

        cleaned_rows.append({
            "filing_id": record["filing_id"],
            "politician": record["politician"],
            "member_status": record["member_status"],
            "state_district": record["state_district"],
            "source_pdf": record["source_pdf"],
            "source_year": record["source_year"],
            "transaction_number_in_filing": transaction_number,
            "owner": owner,
            "asset": asset,
            "ticker": ticker,
            "asset_type": asset_type,
            "transaction_type": record["transaction_type"],
            "transaction_date": transaction_date,
            "notification_date": notification_date,
            "amount_min": amount_info["amount_min"],
            "amount_max": amount_info["amount_max"],
            "amount_exact": amount_info["amount_exact"],
            "amount_category": amount_info["amount_category"],
            "amount_status": amount_info["amount_status"],
            "filing_status": details["filing_status"],
            "subholding": details["subholding"],
            "location": details["location"],
            "description": details["description"],
            "comments": details["comments"],
            "page": record["first_page"],
            "last_page": record["last_page"],
            "page_parse_method": record["page_parse_method"],
            "page_recovery_used": record["page_recovery_used"],
            "filing_page_completeness_issue": bool(suspected_missed_pages),
            "suspected_missed_pages": suspected_missed_pages_text,
            "page_continuation_used": record["page_continuation_used"],
            "duplicate_geometry_rows_merged": record["duplicate_geometry_rows_merged"],
            "possible_adjacent_same_signature": record["possible_adjacent_same_signature"],
            "row_shaded": record["row_shaded"],
            "geometry_quality_score": score,
            "unicode_repair_applied": bool(record["unicode_repair_count"]),
            "unicode_repair_count": int(record["unicode_repair_count"]),
            "needs_review": bool(reasons),
            "original_pdf_url": record["original_pdf_url"],
            "review_level": review_level,
            "review_reason": "; ".join(reasons),
            "owner_raw": record["owner_raw"],
            "asset_raw": record["asset_raw"],
            "asset_lookup_context": asset_lookup_context,
            "transaction_type_raw": record["transaction_type_raw"],
            "transaction_date_raw": record["transaction_date_raw"],
            "notification_date_raw": record["notification_date_raw"],
            "amount_raw": record["amount_raw"],
            "continuation_raw": continuation_raw,
            "detail_raw": detail_raw,
            "_segments": record["segments"],
        })

    return pd.DataFrame(cleaned_rows), {
        "pages_with_geometry": pages_with_geometry,
        "pages_with_words": pages_with_words,
        "table_pages": table_pages,
        "recovered_pages": ",".join(str(x) for x in recovered_pages),
        "suspected_missed_pages": suspected_missed_pages_text,
        "politician": filing_meta["politician"],
        "member_status": filing_meta["member_status"],
        "state_district": filing_meta["state_district"],
    }


# ============================================================================
# STAGE 3 — RESUMABLE ARCHIVE RUNNER
# ============================================================================
# Load prior progress, process only uncompleted PDFs, persist rows atomically, and commit checkpoint records only after data is safely written.

# ---------------------------------------------------------
# Safe read helpers
# ---------------------------------------------------------

# ----------------------------------------------------------------------
# Read an existing persisted CSV using the expected schema, or return
# an empty DataFrame when the file has not been created yet.
# ----------------------------------------------------------------------
def read_csv_if_exists(
    path,
    columns=None
):
    if not path.exists():
        if columns is None:
            return pd.DataFrame()

        return pd.DataFrame(
            columns=columns
        )

    try:
        frame = pd.read_csv(
            path,
            dtype={
                "filing_id": str,
                "source_pdf": str,
                "source_year": str,
                "source_key": str,
            }
        )
    except pd.errors.EmptyDataError:
        frame = pd.DataFrame()

    if columns is not None:
        for column in columns:
            if column not in frame.columns:
                frame[
                    column
                ] = None

    return frame


# ----------------------------------------------------------------------
# Combine DataFrames while preserving the requested column order and
# safely handling empty pieces.
# ----------------------------------------------------------------------
def concat_safe(
    frames
):
    useful = [
        frame
        for frame in frames
        if (
            frame is not None
            and len(
                frame
            )
        )
    ]

    if not useful:
        return pd.DataFrame()

    with warnings.catch_warnings():
        warnings.simplefilter(
            "ignore",
            FutureWarning
        )

        return pd.concat(
            useful,
            ignore_index=True
        )


# ----------------------------------------------------------------------
# Write a CSV to a temporary file first and then replace the
# destination, reducing the chance that an interrupted Colab session
# leaves a half-written production file.
# ----------------------------------------------------------------------
def atomic_write_csv(
    dataframe,
    path
):
    """
    Write to a temporary file in the same Drive directory, then replace.
    """
    temp_path = Path(
        str(
            path
        )
        + ".tmp"
    )

    dataframe.to_csv(
        temp_path,
        index=False
    )

    os.replace(
        temp_path,
        path
    )


# ---------------------------------------------------------
# Load previous V8 progress
# ---------------------------------------------------------

existing_transactions = read_csv_if_exists(
    OUTPUT_CSV,
    TRANSACTION_COLUMNS
)

existing_fallback = read_csv_if_exists(
    FALLBACK_CSV,
    FALLBACK_COLUMNS
)

existing_checkpoint = read_csv_if_exists(
    CHECKPOINT_CSV,
    CHECKPOINT_COLUMNS
)

if len(
    existing_checkpoint
):
    existing_checkpoint = (
        existing_checkpoint
        .drop_duplicates(
            subset=[
                "source_key"
            ],
            keep="last"
        )
        .reset_index(
            drop=True
        )
    )

completed_keys = set(
    existing_checkpoint[
        "source_key"
    ].dropna().astype(
        str
    )
)

remaining_items = [
    (
        source_key,
        pdf_path
    )
    for source_key, pdf_path
    in pdf_items
    if source_key not in completed_keys
]

print("=" * 72)
print("V8 RESUME STATUS")
print("=" * 72)

print(
    "Total source PDFs:",
    len(
        pdf_items
    )
)

print(
    "Already checkpointed:",
    len(
        completed_keys
    )
)

print(
    "Remaining:",
    len(
        remaining_items
    )
)

print(
    "Existing transaction rows:",
    len(
        existing_transactions
    )
)

print(
    "Existing fallback PDFs:",
    len(
        existing_fallback
    )
)

if MAX_NEW_PDFS_THIS_RUN is not None:
    print(
        "Maximum new PDFs this run:",
        MAX_NEW_PDFS_THIS_RUN
    )

print()


# ---------------------------------------------------------
# Current-session additions
# ---------------------------------------------------------

new_transaction_frames = []
new_fallback_rows = []
new_checkpoint_rows = []

processed_this_run = 0


# ----------------------------------------------------------------------
# Build the current checkpoint DataFrame from in-memory checkpoint
# records.
# ----------------------------------------------------------------------
def current_checkpoint_frame():
    additions = pd.DataFrame(
        new_checkpoint_rows
    )

    checkpoint = concat_safe([
        existing_checkpoint,
        additions,
    ])

    if checkpoint.empty:
        return pd.DataFrame(
            columns=CHECKPOINT_COLUMNS
        )

    for column in CHECKPOINT_COLUMNS:
        if column not in checkpoint.columns:
            checkpoint[
                column
            ] = None

    checkpoint = (
        checkpoint[
            CHECKPOINT_COLUMNS
        ]
        .drop_duplicates(
            subset=[
                "source_key"
            ],
            keep="last"
        )
        .sort_values(
            by=[
                "source_year",
                "source_pdf",
            ],
            kind="stable"
        )
        .reset_index(
            drop=True
        )
    )

    return checkpoint


# ----------------------------------------------------------------------
# Build the current transaction DataFrame from already persisted rows
# plus newly parsed rows in this session.
# ----------------------------------------------------------------------
def current_transaction_frame():
    current = concat_safe(
        [
            existing_transactions,
            *new_transaction_frames,
        ]
    )

    if current.empty:
        return pd.DataFrame(
            columns=TRANSACTION_COLUMNS
        )

    for column in TRANSACTION_COLUMNS:
        if column not in current.columns:
            current[
                column
            ] = None

    # Resume safety:
    # a PDF may run again if Colab stopped after a data save but before
    # its next checkpoint save. This stable key removes that repeat.
    current = (
        current[
            TRANSACTION_COLUMNS
        ]
        .drop_duplicates(
            subset=[
                "source_year",
                "filing_id",
                "transaction_number_in_filing",
            ],
            keep="last"
        )
        .sort_values(
            by=[
                "source_year",
                "filing_id",
                "transaction_number_in_filing",
            ],
            kind="stable"
        )
        .reset_index(
            drop=True
        )
    )

    return current


# ----------------------------------------------------------------------
# Build the current fallback DataFrame from persisted and newly
# discovered PDFs needing another extraction method.
# ----------------------------------------------------------------------
def current_fallback_frame():
    additions = pd.DataFrame(
        new_fallback_rows
    )

    current = concat_safe([
        existing_fallback,
        additions,
    ])

    if current.empty:
        current = pd.DataFrame(
            columns=FALLBACK_COLUMNS
        )

    for column in FALLBACK_COLUMNS:
        if column not in current.columns:
            current[
                column
            ] = None

    # Keep only the latest row for each source PDF.
    current = (
        current[
            FALLBACK_COLUMNS
        ]
        .drop_duplicates(
            subset=[
                "source_year",
                "source_pdf",
            ],
            keep="last"
        )
        .reset_index(
            drop=True
        )
    )

    # If a PDF was reprocessed successfully before its checkpoint was written,
    # remove any stale fallback entry.
    checkpoint = current_checkpoint_frame()

    successful_keys = set(
        checkpoint.loc[
            checkpoint[
                "parser_status"
            ]
            == "geometry_v8",
            "source_key"
        ].astype(
            str
        )
    )

    if successful_keys:
        fallback_keys = (
            current[
                "source_year"
            ].astype(
                str
            )
            + "/"
            + current[
                "source_pdf"
            ].astype(
                str
            )
        )

        current = current[
            ~fallback_keys.isin(
                successful_keys
            )
        ].copy()

    current = (
        current
        .sort_values(
            by=[
                "source_year",
                "source_pdf",
            ],
            kind="stable"
        )
        .reset_index(
            drop=True
        )
    )

    return current


# ----------------------------------------------------------------------
# Persist transactions and fallback diagnostics first, then write the
# checkpoint last so a PDF is not marked completed before its data is
# safely on Drive.
# ----------------------------------------------------------------------
def save_v8_progress():
    """
    Save data first; checkpoint LAST.

    That ordering prevents a PDF from being marked complete before its
    transaction/fallback output has been persisted.
    """
    transactions = current_transaction_frame()
    fallback = current_fallback_frame()
    checkpoint = current_checkpoint_frame()

    atomic_write_csv(
        transactions,
        OUTPUT_CSV
    )

    atomic_write_csv(
        fallback,
        FALLBACK_CSV
    )

    # Checkpoint last.
    atomic_write_csv(
        checkpoint,
        CHECKPOINT_CSV
    )

    return (
        transactions,
        fallback,
        checkpoint,
    )


# ----------------------------------------------------------------------
# Convert parser diagnostics into a human-readable explanation of why
# a PDF produced no born-digital geometry transactions and needs
# fallback extraction.
# ----------------------------------------------------------------------
def fallback_reason(
    diagnostic,
    geometry_error
):
    if geometry_error:
        return (
            "geometry parser error; "
            "needs OCR/MinerU/AI fallback: "
            + geometry_error
        )

    if diagnostic[
        "pages_with_words"
    ] == 0:
        return (
            "scanned/image PDF or no embedded text; "
            "needs OCR/MinerU/AI fallback"
        )

    if diagnostic[
        "pages_with_geometry"
    ] == 0:
        return (
            "embedded text exists but PTR table geometry "
            "was not detected; needs OCR/MinerU/AI fallback"
        )

    return (
        "table geometry found but no usable transaction rows; "
        "needs OCR/MinerU/AI fallback"
    )


# ---------------------------------------------------------
# Process remaining PDFs
# ---------------------------------------------------------

for overall_number, (
    source_key,
    pdf_path
) in enumerate(
    pdf_items,
    start=1
):
    if source_key in completed_keys:
        continue

    if (
        MAX_NEW_PDFS_THIS_RUN is not None
        and processed_this_run
        >= MAX_NEW_PDFS_THIS_RUN
    ):
        print()
        print(
            "Reached MAX_NEW_PDFS_THIS_RUN."
        )
        print(
            "Saving a clean checkpoint and stopping this run."
        )
        break

    geometry_error = ""

    try:
        parsed, diagnostic = parse_pdf_geometry_v8(
            pdf_path
        )

    except Exception as exc:
        parsed = pd.DataFrame()

        diagnostic = {
            "pages_with_geometry": 0,
            "pages_with_words": 0,
            "table_pages": 0,
            "recovered_pages": "",
            "suspected_missed_pages": "",
            "politician": "",
            "member_status": "",
            "state_district": "",
        }

        geometry_error = repr(
            exc
        )

    year = pdf_path.parent.name
    filing_id = pdf_path.stem

    transaction_count = len(
        parsed
    )

    if transaction_count:
        parser_status = "geometry_v8"

        new_transaction_frames.append(
            parsed.drop(
                columns=[
                    "_segments"
                ],
                errors="ignore"
            )
        )

    else:
        parser_status = "needs_fallback"

        reason = fallback_reason(
            diagnostic,
            geometry_error
        )

        new_fallback_rows.append({
            "filing_id": filing_id,
            "politician": diagnostic.get(
                "politician",
                ""
            ),
            "member_status": diagnostic.get(
                "member_status",
                ""
            ),
            "state_district": diagnostic.get(
                "state_district",
                ""
            ),
            "source_pdf": pdf_path.name,
            "source_year": year,
            "needs_review": True,
            "original_pdf_url": make_original_pdf_url(
                year,
                filing_id
            ),
            "review_level": "high",
            "review_reason": reason,
            "pages_with_words": diagnostic[
                "pages_with_words"
            ],
            "pages_with_geometry": diagnostic[
                "pages_with_geometry"
            ],
            "table_pages": diagnostic[
                "table_pages"
            ],
            "recovered_pages": diagnostic.get(
                "recovered_pages",
                ""
            ),
            "suspected_missed_pages": diagnostic.get(
                "suspected_missed_pages",
                ""
            ),
            "geometry_error": geometry_error,
        })

    new_checkpoint_rows.append({
        "source_key": source_key,
        "filing_id": filing_id,
        "source_pdf": pdf_path.name,
        "source_year": year,
        "parser_status": parser_status,
        "transaction_rows": transaction_count,
        "pages_with_words": diagnostic[
            "pages_with_words"
        ],
        "pages_with_geometry": diagnostic[
            "pages_with_geometry"
        ],
        "table_pages": diagnostic[
            "table_pages"
        ],
        "recovered_pages": diagnostic.get(
            "recovered_pages",
            ""
        ),
        "suspected_missed_pages": diagnostic.get(
            "suspected_missed_pages",
            ""
        ),
        "geometry_error": geometry_error,
        "processed_at": datetime.now().isoformat(
            timespec="seconds"
        ),
    })

    processed_this_run += 1

    print(
        f"[{year}] "
        f"[overall {overall_number}/{len(pdf_items)}] "
        f"[this run {processed_this_run}] "
        f"{pdf_path.name} | "
        f"rows={transaction_count} | "
        f"{parser_status}"
    )

    if (
        processed_this_run
        % SAVE_EVERY
        == 0
    ):
        (
            saved_transactions,
            saved_fallback,
            saved_checkpoint,
        ) = save_v8_progress()

        print(
            "  checkpoint saved | "
            f"transactions={len(saved_transactions):,} | "
            f"fallback={len(saved_fallback):,} | "
            f"completed={len(saved_checkpoint):,}"
        )


# Always save at the end of a clean run / MAX stop.
(
    transactions_v8,
    fallback_v8,
    checkpoint_v8,
) = save_v8_progress()


print()
print("=" * 72)
print("V8 RUN COMPLETE / PAUSED SAFELY")
print("=" * 72)

print(
    "Processed this run:",
    processed_this_run
)

print(
    "Checkpointed total:",
    len(
        checkpoint_v8
    )
)

print(
    "Total source PDFs:",
    len(
        pdf_items
    )
)

print(
    "Still remaining:",
    max(
        0,
        len(
            pdf_items
        )
        - len(
            checkpoint_v8
        )
    )
)

print(
    "Transaction rows saved:",
    len(
        transactions_v8
    )
)

print(
    "Fallback PDFs saved:",
    len(
        fallback_v8
    )
)

print()
print(
    "You can stop here and rerun the notebook later."
)