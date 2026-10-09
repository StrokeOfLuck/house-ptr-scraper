import json
import shutil
from datetime import datetime

import pandas as pd

from config import (
    PUBLIC_ROOT,
    PUBLIC_LATEST_CSV,
    PUBLIC_METADATA_JSON,
    V81_CHECKPOINT,
    V81_FALLBACK,
    V82_TRANSACTIONS,
    ensure_folders,
)

PUBLIC_WEB_CSV = PUBLIC_ROOT / "house_ptr_transactions_web.csv"


def run() -> None:
    ensure_folders()

    if not V82_TRANSACTIONS.exists():
        raise FileNotFoundError(
            f"Final V8.2 CSV does not exist:\n{V82_TRANSACTIONS}"
        )

    # Keep a stable full-data copy for downstream analysis.
    shutil.copy2(
        V82_TRANSACTIONS,
        PUBLIC_LATEST_CSV,
    )

    transactions = pd.read_csv(
        V82_TRANSACTIONS,
        low_memory=False,
    )

    # A smaller public-facing table for the website.
    #
    # Only columns that exist are selected, so this remains tolerant of
    # minor schema changes later.
    web_columns = [
        "filing_id",
        "politician",
        "member_status",
        "state_district",
        "source_year",
        "transaction_number_in_filing",
        "owner",
        "asset_v8_2_cleaned",
        "ticker_v8_2_cleaned",
        "asset_type",
        "transaction_type",
        "transaction_date",
        "notification_date",
        "amount_category",
        "amount_min",
        "amount_max",
        "amount_exact",
        "filing_status",
        "review_level",
        "needs_review",
        "original_pdf_url",
    ]

    web_columns = [
        column
        for column in web_columns
        if column in transactions.columns
    ]

    web = transactions[web_columns].copy()

    # Newest transactions first when the date parses cleanly.
    if "transaction_date" in web.columns:
        web["_sort_date"] = pd.to_datetime(
            web["transaction_date"],
            errors="coerce",
        )

        web = (
            web
            .sort_values(
                by=["_sort_date", "politician"],
                ascending=[False, True],
                na_position="last",
                kind="stable",
            )
            .drop(columns="_sort_date")
            .reset_index(drop=True)
        )

    web.to_csv(
        PUBLIC_WEB_CSV,
        index=False,
    )

    source_pdfs_accounted_for = None
    if V81_CHECKPOINT.exists():
        checkpoint = pd.read_csv(
            V81_CHECKPOINT,
            low_memory=False,
        )
        source_pdfs_accounted_for = int(
            checkpoint["source_key"].nunique()
        )

    fallback_pdfs = None
    if V81_FALLBACK.exists():
        fallback = pd.read_csv(
            V81_FALLBACK,
            low_memory=False,
        )
        fallback_pdfs = int(len(fallback))

    metadata = {
        "updated_at": datetime.now().astimezone().isoformat(
            timespec="seconds"
        ),
        "transactions": int(len(transactions)),
        "public_table_rows": int(len(web)),
        "public_table_columns": int(len(web.columns)),
        "source_pdfs_accounted_for": source_pdfs_accounted_for,
        "fallback_pdfs": fallback_pdfs,
        "full_csv_filename": PUBLIC_LATEST_CSV.name,
        "web_csv_filename": PUBLIC_WEB_CSV.name,
    }

    PUBLIC_METADATA_JSON.write_text(
        json.dumps(metadata, indent=2),
        encoding="utf-8",
    )

    print("Published full stable CSV:")
    print(PUBLIC_LATEST_CSV)

    print()
    print("Published website CSV:")
    print(PUBLIC_WEB_CSV)
    print(f"{len(web):,} rows x {len(web.columns):,} columns")

    print()
    print("Published metadata:")
    print(PUBLIC_METADATA_JSON)
    print(json.dumps(metadata, indent=2))


if __name__ == "__main__":
    run()
