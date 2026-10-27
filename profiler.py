"""Functions for profiling worksheets and finding data-quality issues."""

from pathlib import Path
import pandas as pd

IDENTIFIER_NAMES = {
    "user_id",
    "event_id",
    "transaction_id",
    "withdrawal_id",
    "referral_id",
}


def _profile_sheet(sheet_name, df, results):
    dataset_rows = results["dataset_rows"]
    column_rows = results["column_rows"]
    category_rows = results["category_rows"]
    quality_rows = results["quality_rows"]

    total_rows = len(df)
    total_columns = len(df.columns)
    total_cells = df.size
    missing_cells = int(df.isna().sum().sum())
    missing_percentage = (
        round(missing_cells / total_cells * 100, 2)
        if total_cells > 0 else 0
    )
    duplicate_rows = int(df.duplicated().sum())

    date_columns = [
        col for col in df.columns
        if any(
            keyword in str(col).lower()
            for keyword in ("date", "time", "_at")
        )
    ]

    parsed_dates = []
    for col in date_columns:
        try:
            parsed = pd.to_datetime(
                df[col], errors="coerce", utc=True
            )
            valid_dates = parsed.dropna()
            if not valid_dates.empty:
                parsed_dates.extend(valid_dates.tolist())
        except (ValueError, TypeError, OverflowError):
            continue

    date_min = (
        min(parsed_dates).strftime("%Y-%m-%d")
        if parsed_dates else ""
    )
    date_max = (
        max(parsed_dates).strftime("%Y-%m-%d")
        if parsed_dates else ""
    )

    dataset_rows.append({
        "dataset": sheet_name,
        "records": total_rows,
        "columns": total_columns,
        "duplicate_rows": duplicate_rows,
        "missing_cells": missing_cells,
        "missing_pct": missing_percentage,
        "date_min_utc": date_min,
        "date_max_utc": date_max,
        "date_like_columns": ", ".join(map(str, date_columns)),
    })

    # Column-level profiling and data-quality checks
    for col in df.columns:
        series = df[col]
        non_null_count = int(series.notna().sum())
        missing_count = int(series.isna().sum())
        unique_count = int(series.nunique(dropna=True))
        duplicate_values = int(series.dropna().duplicated().sum())
        missing_pct = round(series.isna().mean() * 100, 2)

        column_rows.append({
            "dataset": sheet_name,
            "field": col,
            "dtype": str(series.dtype),
            "non_null": non_null_count,
            "missing": missing_count,
            "missing_pct": missing_pct,
            "unique_non_null": unique_count,
            "duplicate_values_beyond_first": duplicate_values,
        })

        if missing_count > 0:
            quality_rows.append({
                "dataset": sheet_name,
                "field": col,
                "issue": "Missing values",
                "records_affected": missing_count,
                "severity": (
                    "Review field requirements; not every missing "
                    "value is necessarily an error"
                ),
            })

        if (
            pd.api.types.is_object_dtype(series.dtype)
            or isinstance(series.dtype, pd.CategoricalDtype)
        ):
            values = series.dropna().astype(str)

            whitespace_count = int(
                (values.str.strip() != values).sum()
            )
            if whitespace_count > 0:
                quality_rows.append({
                    "dataset": sheet_name,
                    "field": col,
                    "issue": "Leading/trailing whitespace",
                    "records_affected": whitespace_count,
                    "severity": "Low",
                })

            distinct_values = values.unique()
            lowercase_values = [v.lower() for v in distinct_values]
            case_variants = (
                len(distinct_values) - len(set(lowercase_values))
            )

            if case_variants > 0:
                quality_rows.append({
                    "dataset": sheet_name,
                    "field": col,
                    "issue": "Possible capitalization variants",
                    "records_affected": case_variants,
                    "severity": "Review controlled taxonomy",
                })

    # Categorical distributions
    categorical_columns = df.select_dtypes(
        include=["object", "category"]
    ).columns

    for col in categorical_columns:
        counts = (
            df[col]
            .fillna("[MISSING]")
            .astype(str)
            .value_counts(dropna=False)
        )

        for category, count in counts.items():
            share_pct = (
                round(count / len(df) * 100, 2)
                if len(df) > 0 else 0
            )
            category_rows.append({
                "dataset": sheet_name,
                "field": col,
                "category": category,
                "count": int(count),
                "share_pct": share_pct,
            })

    # Fully duplicated rows
    if duplicate_rows > 0:
        quality_rows.append({
            "dataset": sheet_name,
            "field": "*",
            "issue": "Fully duplicated rows",
            "records_affected": duplicate_rows,
            "severity": "High",
        })

    # Identifier checks
    for col in df.columns:
        if str(col).lower() in IDENTIFIER_NAMES:
            null_count = int(df[col].isna().sum())
            duplicate_count = int(df[col].dropna().duplicated().sum())

            if null_count > 0 or duplicate_count > 0:
                quality_rows.append({
                    "dataset": sheet_name,
                    "field": col,
                    "issue": "Identifier nulls/duplicates",
                    "records_affected": null_count + duplicate_count,
                    "severity": "High",
                })


def profile_workbook(input_path):
    """Profile all worksheets in an Excel workbook."""
    input_path = Path(input_path)

    results = {
        "dataset_rows": [],
        "column_rows": [],
        "category_rows": [],
        "quality_rows": [],
    }

    with pd.ExcelFile(input_path) as workbook:
        sheet_names = workbook.sheet_names

    for sheet_name in sheet_names:
        print(f"\nProfiling worksheet: {sheet_name}")
        df = pd.read_excel(input_path, sheet_name=sheet_name)
        _profile_sheet(sheet_name, df, results)

    return results
