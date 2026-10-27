"""Export profiling results to a formatted Excel workbook."""

from pathlib import Path
import pandas as pd


QUALITY_COLUMNS = [
    "dataset",
    "field",
    "issue",
    "records_affected",
    "severity",
]


def export_profile(results, output_path):
    """Write profiling results into four Excel worksheets."""
    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)

    tables = {
        "Dataset_Profile": pd.DataFrame(results["dataset_rows"]),
        "Column_Profile": pd.DataFrame(results["column_rows"]),
        "Category_Distributions": pd.DataFrame(
            results["category_rows"]
        ),
        "Quality_Flags": pd.DataFrame(
            results["quality_rows"],
            columns=QUALITY_COLUMNS,
        ),
    }

    with pd.ExcelWriter(
        output_path, engine="openpyxl"
    ) as writer:
        for sheet_name, table in tables.items():
            table.to_excel(
                writer,
                sheet_name=sheet_name,
                index=False,
            )

            worksheet = writer.sheets[sheet_name]
            worksheet.freeze_panes = "A2"
            worksheet.auto_filter.ref = worksheet.dimensions

            for column_cells in worksheet.columns:
                values = [
                    len(str(cell.value or ""))
                    for cell in list(column_cells)[:100]
                ]

                max_length = max(values) + 2 if values else 12
                column_letter = column_cells[0].column_letter

                worksheet.column_dimensions[
                    column_letter
                ].width = min(max(max_length, 12), 45)

    print(f"\nProfile report saved to: {output_path.resolve()}")
