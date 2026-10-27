from pathlib import Path
import argparse

from profiler import profile_workbook
from exporter import export_profile


def main():
    parser = argparse.ArgumentParser(
        description="VELOOP Data Profiling Project"
    )

    parser.add_argument(
        "--input",
        default="data/VELoop_Rewards_Data_01.xlsx",
        help="Path to input Excel file"
    )

    parser.add_argument(
        "--output",
        default="output/VELOOP_Profile_Output.xlsx",
        help="Path to output Excel report"
    )

    args = parser.parse_args()

    input_file = Path(args.input)
    output_file = Path(args.output)

    if not input_file.exists():
        raise FileNotFoundError(
            f"Input file not found: {input_file.resolve()}"
        )

    print("=" * 60)
    print("VELOOP DATA PROFILING STARTED")
    print("=" * 60)

    results = profile_workbook(input_file)

    output_file.parent.mkdir(
        parents=True,
        exist_ok=True
    )

    export_profile(results, output_file)

    print("\nDATA PROFILING COMPLETED SUCCESSFULLY!")
    print(f"Output report: {output_file.resolve()}")


if __name__ == "__main__":
    main()
