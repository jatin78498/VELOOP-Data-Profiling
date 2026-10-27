# VELOOP Rewards Data Profiling

A Python project for analysing Excel datasets and identifying
potential data-quality issues.

## Features

- Profiles every worksheet in an Excel workbook.
- Calculates row and column counts.
- Identifies missing values and duplicate rows.
- Summarizes unique values and categorical distributions.
- Checks leading/trailing whitespace.
- Detects possible capitalization inconsistencies.
- Checks selected identifier columns.
- Exports results to a formatted Excel report.

## Technologies Used

- Python
- Pandas
- OpenPyXL
- Excel

## Project Structure

- `main.py` - Main entry point
- `profiler.py` - Data profiling and quality checks
- `exporter.py` - Excel report generation
- `requirements.txt` - Required libraries
- `data/` - Input dataset folder
- `output/` - Generated report folder

## Setup

Install dependencies:

```bash
python -m pip install -r requirements.txt
