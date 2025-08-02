import os
import argparse
from pathlib import Path
import pandas as pd

def split_excel_file(filename: str, linescount: int = 20, output_dir: str = 'results'):
    # Resolve paths
    input_path = Path(filename)
    if not input_path.is_file():
        raise FileNotFoundError(f"Input file not found: {filename}")

    # Prepare output directory
    output_path = Path(output_dir)
    output_path.mkdir(parents=True, exist_ok=True)

    # Base name without extension
    base_name = input_path.stem

    # Read all sheets
    excel = pd.ExcelFile(input_path)
    for sheet_name in excel.sheet_names:
        # Read sheet into DataFrame
        df = pd.read_excel(excel, sheet_name=sheet_name)
        total_rows = len(df)
        if total_rows == 0:
            continue

        # Chunk and save
        for start in range(0, total_rows, linescount):
            end = min(start + linescount, total_rows)
            lines_from = start + 1
            lines_to = end

            chunk = df.iloc[start:end]

            # Construct output filename
            safe_sheet = sheet_name.replace(' ', '_')
            out_filename = f"{base_name}_{safe_sheet}_{lines_from}_{lines_to}.xlsx"
            out_file_path = output_path / out_filename

            # Save chunk with header
            with pd.ExcelWriter(out_file_path, engine='openpyxl') as writer:
                chunk.to_excel(writer, index=False, sheet_name=sheet_name)

            print(f"Written: {out_file_path}")


def main():
    parser = argparse.ArgumentParser(
        description='Split an Excel file into multiple files per sheet based on a row limit.'
    )
    parser.add_argument(
        'filename',
        help='Path to the source Excel file (relative or absolute).'
    )
    parser.add_argument(
        '-l', '--linescount',
        type=int,
        default=20,
        help='Number of data rows per split file (default: 20).'
    )
    parser.add_argument(
        '-o', '--output',
        default='results',
        help='Directory to place the output files (default: ./results).'
    )

    args = parser.parse_args()
    split_excel_file(args.filename, args.linescount, args.output)

if __name__ == '__main__':
    main()
