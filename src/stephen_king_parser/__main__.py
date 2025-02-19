import argparse
import os

from .core.parser import KingWorksParser
from .exporters.csv_exporter import CSVExporter


def main():
    parser = argparse.ArgumentParser(
        description="Parse Stephen King works from his official website"
    )
    parser.add_argument(
        "--output",
        "-o",
        default=".",
        help='Output directory for the CSV file (a "csv" subdirectory will be created)',
    )
    args = parser.parse_args()

    # Ensure output directory exists
    os.makedirs(args.output, exist_ok=True)

    # Parse works
    king_parser = KingWorksParser()
    works = king_parser.parse_works()

    if works:
        # Export to CSV
        output_file = CSVExporter.export_works(works, args.output)
        print(f"\nSuccessfully exported {len(works)} works to {output_file}")
    else:
        print("No works were processed successfully.")


if __name__ == "__main__":
    main()
