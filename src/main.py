from stephen_king_parser.core.parser import KingWorksParser
import os

def main():
    parser = KingWorksParser()
    works = parser.parse_works()

    print(f"\nFound {len(works)} works")

    # Get the current working directory
    base_dir = os.getcwd()
    output_file = parser.export_to_csv(works, base_dir)
    print(f"\nData exported to {output_file}")

if __name__ == "__main__":
    main()
