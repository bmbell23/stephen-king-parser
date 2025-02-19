import os
import csv
from typing import List
from ..core.work import Work

class CSVExporter:
    @staticmethod
    def create_excel_hyperlink(url: str, text: str) -> str:
        """Create a clean Excel hyperlink formula."""
        return f'=HYPERLINK("{url}", "{text}")'

    @staticmethod
    def export_works(works: List[Work], base_dir: str) -> str:
        """Export works to CSV file with hyperlinked titles."""
        # Create csv directory if it doesn't exist
        csv_dir = os.path.join(base_dir, 'csv')
        os.makedirs(csv_dir, exist_ok=True)

        # Create the CSV file path
        csv_file = os.path.join(csv_dir, 'stephen_king_works.csv')

        with open(csv_file, 'w', newline='', encoding='utf-8') as f:
            writer = csv.writer(f)

            # Write header
            writer.writerow(['Read', 'Published Date', 'Title', 'Type', 'Collection'])

            # Write data
            for work in works:
                title_cell = CSVExporter.create_excel_hyperlink(work.link, work.title) if work.link else work.title
                writer.writerow([
                    '',  # Read status (empty by default)
                    work.published_date,
                    title_cell,
                    work.work_type,
                    work.available_in if work.available_in else ''
                ])

        return csv_file
