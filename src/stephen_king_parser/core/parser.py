import requests
from bs4 import BeautifulSoup
from typing import List
import csv
import os
from datetime import datetime
from urllib.parse import urljoin
from ..models.work import Work

class KingWorksParser:
    """Parser for extracting and organizing Stephen King's literary works"""

    BASE_URL = "https://stephenking.com"
    WORKS_URL = f"{BASE_URL}/works/"

    def __init__(self):
        self.session = requests.Session()

    def extract_collection_info(self, url: str) -> tuple[str, str]:
        """Extract collection information from a work's page."""
        try:
            response = self.session.get(url)
            if not response.ok:
                return ("", "")

            soup = BeautifulSoup(response.text, 'html.parser')

            # Find the "Available In" section
            available_in = soup.find('h2', string='Available In')
            if not available_in:
                return ("", "")

            # Find the collection link in the section following the "Available In" header
            collection_section = available_in.find_next('div', class_='grid-content')
            if not collection_section:
                return ("", "")

            # Find the collection link
            collection_link = collection_section.find('a', class_='text-link')
            if not collection_link:
                return ("", "")

            # Extract collection name and URL
            collection_name = collection_link.text.strip()
            collection_url = collection_link.get('href', '')
            if collection_url and not collection_url.startswith('http'):
                collection_url = urljoin(self.BASE_URL, collection_url)

            # Create the Excel-style hyperlink format
            if collection_url:
                collection_hyperlink = f'=HYPERLINK("{collection_url}", "{collection_name}")'
                return (collection_name, collection_hyperlink)

            return ("", "")

        except Exception as e:
            print(f"Error extracting collection info from {url}: {str(e)}")
            return ("", "")

    def parse_works(self) -> List[Work]:
        """Parse Stephen King works from the website."""
        response = self.session.get(self.WORKS_URL)
        response.raise_for_status()
        soup = BeautifulSoup(response.text, 'html.parser')

        works = soup.find_all('a', class_='row work')
        works_data = []

        for work in works:
            title = work.find('div', class_='works-title').text.strip()
            published = work.get('data-date', '')
            work_type = work.find('div', class_='works-type').text.strip()
            url = self.BASE_URL + work.get('href', '')

            # Get collection information
            _, collection_hyperlink = self.extract_collection_info(url)

            works_data.append(Work(
                title=title,
                published_date=published,
                work_type=work_type,
                url=url,
                available_in=collection_hyperlink
            ))

            print(f"Processed: {title}")

        return works_data

    @staticmethod
    def create_excel_hyperlink(url: str, text: str) -> str:
        """Create Excel hyperlink formula."""
        return f'=HYPERLINK("{url}", "{text}")'

    def export_to_csv(self, works: List[Work], base_dir: str):
        """Export works data to CSV file with hyperlinked titles."""
        csv_dir = os.path.join(base_dir, 'csv')
        os.makedirs(csv_dir, exist_ok=True)

        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f'stephen_king_works_{timestamp}.csv'
        filepath = os.path.join(csv_dir, filename)

        with open(filepath, 'w', newline='', encoding='utf-8') as file:
            writer = csv.writer(file)
            writer.writerow(['Read', 'Owned', 'Published Date', 'Title', 'Type', 'Collection'])

            for work in works:
                title_cell = self.create_excel_hyperlink(work.url, work.title)
                writer.writerow([
                    '',  # Read status
                    '',  # Owned status
                    work.published_date,
                    title_cell,
                    work.work_type,
                    work.available_in  # Collection hyperlink
                ])

        return filepath

def main():
    parser = KingWorksParser()
    works = parser.parse_works()

    print(f"Found {len(works)} works")
    print("\nFirst 5 works:")
    for work in works[:5]:
        print(f"- {work.title} ({work.published_date}) - {work.work_type}")

    output_dir = "output"
    os.makedirs(output_dir, exist_ok=True)
    csv_file = parser.export_to_csv(works, output_dir)
    print(f"\nData exported to {csv_file}")

if __name__ == "__main__":
    main()
