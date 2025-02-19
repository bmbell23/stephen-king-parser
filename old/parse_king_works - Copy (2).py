"""
Module for parsing and organizing Stephen King's literary works from his official website.
Provides functionality to scrape, process, and export work details to various formats including CSV.
"""
import csv
import requests
from bs4 import BeautifulSoup
import time
from datetime import datetime
import re
from typing import Dict, List, Tuple, Optional
from dataclasses import dataclass
from urllib.parse import urljoin
import logging
import threading
from collections import defaultdict
from concurrent.futures import ThreadPoolExecutor
import concurrent.futures

# Set up logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

@dataclass
class WorkData:
    """Data class to store information about a work"""
    title: str
    cleaned_title: str
    link: str
    published_date: str
    work_type: str
    formats: str
    available_in: str
    available_in_link: str

class RequestManager:
    """Manages HTTP requests with rate limiting"""
    def __init__(self, rate_limit: float = 1.0):
        self.rate_limit = rate_limit
        self.last_request_time = 0

    def get(self, url: str) -> Optional[requests.Response]:
        """Make a GET request with rate limiting"""
        current_time = time.time()
        time_since_last_request = current_time - self.last_request_time

        if time_since_last_request < self.rate_limit:
            time.sleep(self.rate_limit - time_since_last_request)

        try:
            response = requests.get(url)
            self.last_request_time = time.time()
            return response
        except requests.RequestException as e:
            logger.error(f"Error making request to {url}: {e}")
            return None

class KingWorksParser:
    """
    Parser for extracting and organizing Stephen King's literary works from his official website.

    This class handles fetching, parsing, and organizing data about Stephen King's works,
    including books, stories, collections, and other publications. It manages web scraping,
    data cleaning, and storage of work details including titles, publication dates, formats,
    and collection relationships.
    """
    BASE_URL = "https://www.stephenking.com"
    WORKS_URL = f"{BASE_URL}/works/"
    MAX_WORKERS = 5  # Limit concurrent threads

    def __init__(self):
        self.request_manager = RequestManager(rate_limit=1.0)  # 1 request per second
        self.works_dict = {}
        self.collection_dates = {}
        self.processed_urls = set()
        self.url_lock = threading.Lock()
        self.data_lock = threading.Lock()

    def clean_title(self, title: str) -> str:
        """
        Clean the title by removing special characters and normalizing format.

        Args:
            title (str): Original title

        Returns:
            str: Cleaned title
        """
        # Remove special characters but keep basic punctuation
        cleaned = re.sub(r'[^\w\s\-\'.,]', '', title)
        # Normalize whitespace
        cleaned = ' '.join(cleaned.split())
        return cleaned

    def is_url_processed(self, url: str) -> bool:
        """Thread-safe check if URL has been processed"""
        with self.url_lock:
            return url in self.processed_urls

    def mark_url_processed(self, url: str):
        """Thread-safe marking of URL as processed"""
        with self.url_lock:
            self.processed_urls.add(url)

    @staticmethod
    def remove_parenthetical_suffix(title: str) -> str:
        """
        Remove parenthetical suffixes from a title string.

        Args:
            title (str): The title string to process

        Returns:
            str: The title with any trailing parenthetical content removed and whitespace trimmed
        """
        # Use regex to remove any text within parentheses at the end of the string
        # \s* matches optional whitespace before and after the parentheses
        # \([^)]*\) matches anything between parentheses
        # $ ensures we only match parentheses at the end of the string
        return re.sub(r'\s*\([^)]*\)\s*$', '', title).strip()

    @staticmethod
    def create_excel_hyperlink(url: str, text: str) -> str:
        """
        Create an Excel-compatible hyperlink formula.

        Args:
            url (str): The URL for the hyperlink
            text (str): The display text for the hyperlink

        Returns:
            str: An Excel formula string in the format '=HYPERLINK("url", "text")'
        """
        # Double up any quotes in the text and URL to escape them for Excel
        text = text.replace('"', '""')
        url = url.replace('"', '""')
        # Create Excel HYPERLINK formula
        return f'=HYPERLINK("{url}", "{text}")'

    @staticmethod
    def convert_to_datetime(date_str: str) -> datetime:
        """
        Convert a date string to a datetime object with error handling.

        Args:
            date_str (str): Date string in 'YYYY-MM-DD' format

        Returns:
            datetime: Parsed datetime object, or far-future date (9999-12-31) for invalid inputs
        """
        # Return far-future date for invalid or empty input
        if not date_str or not isinstance(date_str, str):
            return datetime(9999, 12, 31)

        try:
            # Attempt to parse the date string into a datetime object
            return datetime.strptime(date_str.strip(), '%Y-%m-%d')
        except (ValueError, AttributeError):
            # Return far-future date if parsing fails
            return datetime(9999, 12, 31)

    @staticmethod
    def merge_format_strings(existing: str, new: str) -> str:
        """
        Merge two comma-separated format strings, removing duplicates.

        Args:
            existing (str): Existing format string
            new (str): New format string to merge

        Returns:
            str: Combined, sorted, unique format string
        """
        # Handle cases where either string is empty
        if not existing:
            return new
        if not new:
            return existing

        # Split strings into sets to remove duplicates
        formats = set(format.strip() for format in existing.split(','))
        formats.update(format.strip() for format in new.split(','))
        # Join formats back into sorted, comma-separated string
        return ', '.join(sorted(formats))

    def extract_collection_info(self, work) -> tuple[str, str]:
        """
        Extract collection information from a work's dedicated page.

        Args:
            work: BeautifulSoup element representing the work

        Returns:
            tuple[str, str]: (collection_name, collection_url)
        """
        try:
            # Get the work's specific URL
            work_url = work.get('href', '')
            if not work_url:
                return ("", "")

            # Make sure we have a full URL
            if not work_url.startswith('http'):
                work_url = urljoin(self.BASE_URL, work_url)

            # Fetch the work's dedicated page
            response = self.request_manager.get(work_url)
            if not response:
                return ("", "")

            # Parse the page
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
            print(f"Error extracting collection info from {work_url}: {str(e)}")
            return ("", "")

    def extract_available_formats(self, link: str) -> str:
        """
        Extract available formats from a work's page.
        """
        try:
            if self.is_url_processed(link):
                return ""

            response = self.request_manager.get(link)
            if not response:
                return ""

            self.mark_url_processed(link)

            soup = BeautifulSoup(response.text, 'html.parser')
            formats = set()

            # Enhanced format detection
            format_indicators = {
                'Hardcover': ['hardcover', 'hard cover', 'hard-cover', 'hardback'],
                'Paperback': ['paperback', 'soft cover', 'soft-cover', 'trade paperback', 'mass market'],
                'Ebook': ['ebook', 'e-book', 'kindle', 'digital', 'nook', 'electronic'],
                'Audiobook': ['audiobook', 'audio book', 'audible', 'audio'],
                'Movie': ['movie', 'film', 'feature film', 'motion picture'],
                'Miniseries': ['tv series', 'television series', 'miniseries', 'mini-series', 'mini series']
            }

            # Check all possible containers
            containers = soup.find_all(['div', 'section', 'span', 'p', 'li', 'a'])
            for container in containers:
                text = container.get_text(strip=True).lower()
                for format_type, indicators in format_indicators.items():
                    if any(indicator in text for indicator in indicators):
                        formats.add(format_type)

            # Check metadata
            meta_description = soup.find('meta', {'name': 'description'})
            if meta_description:
                desc_text = meta_description.get('content', '').lower()
                for format_type, indicators in format_indicators.items():
                    if any(indicator in desc_text for indicator in indicators):
                        formats.add(format_type)

            return ', '.join(sorted(formats))

        except Exception as e:
            logger.error(f"Error extracting formats: {str(e)}")
            return ""

    def process_work(self, work) -> Optional[WorkData]:
        """Process a work entry and extract relevant information."""
        try:
            # Extract title
            title_elem = work.find('div', class_='works-title')
            if not title_elem:
                print(f"No title element found for work")
                return None
            title = title_elem.text.strip()

            # Extract date
            published_date = work.get('data-date', 'Unknown')

            # Extract type
            type_elem = work.find('div', class_='works-type')
            work_type = type_elem.text.strip() if type_elem else "Unknown"
            work_type = self.normalize_work_type(work_type)

            # Extract link
            link = work.get('href', '')
            if link and not link.startswith('http'):
                link = urljoin(self.BASE_URL, link)

            # Extract collection info
            collection_name, collection_url = self.extract_collection_info(work)

            # Extract formats
            formats = self.extract_available_formats(link) if link else ""

            work_data = WorkData(
                title=title,
                cleaned_title=self.clean_title(title),
                link=link,
                published_date=published_date,
                work_type=work_type,
                formats=formats,
                available_in=collection_name,
                available_in_link=collection_url
            )

            return work_data

        except Exception as e:
            print(f"Error processing work '{work.get_text().strip()[:100]}': {str(e)}")
            return None

    def normalize_work_type(self, work_type: str) -> str:
        """Normalize work type to standard categories."""
        type_mapping = {
            'novel': 'Novel',
            'short story': 'Short Story',
            'collection': 'Story Collection',
            'anthology': 'Anthology',
            'novella': 'Novella',
            'bachman': 'Bachman Novel',
            'nonfiction': 'Non-Fiction',
            'screenplay': 'Screenplay',
            'poem': 'Poem'
        }

        work_type = work_type.lower()
        for key, value in type_mapping.items():
            if key in work_type:
                return value
        return work_type.title()

    def parse_works(self):
        """
        Main method to parse all works using thread pool
        """
        response = self.request_manager.get(self.WORKS_URL)
        if not response:
            logger.error("Failed to fetch main works page")
            return

        soup = BeautifulSoup(response.text, 'html.parser')
        work_elements = soup.find_all('a', class_='row work')  # Use consistent selector

        # Process works in parallel with limited concurrency
        with ThreadPoolExecutor(max_workers=self.MAX_WORKERS) as executor:
            future_to_work = {executor.submit(self.process_work, work): work
                            for work in work_elements}

            for future in concurrent.futures.as_completed(future_to_work):
                work_data = future.result()
                if work_data:
                    self.add_or_update_work(work_data)

    def add_or_update_work(self, work_data: WorkData):
        """Add or update a work in the works dictionary with deduplication."""
        if not work_data:
            return

        # Create row data
        row_data = [
            f'=HYPERLINK("{work_data.link}", "{work_data.title}")',
            work_data.published_date,
            work_data.work_type,
            work_data.available_in_link,
            work_data.formats
        ]

        # Check for existing entry
        if work_data.cleaned_title in self.works_dict:
            existing_date = self.works_dict[work_data.cleaned_title][1]
            new_date = work_data.published_date

            # Only update if the new date is valid and earlier than existing
            if (new_date != "Unknown" and
                (existing_date == "Unknown" or
                 self.convert_to_datetime(new_date) < self.convert_to_datetime(existing_date))):

                # Combine formats if they exist
                existing_formats = self.works_dict[work_data.cleaned_title][4]
                new_formats = work_data.formats
                combined_formats = self.merge_format_strings(existing_formats, new_formats)  # Changed from combine_formats

                # Use new data but keep combined formats
                row_data[4] = combined_formats
                self.works_dict[work_data.cleaned_title] = row_data
                print(f"Updated: {work_data.title} with earlier date {work_data.published_date}")
        else:
            # Add new work to dictionary
            self.works_dict[work_data.cleaned_title] = row_data

    def sync_collection_dates(self):
        """Synchronize publication dates for works appearing in collections."""
        for title, row_data in self.works_dict.items():
            available_in_hyperlink = row_data[3]
            if available_in_hyperlink:
                collection_name = re.search(r'"[^"]*", "([^"]*)"', available_in_hyperlink)
                if collection_name and collection_name.group(1) in self.collection_dates:
                    collection_date = self.collection_dates[collection_name.group(1)]
                    row_data[1] = collection_date
                    print(f"Updated '{title}' publication date to match collection '{collection_name.group(1)}': {collection_date}")

    def normalize_format(self, format_str: str) -> str:
        """Normalize format strings to standard values."""
        if not format_str:  # Handle None or empty string
            return ''

        format_str = format_str.strip().lower()

        # Format mappings
        if format_str in ['kindle', 'ebook']:
            return 'Yes'
        elif format_str in ['audio', 'audiobook']:
            return 'Yes'
        elif format_str in ['movie', 'tv movie', 'dvd']:
            return 'Yes'
        elif format_str == 'tv miniseries':
            return 'Yes'
        elif format_str in ['hardcover']:
            return 'Yes'
        elif format_str in ['paperback']:
            return 'Yes'
        else:
            return ''

    def process_formats(self, formats_str: str) -> Dict[str, str]:
        """
        Process formats string into a dictionary of format availability.

        Args:
            formats_str (str): Comma-separated string of formats

        Returns:
            Dict[str, str]: Dictionary with format types as keys and '✓' or '' as values
        """
        formats_dict = {
            'Hardcover': '',
            'Paperback': '',
            'Ebook': '',
            'Audiobook': '',
            'Movie': '',
            'Miniseries': ''
        }

        if not formats_str:
            return formats_dict

        format_list = formats_str.split(',')
        for fmt in format_list:
            fmt = fmt.strip()
            if 'Hardcover' in fmt:
                formats_dict['Hardcover'] = '✓'
            if 'Paperback' in fmt:
                formats_dict['Paperback'] = '✓'
            if 'Kindle' in fmt or 'eBook' in fmt:
                formats_dict['Ebook'] = '✓'
            if 'Audio' in fmt or 'Audiobook' in fmt:
                formats_dict['Audiobook'] = '✓'
            if 'Movie' in fmt:
                formats_dict['Movie'] = '✓'
            if 'TV' in fmt or 'Miniseries' in fmt:
                formats_dict['Miniseries'] = '✓'

        return formats_dict

    def format_published_date(self, date_str: str) -> str:
        """Convert '0000-00-00' dates to empty string."""
        return '' if date_str == '0000-00-00' else date_str

    def export_to_csv(self, filename: str, works_data: List[List[str]]):
        """Export works data to CSV file with separate format columns."""
        with open(filename, 'w', newline='', encoding='utf-8') as file:
            writer = csv.writer(file)
            writer.writerow([
                'Read', 'Owned', 'Published',  # First three columns
                'Title', 'Type', 'Available In',
                'Hardcover', 'Paperback', 'Ebook', 'Audiobook', 'Movie', 'Miniseries'
            ])

            for row in works_data:
                formats_dict = self.process_formats(row[4])
                writer.writerow([
                    '',  # Read
                    '',  # Owned
                    self.format_published_date(row[1]),  # Published
                    row[0],  # Title
                    row[2],  # Type
                    row[3],  # Available In
                    formats_dict['Hardcover'],
                    formats_dict['Paperback'],
                    formats_dict['Ebook'],
                    formats_dict['Audiobook'],
                    formats_dict['Movie'],
                    formats_dict['Miniseries']
                ])

    def parse_excel_hyperlink(self, excel_formula: str) -> tuple[str, str]:
        """
        Parse Excel HYPERLINK formula into URL and text components.

        Args:
            excel_formula (str): Excel formula string in format '=HYPERLINK("url", "text")'

        Returns:
            tuple[str, str]: (url, text) tuple
        """
        if not excel_formula.startswith('=HYPERLINK('):
            return ('', excel_formula)

        # Extract URL and text from HYPERLINK formula
        match = re.match(r'=HYPERLINK\("([^"]+)",\s*"([^"]+)"\)', excel_formula)
        if match:
            return (match.group(1), match.group(2))
        return ('', excel_formula)

    def excel_hyperlink_to_html(self, excel_formula: str) -> str:
        """
        Convert Excel HYPERLINK formula to HTML anchor tag.

        Args:
            excel_formula (str): Excel formula string in format '=HYPERLINK("url", "text")'

        Returns:
            str: HTML anchor tag
        """
        if not excel_formula.startswith('=HYPERLINK('):
            return excel_formula

        url, text = self.parse_excel_hyperlink(excel_formula)
        if url and text:
            return f'<a href="{url}">{text}</a>'
        return excel_formula

    def generate_html_table(self, rows: List[List[str]]) -> str:
        """Generate HTML table with proper formatting and structure."""
        collections = defaultdict(list)
        standalone_works = []

        # Sort works into collections and standalone
        for row in rows:
            collection = row[3]  # Collection column
            if collection:
                collections[collection].append(row)
            else:
                standalone_works.append(row)

        # Generate HTML
        html_parts = []

        # Add collections first
        for collection_name, collection_works in collections.items():
            html_parts.append(f'<div class="collection-section">')
            html_parts.append(f'<div class="collection-header">{collection_name}</div>')
            html_parts.append(self._generate_table_content(collection_works))
            html_parts.append('</div>')

        # Add standalone works
        if standalone_works:
            html_parts.append('<div class="standalone-section">')
            html_parts.append('<div class="standalone-header">Standalone Works</div>')
            html_parts.append(self._generate_table_content(standalone_works))
            html_parts.append('</div>')

        return '\n'.join(html_parts)

    def _generate_table_content(self, rows: List[List[str]]) -> str:
        """Generate the actual table HTML for a set of works."""
        table_html = [
            '<table class="works-table">',
            '<thead>',
            '<tr>',
            '<th>Read</th>',
            '<th>Owned</th>',
            '<th>Published</th>',
            '<th>Title</th>',
            '<th>Type</th>',
            '<th>Collection</th>',
            '<th>Hardcover</th>',
            '<th>Paperback</th>',
            '<th>Ebook</th>',
            '<th>Audiobook</th>',
            '<th>Movie</th>',
            '<th>Miniseries</th>',
            '</tr>',
            '</thead>',
            '<tbody>'
        ]

        for row in rows:
            formats_dict = self.process_formats(row[4])
            url, text = self.parse_excel_hyperlink(row[0])
            collection_url, collection_text = self.parse_excel_hyperlink(row[3])

            table_html.append('<tr>')
            table_html.extend([
                f'<td><input type="checkbox" class="status-checkbox" data-title="{text}" data-type="read"></td>',
                f'<td><input type="checkbox" class="status-checkbox" data-title="{text}" data-type="owned"></td>',
                f'<td>{self.format_published_date(row[1])}</td>',
                f'<td><a href="{url}">{text}</a></td>',
                f'<td>{row[2]}</td>'
            ])

            # Collection column
            if collection_url and collection_text:
                table_html.append(f'<td><a href="{collection_url}">{collection_text}</a></td>')
            else:
                table_html.append('<td></td>')

            # Format columns
            for format_type in ['Hardcover', 'Paperback', 'Ebook', 'Audiobook', 'Movie', 'Miniseries']:
                is_available = formats_dict[format_type] == '✓'
                cell_class = 'format-cell yes' if is_available else 'format-cell'
                table_html.append(f'<td class="{cell_class}">{formats_dict[format_type]}</td>')

            table_html.append('</tr>')

        table_html.extend(['</tbody>', '</table>'])
        return '\n'.join(table_html)

    def export_to_html(self, filename: str, works_data: List[List[str]]):
        """Export works data to HTML."""
        table_content = self.generate_html_table(works_data)

        html_content = f"""<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Stephen King Works</title>
    <link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600&display=swap" rel="stylesheet">
    <link rel="stylesheet" type="text/css" href="https://cdn.datatables.net/1.11.5/css/jquery.dataTables.css">
    <link rel="stylesheet" type="text/css" href="https://cdn.datatables.net/fixedheader/3.2.2/css/fixedHeader.dataTables.min.css">
    <script type="text/javascript" src="https://code.jquery.com/jquery-3.5.1.min.js"></script>
    <script type="text/javascript" src="https://cdn.datatables.net/1.11.5/js/jquery.dataTables.min.js"></script>
    <script type="text/javascript" src="https://cdn.datatables.net/fixedheader/3.2.2/js/dataTables.fixedHeader.min.js"></script>
    <style>
        :root {{
            --primary-color: #2c3e50;
            --secondary-color: #34495e;
            --accent-color: #3498db;
            --hover-color: #f8f9fa;
            --border-color: #e9ecef;
            --success-color: #2ecc71;
        }}

        * {{
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }}

        body {{
            font-family: 'Inter', -apple-system, BlinkMacSystemFont, sans-serif;
            line-height: 1.6;
            color: var(--primary-color);
            background-color: #f8f9fa;
            padding: 2rem;
        }}

        .container {{
            max-width: 1400px;
            margin: 0 auto;
            background-color: white;
            border-radius: 12px;
            box-shadow: 0 2px 12px rgba(0, 0, 0, 0.1);
            padding: 2rem;
        }}

        h1 {{
            font-size: 2.5rem;
            color: var(--primary-color);
            margin-bottom: 2rem;
            text-align: center;
            font-weight: 600;
        }}

        .works-table {{
            width: 100% !important;
            border-collapse: separate;
            border-spacing: 0;
            margin-bottom: 2rem;
            border-radius: 8px;
            overflow: hidden;
        }}

        .works-table th,
        .works-table td {{
            padding: 1rem !important;
            text-align: left;
            border-bottom: 1px solid var(--border-color);
        }}

        .works-table thead th {{
            background-color: var(--secondary-color);
            color: white !important;
            font-weight: 500;
            text-transform: uppercase;
            font-size: 0.85rem;
            letter-spacing: 0.5px;
        }}

        .works-table tbody tr:hover {{
            background-color: var(--hover-color);
        }}

        .format-cell {{
            text-align: center;
            font-weight: 500;
        }}

        .format-cell.yes {{
            color: var(--success-color);
        }}

        a {{
            color: var(--accent-color);
            text-decoration: none;
            transition: color 0.2s ease;
        }}

        a:hover {{
            color: #2980b9;
        }}

        .status-checkbox {{
            width: 18px;
            height: 18px;
            cursor: pointer;
            border-radius: 4px;
            border: 2px solid var(--border-color);
            transition: all 0.2s ease;
        }}

        .status-checkbox:checked {{
            background-color: var(--success-color);
            border-color: var(--success-color);
        }}

        @media (max-width: 1200px) {{
            .container {{
                padding: 1rem;
            }}
        }}

        @media (max-width: 768px) {{
            body {{
                padding: 1rem;
            }}

            h1 {{
                font-size: 2rem;
            }}
        }}
    </style>
    <script>
        $(document).ready(function() {{
            $('.works-table').DataTable({{
                fixedHeader: true,
                pageLength: 25,
                order: [[2, 'asc']],
                columnDefs: [
                    {{
                        targets: [0, 1],
                        orderable: false
                    }},
                    {{
                        targets: [6, 7, 8, 9, 10, 11],
                        className: 'text-center'
                    }}
                ],
                drawCallback: function() {{
                    $('.status-checkbox').each(function() {{
                        const title = $(this).data('title');
                        const type = $(this).data('type');
                        const key = `${{title}}-${{type}}`;
                        const checked = localStorage.getItem(key) === 'true';
                        $(this).prop('checked', checked);
                    }});
                }}
            }});

            $(document).on('change', '.status-checkbox', function() {{
                const title = $(this).data('title');
                const type = $(this).data('type');
                const key = `${{title}}-${{type}}`;
                localStorage.setItem(key, $(this).prop('checked'));
            }});
        }});
    </script>
</head>
<body>
    <div class="container">
        <h1>Stephen King Works</h1>
        {table_content}
    </div>
</body>
</html>"""

        with open(filename, 'w', encoding='utf-8') as f:
            f.write(html_content)

    def parse_and_export(self):
        """Parse website and export data to CSV and HTML."""
        print("Starting to fetch data from website...")
        response = self.request_manager.get(self.WORKS_URL)
        if not response:
            print(f"Failed to get response from {self.WORKS_URL}")
            return

        print(f"Got response with status code: {response.status_code}")
        soup = BeautifulSoup(response.text, 'html.parser')

        # Look specifically for work elements with the correct class
        works = soup.find_all('a', class_='row work')

        if not works:
            print("No works found using primary selector. HTML content sample:")
            print(response.text[:1000])
            return

        print(f"Found {len(works)} works to process")

        # Process all works
        processed_count = 0
        failed_count = 0
        for work in works:
            try:
                work_data = self.process_work(work)
                if work_data:
                    self.add_or_update_work(work_data)
                    processed_count += 1
                    if processed_count % 10 == 0:
                        print(f"Processed {processed_count} works so far...")
                else:
                    failed_count += 1
                    print(f"Failed to process work: {work.get_text().strip()[:100]}")
            except Exception as e:
                failed_count += 1
                print(f"Error processing work: {str(e)}")

        print(f"Successfully processed {processed_count} works, failed to process {failed_count} works")

        # Sort and prepare works data
        works_data = sorted(
            self.works_dict.values(),
            key=lambda x: (self.convert_to_datetime(x[1]), x[0])  # Sort by date, then title
        )

        if not works_data:
            print("No works data to export!")
            return

        # Generate timestamp for filenames
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")

        # Export to CSV
        csv_file = f'stephen_king_works_{timestamp}.csv'
        self.export_to_csv(csv_file, works_data)
        print(f"CSV file '{csv_file}' created successfully!")

        # Export to HTML
        html_file = f'stephen_king_works_{timestamp}.html'
        self.export_to_html(html_file, works_data)
        print(f"HTML file '{html_file}' created successfully!")

def main():
    """Main entry point for the Stephen King works parser application.

    Creates a KingWorksParser instance and executes the parsing and export process
    to generate CSV and HTML files containing Stephen King's bibliography.
    """
    parser = KingWorksParser()
    parser.parse_and_export()

if __name__ == "__main__":
    main()
