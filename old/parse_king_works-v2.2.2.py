"""
Module for parsing and organizing Stephen King's literary works from his official website.
Provides functionality to scrape, process, and export work details to various formats including CSV.
"""

import argparse
import concurrent.futures
import csv
import glob
import logging
import re
import threading
import time
from collections import defaultdict
from concurrent.futures import ThreadPoolExecutor
from dataclasses import dataclass
from datetime import datetime
from typing import Dict, List, Optional, Union
from urllib.parse import urljoin

import pandas as pd
import requests
from bs4 import BeautifulSoup
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

#   Version Ntes
#   1.3   - Mostly working well, but the CSV being generated was wrong
#   1.3.1 - ???
#   1.3.2 - ???
#   1.3.3 - Finally fixed the hyperlink issue in the CSV, but still present in the HTML
#   1.4   - Finally fixed all hyperink issues!
#   1.4.1 - Hyperlinks should work for collction titles as well
#   1.4.2 - Cut the run time down to around 5 minutes fro 17! But just realized that we are missing formats.

#   Version Notes
#   2.0.0   - Fixed a lot of issues we had, and this is a great working version
#               Things working well:
#                   - Runtime has been significantly reduced to about 1 minute
#                   - CSV is working and have proper hyperlinks
#                   - HTML is working and have proper hyperlinks
#                   - Collections are working
#                   - HTML is sortable and searching
#                   - Formats are working
#                   - Hyperlinks are working
#                   - Dates with no publication
#               Things to improve:
#                   - There are still various duplicates in the data
#                   - There are still 0000-00-00 dates
#                   - There are still Titles in collections that could get pub dates inherited from the collection
#                   - Logging is still a bit verbose
#                   - Would be nice to make the HTML look a bit better or more modern
#                   - Formatting issues with The Genius of "The Tell-Tale Heart"
#
#   2.0.1       Fixed the Titles with 0000-00-00 dates that could inherit from collections, and now do.
#   2.0.2       Fixed the issue with the HTML table not being sortable and searchable. Removed Logging.
#   2.0.3       Removed logging for collections found
#   2.0.4       Not much change, tried to fix some things but functionally the same
#   2.0.5       Dates with 0000-00-00 are now blank in the HTML table
#   2.1.0       Updated HTML that looks a lot more spooky and modern
#   2.1.1       Updated a bit more and looks decent, but want to save it here.
#   2.1.2       Fixed the issue of multipl HTML templates
#   2.1.3       Very good clean run with HTML looking very nice
#   2.1.4       Even better. Jut noticed there are no ebook entries.
#   2.2.0       Removed formats. Amazing now.
#   2.2.2       AMAZING.

# Set up logging to only show WARNING and ERROR level messages
logging.basicConfig(
    level=logging.WARNING,
    format="%(message)s",  # Simplified format to only show the message
)
logger = logging.getLogger(__name__)


@dataclass
class WorkData:
    """Data class to store information about a work"""

    title: str
    cleaned_title: str
    link: str
    published_date: str
    work_type: str
    available_in: str
    available_in_link: str


class RequestManager:
    """Manages HTTP requests with rate limiting"""

    def __init__(self, rate_limit: float = 1.0):
        self.rate_limit = rate_limit
        self.last_request_time = 0
        self.logger = logging.getLogger(__name__)
        self.logger.setLevel(logging.WARNING)  # Keep only warnings and errors

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
        except Exception as e:
            self.logger.warning(f"Request failed: {str(e)}")
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
    MAX_WORKERS = 10  # Increased from 5 for better parallelization

    def __init__(self):
        self.request_manager = RequestManager(
            rate_limit=0.5
        )  # Decreased from 1.0 to 0.5 seconds
        self.works_dict = {}
        self.collection_dates = {}
        self.processed_urls = set()
        self.url_lock = threading.Lock()
        self.data_lock = threading.Lock()
        self.session = requests.Session()  # Add session for connection pooling

    def clean_title(self, title: str) -> str:
        """
        Clean the title by removing special characters and normalizing format.
        Also handles common variations of the same title.

        Args:
            title (str): Original title

        Returns:
            str: Cleaned title for comparison
        """
        # Remove special characters but keep basic punctuation
        cleaned = re.sub(r"[^\w\s\-\'.,]", "", title)

        # Convert to lowercase for comparison
        cleaned = cleaned.lower()

        # Remove common suffixes and variations
        cleaned = re.sub(
            r"\s*:\s*the\s+complete\s+(?:&|and)\s+uncut\s+edition\s*$", "", cleaned
        )
        cleaned = re.sub(
            r"\s*:\s*(?:expanded|limited|special|collectors?)\s+edition\s*$",
            "",
            cleaned,
        )
        cleaned = re.sub(r"\s+edition\s*$", "", cleaned)

        # Normalize whitespace
        cleaned = " ".join(cleaned.split())
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
        return re.sub(r"\s*\([^)]*\)\s*$", "", title).strip()

    @staticmethod
    def create_excel_hyperlink(url: str, text: str) -> str:
        """
        Create a clean Excel hyperlink formula.

        Args:
            url: The URL to link to
            text: The display text

        Returns:
            Excel HYPERLINK formula with proper quote handling
        """
        # First, escape any existing double quotes in both url and text
        escaped_url = url.replace('"', '""')
        escaped_text = text.replace('"', '""')

        # Create the Excel formula with the escaped strings
        return f'=HYPERLINK("{escaped_url}", "{escaped_text}")'

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
            return datetime.strptime(date_str.strip(), "%Y-%m-%d")
        except (ValueError, AttributeError):
            # Return far-future date if parsing fails
            return datetime(9999, 12, 31)

    def extract_collection_info(self, work) -> tuple[str, str]:
        """Extract collection information from a work's dedicated page."""
        try:
            # Get the work's specific URL
            work_url = work.get("href", "")
            if not work_url:
                return ("", "")

            # Make sure we have a full URL
            if not work_url.startswith("http"):
                work_url = urljoin(self.BASE_URL, work_url)

            # Fetch the work's dedicated page
            response = self.request_manager.get(work_url)
            if not response:
                return ("", "")

            # Parse the page
            soup = BeautifulSoup(response.text, "html.parser")

            # Find the "Available In" section
            available_in = soup.find("h2", string="Available In")
            if not available_in:
                return ("", "")

            # Find the collection link in the section following the "Available In" header
            collection_section = available_in.find_next("div", class_="grid-content")
            if not collection_section:
                return ("", "")

            # Find the collection link
            collection_link = collection_section.find("a", class_="text-link")
            if not collection_link:
                return ("", "")

            # Extract collection name and URL
            collection_name = collection_link.text.strip()
            collection_url = collection_link.get("href", "")
            if collection_url and not collection_url.startswith("http"):
                collection_url = urljoin(self.BASE_URL, collection_url)

            return (collection_name, collection_url)

        except Exception as e:
            logger.error(f"Error extracting collection info from {work_url}: {str(e)}")
            return ("", "")

    def process_work(self, work: BeautifulSoup) -> Optional[WorkData]:
        """Process a single work with improved error handling and caching"""
        try:
            # Extract work URL
            work_url = work.get("href", "")
            if not work_url:
                return None

            # Check if already processed
            if self.is_url_processed(work_url):
                return None

            # Make sure we have a full URL
            if not work_url.startswith("http"):
                work_url = urljoin(self.BASE_URL, work_url)

            # Mark as processed
            self.mark_url_processed(work_url)

            # Extract work data
            title = work.find("div", class_="works-title").text.strip()
            published_date = work.get("data-date", "").strip()
            work_type = work.find("div", class_="works-type").text.strip()

            # Get collection info
            available_in, available_in_link = self.extract_collection_info(work)

            return WorkData(
                title=title,
                cleaned_title=self.clean_title(title),
                link=work_url,
                published_date=published_date,
                work_type=work_type,
                available_in=available_in,
                available_in_link=available_in_link,
            )

        except Exception as e:
            logger.error(f"Error processing work: {e}")
            return None

    def normalize_work_type(self, work_type: str) -> str:
        """Normalize work type to standard categories."""
        type_mapping = {
            "novel": "Novel",
            "short story": "Short Story",
            "collection": "Story Collection",
            "anthology": "Anthology",
            "novella": "Novella",
            "bachman": "Bachman Novel",
            "nonfiction": "Non-Fiction",
            "screenplay": "Screenplay",
            "poem": "Poem",
        }

        work_type = work_type.lower()
        for key, value in type_mapping.items():
            if key in work_type:
                return value
        return work_type.title()

    def batch_process_works(
        self, works: List[BeautifulSoup], batch_size: int = 20
    ) -> List[WorkData]:
        """Process works in batches using ThreadPoolExecutor"""
        results = []
        for i in range(0, len(works), batch_size):
            batch = works[i : i + batch_size]
            with ThreadPoolExecutor(max_workers=self.MAX_WORKERS) as executor:
                futures = [executor.submit(self.process_work, work) for work in batch]
                for future in concurrent.futures.as_completed(futures):
                    if work_data := future.result():
                        results.append(work_data)
        return results

    def parse_works(self):
        """Main method to parse all works using batched processing"""
        response = self.request_manager.get(self.WORKS_URL)
        if not response:
            logger.error("Failed to fetch main works page")
            return

        soup = BeautifulSoup(response.text, "html.parser")
        work_elements = soup.find_all("a", class_="row work")

        # Process works in batches
        processed_works = self.batch_process_works(work_elements)

        # Update works dictionary
        for work_data in processed_works:
            self.add_or_update_work(work_data)

        # Synchronize collection dates BEFORE exporting
        logger.info("Starting collection date synchronization...")
        works_list = list(self.works_dict.values())
        updated_count = self.sync_collection_dates(works_list)
        logger.info(f"Updated {updated_count} works with collection dates")

        return processed_works

    def add_or_update_work(self, work_data: WorkData):
        """Add or update a work in the works dictionary with enhanced deduplication."""
        if not work_data:
            return

        # Create row data
        row_data = [
            f'=HYPERLINK("{work_data.link}", "{work_data.title}")',
            work_data.published_date,
            work_data.work_type,
            work_data.available_in_link,
        ]

        # Check for existing entry
        if work_data.cleaned_title in self.works_dict:
            existing_date = self.works_dict[work_data.cleaned_title][1]
            new_date = work_data.published_date

            # If this is a special edition or variant, prefer the more detailed title
            existing_title = self.works_dict[work_data.cleaned_title][0]
            if (
                "complete" in work_data.title.lower()
                or "uncut" in work_data.title.lower()
                or "expanded" in work_data.title.lower()
            ):
                row_data[0] = f'=HYPERLINK("{work_data.link}", "{work_data.title}")'

            # Keep the earliest date
            if new_date != "Unknown" and (
                existing_date == "Unknown"
                or self.convert_to_datetime(new_date)
                < self.convert_to_datetime(existing_date)
            ):
                row_data[1] = new_date

            # Update the entry
            self.works_dict[work_data.cleaned_title] = row_data
            print(f"Updated: {work_data.title}")
        else:
            # Add new work to dictionary
            self.works_dict[work_data.cleaned_title] = row_data

    def sync_collection_dates(self, works_list):
        """
        Synchronize publication dates for works appearing in collections.
        Updates works that have no date or '0000-00-00' date with their collection's date.
        """
        # First, build a dictionary of collection titles and their dates
        collection_dates = {}
        for work in works_list:
            if work.work_type.lower() in [
                "collection",
                "anthology",
                "story collection",
            ]:
                if work.published_date and work.published_date != "0000-00-00":
                    collection_dates[work.title] = work.published_date
                    collection_dates[self.clean_title(work.title)] = work.published_date

        # Then update works that appear in collections but have no date
        updated_count = 0
        for work in works_list:
            if (
                not work.published_date or work.published_date == "0000-00-00"
            ) and work.available_in:
                collection_name = work.available_in
                cleaned_collection_name = self.clean_title(collection_name)

                # Try all possible matches
                if collection_name in collection_dates:
                    work.published_date = collection_dates[collection_name]
                    updated_count += 1
                elif cleaned_collection_name in collection_dates:
                    work.published_date = collection_dates[cleaned_collection_name]
                    updated_count += 1

        return updated_count

    def get_sort_key(self, work_data: WorkData):
        """
        Create a sort key for a WorkData object.
        Returns tuple of (has_date, date_value, title) where:
        - has_date is True for valid dates (to sort them first)
        - date_value is the actual date or max date for empty/invalid dates
        - title is used as secondary sort key
        """
        date_str = work_data.published_date.strip() if work_data.published_date else ""
        title = work_data.title

        # Handle empty or invalid dates - use max date to sort them to the end
        if not date_str or date_str == "0000-00-00":
            return (False, datetime.max, title)

        try:
            # Try to parse the date
            parsed_date = datetime.strptime(date_str, "%Y-%m-%d")
            return (True, parsed_date, title)
        except (ValueError, AttributeError):
            # Invalid date format - sort to end
            return (False, datetime.max, title)

    def export_to_csv(self, filename: str, works_data: List[List[str]]):
        """Export works data to CSV file."""
        # Prepare header row
        header = ["Read", "Owned", "Published", "Title", "Type", "Available In"]

        with open(filename, "w", newline="", encoding="utf-8") as file:
            writer = csv.writer(file, quoting=csv.QUOTE_MINIMAL)
            writer.writerow(header)

            for row in works_data:
                # Process each cell to ensure proper formatting
                processed_row = []
                for item in row:
                    if isinstance(item, str) and item.startswith("=HYPERLINK"):
                        # Handle hyperlinks with single quotes
                        item = item.replace(
                            '""', '"'
                        )  # Remove any existing double quotes
                        processed_row.append(item)
                    else:
                        processed_row.append(item)
                writer.writerow(processed_row)

    def parse_excel_hyperlink(self, excel_formula: str) -> tuple[str, str]:
        """
        Parse Excel HYPERLINK formula into URL and text components.

        Args:
            excel_formula (str): Excel formula string in format '=HYPERLINK("url", "text")'

        Returns:
            tuple[str, str]: (url, text) tuple
        """
        # Convert float or other types to string
        excel_formula = str(excel_formula) if excel_formula is not None else ""

        if not excel_formula.startswith("=HYPERLINK("):
            return ("", excel_formula)

        # Use a more robust regex that handles escaped quotes
        match = re.match(
            r'=HYPERLINK\("((?:[^"]|"")+)",\s*"((?:[^"]|"")+)"\)', excel_formula
        )
        if match:
            url = match.group(1).replace('""', '"')
            text = match.group(2).replace('""', '"')
            return (url, text)
        return ("", excel_formula)

    def excel_hyperlink_to_html(self, excel_formula: str) -> str:
        """Convert Excel HYPERLINK formula to HTML anchor tag with bold text."""
        if not excel_formula or not excel_formula.startswith("=HYPERLINK"):
            return excel_formula

        # Extract URL and text from Excel formula
        match = re.search(r'HYPERLINK\("([^"]+)",\s*"([^"]+)"\)', excel_formula)
        if match:
            url, text = match.groups()
            return f'<a href="{url}"><strong>{text}</strong></a>'
        return excel_formula

    def format_row_for_export(self, work_data: WorkData) -> List[str]:
        """Format a single work's data for export"""
        # Create hyperlink for available_in if there's a link
        available_in = (
            self.create_excel_hyperlink(
                work_data.available_in_link, work_data.available_in
            )
            if work_data.available_in and work_data.available_in_link
            else work_data.available_in
        )

        return [
            "",  # Read
            "",  # Owned
            work_data.published_date.strip(),
            self.create_excel_hyperlink(
                work_data.link, work_data.title
            ),  # Title with hyperlink
            work_data.work_type,
            available_in,  # Collection with hyperlink
        ]

    def generate_html_table(self, rows: List[List[str]]) -> str:
        """Generate HTML table with proper formatting and structure."""
        table_html = [
            '<table class="works-table">',
            "<thead>",
            "<tr>",
            '<th class="narrow-col">Read</th>',
            '<th class="narrow-col">Owned</th>',
            '<th class="date-col">Published</th>',
            '<th class="title-col">Title</th>',
            '<th class="type-col">Type</th>',
            '<th class="collection-col">Collection</th>',
            "</tr>",
            "</thead>",
            "<tbody>",
        ]

        for row in rows:
            read = row[0]
            owned = row[1]
            published_date = row[2]
            title_formula = row[3]
            work_type = row[4]
            collection = row[5]

            # Handle the published date
            display_date = ""
            sort_date = "9999-99-99"  # Default sort value for empty dates

            if published_date and published_date not in ["0000-00-00", "9999-99-99"]:
                try:
                    parsed_date = datetime.strptime(published_date, "%Y-%m-%d")
                    display_date = parsed_date.strftime("%Y-%m-%d")
                    sort_date = display_date
                except ValueError:
                    pass

            # Convert Excel hyperlinks to HTML and wrap in bold tags
            title_html = self.excel_hyperlink_to_html(title_formula)
            title_html = f"<strong>{title_html}</strong>" if title_html else ""

            collection_html = self.excel_hyperlink_to_html(collection)
            collection_html = (
                f"<strong>{collection_html}</strong>" if collection_html else ""
            )

            table_html.append("<tr>")
            table_html.extend(
                [
                    f'<td class="narrow-col"><input type="checkbox" class="status-checkbox" data-title="{self.parse_excel_hyperlink(title_formula)[1]}" data-type="read"{" checked" if read else ""}></td>',
                    f'<td class="narrow-col"><input type="checkbox" class="status-checkbox" data-title="{self.parse_excel_hyperlink(title_formula)[1]}" data-type="owned"{" checked" if owned else ""}></td>',
                    f'<td class="date-col" data-sort="{sort_date}">{display_date}</td>',
                    f'<td class="title-col">{title_html}</td>',
                    f'<td class="type-col">{work_type}</td>',
                    f'<td class="collection-col">{collection_html}</td>',
                ]
            )
            table_html.append("</tr>")

        table_html.extend(["</tbody>", "</table>"])

        return "\n".join(table_html)

    def generate_html_table(self, rows: List[List[str]]) -> str:
        """Generate HTML table with proper formatting and structure."""
        table_html = [
            '<table class="works-table">',
            "<thead>",
            "<tr>",
            '<th class="narrow-col">Read</th>',
            '<th class="narrow-col">Owned</th>',
            '<th class="date-col">Published</th>',
            '<th class="title-col">Title</th>',
            '<th class="type-col">Type</th>',
            '<th class="collection-col">Collection</th>',
            "</tr>",
            "</thead>",
            "<tbody>",
        ]

        for row in rows:
            read = row[0]
            owned = row[1]
            published_date = row[2]
            title_formula = row[3]
            work_type = row[4]
            collection = row[5]

            # Handle the published date
            display_date = ""
            sort_date = "9999-99-99"  # Default sort value for empty dates

            if published_date and published_date not in ["0000-00-00", "9999-99-99"]:
                try:
                    parsed_date = datetime.strptime(published_date, "%Y-%m-%d")
                    display_date = parsed_date.strftime("%Y-%m-%d")
                    sort_date = display_date
                except ValueError:
                    pass

            # Convert Excel hyperlinks to HTML
            title_html = self.excel_hyperlink_to_html(title_formula)
            collection_html = (
                self.excel_hyperlink_to_html(collection) if collection else ""
            )

            table_html.append("<tr>")
            table_html.extend(
                [
                    f'<td class="narrow-col"><input type="checkbox" class="status-checkbox" data-title="{self.parse_excel_hyperlink(title_formula)[1]}" data-type="read"{" checked" if read else ""}></td>',
                    f'<td class="narrow-col"><input type="checkbox" class="status-checkbox" data-title="{self.parse_excel_hyperlink(title_formula)[1]}" data-type="owned"{" checked" if owned else ""}></td>',
                    f'<td class="date-col" data-sort="{sort_date}">{display_date}</td>',
                    f'<td class="title-col">{title_html}</td>',
                    f'<td class="type-col">{work_type}</td>',
                    f'<td class="collection-col">{collection_html}</td>',
                ]
            )
            table_html.append("</tr>")

        table_html.extend(["</tbody>", "</table>"])

        return "\n".join(table_html)

    def generate_html_table(self, rows: List[List[str]]) -> str:
        """Generate HTML table with proper formatting and structure."""
        table_html = [
            '<table class="works-table">',
            "<thead>",
            "<tr>",
            '<th class="narrow-col">Read</th>',
            '<th class="narrow-col">Owned</th>',
            '<th class="date-col">Published</th>',
            '<th class="title-col">Title</th>',
            '<th class="type-col">Type</th>',
            '<th class="collection-col">Collection</th>',
            "</tr>",
            "</thead>",
            "<tbody>",
        ]

        for row in rows:
            read = row[0]
            owned = row[1]
            published_date = row[2]
            title_formula = row[3]
            work_type = row[4]
            collection = row[5]

            # Handle the published date
            display_date = ""
            sort_date = "9999-99-99"  # Default sort value for empty dates

            if published_date and published_date not in ["0000-00-00", "9999-99-99"]:
                try:
                    parsed_date = datetime.strptime(published_date, "%Y-%m-%d")
                    display_date = parsed_date.strftime("%Y-%m-%d")
                    sort_date = display_date
                except ValueError:
                    pass

            # Convert Excel hyperlinks to HTML
            title_html = self.excel_hyperlink_to_html(title_formula)
            collection_html = (
                self.excel_hyperlink_to_html(collection) if collection else ""
            )

            table_html.append("<tr>")
            table_html.extend(
                [
                    f'<td class="narrow-col"><input type="checkbox" class="status-checkbox" data-title="{self.parse_excel_hyperlink(title_formula)[1]}" data-type="read"{" checked" if read else ""}></td>',
                    f'<td class="narrow-col"><input type="checkbox" class="status-checkbox" data-title="{self.parse_excel_hyperlink(title_formula)[1]}" data-type="owned"{" checked" if owned else ""}></td>',
                    f'<td class="date-col" data-sort="{sort_date}">{display_date}</td>',
                    f'<td class="title-col">{title_html}</td>',
                    f'<td class="type-col">{work_type}</td>',
                    f'<td class="collection-col">{collection_html}</td>',
                ]
            )
            table_html.append("</tr>")

        table_html.extend(["</tbody>", "</table>"])

        return "\n".join(table_html)

    def generate_html_table(self, rows: List[List[str]]) -> str:
        """Generate HTML table with proper formatting and structure."""
        table_html = [
            '<table class="works-table">',
            "<thead>",
            "<tr>",
            '<th class="narrow-col">Read</th>',
            '<th class="narrow-col">Owned</th>',
            '<th class="date-col">Published</th>',
            '<th class="title-col">Title</th>',
            '<th class="type-col">Type</th>',
            '<th class="collection-col">Collection</th>',
            "</tr>",
            "</thead>",
            "<tbody>",
        ]

        for row in rows:
            read = row[0]
            owned = row[1]
            published_date = row[2]
            title_formula = row[3]
            work_type = row[4]
            collection = row[5]

            # Handle the published date
            display_date = ""
            sort_date = "9999-99-99"  # Default sort value for empty dates

            if published_date and published_date not in ["0000-00-00", "9999-99-99"]:
                try:
                    parsed_date = datetime.strptime(published_date, "%Y-%m-%d")
                    display_date = parsed_date.strftime("%Y-%m-%d")
                    sort_date = display_date
                except ValueError:
                    pass

            # Convert Excel hyperlinks to HTML
            title_html = self.excel_hyperlink_to_html(title_formula)
            collection_html = (
                self.excel_hyperlink_to_html(collection) if collection else ""
            )

            table_html.append("<tr>")
            table_html.extend(
                [
                    f'<td class="narrow-col"><input type="checkbox" class="status-checkbox" data-title="{self.parse_excel_hyperlink(title_formula)[1]}" data-type="read"{" checked" if read else ""}></td>',
                    f'<td class="narrow-col"><input type="checkbox" class="status-checkbox" data-title="{self.parse_excel_hyperlink(title_formula)[1]}" data-type="owned"{" checked" if owned else ""}></td>',
                    f'<td class="date-col" data-sort="{sort_date}">{display_date}</td>',
                    f'<td class="title-col">{title_html}</td>',
                    f'<td class="type-col">{work_type}</td>',
                    f'<td class="collection-col">{collection_html}</td>',
                ]
            )
            table_html.append("</tr>")

        table_html.extend(["</tbody>", "</table>"])

        return "\n".join(table_html)

    def extract_title_from_hyperlink(self, hyperlink: str) -> str:
        """Extract the title from an Excel or HTML hyperlink."""
        if hyperlink.startswith("=HYPERLINK"):
            # Extract from Excel hyperlink
            match = re.search(r'HYPERLINK\("[^"]*",\s*"([^"]+)"\)', hyperlink)
            return match.group(1) if match else ""
        elif "<a href=" in hyperlink:
            # Extract from HTML hyperlink
            match = re.search(r">([^<]+)</a>", hyperlink)
            return match.group(1) if match else ""
        return hyperlink

    def export_to_html(self, filename: str, works_data: List[List[str]]):
        """Export works data to HTML file."""
        # Convert all data to strings first
        works_data = [
            [str(cell) if cell is not None else "" for cell in row]
            for row in works_data
        ]

        table_content = self.generate_html_table(works_data)

        html_content = self.generate_html_content(table_content)

        with open(filename, "w", encoding="utf-8") as f:
            f.write(html_content)

    def generate_html_content(self, table_content: str) -> str:
        """Generate complete HTML document with modern, Stephen King-inspired styling."""
        return f"""<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Stephen King Bibliography</title>
    <link href="https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;800&display=swap" rel="stylesheet">
    <link href="https://fonts.googleapis.com/css2?family=Special+Elite&display=swap" rel="stylesheet">
    <link href="https://fonts.cdnfonts.com/css/portrait-condensed" rel="stylesheet">
    <link rel="stylesheet" type="text/css" href="https://cdn.datatables.net/1.13.4/css/jquery.dataTables.css">
    <style>
        :root {{
            --blood-red: #8B0000;
            --blood-red-hover: #A00000;
            --border-color: #2c3e50;
            --hover-color: #2a2a2a;
            --text-secondary: #b3b3b3;
            --background-dark: #1e1e1e;
            --background-darker: #252525;
            --text-primary: #ffffff;
        }}

        body {{
            font-family: 'Plus Jakarta Sans', sans-serif;
            margin: 0;
            padding: 20px;
            background-color: var(--background-dark);
            color: var(--text-primary);
        }}

        .container {{
            max-width: 1200px;
            margin: 0 auto;
        }}

        h1 {{
            font-family: 'Special Elite', cursive;
            color: var(--blood-red);
            text-align: center;
            margin-bottom: 30px;
            text-shadow: 2px 2px 4px rgba(0,0,0,0.5);
            font-size: 3.5rem;
            letter-spacing: 2px;
        }}

        /* DataTables Customization */
        .dataTables_wrapper {{
            margin-top: 20px;
            padding: 20px;
            background-color: var(--background-darker);
            border-radius: 8px;
            box-shadow: 0 4px 6px rgba(0,0,0,0.3);
        }}

        .dataTables_filter input {{
            border: 1px solid #2c3e50 !important;
            border-radius: 4px !important;
            padding: 6px 10px !important;
            background-color: var(--background-dark) !important;
            color: var(--text-primary) !important;
        }}

        .dataTables_length select {{
            border: 1px solid #2c3e50 !important;
            border-radius: 4px !important;
            padding: 4px 8px !important;
            background-color: var(--background-dark) !important;
            color: var(--text-primary) !important;
        }}

        /* Table Styling */
        .works-table {{
            background-color: var(--background-darker) !important;
            color: var(--text-primary) !important;
        }}

        .works-table thead th {{
            background-color: #990000 !important;  /* Deeper blood red for headers */
            color: var(--text-primary) !important;
            border-bottom: 2px solid var(--border-color) !important;
        }}

        .works-table tbody td {{
            background-color: var(--background-darker) !important;
            color: var(--text-primary) !important;
            border-bottom: 1px solid var(--border-color) !important;
        }}

        .works-table tbody tr:hover td {{
            background-color: var(--hover-color) !important;
        }}

        /* DataTables specific styling */
        .dataTables_info,
        .dataTables_length label,
        .dataTables_filter label {{
            color: var(--text-secondary) !important;
        }}

        .paginate_button {{
            color: var(--text-secondary) !important;
            background-color: var(--background-darker) !important;
        }}

        .paginate_button.current {{
            color: var(--text-primary) !important;
            background-color: #990000 !important;  /* Matching header color */
            border: 1px solid #990000 !important;
        }}

        /* Links styling */
        .works-table a {{
            color: #cc0000 !important;  /* Brighter red for links */
            text-decoration: none !important;
        }}

        .works-table a:hover {{
            color: #ff0000 !important;  /* Even brighter on hover */
            text-decoration: underline !important;
        }}

        /* Completely remove sort arrows */
        .works-table thead th,
        table.dataTable thead th,
        .works-table thead td,
        table.dataTable thead td {{
            background-image: none !important;
            padding-right: 8px !important;  /* Remove extra padding that was for arrows */
        }}

        /* Override all DataTables sorting classes */
        table.dataTable thead .sorting::before,
        table.dataTable thead .sorting::after,
        table.dataTable thead .sorting_asc::before,
        table.dataTable thead .sorting_asc::after,
        table.dataTable thead .sorting_desc::before,
        table.dataTable thead .sorting_desc::after,
        table.dataTable thead .sorting_asc_disabled::before,
        table.dataTable thead .sorting_asc_disabled::after,
        table.dataTable thead .sorting_desc_disabled::before,
        table.dataTable thead .sorting_desc_disabled::after {{
            display: none !important;
            content: "" !important;
        }}

        /* Optional: Add a subtle cursor change to indicate sortable columns */
        .works-table thead th,
        table.dataTable thead th {{
            cursor: pointer;
        }}
    </style>
    <script type="text/javascript" src="https://code.jquery.com/jquery-3.7.0.min.js"></script>
    <script type="text/javascript" src="https://cdn.datatables.net/1.13.4/js/jquery.dataTables.min.js"></script>
    <script>
        $(document).ready(function() {{
            $('.works-table').DataTable({{
                pageLength: 50,
                order: [[2, 'asc']],
                fixedHeader: true
            }});

            $('.status-checkbox').change(function() {{
                const checkbox = $(this);
                const title = checkbox.data('title');
                const type = checkbox.data('type');
                const isChecked = checkbox.prop('checked');
                console.log(`${{type}} status for "${{title}}" changed to: ${{isChecked}}`);
            }});
        }});
    </script>
</head>
<body>
    <div class="container">
        <h1>Stephen King Bibliography</h1>
        {table_content}
    </div>
</body>
</html>"""

    def fix_missing_dates(self, works_list):
        """
        Fix missing dates (0000-00-00) for works in collections by using the collection's publication date.
        """
        # First pass: Build collection dates dictionary
        collection_dates = {}
        for work in works_list:
            if work.work_type.lower() in [
                "collection",
                "anthology",
                "story collection",
            ]:
                if work.published_date and work.published_date != "0000-00-00":
                    print(
                        f"Found collection: {work.title} with date {work.published_date}"
                    )  # Debug print
                    collection_dates[work.title] = work.published_date

        # Second pass: Update works with missing dates
        for work in works_list:
            if work.published_date == "0000-00-00" and work.available_in:
                collection_title = work.available_in

                # Remove any Excel HYPERLINK formatting if present
                if "=HYPERLINK" in collection_title:
                    match = re.search(
                        r'=HYPERLINK\("[^"]*",\s*"([^"]+)"\)', collection_title
                    )
                    if match:
                        collection_title = match.group(1)

                if collection_title in collection_dates:
                    work.published_date = collection_dates[collection_title]
                    print(
                        f"Updated {work.title} with date {work.published_date} from collection {collection_title}"
                    )  # Debug print

    def parse_and_export(self):
        """Parse works and export to both CSV and HTML formats."""
        response = self.request_manager.get(self.WORKS_URL)
        if not response:
            return

        soup = BeautifulSoup(response.text, "html.parser")
        works = soup.find_all("a", class_="row work")

        formatted_data = []
        headers = ["Read", "Owned", "Published", "Title", "Type", "Available In"]

        # Process works with ThreadPoolExecutor
        processed_works = []
        with ThreadPoolExecutor(max_workers=self.MAX_WORKERS) as executor:
            futures = [executor.submit(self.process_work, work) for work in works]
            for future in concurrent.futures.as_completed(futures):
                if work_data := future.result():
                    processed_works.append(work_data)

        # Fix missing dates before formatting
        self.fix_missing_dates(processed_works)

        # Format data for export
        for work_data in processed_works:
            formatted_row = self.format_row_for_export(work_data)
            formatted_data.append(formatted_row)

        # Generate timestamp for filenames
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        csv_file = f"stephen_king_works_{timestamp}.csv"

        # Export to CSV
        with open(csv_file, "w", newline="", encoding="utf-8") as f:
            writer = csv.writer(f)
            writer.writerow(headers)
            writer.writerows(formatted_data)

        print(f"CSV file '{csv_file}' created successfully!")

        # Export to HTML
        html_file = f"stephen_king_works_{timestamp}.html"
        self.export_to_html(
            html_file, formatted_data
        )  # Changed from save_to_html to export_to_html
        print(f"HTML file '{html_file}' created successfully!")


def main():
    """Main entry point for the Stephen King works parser application.

    Usage:
        python parse_king_works.py           # Fetch new data and generate files
        python parse_king_works.py --html    # Generate HTML from existing CSV
    """
    parser = argparse.ArgumentParser(description="Stephen King Works Parser")
    parser.add_argument(
        "--html", action="store_true", help="Generate HTML from existing CSV only"
    )
    parser.add_argument(
        "--csv",
        type=str,
        help="Input CSV file (default: most recent stephen_king_works_*.csv)",
        default=None,
    )
    args = parser.parse_args()

    if args.html:
        try:
            # Find most recent CSV if not specified
            if not args.csv:
                csv_files = glob.glob("stephen_king_works_*.csv")
                if not csv_files:
                    print(
                        "No CSV files found! Please run without --html first or specify a CSV file."
                    )
                    return
                args.csv = max(csv_files)  # Gets most recent file by name

            print(f"Reading CSV file: {args.csv}")

            # Read CSV with explicit string conversion
            df = pd.read_csv(args.csv, dtype=str, na_filter=False)
            # Replace NaN with empty string
            df = df.fillna("")
            works_data = df.values.tolist()

            # Generate timestamp for HTML filename
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            html_file = f"stephen_king_works_{timestamp}.html"

            print(f"Generating HTML file: {html_file}")

            # Create parser instance and export to HTML
            parser = KingWorksParser()
            parser.export_to_html(html_file, works_data)
            print(f"HTML file '{html_file}' created successfully!")

        except Exception as e:
            print(f"Error generating HTML from CSV: {e}")
            import traceback

            traceback.print_exc()
    else:
        # Original functionality - fetch new data and generate both files
        parser = KingWorksParser()
        parser.parse_and_export()


if __name__ == "__main__":
    main()
