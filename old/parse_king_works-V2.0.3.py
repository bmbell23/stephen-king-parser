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
from typing import Dict, List, Optional, Tuple
from urllib.parse import urljoin

import pandas as pd  # Move pandas import here with other imports
import requests
from bs4 import BeautifulSoup

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
    formats: str
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
            return datetime.strptime(date_str.strip(), "%Y-%m-%d")
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
        formats = set(format.strip() for format in existing.split(","))
        formats.update(format.strip() for format in new.split(","))
        # Join formats back into sorted, comma-separated string
        return ", ".join(sorted(formats))

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

    def extract_available_formats(self, link: str) -> str:
        """Extract available formats from a work's page."""
        try:
            logger.info(f"Starting format extraction for: {link}")
            response = self.request_manager.get(link)
            if not response:
                logger.warning(f"No response from: {link}")
                return ""

            soup = BeautifulSoup(response.text, "html.parser")
            formats = set()

            # Enhanced format detection
            format_indicators = {
                "Hardcover": ["hardcover", "hard cover", "hard-cover", "hardback"],
                "Paperback": [
                    "paperback",
                    "soft cover",
                    "soft-cover",
                    "trade paperback",
                    "mass market",
                ],
                "Ebook": ["ebook", "e-book", "kindle", "digital", "nook", "electronic"],
                "Audiobook": ["audiobook", "audio book", "audible", "audio"],
                "Movie": ["movie", "film", "feature film", "motion picture"],
                "Miniseries": [
                    "tv series",
                    "television series",
                    "miniseries",
                    "mini-series",
                    "mini series",
                ],
            }

            # Check specific sections first
            format_sections = soup.find_all(
                ["div", "section"], class_=lambda x: x and "format" in x.lower()
            )
            for section in format_sections:
                text = section.get_text(strip=True).lower()
                for format_type, indicators in format_indicators.items():
                    if any(indicator in text for indicator in indicators):
                        formats.add(format_type)
                        logger.info(f"Found format {format_type} in format section")

            # Check all possible containers if we haven't found all formats
            if len(formats) < len(format_indicators):
                containers = soup.find_all(["div", "section", "span", "p", "li", "a"])
                for container in containers:
                    text = container.get_text(strip=True).lower()
                    for format_type, indicators in format_indicators.items():
                        if format_type not in formats and any(
                            indicator in text for indicator in indicators
                        ):
                            formats.add(format_type)
                            logger.info(
                                f"Found format {format_type} in general content"
                            )

            # Check metadata
            meta_description = soup.find("meta", {"name": "description"})
            if meta_description:
                desc_text = meta_description.get("content", "").lower()
                for format_type, indicators in format_indicators.items():
                    if format_type not in formats and any(
                        indicator in desc_text for indicator in indicators
                    ):
                        formats.add(format_type)
                        logger.info(f"Found format {format_type} in metadata")

            result = ", ".join(sorted(formats))
            logger.info(f"Final formats for {link}: {result}")
            return result

        except Exception as e:
            logger.error(f"Error extracting formats from {link}: {str(e)}")
            return ""

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

            # Get formats from the work's individual page
            formats = self.extract_available_formats(work_url)

            logger.info(f"Extracted formats for {title}: {formats}")  # Add this log

            return WorkData(
                title=title,
                cleaned_title=self.clean_title(title),
                link=work_url,
                published_date=published_date,
                work_type=work_type,
                formats=formats,  # Use the extracted formats here
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
            work_data.formats,
        ]

        # Check for existing entry
        if work_data.cleaned_title in self.works_dict:
            existing_date = self.works_dict[work_data.cleaned_title][1]
            new_date = work_data.published_date
            existing_formats = self.works_dict[work_data.cleaned_title][4]
            new_formats = work_data.formats

            # Combine formats
            combined_formats = self.merge_format_strings(existing_formats, new_formats)

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

            # Use combined formats
            row_data[4] = combined_formats

            # Update the entry
            self.works_dict[work_data.cleaned_title] = row_data
            print(f"Updated: {work_data.title} with combined formats and earliest date")
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

    def normalize_format(self, format_str: str) -> str:
        """Normalize format strings to standard values."""
        if not format_str:  # Handle None or empty string
            return ""

        format_str = format_str.strip().lower()

        # Format mappings
        if format_str in ["kindle", "ebook"]:
            return "Yes"
        elif format_str in ["audio", "audiobook"]:
            return "Yes"
        elif format_str in ["movie", "tv movie", "dvd"]:
            return "Yes"
        elif format_str == "tv miniseries":
            return "Yes"
        elif format_str in ["hardcover"]:
            return "Yes"
        elif format_str in ["paperback"]:
            return "Yes"
        else:
            return ""

    def process_formats(self, formats_str: str) -> Dict[str, str]:
        """
        Process formats string into a dictionary of format availability.

        Args:
            formats_str (str): Comma-separated string of formats

        Returns:
            Dict[str, str]: Dictionary with format types as keys and '✓' or '' as values
        """
        formats_dict = {
            "Hardcover": "",
            "Paperback": "",
            "Ebook": "",
            "Audiobook": "",
            "Movie": "",
            "Miniseries": "",
        }

        if not formats_str:
            return formats_dict

        format_list = [fmt.strip() for fmt in formats_str.split(",")]
        for fmt in format_list:
            fmt = fmt.strip()
            if "Hardcover" in fmt:
                formats_dict["Hardcover"] = "✓"
            if "Paperback" in fmt:
                formats_dict["Paperback"] = "✓"
            if "Kindle" in fmt or "eBook" in fmt:
                formats_dict["Ebook"] = "✓"
            if "Audio" in fmt or "Audiobook" in fmt:
                formats_dict["Audiobook"] = "✓"
            if "Movie" in fmt:
                formats_dict["Movie"] = "✓"
            if "TV" in fmt or "Miniseries" in fmt:
                formats_dict["Miniseries"] = "✓"

        return formats_dict

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
        header = [
            "Read",
            "Owned",
            "Published",
            "Title",
            "Type",
            "Available In",
            "Hardcover",
            "Paperback",
            "Ebook",
            "Audiobook",
            "Movie",
            "Miniseries",
        ]

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

        # Extract URL and text from HYPERLINK formula
        match = re.match(r'=HYPERLINK\("([^"]+)",\s*"([^"]+)"\)', excel_formula)
        if match:
            return (match.group(1), match.group(2))
        return ("", excel_formula)

    def excel_hyperlink_to_html(self, excel_formula: str) -> str:
        """Convert Excel HYPERLINK formula to HTML anchor tag."""
        if not excel_formula or not isinstance(excel_formula, str):
            return ""

        if not excel_formula.startswith("=HYPERLINK("):
            return excel_formula

        # Extract URL and text from HYPERLINK("url", "text")
        match = re.match(r'=HYPERLINK\("([^"]+)",\s*"([^"]+)"\)', excel_formula)
        if match:
            url, text = match.groups()
            return f'<a href="{url}">{text}</a>'
        return excel_formula

    def format_row_for_export(self, work_data: WorkData) -> List[str]:
        """Format a single work's data for export"""
        formats_dict = self.process_formats(work_data.formats)

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
            formats_dict["Hardcover"],
            formats_dict["Paperback"],
            formats_dict["Ebook"],
            formats_dict["Audiobook"],
            formats_dict["Movie"],
            formats_dict["Miniseries"],
        ]

    def generate_html_table(self, rows: List[List[str]]) -> str:
        """Generate HTML table with proper formatting."""
        headers = [
            "Read",
            "Owned",
            "Published",
            "Title",
            "Type",
            "Collection",
            "Hardcover",
            "Paperback",
            "Ebook",
            "Audiobook",
            "Movie",
            "Miniseries",
        ]

        html = ['<table class="works-table">']

        # Add header row
        html.append("<thead><tr>")
        for header in headers:
            html.append(f"<th>{header}</th>")
        html.append("</tr></thead>")

        # Add data rows
        html.append("<tbody>")
        for row in rows:
            html.append("<tr>")
            # Handle Read checkbox (column 0)
            title = self.extract_title_from_hyperlink(
                row[3]
            )  # Extract title from the hyperlink in column 3
            html.append(
                f'<td><input type="checkbox" class="status-checkbox" data-type="read" data-title="{title}"></td>'
            )
            # Handle Owned checkbox (column 1)
            html.append(
                f'<td><input type="checkbox" class="status-checkbox" data-type="owned" data-title="{title}"></td>'
            )
            # Handle remaining columns
            for cell in row[2:]:  # Start from column 2 (Published)
                cell_content = self.excel_hyperlink_to_html(cell) if cell else ""
                html.append(f"<td>{cell_content}</td>")
            html.append("</tr>")
        html.append("</tbody>")

        html.append("</table>")
        return "\n".join(html)

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
            -webkit-appearance: none;
            -moz-appearance: none;
            appearance: none;
            position: relative;
            margin: 0;
            padding: 0;
            display: inline-block;
            vertical-align: middle;
        }}

        .status-checkbox:checked {{
            background-color: var(--success-color);
            border-color: var(--success-color);
        }}

        .status-checkbox:checked::after {{
            content: '✓';
            position: absolute;
            color: white;
            font-size: 14px;
            left: 2px;
            top: -1px;
        }}

        .status-checkbox:focus {{
            outline: none;
            box-shadow: 0 0 0 2px rgba(46, 204, 113, 0.2);
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
                order: [[2, 'asc']], // Sort by Published date (column index 2) in ascending order
                columnDefs: [{{
                    targets: 2, // Published date column
                    type: 'date'
                }}]
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

        with open(filename, "w", encoding="utf-8") as f:
            f.write(html_content)

    def save_to_html(self, formatted_data: List[List[str]], output_file: str):
        """Save works to HTML file with DataTables functionality."""
        logger.info(f"Generating HTML output to {output_file}")

        # Generate HTML table
        html_content = self.generate_html_table(formatted_data)

        # Complete HTML document with DataTables
        full_html = """<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Stephen King Works</title>
    <link rel="stylesheet" type="text/css" href="https://cdn.datatables.net/1.10.24/css/jquery.dataTables.css">
    <link rel="stylesheet" type="text/css" href="https://cdn.datatables.net/fixedheader/3.1.8/css/fixedHeader.dataTables.min.css">
    <script type="text/javascript" src="https://code.jquery.com/jquery-3.5.1.min.js"></script>
    <script type="text/javascript" src="https://cdn.datatables.net/1.10.24/js/jquery.dataTables.min.js"></script>
    <script type="text/javascript" src="https://cdn.datatables.net/fixedheader/3.1.8/js/dataTables.fixedHeader.min.js"></script>
    <style>
        .works-table {
            width: 100% !important;
            margin: 0 !important;
            border-collapse: collapse;
            table-layout: fixed;
        }

        .works-table td,
        .works-table th {
            padding: 0.6rem 0.4rem !important;
            font-size: 0.95rem;
            overflow: hidden;
            text-overflow: ellipsis;
        }

        /* READ/OWNED columns */
        .works-table th:nth-child(1),
        .works-table td:nth-child(1),
        .works-table th:nth-child(2),
        .works-table td:nth-child(2) {
            width: 40px;
            min-width: auto;
            padding: 0.6rem 0.2rem !important;
            text-align: center;
        }

        /* Published Date */
        .works-table th:nth-child(3),
        .works-table td:nth-child(3) {
            width: 80px;
            white-space: nowrap;
        }

        /* Title */
        .works-table th:nth-child(4),
        .works-table td:nth-child(4) {
            width: 35%;
            white-space: normal;
        }

        /* Description */
        .works-table th:nth-child(5),
        .works-table td:nth-child(5) {
            width: 35%;
            white-space: normal;
        }

        /* Format columns */
        .works-table th:nth-child(6),
        .works-table td:nth-child(6),
        .works-table th:nth-child(7),
        .works-table td:nth-child(7),
        .works-table th:nth-child(8),
        .works-table td:nth-child(8),
        .works-table th:nth-child(9),
        .works-table td:nth-child(9),
        .works-table th:nth-child(10),
        .works-table td:nth-child(10),
        .works-table th:nth-child(11),
        .works-table td:nth-child(11) {
            width: 45px;
            white-space: nowrap;
            text-align: center;
            padding: 0.6rem 0.1rem !important;
            font-size: 0.8rem;
        }

        .container {
            max-width: 99%;
            margin: 0 auto;
            padding: 0.8rem;
            overflow-x: hidden;
        }

        @media (max-width: 1200px) {
            .works-table td,
            .works-table th {
                font-size: 0.9rem;
            }

            .works-table td:nth-child(2) {
                font-size: 0.75rem;
            }
        }

        @media (max-width: 992px) {
            .container {
                max-width: 100%;
                padding: 0.4rem;
            }

            .works-table td,
            .works-table th {
                padding: 0.5rem 0.3rem !important;
            }
        }
    </style>
    <script>
        $(document).ready(function() {
            $('.works-table').DataTable({
                responsive: true,
                autoWidth: false,
                pageLength: 25,
                order: [[2, 'asc']],
                columnDefs: [
                    {
                        targets: 2,
                        type: 'date'
                    },
                    {
                        targets: [0, 1],
                        orderable: false,
                        width: 'auto'
                    }
                ],
                scrollX: false
            });
        });
    </script>
</head>
<body>
    <div class="container">
        <h1>Stephen King Works</h1>
        {html_content}
    </div>
</body>
</html>"""

        # Write to file
        try:
            with open(output_file, "w", encoding="utf-8") as f:
                f.write(full_html)
            logger.info(f"Successfully wrote HTML to {output_file}")
        except Exception as e:
            logger.error(f"Error writing HTML file: {str(e)}")

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
        """Main method to parse works and export to CSV and HTML"""
        response = self.request_manager.get(self.WORKS_URL)
        if not response:
            return

        soup = BeautifulSoup(response.text, "html.parser")
        works = soup.find_all("a", class_="row work")

        formatted_data = []
        headers = [
            "Read",
            "Owned",
            "Published",
            "Title",
            "Type",
            "Available In",
            "Hardcover",
            "Paperback",
            "Ebook",
            "Audiobook",
            "Movie",
            "Miniseries",
        ]

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
        self.save_to_html(formatted_data, html_file)  # Removed headers argument
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
