from bs4 import BeautifulSoup
from typing import Optional
from ..models.work import Work

class WorkProcessor:
    @staticmethod
    def process_work(soup: BeautifulSoup) -> Optional[Work]:
        """Process a single work element more efficiently"""
        try:
            # Extract data using direct selectors
            title = soup.select_one('.title').get_text(strip=True)
            link = f"{base_url}{soup.get('href', '')}"

            # Use more efficient date extraction
            date_elem = soup.select_one('.date')
            published_date = None
            if date_elem:
                date_str = date_elem.get_text(strip=True)
                try:
                    published_date = datetime.strptime(date_str, '%Y-%m-%d')
                except ValueError:
                    published_date = datetime.max

            # Optimize type extraction
            type_elem = soup.select_one('.type')
            work_type = type_elem.get_text(strip=True) if type_elem else "Unknown"

            # Efficient collection processing
            collection_elem = soup.select_one('.collection')
            collection = collection_elem.get_text(strip=True) if collection_elem else None
            collection_link = None
            if collection_elem and collection_elem.find('a'):
                collection_link = f"{base_url}{collection_elem.find('a').get('href', '')}"

            return Work(
                title=title,
                link=link,
                published_date=published_date,
                work_type=work_type,
                collection=collection,
                collection_link=collection_link
            )

        except Exception as e:
            logging.error(f"Error processing work element: {e}")
            return None
