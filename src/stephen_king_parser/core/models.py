from dataclasses import dataclass
from datetime import datetime


@dataclass
class Work:
    title: str
    published_date: str
    work_type: str
    url: str
