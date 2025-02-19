from dataclasses import dataclass

@dataclass
class Work:
    title: str
    published_date: str
    work_type: str
    url: str
    available_in: str = None
