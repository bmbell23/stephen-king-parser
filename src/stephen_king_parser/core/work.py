class Work:
    def __init__(
        self,
        title: str,
        published_date: str = None,
        work_type: str = None,
        url: str = None,
        available_in: str = None,
    ):
        self.title = title
        self.published_date = published_date
        self.work_type = work_type
        self.url = url  # Changed from 'link' to 'url' to match usage
        self.available_in = available_in

    def __str__(self):
        return f"{self.title} ({self.published_date}) - {self.work_type}"

    def __repr__(self):
        return f"Work(title='{self.title}', published_date='{self.published_date}', work_type='{self.work_type}')"
