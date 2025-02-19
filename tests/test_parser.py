import pytest
from pathlib import Path
from stephen_king_parser.core.parser import KingWorksParser
from stephen_king_parser.core.models import Work, WorkType
from stephen_king_parser.utils.config import Settings

@pytest.fixture
def sample_html():
    return Path("tests/fixtures/sample_pages/work_list.html").read_text()

@pytest.fixture
def parser():
    config = Settings()
    return KingWorksParser(config)

@pytest.mark.asyncio
async def test_parse_work_list(parser, sample_html):
    works = await parser.parse_work_list(sample_html)
    assert len(works) > 0
    assert all(isinstance(work, Work) for work in works)

@pytest.mark.asyncio
async def test_parse_work_details(parser, sample_html):
    work = await parser.parse_work_details(sample_html)
    assert isinstance(work, Work)
    assert work.title
    assert work.published_date