from pathlib import Path

import pytest


@pytest.fixture
def template() -> Path:
    return Path(__file__).parent / 'fixtures' / 'resume-template.docx'
