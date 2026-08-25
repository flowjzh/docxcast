"""DocXCast: Turn Word content controls into a typed schema, then cast data back into documents."""

from importlib.metadata import PackageNotFoundError, version

from docxcast.derive import DeriveError, derive_schema
from docxcast.render import RenderError, RenderResult, render
from docxcast.schema import (
    FORMAT_DATE,
    FORMAT_DATE_MONTH,
    FORMAT_DATE_TIME,
    FORMAT_DATE_YEAR,
    Field,
    Issue,
    Repeat,
    Schema,
    Section,
    ValidationResult,
)

try:
    __version__ = version('docxcast')
except PackageNotFoundError:
    __version__ = '0.1.0'

__all__ = [
    'derive_schema',
    'DeriveError',
    'render',
    'RenderError',
    'RenderResult',
    'Schema',
    'Section',
    'Repeat',
    'Field',
    'Issue',
    'ValidationResult',
    'FORMAT_DATE',
    'FORMAT_DATE_MONTH',
    'FORMAT_DATE_YEAR',
    'FORMAT_DATE_TIME',
]
