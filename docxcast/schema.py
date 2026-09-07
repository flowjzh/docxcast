"""Schema model: the contract derived from a docx template and validated against data."""

from __future__ import annotations

import re
from dataclasses import asdict, dataclass, field
from typing import Optional, Union

class _Missing:
    """Sentinel that survives deepcopy so asdict-based filtering works."""

    def __deepcopy__(self, memo):
        return self


MISSING = _Missing()

MISSING_REQUIRED = 'missing_required'
TYPE_MISMATCH = 'type_mismatch'
ENUM_INVALID = 'enum_invalid'
FORMAT_INVALID = 'format_invalid'
UNKNOWN_KEY = 'unknown_key'

FORMAT_DATE = 'date'
FORMAT_DATE_MONTH = 'date-month'
FORMAT_DATE_YEAR = 'date-year'
FORMAT_DATE_TIME = 'date-time'
TYPE_BOOLEAN = 'boolean'

_MESSAGES = {
    MISSING_REQUIRED: 'Required value is missing.',
    TYPE_MISMATCH: 'Value has the wrong type.',
    ENUM_INVALID: 'Value is not among the allowed options.',
    FORMAT_INVALID: 'Value does not match the expected format.',
    UNKNOWN_KEY: 'Key is not defined in the schema.',
}

# everything except the warning-level code is an error
ERROR_CODES = frozenset(_MESSAGES) - {UNKNOWN_KEY}

_DATE_RE = re.compile(r'^\d{4}-\d{2}-\d{2}$')
_DATE_MONTH_RE = re.compile(r'^\d{4}-\d{2}$')
_DATE_YEAR_RE = re.compile(r'^\d{4}$')
_DATETIME_RE = re.compile(
    r'^\d{4}-\d{2}-\d{2}[T ]\d{2}:\d{2}(:\d{2})?(\.\d+)?(Z|[+-]\d{2}:?\d{2})?$'
)

_FORMATS = (
    (FORMAT_DATE, _DATE_RE, 'ISO 8601 date (YYYY-MM-DD)'),
    (FORMAT_DATE_MONTH, _DATE_MONTH_RE, 'ISO 8601 month (YYYY-MM)'),
    (FORMAT_DATE_YEAR, _DATE_YEAR_RE, 'year (YYYY)'),
    (FORMAT_DATE_TIME, _DATETIME_RE, 'ISO 8601 date-time'),
)

# the regex source of truth, emitted as JSON Schema `pattern` so the
# format branch is enforced by any validator, not just this one
_PATTERNS = {fmt: rx.pattern for fmt, rx, _ in _FORMATS}


@dataclass
class Issue:
    """A single validation issue. `message` is English and self-contained."""

    path: str
    code: str
    message: str
    expected: object = MISSING
    got: object = MISSING

    def to_dict(self) -> dict:
        return {k: v for k, v in asdict(self).items() if v is not MISSING}


@dataclass
class Field:
    name: str
    remark: Optional[str] = None
    type: str = 'string'  # 'string' | 'boolean'
    format: Optional[str] = None  # 'date' | 'date-month' | 'date-year' | 'date-time'
    enum: Optional[list] = None
    alternatives: Optional[list] = None  # literal values a group shows instead of the format
    enum_labels: Optional[list] = None  # display labels parallel to enum
    required: bool = False

    def to_json_schema(self) -> dict:
        node = self._value_branch()
        if self.alternatives:
            # a union carries no type of its own — each branch does
            node = {'anyOf': [node, {'enum': self.alternatives}]}
        if self.enum is not None and not self.required:
            # an optional field may stay unfilled — '' rides the emitted
            # contract so strict validators (jsonschema etc.) agree with
            # _validate_field, not just this one
            node = {'anyOf': [node, {'enum': ['']}]}
        if (desc := self._description()) is not None:
            node['description'] = desc
        return node

    def _value_branch(self) -> dict:
        """The field's value node — a union wraps it in anyOf."""
        branch = {'type': self.type}
        for key, val in (
            ('format', self.format),
            ('enum', self.enum),
            ('pattern', _PATTERNS.get(self.format)),
        ):
            if val is not None:
                branch[key] = val
        return branch

    def _description(self) -> Optional[str]:
        if self.enum_labels and self.enum:
            mapping = '; '.join(f'{d}={v}' for d, v in zip(self.enum_labels, self.enum))
            return f'{self.remark} ({mapping})' if self.remark else mapping
        return self.remark


@dataclass
class SortSpec:
    """One repeat-level sort directive: ``@sort=<key>[:asc|desc]`` in the
    repeat control's tag."""

    key: str
    desc: bool
    kind: str


def sort_kind(fmt: Optional[str]) -> str:
    """The sort kind a field's declared format names — the one
    format→kind mapping, beside the FORMAT_* constants."""
    if fmt == FORMAT_DATE_TIME:
        return 'datetime'
    if fmt in (FORMAT_DATE, FORMAT_DATE_MONTH, FORMAT_DATE_YEAR):
        return 'date'
    return 'alphanumeric'


@dataclass
class Repeat:
    name: str
    remark: Optional[str] = None
    fields: list = field(default_factory=list)
    required: bool = False
    item: Optional[Field] = None  # a '$'-addressed item: the array holds its values
    sort: Optional[SortSpec] = None  # @sort directive from the control's tag

    def to_json_schema(self) -> dict:
        items = (self.item.to_json_schema() if self.item is not None
                 else _object_schema(self.fields))
        node = {'type': 'array', 'items': items}
        if self.remark is not None:
            node['description'] = self.remark
        return node


@dataclass
class Section:
    name: str
    remark: Optional[str] = None
    children: list = field(default_factory=list)
    required: bool = False

    def to_json_schema(self) -> dict:
        return _object_schema(self.children, self.remark)


Node = Union[Field, Repeat, Section]


@dataclass
class Schema:
    children: list = field(default_factory=list)
    # the template's Comments core property; guidance for the extraction
    # model, never a constraint
    description: Optional[str] = None

    def to_json_schema(self) -> dict:
        return _object_schema(self.children, self.description)

    def to_dict(self) -> dict:
        d = asdict(self)
        if d['description'] is None:  # omit-when-absent, like the emission path
            del d['description']
        return d

    def validate(self, data: object) -> ValidationResult:
        issues: list = []
        if data is not None:
            _validate_object(self.children, data, '', issues)
        return ValidationResult(issues=issues)


@dataclass
class ValidationResult:
    issues: list = field(default_factory=list)

    @property
    def errors(self) -> list:
        return [i for i in self.issues if i.code in ERROR_CODES]

    @property
    def warnings(self) -> list:
        return [i for i in self.issues if i.code not in ERROR_CODES]

    @property
    def ok(self) -> bool:
        return not self.errors

    def to_dict(self) -> dict:
        return {'ok': self.ok, 'issues': [i.to_dict() for i in self.issues]}


def _nullable(schema: dict) -> dict:
    """An optional object may also be null — the shape the model reads:
    an omitted structure field is absent, not filled with an empty
    object. Arrays keep a single type: empty already reads falsy to
    the renderer, and provider-side schema validators may reject type
    unions on arrays."""
    if schema.get('type') == 'object':
        schema['type'] = ['object', 'null']
    return schema


def _object_schema(children: list, description: Optional[str] = None) -> dict:
    node = {'type': 'object'}
    if children:
        node['properties'] = {
            c.name: c.to_json_schema() if c.required
            else _nullable(c.to_json_schema()) for c in children}
    required = [c.name for c in children if c.required]
    if required:
        node['required'] = required
    if description is not None:
        node['description'] = description
    return node


def _validate_children(nodes: list, data: dict, prefix: str, issues: list) -> None:
    known = set()
    for node in nodes:
        known.add(node.name)
        value = data.get(node.name, MISSING)
        if value is MISSING:
            if node.required:
                issues.append(_issue(_path(prefix, node.name), MISSING_REQUIRED))
            continue
        _validate_node(node, value, _path(prefix, node.name), issues)
    for key in data:
        if key not in known:
            issues.append(_issue(_path(prefix, key), UNKNOWN_KEY))


def _validate_node(node: Node, value: object, path: str, issues: list) -> None:
    if value is None:
        if node.required:
            issues.append(_issue(path, MISSING_REQUIRED))
        return  # null is an absent value, allowed on every optional node
    if isinstance(node, Section):
        _validate_object(node.children, value, path, issues)
    elif isinstance(node, Repeat):
        if not isinstance(value, list):
            issues.append(_issue(path, TYPE_MISMATCH, expected='array', got=value))
            return
        for i, item in enumerate(value):
            if node.item is not None:
                _validate_field(node.item, item, f'{path}[{i}]', issues)
            else:
                _validate_object(node.fields, item, f'{path}[{i}]', issues)
    else:
        _validate_field(node, value, path, issues)


def _validate_object(nodes: list, value: object, path: str, issues: list) -> None:
    if not isinstance(value, dict):
        issues.append(_issue(path, TYPE_MISMATCH, expected='object', got=value))
        return
    _validate_children(nodes, value, path, issues)


def _validate_field(field: Field, value: object, path: str, issues: list) -> None:
    if field.alternatives and value in field.alternatives:
        return  # a literal alternative satisfies the union on its own
    if field.type == TYPE_BOOLEAN:
        if not isinstance(value, bool):
            issues.append(_issue(path, TYPE_MISMATCH, expected=TYPE_BOOLEAN, got=value))
        return
    if not isinstance(value, str):
        issues.append(_issue(path, TYPE_MISMATCH, expected=field.type, got=value))
        return
    if value == '' and not field.required:
        return  # an optional field may stay unfilled: emptiness is governed
                # by required (the control's lock), never by enum membership
    if field.enum is not None and value not in field.enum:
        issues.append(_issue(path, ENUM_INVALID, expected=field.enum, got=value))
    for fmt, rx, label in _FORMATS:
        if field.format == fmt and not rx.match(value):
            issues.append(_issue(path, FORMAT_INVALID, expected=label, got=value))


def _issue(path: str, code: str, expected: object = MISSING, got: object = MISSING) -> Issue:
    return Issue(path=path, code=code, message=_MESSAGES[code], expected=expected, got=got)


def _path(prefix: str, name: str) -> str:
    return f'{prefix}.{name}' if prefix else name
