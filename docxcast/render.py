"""Render data back into a docx template: fill controls, clone repeats, unwrap."""

from __future__ import annotations

import re
from calendar import month_abbr, month_name
from copy import deepcopy
from dataclasses import dataclass
from io import BytesIO
from pathlib import Path
from typing import Optional

from docx import Document
from lxml import etree

from docxcast._ooxml import (
    CONTAINER_TAGS,
    KIND_CHECKBOX,
    KIND_DATE,
    KIND_DROPDOWN,
    KIND_GROUP,
    KIND_ITEM,
    KIND_PICTURE,
    group_members,
    KIND_REPEAT,
    KIND_SECTION,
    SELF,
    XML_SPACE,
    dropdown_items,
    date_mask_runs,
    group_address,
    group_chunks,
    group_literals,
    group_members,
    kind_of,
    qn,
    repeating_item,
    scope_guards,
    split_conditional,
    text_attr,
)
from docxcast.derive import _derive
from docxcast.schema import Issue, ValidationResult
from docxcast.sorting import apply_sorts


class RenderError(Exception):
    """Raised by render(strict=True) on the first error-level validation issue."""

    def __init__(self, issue: Issue):
        self.issue = issue
        super().__init__(f'{issue.path}: {issue.message}')


@dataclass
class RenderResult:
    document: bytes
    report: ValidationResult

    def save(self, path) -> None:
        Path(path).write_bytes(self.document)


def render(template, data, *, strict: bool = False, keep_controls: bool = False) -> RenderResult:
    """Fill data into the template's content controls and return the rendered docx.

    Controls are unwrapped by default so the output is a clean document; pass
    keep_controls=True to keep the control structure for round-trip editing.
    The template is re-parsed on every call, data is validated against it,
    and the repeat-level @sort directives are applied — the rule-carrying
    arrays of ``data`` are reordered in place, so the caller's dict and
    the rendered document agree on order.
    """
    doc = Document(template)
    schema = _derive(doc)
    report = schema.validate(data)
    if strict and report.errors:
        raise RenderError(report.errors[0])
    apply_sorts(schema, data)
    _fill_children(doc.element.body, _as_dict(data), keep_controls)
    buf = BytesIO()
    doc.save(buf)
    return RenderResult(document=buf.getvalue(), report=report)


def _fill_children(container, data: dict, keep_controls: bool) -> None:
    for child in list(container):
        if child.tag == qn('w:sdt'):
            _handle_sdt(child, data, keep_controls)
        elif child.tag in CONTAINER_TAGS:
            _fill_children(child, data, keep_controls)


def _handle_sdt(sdt, data: dict, keep_controls: bool) -> None:
    pr = sdt.find(qn('w:sdtPr'))
    content = sdt.find(qn('w:sdtContent'))
    if pr is None or content is None:
        return
    # derive rejects '?' outside conditional groups, so by fill time a
    # marked address is always a conditional group's member
    alias, _ = split_conditional(text_attr(pr, 'w:alias'))
    kind = kind_of(pr)
    value = data.get(alias) if alias else None

    if kind == KIND_REPEAT:
        if value is None:
            _unwrap(sdt, [], keep_controls)
            return
        item = repeating_item(content)
        if item is None:
            return
        item_content = item.find(qn('w:sdtContent'))
        scalar = _has_self(item_content)
        blocks = []
        for i, entry in enumerate(value):
            clone = deepcopy(item)
            clone_content = clone.find(qn('w:sdtContent'))
            scope = ({SELF: entry} if scalar else _as_dict(entry)) | scope_guards(i, len(value))
            _fill_children(clone_content, scope, keep_controls)
            if keep_controls:
                blocks.append(clone)
            else:
                blocks.extend(clone_content)
        if keep_controls:
            content[:] = blocks
        else:
            _splice(sdt, blocks)
        return

    if kind in (KIND_SECTION, KIND_ITEM):
        if kind == KIND_SECTION and value is None:
            _unwrap(sdt, [], keep_controls)
            return
        _fill_children(content, _as_dict(value) if kind == KIND_SECTION else data, keep_controls)
        _unwrap(sdt, content, keep_controls)
        return

    if kind == KIND_PICTURE:
        _unwrap(sdt, [], False)  # picture controls are always removed
        return

    if kind == KIND_GROUP:
        _fill_group(sdt, pr, content, data, keep_controls)
        return

    if value is not None:
        _fill_leaf(content, kind, value, pr)
    _unwrap(sdt, content, keep_controls)


def _fill_group(sdt, pr, content, data: dict, keep_controls: bool) -> None:
    """A group renders exactly one of its branches: the value either
    fills the control (literal text dropped) or is one of the literals
    (control dropped). keep_controls=True keeps both sides editable
    instead of erasing the unchosen branch. A conditional group (its
    address ends in '?') instead gates its whole block on the value:
    falsy drops the block — label and all — truthy fills transparently
    with the literal text kept as decoration."""
    members = group_members(content)
    name, _, _, conditional = group_address(pr, members)
    value = data.get(name) if name else None
    if conditional:
        if not value:
            _unwrap(sdt, [], keep_controls)
            return
        _fill_children(content, data, keep_controls)
        _unwrap(sdt, content, keep_controls)
        return
    literals = group_literals(content)
    if value in literals:
        for msdt, *_ in members:
            _unwrap(msdt, [], keep_controls)
    elif value is not None and members:
        if not keep_controls:
            _drop_literal_runs(content, set(literals))  # before the fill unwraps the leaf
        msdt, mpr, _, _ = members[0]
        mcontent = msdt.find(qn('w:sdtContent'))
        if mcontent is not None:
            _fill_leaf(mcontent, kind_of(mpr), value, mpr)
            _unwrap(msdt, mcontent, keep_controls)
    _unwrap(sdt, content, keep_controls)


def _drop_literal_runs(content, texts: set) -> None:
    for holder, text in group_chunks(content):
        if text in texts:
            for r in [c for c in holder if c.tag == qn('w:r')]:
                holder.remove(r)


def _fill_leaf(content, kind: str, value, pr) -> None:
    if kind == KIND_CHECKBOX:
        _set_checkbox(content, pr, bool(value))
    elif kind == KIND_DROPDOWN:
        _set_text(content, _display_text(pr, value) or str(value))
    elif kind == KIND_DATE:
        _set_text(content, _date_text(pr, value))
    else:
        _set_text(content, str(value))


def _set_checkbox(content, pr, checked: bool) -> None:
    checkbox = pr.find(qn('w14:checkbox'))
    state = None
    if checkbox is not None:
        checked_el = checkbox.find(qn('w14:checked'))
        if checked_el is None:
            checked_el = etree.SubElement(checkbox, qn('w14:checked'))
        checked_el.set(qn('w14:val'), '1' if checked else '0')
        state = checkbox.find(qn('w14:checkedState') if checked else qn('w14:uncheckedState'))
    glyph = '☒' if checked else '☐'
    if state is not None and (val := state.get(qn('w14:val'))):
        glyph = chr(int(val, 16))
    _set_text(content, glyph)


def _display_text(pr, value) -> Optional[str]:
    target = str(value)
    for display, item_value in dropdown_items(pr):
        if item_value == target:
            return display
    return None


def _date_text(pr, value) -> str:
    """A date control fills with its own display mask — the same text Word
    would write had the user picked the date in the picker. Values that
    don't parse fall back to the raw string. The w:date element is
    guaranteed: _fill_leaf only routes KIND_DATE controls here."""
    s = str(value)
    mask = text_attr(pr.find(qn('w:date')), 'w:dateFormat')
    if not mask or (parts := _date_parts(s)) is None:
        return s
    return _apply_mask(mask, *parts)


# the shapes schema's four format regexes accept, capture-grouped: month
# only opens day, day only opens the clock — a time without a date never
# parses
_ISO_RE = re.compile(r'^(\d{4})(?:-(\d{2})(?:-(\d{2})'
                     r'(?:[T ](\d{2}):(\d{2})(?::(\d{2}))?)?)?)?')


def _date_parts(value: str) -> Optional[tuple]:
    m = _ISO_RE.match(value)
    if m is None:
        return None
    return tuple(int(p) if p else None for p in m.groups())


def _apply_mask(mask: str, year: int, month: Optional[int], day: Optional[int],
                hour: Optional[int], minute: Optional[int],
                second: Optional[int]) -> str:
    """Render a date-time through a Word dateFormat mask, run by run. A
    run's letter dispatches the field-code units (y/M/d, H/h 24-12-hour,
    m minute, s second — the units a mask's own granularity guarantees
    are present); unrecognized letters and literals pass through. Month
    names are the locale-independent English tables."""
    out = []
    for letter, text in date_mask_runs(mask):
        width = len(text)
        if letter == 'y':
            out.append(f'{year:04d}' if width > 2 else f'{year % 100:02d}')
        elif letter == 'M' and month is not None:
            out.append(str(month) if width == 1 else month_abbr[month] if width == 3
                       else month_name[month] if width == 4 else f'{month:02d}')
        elif letter == 'd' and day is not None:
            out.append(str(day) if width == 1 else f'{day:02d}')
        elif letter in ('h', 'H') and hour is not None:
            h = hour if letter == 'H' else (hour % 12 or 12)
            out.append(str(h) if width == 1 else f'{h:02d}')
        elif letter == 'm' and minute is not None:
            out.append(str(minute) if width == 1 else f'{minute:02d}')
        elif letter == 's' and second is not None:
            out.append(str(second) if width == 1 else f'{second:02d}')
        else:
            out.append(text)
    return ''.join(out)


def _set_text(content, text: str) -> None:
    runs = content.findall(qn('w:r'))
    r = runs[0] if runs else etree.SubElement(content, qn('w:r'))
    r_pr = r.find(qn('w:rPr'))
    for extra in runs[1:]:
        content.remove(extra)
    r.clear()
    if r_pr is not None:
        r.append(r_pr)
    t = etree.SubElement(r, qn('w:t'))
    t.set(XML_SPACE, 'preserve')
    t.text = text


def _unwrap(sdt, nodes, keep_controls: bool) -> None:
    if not keep_controls:
        _splice(sdt, list(nodes))


def _splice(old, nodes) -> None:
    parent = old.getparent()
    if parent is None:
        return
    for node in nodes:
        old.addprevious(node)
    parent.remove(old)
    if parent.tag == qn('w:tc') and not parent.findall(qn('w:p')):
        etree.SubElement(parent, qn('w:p'))  # a cell must keep a paragraph


def _as_dict(value) -> dict:
    return value if isinstance(value, dict) else {}


def _has_self(item_content) -> bool:
    """The item carries the '$' self address at its OWN level: each
    entry fills that control directly. A nested repeat's own '$' leaves
    belong to it, so the walk never enters a control's content —
    group_members' descent rule, shared with derive."""
    return any(alias == SELF for _, _, alias, _ in group_members(item_content))
