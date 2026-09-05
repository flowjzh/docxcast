"""Derive a Schema from the content controls of a docx template."""

from __future__ import annotations

import re

from docx import Document

from docxcast._ooxml import (
    CONTAINER_TAGS,
    KIND_CHECKBOX,
    KIND_DATE,
    KIND_DROPDOWN,
    KIND_GROUP,
    KIND_ITEM,
    KIND_REPEAT,
    KIND_SECTION,
    RESERVED,
    SELF,
    date_mask_runs,
    dropdown_items,
    group_address,
    group_literals,
    group_members,
    is_block,
    kind_of,
    qn,
    repeating_item,
    split_conditional,
    text_attr,
)

from docxcast.schema import (
    FORMAT_DATE,
    FORMAT_DATE_MONTH,
    FORMAT_DATE_TIME,
    FORMAT_DATE_YEAR,
    TYPE_BOOLEAN,
    Field,
    Repeat,
    Schema,
    Section,
    SortSpec,
    sort_kind,
)


def derive_schema(template) -> Schema:
    """Read the content controls of a docx template and return its Schema.

    template: a path-like or file-like object holding a .docx file.
    """
    return _derive(Document(template))


class DeriveError(ValueError):
    """Raised when a template addresses one field twice — e.g. two named
    members inside a single group. A ValueError so host apps map it to
    their bad-input response, not a crash."""


def _derive(doc: Document) -> Schema:
    children = []
    _parse_children(doc.element.body, children)
    return Schema(children=children)


def _parse_children(container, out: list, repeat: bool = False) -> None:
    for child in container:
        if child.tag == qn('w:sdt'):
            _parse_sdt(child, out, repeat)
        elif child.tag in CONTAINER_TAGS:
            _parse_children(child, out, repeat)


def _parse_sdt(sdt, out: list, repeat: bool = False) -> None:
    pr = sdt.find(qn('w:sdtPr'))
    if pr is None:
        return
    name = text_attr(pr, 'w:alias')
    remark = text_attr(pr, 'w:tag')
    kind = kind_of(pr)
    if kind in (KIND_REPEAT, KIND_SECTION) and name is None:
        return  # unaddressable, skip entirely
    if kind != KIND_GROUP:
        if split_conditional(name)[1]:
            raise DeriveError(f'{name!r}: a trailing "?" marks a conditional block — '
                              'wrap the block in a group control')
        if name in RESERVED:
            raise DeriveError(f'{name!r} is reserved for the positional guards '
                              f'({"/".join(sorted(RESERVED))}) inside repeats')
        if name == SELF and (kind in (KIND_REPEAT, KIND_SECTION, KIND_GROUP) or not repeat):
            raise DeriveError(f'{SELF!r} addresses a repeating item\'s own value — '
                              'it belongs on the control inside a repeating item')
    content = sdt.find(qn('w:sdtContent'))
    if kind == KIND_REPEAT:
        out.append(_parse_repeat(content, name, remark, _is_locked(pr)))
    elif kind == KIND_ITEM:
        _parse_children(content, out, repeat)
    elif kind == KIND_SECTION:
        out.append(_parse_section(content, name, remark, _is_locked(pr)))
    elif kind == KIND_GROUP:
        if (group := _parse_group(pr, content, repeat)) is not None:
            out.append(group)
    elif name is not None:
        out.append(_parse_field(pr, name, remark, _is_locked(pr), kind))


def _parse_repeat(content, name, remark, required) -> Repeat:
    fields = []
    if (item := repeating_item(content)) is not None:
        _parse_children(item.find(qn('w:sdtContent')), fields, repeat=True)
    remark, sort = _split_sort(name, remark, fields)
    if selfs := [f for f in fields if f.name == SELF]:
        if len(fields) > 1:
            raise DeriveError(
                f'{name!r}: {SELF!r} addresses the item itself — no other '
                'addressed control may share the item')
        selfs[0].required = True  # every entry is a value; there is no way to omit it
        return Repeat(name=name, remark=remark, item=selfs[0], required=required)
    return Repeat(name=name, remark=remark, fields=fields, required=required,
                  sort=sort)


_SORT_RE = re.compile(r'\s*@sort=([\w.]+)(?::(asc|desc))?\s*$')


def _split_sort(name: str, remark, fields: list) -> tuple[str | None, SortSpec | None]:
    """The repeat tag's trailing ``@sort=<key>[:asc|desc]`` directive,
    split off the remark — the schema description never carries it,
    ordering is code's job, not the model's. An unknown key or a
    mangled directive fails at derive time."""
    if remark is None:
        return None, None
    m = _SORT_RE.search(remark)
    if m is None:
        if '@sort' in remark:
            raise DeriveError(f'{name!r}: unparseable @sort directive')
        return remark, None
    key = m.group(1)
    field = _sort_field(name, fields, key)
    return (remark[:m.start()].rstrip(),
            SortSpec(key=key, desc=m.group(2) == 'desc',
                     kind=sort_kind(field.format)))


def _sort_field(name: str, fields: list, key: str) -> Field:
    """The field a sort key addresses. The key may walk sections; a
    repeat ends the walk — its item's fields are another repeat's own
    scope."""
    nodes = fields
    node = None
    for part in key.split('.'):
        node = next((f for f in nodes if f.name == part), None)
        if node is None:
            raise DeriveError(f'{name!r}: @sort key {key!r} is not a field of the repeat item')
        if isinstance(node, Repeat):
            raise DeriveError(f'{name!r}: @sort key {key!r} walks through a repeat')
        nodes = node.children if isinstance(node, Section) else []
    if not isinstance(node, Field):
        raise DeriveError(f'{name!r}: @sort key {key!r} must address a field')
    return node


def _parse_section(content, name, remark, required) -> Section:
    # a section fills from its own dict at render, so a repeat item's
    # positional-guard scope ends here
    children = []
    _parse_children(content, children)
    return Section(name=name, remark=remark, children=children, required=required)


def _parse_field(pr, name: str, remark, required: bool, kind: str) -> Field:
    kwargs = {}
    if kind == KIND_DROPDOWN:
        pairs = dropdown_items(pr)
        kwargs.update(enum=[v for _, v in pairs], enum_labels=[d for d, _ in pairs])
    elif kind == KIND_DATE:
        date = pr.find(qn('w:date'))
        kwargs['format'] = (
            FORMAT_DATE_TIME if text_attr(date, 'w:storeMappedDataAs') == 'dateTime'
            else _granularity(text_attr(date, 'w:dateFormat'))
        )
    elif kind == KIND_CHECKBOX:
        kwargs['type'] = TYPE_BOOLEAN
    return Field(name=name, remark=remark, required=required, **kwargs)


def _parse_group(pr, content, repeat: bool = False):
    """One field assembled from a group's members: the group's own
    address if it carries one, else the first member that does; every
    other member must stay anonymous and no second control is allowed.
    The control member gives type/format; the group's literal text
    becomes the value alternatives. A trailing '?' on the address turns
    the group into a conditional block instead: the literal text stays
    decoration (never alternatives), render drops the whole block when
    the value is falsy, and the guarded member may be a repeat or
    section, not just a control."""
    members = group_members(content)
    name, remark, source, conditional = group_address(pr, members)
    for i, (_, _, alias, tag) in enumerate(members):
        if (alias or tag) and i != source:
            raise DeriveError(
                f'group {name!r}: member {alias or tag!r} also carries an address '
                f'(w:alias/w:tag); a group is one field, address it once')
    if len(members) > 1:
        raise DeriveError(
            f'group {name!r}: {len(members)} control members; a group is one '
            f'field — one control, the rest literal text')
    if name is None:
        return None  # unaddressable, skipped like unaddressable repeats
    if name in RESERVED:
        if not conditional:
            raise DeriveError(f'{name!r} is reserved for the positional guards '
                              f'({"/".join(sorted(RESERVED))}) inside repeats')
        if not repeat:
            raise DeriveError(f'{name}? is a positional guard; it only exists '
                              'inside a repeating section item')
        return None  # a render directive, never a schema node
    if not conditional and is_block(content):
        raise DeriveError(
            f'group {name!r}: a block-level group is a conditional block — mark '
            f'the guarded control\'s Title with a trailing \'?\'')
    required = _is_locked(pr) or (bool(members) and _is_locked(members[0][1]))
    if conditional:
        if not members:  # the group itself is the guard of static content
            return Field(name=name, remark=remark, required=required)
        msdt, mpr, _, _ = members[0]
        mcontent = msdt.find(qn('w:sdtContent'))
        if (mkind := kind_of(mpr)) == KIND_REPEAT:
            return _parse_repeat(mcontent, name, remark, required)
        if mkind == KIND_SECTION:
            return _parse_section(mcontent, name, remark, required)
        return _parse_field(mpr, name, remark, required, mkind)
    if members:
        field = _parse_field(members[0][1], name, remark, required, kind_of(members[0][1]))
    else:
        field = Field(name=name, remark=remark, required=required)
    if literals := group_literals(content):
        field.alternatives = literals
    return field


def _granularity(date_format) -> str:
    """The picker's own display format names the finest unit the value
    carries; quoted literals in the mask are decoration, not tokens."""
    if not date_format:
        return FORMAT_DATE  # a bare picker collects a full date
    letters = {letter for letter, _ in date_mask_runs(date_format) if letter}
    if 'd' in letters:
        return FORMAT_DATE
    return FORMAT_DATE_MONTH if 'M' in letters else FORMAT_DATE_YEAR


def _is_locked(pr) -> bool:
    lock = pr.find(qn('w:lock'))
    if lock is None:
        return False
    vals = {lock.get(qn('w:val')), lock.get(qn('w:lockType'))}
    return 'sdtLocked' in vals
