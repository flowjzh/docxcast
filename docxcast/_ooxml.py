"""OOXML namespace constants and shared sdt helpers used by both derive and render."""

from functools import lru_cache
from typing import Optional

NS = {
    'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
    'w14': 'http://schemas.microsoft.com/office/word/2010/wordml',
    'w15': 'http://schemas.microsoft.com/office/word/2012/wordml',
}

XML_SPACE = '{http://www.w3.org/XML/1998/namespace}space'


@lru_cache(maxsize=None)
def qn(tag: str) -> str:
    prefix, local = tag.split(':')
    return f'{{{NS[prefix]}}}{local}'


# tags the sdt walker descends into (paragraphs and table cells)
CONTAINER_TAGS = frozenset(qn(t) for t in ('w:p', 'w:tbl', 'w:tr', 'w:tc'))

KIND_REPEAT = 'repeat'
KIND_ITEM = 'item'
KIND_SECTION = 'section'
KIND_GROUP = 'group'
KIND_PICTURE = 'picture'
KIND_DROPDOWN = 'dropdown'
KIND_DATE = 'date'
KIND_CHECKBOX = 'checkbox'
KIND_TEXT = 'text'

_TAG_KINDS = {
    qn('w15:repeatingSection'): KIND_REPEAT,
    qn('w15:repeatingSectionItem'): KIND_ITEM,
    qn('w:docPartList'): KIND_SECTION,
    qn('w:group'): KIND_GROUP,
    qn('w:picture'): KIND_PICTURE,
    qn('w:dropDownList'): KIND_DROPDOWN,
    qn('w:date'): KIND_DATE,
    qn('w14:checkbox'): KIND_CHECKBOX,
}


def kind_of(pr) -> str:
    for child in pr:
        kind = _TAG_KINDS.get(child.tag)
        if kind is not None:
            return kind
    return KIND_TEXT


def repeating_item(content):
    for child in content:
        if child.tag != qn('w:sdt'):
            continue
        pr = child.find(qn('w:sdtPr'))
        if pr is not None and pr.find(qn('w15:repeatingSectionItem')) is not None:
            return child
    return None


def dropdown_items(pr) -> list:
    """(display, value) pairs of a dropdown, skipping placeholder items without display text."""
    dropdown = pr.find(qn('w:dropDownList'))
    if dropdown is None:
        return []
    items = []
    for item in dropdown.findall(qn('w:listItem')):
        display = item.get(qn('w:displayText'))
        if display is not None:
            items.append((display, item.get(qn('w:value'))))
    return items


def text_attr(pr, tag: str) -> Optional[str]:
    el = pr.find(qn(tag))
    return el.get(qn('w:val')) if el is not None else None


def date_mask_runs(mask: str) -> list:
    """Tokenize a Word dateFormat mask into (letter, text) runs: quoted
    spans and non-ASCII-letter text are literals (letter None), an
    unquoted ASCII letter run is a token carrying its letter. Shared by
    derive (granularity) and render (formatting) so both read one
    grammar and can't drift apart."""
    runs = []
    i = 0
    while i < len(mask):
        c = mask[i]
        if c == "'":  # a quoted literal
            end = mask.find("'", i + 1)
            if end == -1:
                runs.append((None, mask[i + 1:]))
                break
            runs.append((None, mask[i + 1:end]))
            i = end + 1
            continue
        j = i
        while j < len(mask) and mask[j] == c:
            j += 1
        runs.append((c if c.isascii() and c.isalpha() else None, mask[i:j]))
        i = j
    return runs


# guard names injected into every repeating-section item's scope; a real
# field may not take them
RESERVED = frozenset(('has_prev', 'has_next'))

# the self address: a repeating item's sole addressed control named '$'
# fills from the entry itself, making the repeat an array of that value
SELF = '$'


def split_conditional(alias: Optional[str]) -> tuple[Optional[str], bool]:
    """A trailing '?' on an address (TS-style ``name?``) marks its block
    conditionally displayed: a falsy value removes the block."""
    if alias and alias.endswith('?'):
        return alias[:-1], True
    return alias, False


def sdt_alias(sdt) -> Optional[str]:
    """The control's unmarked address — the alias with any conditional
    marker stripped, None when the sdt carries no sdtPr or alias."""
    pr = sdt.find(qn('w:sdtPr'))
    return split_conditional(text_attr(pr, 'w:alias'))[0] if pr is not None else None


def scope_guards(i: int, n: int) -> dict:
    """Positional guard values for item ``i`` of ``n`` in a repeat."""
    return {'has_prev': i > 0, 'has_next': i < n - 1}


def is_block(content) -> bool:
    """Group content is block-level when the outside-sdt walk meets a
    paragraph — the same walk group_literals chunks, so block and
    inline cannot be classified two ways."""
    return any(el.tag == qn('w:p') for el in outside_sdts(content))


def outside_sdts(container):
    """Iterate elements of ``container`` without descending into nested
    content controls — their content belongs to the control, not the
    group around it."""
    for child in container:
        if child.tag == qn('w:sdt'):
            continue
        yield child
        yield from outside_sdts(child)


def group_members(content) -> list:
    """(sdt, pr, alias, tag) of a group's control members, in document
    order."""
    members = []
    for child in content:
        if child.tag == qn('w:sdt'):
            if (pr := child.find(qn('w:sdtPr'))) is not None:
                members.append((child, pr, text_attr(pr, 'w:alias'), text_attr(pr, 'w:tag')))
        elif child.tag in CONTAINER_TAGS:
            members.extend(group_members(child))
    return members


def group_address(pr, members) -> tuple:
    """A group's (alias, tag, member index, conditional): the group's own
    address if it carries one, else the first member that does (-1 if
    none). The alias comes back unmarked — split_conditional's trailing
    '?' marks the group a conditional block. Shared by derive and
    render so the two can't drift apart."""
    if text_attr(pr, 'w:alias') or text_attr(pr, 'w:tag'):
        name, conditional = split_conditional(text_attr(pr, 'w:alias'))
        return name, text_attr(pr, 'w:tag'), -1, conditional
    for i, (_, mpr, alias, tag) in enumerate(members):
        if alias or tag:
            name, conditional = split_conditional(alias)
            return name, tag, i, conditional
    return None, None, -1, False


def outside_text(el) -> str:
    return ''.join(t.text or '' for t in outside_sdts(el) if t.tag == qn('w:t')).strip()


def group_chunks(content) -> list:
    """A group's literal text as (holder, text) pairs: one per
    paragraph outside nested controls, or the joined run text of an
    inline group. Derive advertises the texts, render deletes from the
    holders — one chunking rule for both."""
    els = list(outside_sdts(content))
    if paras := [p for p in els if p.tag == qn('w:p')]:
        return [(p, outside_text(p)) for p in paras]
    return [(content, outside_text(content))]


def group_literals(content) -> list:
    """A group's literal value alternatives: the text it shows when no
    control applies."""
    return [text for _, text in group_chunks(content) if text]
