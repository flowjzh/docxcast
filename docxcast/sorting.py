"""Deterministic sorts over extracted data, driven by repeat-level
``@sort`` directives (docxcast/schema.SortSpec). The kinds, tail and
stability semantics are the template contract — see SYNTAX.md.
"""

from __future__ import annotations

import re

from docxcast.schema import Repeat, Section, _path

_GROUPS = re.compile(r'\d+')
_SEGMENT_SPLIT = re.compile(r'(\d+)')


def apply_sorts(schema, data: dict) -> dict:
    """Sort the rule-carrying arrays of ``data`` in place — repeats at
    any nesting depth, sections included. Returns the per-path
    permutations the sort applied: ``perm[i]`` is the original index
    of the item now at position ``i`` (dotted paths through sections;
    repeats nested inside another repeat's items sort per parent
    instance and hold no records of their own — excluded). Hosts
    re-align per-item records written in extraction order (budget
    arrangements) by it. Sorts never enter the JSON schema: the model
    is not asked to order anything."""
    perms = {}
    _sort_children(schema.children, data, '', perms)
    return perms


def _sort_children(nodes: list, data, prefix: str, perms: dict) -> None:
    for node in nodes:
        path = _path(prefix, node.name)
        if isinstance(node, Repeat):
            items = data.get(node.name) if isinstance(data, dict) else None
            if node.sort is not None and isinstance(items, list):
                order = _order(items, node.sort)
                data[node.name] = [items[i] for i in order]
                # a derive that emits one repeat twice (a template
                # duplicate) walks the now-sorted array again — the
                # first observation is the pre-sort permutation
                perms.setdefault(path, order)
            for item in items or []:
                if isinstance(item, dict):
                    # nested repeats sort per parent instance and hold
                    # no budget records of their own — discarded
                    _sort_children(node.fields, item, '', {})
        elif isinstance(node, Section):
            scoped = data.get(node.name) if isinstance(data, dict) else None
            if isinstance(scoped, dict):
                _sort_children(node.children, scoped, path, perms)


def _order(items: list, spec) -> list:
    """The sorted permutation — ``order[i]`` is the original index of
    the item now at position i. Each item's value is parsed once;
    unparsable items partition to the tail in their original order,
    both directions."""
    parts = tuple(spec.key.split('.'))
    good, tail = [], []
    for i, item in enumerate(items):
        parsed = _parse(spec.kind, _value(item, parts))
        (tail if parsed is None else good).append((i, parsed))
    good.sort(key=lambda pair: pair[1], reverse=spec.desc)
    return [i for i, _ in good] + [i for i, _ in tail]


def _value(item, parts: tuple):
    for part in parts:
        item = item.get(part) if isinstance(item, dict) else None
        if item is None:
            return None
    return item


def _parse(kind: str, value):
    if value is None:
        return None
    text = value if isinstance(value, str) else str(value)
    if not text.strip():
        return None
    if kind in ('date', 'datetime'):
        return _date(text, time=kind == 'datetime')
    return _segments(text)


def _date(text: str, time: bool = False) -> tuple | None:
    """(year, month, day[, hour, minute, second]) from the leading digit
    groups; out-of-range components drop to 0. None when no plausible
    year leads the value."""
    groups = [int(g) for g in _GROUPS.findall(text)][:6 if time else 3]
    if not groups or not 1000 <= groups[0] <= 9999:
        return None
    y, m, d, h, mi, s = (groups + [0] * 6)[:6]
    if not 1 <= m <= 12:
        m = 0
    if not 1 <= d <= 31:
        d = 0
    if not time:
        return (y, m, d)
    return (y, m, d, h if h <= 23 else 0, mi if mi <= 59 else 0,
            s if s <= 59 else 0)


def _segments(text: str) -> list:
    """Natural-order segments as (rank, value) pairs — digits (rank 0)
    sort before text (rank 1) — so tuple comparison never mixes int
    and str."""
    return [(0, int(p)) if p.isdigit() else (1, p.casefold())
            for p in _SEGMENT_SPLIT.split(text) if p != '']
