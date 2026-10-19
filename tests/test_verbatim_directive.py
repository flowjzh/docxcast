"""The @verbatim directive: stripped off the Tag at derive like @sort
and @override, carried as the schema's ``x-verbatim`` keyword — the
structured channel the extraction pipeline reads instead of scanning
descriptions."""

from io import BytesIO

import pytest

from docxcast import DeriveError, derive_schema
from docxcast._ooxml import SELF
from tests.helpers import (build_doc, group_pr, make_sdt, para,
                           repeating_item_pr, repeating_pr, run)


def _marked_template(tag: str) -> bytes:
    """A '$'-addressed string repeat — the verbatim contract's shape."""
    item = make_sdt(pr_extra=repeating_item_pr(), content=
                    para(make_sdt(SELF, '一条描述文字', content=run('x'))))
    return build_doc(make_sdt('description', tag,
                              pr_extra=repeating_pr(), content=item))


def _marked_schema(tag: str):
    return derive_schema(BytesIO(_marked_template(tag)))


def test_bare_marker_strips_the_remark_and_defaults_the_ratio():
    repeat = _marked_schema('工作职责：逐条照抄 @verbatim').children[0]
    assert repeat.remark == '工作职责：逐条照抄'
    assert repeat.verbatim == 0.9
    node = repeat.to_json_schema()
    assert '@verbatim' not in node['description']
    assert node['x-verbatim'] == 0.9


def test_valued_and_percent_markers_normalize_to_a_fraction():
    assert _marked_schema('h @verbatim=0.85').children[0].verbatim == 0.85
    assert _marked_schema('h @verbatim=85').children[0].verbatim == 0.85
    assert _marked_schema('h @verbatim=100').children[0].verbatim == 1.0


@pytest.mark.parametrize('tag', [
    'h @verbatim=abc',   # garbage value
    'h @verbatim=',      # empty value
    'h @verbatim=0',     # out of range
    '提及 @verbatim 的说明文字',  # mid-prose: unparseable directive
])
def test_mangled_directives_fail_at_derive(tag):
    with pytest.raises(DeriveError, match='verbatim'):
        _marked_schema(tag)


def test_field_tag_carries_the_directive_too():
    template = build_doc(make_sdt('duties', '逐条照抄 @verbatim=0.8',
                                  content=run('x')))
    field = derive_schema(BytesIO(template)).children[0]
    assert field.remark == '逐条照抄'
    assert field.verbatim == 0.8
    assert field.to_json_schema()['x-verbatim'] == 0.8


def test_group_wrapped_repeat_keeps_its_own_directive():
    """The bullet-block shape: a conditional group wraps the repeat, and
    the address (with the directive) comes from the repeat's own tag —
    the group layer must hand it to the repeat's parser unstripped."""
    item = make_sdt(pr_extra=repeating_item_pr(), content=
                    para(make_sdt(SELF, '一条描述文字', content=run('x'))))
    repeat = make_sdt('description?', '逐条照抄 @verbatim',
                      pr_extra=repeating_pr(), content=item)
    group = make_sdt(None, None, pr_extra=group_pr(), content=[repeat])
    node = derive_schema(BytesIO(build_doc(group))).children[0]
    assert node.verbatim == 0.9
    assert node.remark == '逐条照抄'
    assert node.to_json_schema()['x-verbatim'] == 0.9


def test_unmarked_tags_carry_no_key():
    repeat = _marked_schema('普通描述').children[0]
    assert repeat.verbatim is None
    assert 'x-verbatim' not in repeat.to_json_schema()
