"""Boundary tests on in-memory constructed templates."""

from io import BytesIO

import pytest

from docx import Document
from lxml import etree

from docxcast import DeriveError, Field, Repeat, Schema, Section, derive_schema, render
from docxcast._ooxml import qn
from tests.helpers import (
    build_cell_doc,
    build_doc,
    checkbox_pr,
    count_sdts,
    date_pr,
    docpartlist_pr,
    dropdown_pr,
    full_text,
    group_pr,
    lock_pr,
    make_sdt,
    para,
    picture_pr,
    reopen,
    repeating_item_pr,
    repeating_pr,
    run,
    save_bytes,
)


def paragraph_texts(doc: Document) -> list:
    return [''.join(t.text or '' for t in p.iter(qn('w:t')))
            for p in doc.element.body.findall(qn('w:p'))]


def test_repeat_clones_items_and_unwraps():
    item = make_sdt(pr_extra=repeating_item_pr(), content=para(
        make_sdt('a', 'field a', content=run('x')),
        run(' '),
        make_sdt('b', 'field b', content=run('y')),
    ))
    template = build_doc(make_sdt('r', 'repeat', pr_extra=repeating_pr(), content=item))
    schema = derive_schema(BytesIO(template))
    assert schema.children == [Repeat(name='r', remark='repeat', fields=[
        Field(name='a', remark='field a'),
        Field(name='b', remark='field b'),
    ])]

    result = render(BytesIO(template), {'r': [{'a': '1', 'b': '2'}, {'a': '3', 'b': '4'}]})
    out = reopen(result)
    assert count_sdts(out) == 0
    assert paragraph_texts(out) == ['1 2', '3 4']


def test_dropdown_skips_placeholder_item():
    dd = dropdown_pr([('Employed', 'employed'), ('Unemployed', 'unemployed')])
    placeholder = etree.SubElement(dd, qn('w:listItem'))
    placeholder.set(qn('w:value'), 'Choose an item.')
    template = build_doc(make_sdt('status', 'status', pr_extra=dd, content=run('Choose an item.')))
    field = derive_schema(BytesIO(template)).children[0]
    assert field.enum == ['employed', 'unemployed']
    assert 'Employed=employed' in field.to_json_schema()['description']


def _dropdown_doc(items, *, locked=False):
    """A doc carrying one 'status' dropdown; locked marks it required."""
    pr_extra = [lock_pr(), dropdown_pr(items)] if locked else dropdown_pr(items)
    return build_doc(make_sdt('status', 'status', pr_extra=pr_extra,
                              content=run('Choose an item.')))


def test_dropdown_renders_display_text():
    template = _dropdown_doc([('Employed', 'employed')])
    result = render(BytesIO(template), {'status': 'employed'})
    assert 'Employed' in full_text(reopen(result))


def test_optional_dropdown_accepts_empty():
    schema = derive_schema(BytesIO(_dropdown_doc(
        [('Employed', 'employed'), ('Unemployed', 'unemployed')])))
    assert schema.validate({'status': ''}).ok
    assert schema.validate({'status': 'employed'}).ok
    assert schema.validate({'status': 'freelancing'}).errors[0].code == 'enum_invalid'
    assert schema.children[0].to_json_schema()['anyOf'][1] == {'enum': ['']}


def test_required_dropdown_rejects_empty():
    schema = derive_schema(BytesIO(_dropdown_doc([('Employed', 'employed')], locked=True)))
    assert schema.validate({'status': ''}).errors[0].code == 'enum_invalid'
    assert schema.validate({}).errors[0].code == 'missing_required'
    assert 'anyOf' not in schema.children[0].to_json_schema()


def test_date_granularity_and_datetime_subtype():
    template = build_doc(
        make_sdt('d', 'date', pr_extra=date_pr(), content=run('')),
        make_sdt('m', 'month', pr_extra=date_pr(date_format="yyyy-MM"), content=run('')),
        make_sdt('y', 'year', pr_extra=date_pr(date_format='yyyy'), content=run('')),
        make_sdt('dt', 'time', pr_extra=date_pr(store='dateTime'), content=run('')),
    )
    schema = derive_schema(BytesIO(template))
    assert [f.format for f in schema.children] == \
        ['date', 'date-month', 'date-year', 'date-time']


def test_date_control_fills_with_its_display_mask():
    template = build_doc(
        make_sdt('a', 'dot month', pr_extra=date_pr(date_format='yyyy.M'), content=run('')),
        make_sdt('b', 'month name', pr_extra=date_pr(date_format='MMM yyyy'), content=run('')),
        make_sdt('c', 'slash month', pr_extra=date_pr(date_format='yyyy/M'), content=run('')),
        make_sdt('d', 'padded', pr_extra=date_pr(date_format='yyyy-MM'), content=run('')),
        make_sdt('e', 'full month name', pr_extra=date_pr(date_format='yyyy年MMMM'), content=run('')),
    )
    filled = reopen(render(BytesIO(template), {
        'a': '2025-01', 'b': '2024-12', 'c': '2013-05',
        'd': '2014-07', 'e': '1999-07'}))
    text = full_text(filled)
    assert '2025.1' in text          # M drops the pad
    assert 'Dec 2024' in text       # MMM abbreviates
    assert '2013/5' in text         # the separator passes through
    assert '2014-07' in text        # MM keeps the pad
    assert '1999年July' in text     # quoted-free literals and MMMM


def test_date_without_mask_fills_raw_value():
    template = build_doc(make_sdt('d', 'date', pr_extra=date_pr(), content=run('')))
    filled = reopen(render(BytesIO(template), {'d': '2025-01-05'}))
    assert '2025-01-05' in full_text(filled)


def test_date_mask_fill_keeps_placeholder_formatting():
    # the control carries an empty formatted run; the filled date must keep
    # that run's rPr, not drop to the paragraph default
    holder = etree.Element(qn('w:r'))
    etree.SubElement(etree.SubElement(holder, qn('w:rPr')), qn('w:b'))
    etree.SubElement(holder, qn('w:t')).text = ''
    control = make_sdt('end_date', 'job end date',
                       pr_extra=date_pr(date_format='yyyy.M'), content=holder)
    group = make_sdt(pr_extra=group_pr(), content=[control, run('Present')])
    filled = reopen(render(BytesIO(build_doc(group)), {'end_date': '2025-01'}))
    r = filled.element.body.find(f'.//{qn("w:r")}')
    assert ''.join(t.text or '' for t in r.findall(qn('w:t'))) == '2025.1'
    assert 'Present' not in full_text(filled)
    assert r.find(qn('w:rPr')).find(qn('w:b')) is not None


def _end_date_group(alias_on='child'):
    child, grp = ((('end_date', 'job end date'), (None, None)) if alias_on == 'child'
                  else ((None, None), ('end_date', 'job end date')))
    control = make_sdt(*child, pr_extra=date_pr(date_format='yyyy-MM'), content=run('2025-01'))
    return make_sdt(*grp, pr_extra=group_pr(), content=[control, run('Present')])


def test_group_derives_union_from_first_member():
    schema = derive_schema(BytesIO(build_doc(_end_date_group())))
    assert schema.children == [Field(name='end_date', remark='job end date',
                                     format='date-month', alternatives=['Present'])]
    node = schema.children[0].to_json_schema()
    assert 'type' not in node  # the union carries no type of its own
    assert node['anyOf'] == [
        {'type': 'string', 'format': 'date-month', 'pattern': r'^\d{4}-\d{2}$'},
        {'enum': ['Present']},
    ]


def test_group_address_prefers_the_group_itself():
    schema = derive_schema(BytesIO(build_doc(_end_date_group(alias_on='group'))))
    assert schema.children == [Field(name='end_date', remark='job end date',
                                     format='date-month', alternatives=['Present'])]


def test_group_with_two_addressed_members_fails_derive():
    control = make_sdt('end_date', 'job end date',
                       pr_extra=date_pr(date_format='yyyy-MM'), content=run(''))
    second = make_sdt('end_note', 'note', content=run(''))
    template = build_doc(make_sdt(pr_extra=group_pr(), content=[control, second, run('Present')]))
    with pytest.raises(DeriveError, match='end_note'):
        derive_schema(BytesIO(template))


def test_group_with_two_control_members_fails_derive():
    control = make_sdt('end_date', 'job end date',
                       pr_extra=date_pr(date_format='yyyy-MM'), content=run(''))
    anonymous = make_sdt(content=run(''))
    template = build_doc(make_sdt(pr_extra=group_pr(), content=[control, anonymous, run('Present')]))
    with pytest.raises(DeriveError, match='2 control members'):
        derive_schema(BytesIO(template))


def test_group_union_validates_both_branches():
    schema = derive_schema(BytesIO(build_doc(_end_date_group())))
    assert schema.validate({'end_date': '2025-01'}).ok
    assert schema.validate({'end_date': 'Present'}).ok
    assert schema.validate({'end_date': '2025/01'}).errors[0].code == 'format_invalid'
    assert schema.validate({'end_date': 3}).errors[0].code == 'type_mismatch'


def test_group_renders_either_branch():
    template = build_doc(_end_date_group())
    as_date = reopen(render(BytesIO(template), {'end_date': '2025-01'}))
    assert '2025-01' in full_text(as_date) and 'Present' not in full_text(as_date)
    as_literal = reopen(render(BytesIO(template), {'end_date': 'Present'}))
    assert 'Present' in full_text(as_literal) and '2025-01' not in full_text(as_literal)
    assert count_sdts(as_date) == count_sdts(as_literal) == 0


def test_group_union_validates_literal_before_type_dispatch():
    # a checkbox member plus a literal: the emitted anyOf accepts the
    # literal, so the internal validator must too (not type_mismatch)
    control = make_sdt('marital', 'marital status', pr_extra=checkbox_pr(), content=run('☐'))
    template = build_doc(make_sdt(pr_extra=group_pr(), content=[control, run('Single')]))
    schema = derive_schema(BytesIO(template))
    assert schema.children[0].type == 'boolean' \
        and schema.children[0].alternatives == ['Single']
    assert schema.validate({'marital': True}).ok
    assert schema.validate({'marital': 'Single'}).ok
    assert schema.validate({'marital': 3}).errors[0].code == 'type_mismatch'


def _nationality_line():
    """Nationality:  [nationality?] wrapped in a block-level group."""
    control = make_sdt('nationality?', "the candidate's nationality", content=run(''))
    return make_sdt(pr_extra=group_pr(), content=para(run('Nationality: '), control))


def test_conditional_block_derives_plain_field():
    schema = derive_schema(BytesIO(build_doc(_nationality_line())))
    assert schema.children == [Field(name='nationality', remark="the candidate's nationality")]


def test_conditional_block_truthy_renders_label_and_value():
    out = reopen(render(BytesIO(build_doc(_nationality_line())), {'nationality': 'Chinese'}))
    assert count_sdts(out) == 0
    assert 'Nationality: Chinese' in full_text(out)


def test_conditional_block_falsy_drops_the_whole_paragraph():
    template = build_doc(_nationality_line())
    for data in ({}, {'nationality': None}, {'nationality': ''}):
        out = reopen(render(BytesIO(template), data))
        assert 'nationality' not in full_text(out)
        assert paragraph_texts(out) == []


def test_conditional_block_falsy_keeps_structure_with_keep_controls():
    result = render(BytesIO(build_doc(_nationality_line())), {}, keep_controls=True)
    assert count_sdts(reopen(result)) == 2  # the group and its member survive


def test_conditional_inline_span_drops_only_the_span():
    control = make_sdt('email?', 'email', content=run(''))
    span = make_sdt(pr_extra=group_pr(), content=[run('Email: '), control])
    template = build_doc(para(run('Alex '), span, run(' end')))
    out = reopen(render(BytesIO(template), {}))
    assert 'email' not in full_text(out)
    assert paragraph_texts(out) == ['Alex  end']
    out = reopen(render(BytesIO(template), {'email': 'a@b.c'}))
    assert paragraph_texts(out) == ['Alex Email: a@b.c end']


def _experiences_block():
    item = make_sdt(pr_extra=repeating_item_pr(),
                    content=para(make_sdt('company', 'company', content=run('x'))))
    repeat = make_sdt('work_experiences?', 'work history', pr_extra=repeating_pr(), content=item)
    return make_sdt(pr_extra=group_pr(), content=[para(run('Work Experiences')), repeat])


def test_conditional_block_guards_a_repeat_with_its_heading():
    template = build_doc(_experiences_block())
    schema = derive_schema(BytesIO(template))
    assert schema.children == [Repeat(name='work_experiences', remark='work history',
                                      fields=[Field(name='company', remark='company')])]
    for data in ({}, {'work_experiences': []}):
        out = reopen(render(BytesIO(template), data))
        assert 'Work Experiences' not in full_text(out)
        assert paragraph_texts(out) == []
    out = reopen(render(BytesIO(template),
                        {'work_experiences': [{'company': 'A'}, {'company': 'B'}]}))
    assert 'Work Experiences' in full_text(out)
    assert paragraph_texts(out) == ['Work Experiences', 'A', 'B']


def test_conditional_block_guards_a_section():
    section = make_sdt('education?', 'education', pr_extra=docpartlist_pr(),
                       content=para(make_sdt('school', 'school', content=run(''))))
    template = build_doc(make_sdt(pr_extra=group_pr(),
                                  content=[para(run('Education')), section]))
    assert derive_schema(BytesIO(template)).children == [
        Section(name='education', remark='education', children=[Field(name='school', remark='school')])]
    assert 'Education' not in full_text(reopen(render(BytesIO(template), {})))
    out = reopen(render(BytesIO(template), {'education': {'school': 'MIT'}}))
    assert 'Education' in full_text(out) and 'MIT' in full_text(out)


def test_conditional_marker_on_the_group_guards_static_text():
    template = build_doc(make_sdt('nationality?', 'foreign', pr_extra=group_pr(),
                                  content=para(run('Foreign candidate'))))
    assert derive_schema(BytesIO(template)).children == [Field(name='nationality', remark='foreign')]
    assert 'Foreign candidate' not in full_text(reopen(render(BytesIO(template), {})))
    assert 'Foreign candidate' in full_text(reopen(render(BytesIO(template),
                                                  {'nationality': 'yes'})))


def test_conditional_inside_repeat_item_reads_item_data():
    control = make_sdt('nationality?', None, content=run(''))
    line = make_sdt(pr_extra=group_pr(), content=para(run('Nationality: '), control))
    item = make_sdt(pr_extra=repeating_item_pr(),
                    content=[para(make_sdt('name', None, content=run(''))), line])
    template = build_doc(make_sdt('people', None, pr_extra=repeating_pr(), content=item))
    data = {'people': [{'name': 'A', 'nationality': 'Chinese'}, {'name': 'B'}]}
    assert paragraph_texts(reopen(render(BytesIO(template), data))) \
        == ['A', 'Nationality: Chinese', 'B']


@pytest.mark.parametrize('guard', ('has_next?', 'has_prev?'))
def test_separator_shows_between_items_only(guard):
    sep = make_sdt(guard, None, pr_extra=group_pr(), content=para(run(',')))
    line = para(make_sdt('skill', 'skill', content=run('x')))
    children = [line, sep] if guard == 'has_next?' else [sep, line]
    item = make_sdt(pr_extra=repeating_item_pr(), content=children)
    template = build_doc(make_sdt('skills', 'skills', pr_extra=repeating_pr(), content=item))
    assert derive_schema(BytesIO(template)).children == [
        Repeat(name='skills', remark='skills', fields=[Field(name='skill', remark='skill')])]
    single = reopen(render(BytesIO(template), {'skills': [{'skill': 'Python'}]}))
    assert ',' not in full_text(single)
    pair = reopen(render(BytesIO(template),
                         {'skills': [{'skill': 'Python'}, {'skill': 'Go'}]}))
    assert full_text(pair).count(',') == 1


def test_block_group_without_marker_fails_derive():
    control = make_sdt('nationality', 'nationality', content=run(''))
    template = build_doc(make_sdt(pr_extra=group_pr(), content=para(run('Nationality: '), control)))
    with pytest.raises(DeriveError, match='trailing'):
        derive_schema(BytesIO(template))


def test_question_outside_a_group_fails_derive():
    template = build_doc(make_sdt('nationality?', 'nationality', content=run('')))
    with pytest.raises(DeriveError, match='group'):
        derive_schema(BytesIO(template))


def test_reserved_field_name_fails_derive():
    template = build_doc(make_sdt('has_next', 'separator', content=run('')))
    with pytest.raises(DeriveError, match='reserved'):
        derive_schema(BytesIO(template))


def test_positional_guard_outside_a_repeat_fails_derive():
    template = build_doc(make_sdt('has_next?', None, pr_extra=group_pr(),
                                  content=para(run(','))))
    with pytest.raises(DeriveError, match='repeating'):
        derive_schema(BytesIO(template))


def test_conditional_block_in_a_cell_leaves_a_paragraph():
    control = make_sdt('note?', 'note', content=run(''))
    block = make_sdt(pr_extra=group_pr(), content=para(run('Note: '), control))
    template = build_cell_doc(block, replace_paragraph=True)
    out = reopen(render(BytesIO(template), {}))
    tc = out.element.body.find(qn('w:tbl')).find(qn('w:tr')).find(qn('w:tc'))
    assert len(tc.findall(qn('w:p'))) == 1  # the cell stays a valid cell
    assert 'note' not in full_text(out)


def test_format_fields_carry_a_pattern():
    template = build_doc(make_sdt('y', 'year', pr_extra=date_pr(date_format='yyyy'), content=run('')))
    node = derive_schema(BytesIO(template)).children[0].to_json_schema()
    assert node['format'] == 'date-year' and node['pattern'] == r'^\d{4}$'


def test_checkbox_renders_both_states():
    template = build_doc(make_sdt('flag', 'flag', pr_extra=checkbox_pr(), content=run('☐')))
    for value, glyph in ((True, '☒'), (False, '☐')):
        result = render(BytesIO(template), {'flag': value})
        assert glyph in full_text(reopen(result))
    # keep_controls keeps the checkbox state
    result = render(BytesIO(template), {'flag': True}, keep_controls=True)
    out = reopen(result)
    checked_el = next(out.element.body.iter(qn('w14:checked')))
    assert checked_el.get(qn('w14:val')) == '1'


def test_lock_variants_mark_required():
    cases = [
        (lock_pr(), True),
        (lock_pr('sdtContentLocked', 'sdtLocked'), True),
        (lock_pr('sdtContentLocked'), False),
        (None, False),
    ]
    for extra, expected in cases:
        template = build_doc(make_sdt('f', 'field', pr_extra=extra, content=run('v')))
        field = derive_schema(BytesIO(template)).children[0]
        assert field.required is expected


def test_unnamed_control_is_unwrapped_with_text_preserved():
    template = build_doc(make_sdt(None, None, content=run('original')))
    assert derive_schema(BytesIO(template)).children == []
    result = render(BytesIO(template), {})
    out = reopen(result)
    assert count_sdts(out) == 0
    assert 'original' in full_text(out)


def test_picture_is_derived_and_removed():
    template = build_doc(make_sdt('photo', 'photo path', pr_extra=picture_pr(), content=run('')))
    field = derive_schema(BytesIO(template)).children[0]
    assert field.name == 'photo' and field.type == 'string'
    result = render(BytesIO(template), {'photo': '/tmp/x.png'})
    assert count_sdts(reopen(result)) == 0


def test_nested_sections():
    inner = make_sdt('inner', 'inner rule', pr_extra=docpartlist_pr(),
                     content=para(make_sdt('name', 'display name', content=run('old name'))))
    outer = make_sdt('outer', 'outer rule', pr_extra=docpartlist_pr(), content=inner)
    template = build_doc(outer)
    schema = derive_schema(BytesIO(template))
    assert schema.children == [Section(name='outer', remark='outer rule', children=[
        Section(name='inner', remark='inner rule', children=[
            Field(name='name', remark='display name'),
        ]),
    ])]
    result = render(BytesIO(template), {'outer': {'inner': {'name': 'new name'}}})
    assert 'new name' in full_text(reopen(result))


def test_control_inside_table_cell():
    template = build_cell_doc(make_sdt('age', 'age', content=run('30')))

    schema = derive_schema(BytesIO(template))
    assert schema.children[0].name == 'age'

    result = render(BytesIO(template), {'age': '31'})
    out = reopen(result)
    text = full_text(out)
    assert '31' in text and '30' not in text
    assert count_sdts(out) == 0


def test_missing_repeat_renders_zero_items():
    item = make_sdt(pr_extra=repeating_item_pr(), content=para(make_sdt('a', 'a', content=run('x'))))
    template = build_doc(make_sdt('r', 'repeat', pr_extra=repeating_pr(), content=item))
    result = render(BytesIO(template), {})
    assert 'x' not in full_text(reopen(result))


def test_scalar_repeat_derives_a_string_array():
    item = make_sdt(pr_extra=repeating_item_pr(),
                    content=para(make_sdt('$', 'a single responsibility item', content=run('x'))))
    template = build_doc(make_sdt('responsibilities', 'resp', pr_extra=repeating_pr(), content=item))
    rep = derive_schema(BytesIO(template)).children[0]
    assert rep.item == Field(name='$', remark='a single responsibility item', required=True)
    node = rep.to_json_schema()
    assert node['type'] == 'array'
    assert node['items'] == {'type': 'string', 'description': 'a single responsibility item'}


def test_scalar_repeat_renders_one_entry_per_item():
    item = make_sdt(pr_extra=repeating_item_pr(),
                    content=para(make_sdt('$', None, content=run('Details'))))
    template = build_doc(make_sdt('skills', None, pr_extra=repeating_pr(), content=item))
    out = reopen(render(BytesIO(template), {'skills': ['Python', 'Go', 'Rust']}))
    assert paragraph_texts(out) == ['Python', 'Go', 'Rust']


def test_scalar_repeat_validates_string_entries_only():
    item = make_sdt(pr_extra=repeating_item_pr(),
                    content=para(make_sdt('$', None, content=run('x'))))
    template = build_doc(make_sdt('skills', None, pr_extra=repeating_pr(), content=item))
    schema = derive_schema(BytesIO(template))
    assert schema.validate({'skills': ['a', '']}).ok
    assert schema.validate({'skills': [{'$': 'a'}]}).errors[0].code == 'type_mismatch'


def test_scalar_repeat_keeps_positional_guards():
    sep = make_sdt('has_next?', None, pr_extra=group_pr(), content=para(run(',')))
    line = para(make_sdt('$', None, content=run('x')))
    item = make_sdt(pr_extra=repeating_item_pr(), content=[line, sep])
    template = build_doc(make_sdt('skills', None, pr_extra=repeating_pr(), content=item))
    schema = derive_schema(BytesIO(template))
    assert schema.children[0].item is not None  # the guard is no second address
    out = reopen(render(BytesIO(template), {'skills': ['Python', 'Go', 'Rust']}))
    assert full_text(out) == 'Python,Go,Rust'


def test_self_address_outside_a_repeating_item_fails_derive():
    template = build_doc(make_sdt('$', 'value', content=run('x')))
    with pytest.raises(DeriveError, match='repeating item'):
        derive_schema(BytesIO(template))


def test_self_address_beside_other_fields_fails_derive():
    item = make_sdt(pr_extra=repeating_item_pr(), content=para(
        make_sdt('$', None, content=run('x')), run(' '), make_sdt('note', None, content=run('y'))))
    template = build_doc(make_sdt('r', None, pr_extra=repeating_pr(), content=item))
    with pytest.raises(DeriveError, match='no other addressed control'):
        derive_schema(BytesIO(template))


def test_scalar_repeat_nested_in_an_object_repeat_fills_both_levels():
    bullets = make_sdt('lines', None, pr_extra=repeating_pr(), content=make_sdt(
        pr_extra=repeating_item_pr(), content=para(make_sdt('$', None, content=run('·')))))
    item = make_sdt(pr_extra=repeating_item_pr(),
                    content=[para(make_sdt('name', None, content=run(''))), bullets])
    template = build_doc(make_sdt('people', None, pr_extra=repeating_pr(), content=item))
    data = {'people': [{'name': 'A', 'lines': ['x', 'y']}, {'name': 'B'}]}
    assert paragraph_texts(reopen(render(BytesIO(template), data))) \
        == ['A', 'x', 'y', 'B']


def test_schema_to_dict_roundtrip():
    template = build_doc(make_sdt('a', 'field a', content=run('1')))
    d = derive_schema(BytesIO(template)).to_dict()
    assert d['children'][0]['name'] == 'a'
    assert Schema(**d)  # dict is constructor-compatible


@pytest.mark.parametrize('comments, expected', [
    (None, None),  # untouched: the default template's sentinel
    ('', None),
    ('Chinese resumes only', 'Chinese resumes only'),
])
def test_comments_core_property_becomes_schema_description(comments, expected):
    doc = Document()
    if comments is not None:
        doc.core_properties.comments = comments
    assert derive_schema(BytesIO(save_bytes(doc))).to_json_schema() \
        .get('description') == expected


def test_checkbox_renders_fallback_glyphs_without_states():
    cb = etree.Element(qn('w14:checkbox'))
    checked = etree.SubElement(cb, qn('w14:checked'))
    checked.set(qn('w14:val'), '0')
    template = build_doc(make_sdt('flag', 'flag', pr_extra=cb, content=run('☐')))
    assert '☒' in full_text(reopen(render(BytesIO(template), {'flag': True})))
    assert '☐' in full_text(reopen(render(BytesIO(template), {'flag': False})))


def test_render_preserves_exact_run_formatting():
    r = etree.Element(qn('w:r'))
    r_pr = etree.SubElement(r, qn('w:rPr'))
    etree.SubElement(r_pr, qn('w:b'))
    t = etree.SubElement(r, qn('w:t'))
    t.text = 'Old'
    template = build_doc(make_sdt('f', 'field', content=r))

    result = render(BytesIO(template), {'f': 'New'})
    out = reopen(result)
    runs = [el for el in out.element.body.iter(qn('w:r')) if el.find(qn('w:t')) is not None]
    assert len(runs) == 1
    assert runs[0].find(qn('w:rPr')).find(qn('w:b')) is not None
    assert runs[0].find(qn('w:t')).text == 'New'


# --- @override: variant layouts where data picks the rendering ---

def _variant_item(whole_tag='single position @override=sub_experiences'):
    # both variants ride in anonymous groups, the drop units — the
    # whole layout as gen_templates.wrap_section builds it
    whole = make_sdt('whole_experience?', whole_tag, pr_extra=docpartlist_pr(),
                     content=para(make_sdt('title', 'position title', content=run(''))))
    whole = make_sdt(pr_extra=group_pr(), content=[para(run('Whole: ')), whole])
    sub_repeat = make_sdt(
        'sub_experiences?', 'dated sub-positions', pr_extra=repeating_pr(),
        content=make_sdt(pr_extra=repeating_item_pr(),
                         content=para(make_sdt('title', 'position title', content=run('')))))
    sub = make_sdt(pr_extra=group_pr(), content=[para(run('')), sub_repeat])
    item = make_sdt(pr_extra=repeating_item_pr(), content=[whole, sub])
    return build_doc(make_sdt('work', 'jobs', pr_extra=repeating_pr(), content=item))


def _variant_data(whole=None, subs=None):
    entry = {}
    if whole is not None:
        entry['whole_experience'] = {'title': whole}
    if subs is not None:
        entry['sub_experiences'] = [{'title': t} for t in subs]
    return {'work': [entry]}


def test_override_directive_is_stripped_from_the_schema():
    schema = derive_schema(BytesIO(_variant_item()))
    [work] = schema.children
    whole, sub = work.fields
    assert whole.remark == 'single position' and whole.overrides == 'sub_experiences'
    assert sub.overrides is None
    assert '@override' not in str(schema.to_json_schema())


def test_override_drops_the_named_sibling_when_truthy():
    template = _variant_item()
    out = reopen(render(BytesIO(template), _variant_data(whole='Senior X', subs=['a', 'b'])))
    assert 'Senior X' in full_text(out)
    assert 'a' not in full_text(out) and 'Whole: ' in full_text(out)
    assert paragraph_texts(out) == ['Whole: ', 'Senior X']  # the blank rides the group out


def test_override_keeps_the_sibling_when_the_winner_is_falsy():
    template = _variant_item()
    out = reopen(render(BytesIO(template), _variant_data(subs=['a', 'b'])))
    # the sub group's own leading blank renders with it
    assert paragraph_texts(out) == ['', 'a', 'b']


def test_override_without_a_target_value_changes_nothing():
    # sub_experiences falsy: its own ? gate already dropped it
    template = _variant_item()
    out = reopen(render(BytesIO(template), _variant_data(whole='Senior X')))
    assert paragraph_texts(out) == ['Whole: ', 'Senior X']


def test_unparseable_override_directive_fails_derive():
    with pytest.raises(DeriveError, match='unparseable @override'):
        derive_schema(BytesIO(_variant_item(whole_tag='single position @override=')))


def test_override_target_not_a_sibling_fails_derive():
    with pytest.raises(DeriveError, match='not a sibling'):
        derive_schema(BytesIO(_variant_item(whole_tag='single @override=nosuch')))


def test_override_removes_the_sibling_under_keep_controls():
    result = render(BytesIO(_variant_item()),
                    _variant_data(whole='Senior X', subs=['a']), keep_controls=True)
    assert 'a' not in full_text(reopen(result))


def test_override_naming_itself_fails_derive():
    with pytest.raises(DeriveError, match='names the node itself'):
        derive_schema(BytesIO(_variant_item(
            whole_tag='single @override=whole_experience')))
