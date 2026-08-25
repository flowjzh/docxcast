"""Acceptance tests against the real resume template fixture."""

import pytest

from docxcast import Field, RenderError, Repeat, Schema, Section, derive_schema, render
from docxcast._ooxml import qn
from tests.helpers import count_sdts, full_text, reopen

EXPECTED_SCHEMA = Schema(children=[
    Section(name='basic_info', remark='the candidate\'s basic information', children=[
        Field(name='name', remark='the candidate\'s full name'),
        Field(name='employment_status', remark='current employment status',
              enum=['employed', 'unemployed'],
              enum_labels=['Employed', 'Unemployed'], required=True),
        Field(name='start_work_date', remark='year the candidate started working',
              format='date-year'),
        Field(name='is_married_with_children', remark='married with children', type='boolean'),
    ]),
    Section(name='career', remark='career history', children=[
        Repeat(name='work_experiences', remark='work experiences, one item per job', fields=[
            Field(name='start_date', remark='job start date', format='date-month'),
            Field(name='end_date', remark='job end date', format='date-month',
                  alternatives=['Present']),
            Field(name='company', remark='company'),
            Field(name='title', remark='job title'),
        ]),
    ]),
    Field(name='creation_time', remark='resume generation time', format='date-time'),
])

VALID_DATA = {
    'basic_info': {
        'name': 'Alex Chen',
        'employment_status': 'employed',
        'start_work_date': '2015',
        'is_married_with_children': True,
    },
    'career': {
        'work_experiences': [
            {'start_date': '2021-07', 'end_date': '2023-03',
             'company': 'Acme Analytics Ltd.', 'title': 'Backend Engineer'},
            {'start_date': '2023-04', 'end_date': '2025-08',
             'company': 'Globex Cloud Inc.', 'title': 'Senior Staff Engineer'},
            {'start_date': '2025-09', 'end_date': 'Present',
             'company': 'Initech Software', 'title': 'Engineering Director'},
        ],
    },
    'creation_time': '2026-08-13 10:30:00',
}

VALID_DATA_MISSING_REQUIRED = {
    **VALID_DATA,
    'basic_info': {k: v for k, v in VALID_DATA['basic_info'].items() if k != 'employment_status'},
}


def test_derive_matches_expected_schema(template):
    assert derive_schema(template) == EXPECTED_SCHEMA


def test_to_json_schema(template):
    json_schema = derive_schema(template).to_json_schema()
    # the structure mirrors the model golden; pin only the composed values
    # that cannot be reconstructed from the Schema object itself
    assert json_schema == EXPECTED_SCHEMA.to_json_schema()
    status = json_schema['properties']['basic_info']['properties']['employment_status']
    assert status['description'] == 'current employment status (Employed=employed; Unemployed=unemployed)'
    assert json_schema['properties']['basic_info']['required'] == ['employment_status']


def test_validate_accepts_valid_data(template):
    result = derive_schema(template).validate(VALID_DATA)
    assert result.ok
    assert result.issues == []


def test_validate_reports_each_issue_code(template):
    schema = derive_schema(template)

    missing = schema.validate({'basic_info': {}, 'career': {}, 'creation_time': '2026-08-13 10:30:00'})
    assert missing.errors[0].code == 'missing_required'
    assert missing.errors[0].path == 'basic_info.employment_status'

    bad_enum = {**VALID_DATA, 'basic_info': {**VALID_DATA['basic_info'], 'employment_status': 'freelancing'}}
    assert schema.validate(bad_enum).errors[0].code == 'enum_invalid'

    bad_format = {**VALID_DATA, 'creation_time': '2026/08/13'}
    assert schema.validate(bad_format).errors[0].code == 'format_invalid'

    bad_month = {**VALID_DATA, 'career': {'work_experiences': [
        {**VALID_DATA['career']['work_experiences'][0], 'start_date': '2021/07'}]}}
    month = schema.validate(bad_month).errors[0]
    assert month.code == 'format_invalid' and month.path == 'career.work_experiences[0].start_date'

    bad_type = {**VALID_DATA, 'basic_info': {**VALID_DATA['basic_info'], 'is_married_with_children': 'yes'}}
    assert schema.validate(bad_type).errors[0].code == 'type_mismatch'

    extra = {**VALID_DATA, 'extra': 1}
    warned = schema.validate(extra)
    assert warned.ok
    assert warned.warnings[0].code == 'unknown_key'
    assert warned.warnings[0].path == 'extra'


def test_validate_optional_format_field_accepts_empty(template):
    # the required side of the rule is pinned by test_required_dropdown_rejects_empty
    schema = derive_schema(template)
    optional = {**VALID_DATA,
                'basic_info': {**VALID_DATA['basic_info'], 'start_work_date': ''}}
    assert schema.validate(optional).ok


def test_render_fills_and_unwraps(template):
    result = render(template, VALID_DATA)
    out = reopen(result)
    assert count_sdts(out) == 0
    text = full_text(out)
    assert 'John Doe' not in text
    assert 'Alex Chen' in text
    assert 'Employed' in text and 'employed' not in text  # displayText written, not the value
    assert '2015' in text
    assert '2021-07 — 2023-03' in text.replace(' ', ' ')
    assert '2025-09 — Present' in text  # the ongoing job renders the literal branch
    assert '2026-08-13 10:30:00' in text
    assert text.index('Backend Engineer') < text.index('Senior Staff Engineer')  # two cloned paragraphs in order


def test_render_preserves_run_formatting(template):
    result = render(template, VALID_DATA)
    out = reopen(result)
    for run in out.element.body.iter(qn('w:r')):
        texts = run.findall(qn('w:t'))
        if texts and texts[0].text == 'Alex Chen':
            r_pr = run.find(qn('w:rPr'))
            assert r_pr is not None  # the template run keeps its rPr (exact content covered synthetically)
            assert r_pr.find(qn('w:rFonts')) is not None
            break
    else:
        pytest.fail('name run not found')


def test_render_lenient_keeps_missing_required(template):
    result = render(template, VALID_DATA_MISSING_REQUIRED)
    assert not result.report.ok
    assert result.report.errors[0].code == 'missing_required'
    out = reopen(result)
    assert 'Alex Chen' in full_text(out)  # the rest still renders


def test_render_strict_raises(template):
    with pytest.raises(RenderError) as exc:
        render(template, VALID_DATA_MISSING_REQUIRED, strict=True)
    assert exc.value.issue.path == 'basic_info.employment_status'


def test_render_omits_missing_section(template):
    data = {k: v for k, v in VALID_DATA.items() if k != 'career'}
    result = render(template, data)
    text = full_text(reopen(result))
    assert 'Engineering Director' not in text  # career section removed entirely
    assert '2024-02' not in text
    assert result.report.ok


def test_render_keep_controls_round_trip(template):
    result = render(template, VALID_DATA, keep_controls=True)
    out = reopen(result)
    assert count_sdts(out) > 0
    text = full_text(out)
    assert 'Alex Chen' in text
    assert text.count('Backend Engineer') == 1  # both experience items kept as controls
    assert text.count('Senior Staff Engineer') == 1
