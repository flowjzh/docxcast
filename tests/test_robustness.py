"""Robustness and error-path tests."""

from io import BytesIO
from zipfile import BadZipFile

import pytest

from docxcast import RenderError, Schema, derive_schema, render
from tests.helpers import build_doc, empty_template, lock_pr, make_sdt, run


def test_derive_rejects_garbage():
    with pytest.raises(BadZipFile):
        derive_schema(BytesIO(b'not a docx at all'))


def test_derive_rejects_truncated_docx():
    full = empty_template()
    with pytest.raises(BadZipFile):
        derive_schema(BytesIO(full[: len(full) // 2]))


def test_derive_empty_body():
    schema = derive_schema(BytesIO(empty_template()))
    assert schema == Schema(children=[])
    assert schema.to_json_schema() == {'type': 'object'}
    assert schema.validate({}).ok


def test_validate_non_dict_data():
    schema = Schema(children=[])
    result = schema.validate([1, 2])
    assert not result.ok
    assert result.errors[0].path == ''
    assert result.errors[0].expected == 'object'
    assert result.errors[0].got == [1, 2]
    assert schema.validate(None).issues == []


def test_render_with_non_dict_data_still_produces_document():
    result = render(BytesIO(empty_template()), [1, 2])
    assert not result.report.ok
    assert result.document


def test_strict_raises_first_error_in_document_order():
    template = build_doc(
        make_sdt('first', 'field one', pr_extra=lock_pr(), content=run('')),
        make_sdt('second', 'field two', pr_extra=lock_pr(), content=run('')),
    )
    with pytest.raises(RenderError) as exc:
        render(BytesIO(template), {}, strict=True)
    assert exc.value.issue.path == 'first'
    assert exc.value.issue.code == 'missing_required'


def test_validate_never_raises_on_weird_values():
    schema = derive_schema(BytesIO(empty_template()))
    for weird in (42, 'str', [None], {1: 2}):
        assert isinstance(schema.validate(weird).issues, list)


def test_unknown_keys_are_warnings_inside_nested_objects():
    template = build_doc(make_sdt('f', 'field', content=run('v')))
    result = derive_schema(BytesIO(template)).validate({'f': 'x', 'extra': 1})
    assert result.ok
    assert result.warnings[0].code == 'unknown_key'
    assert result.warnings[0].path == 'extra'
    assert result.to_dict()['ok'] is True


def test_issue_to_dict_omits_absent_fields():
    from docxcast.schema import Issue

    assert Issue(path='a', code='unknown_key', message='m').to_dict() == {
        'path': 'a', 'code': 'unknown_key', 'message': 'm',
    }
    d = Issue(path='a', code='type_mismatch', message='m', expected='string', got=1).to_dict()
    assert d['expected'] == 'string' and d['got'] == 1
