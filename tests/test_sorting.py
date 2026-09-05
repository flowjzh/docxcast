"""The repeat-level @sort directive: derive-side parsing and the
deterministic sort over extracted data.

Sort-semantics fixtures derive from real templates — the sort kind
comes from the inference path the tests are meant to protect.
"""

from io import BytesIO

import pytest

from docxcast import DeriveError, Schema, derive_schema, render
from docxcast.sorting import apply_sorts
from tests.helpers import (build_doc, date_pr, docpartlist_pr, full_text,
                           make_sdt, para, reopen, repeating_item_pr,
                           repeating_pr, run)


def _work_template(tag: str, *, date_format: str = 'yyyy.M',
                   store: str = 'date', nested: bool = False) -> bytes:
    """A work-experience repeat over company + start_date (the given
    date format); ``nested`` opts into a sub repeat whose own tag
    carries a sort."""
    content = [
        para(make_sdt('company', None, content=run('x'))),
        para(make_sdt('start_date', None,
                      pr_extra=date_pr(store, date_format), content=run('x')))]
    if nested:
        content.append(para(make_sdt('sub', '@sort=sub_date:desc',
                                     pr_extra=repeating_pr(), content=
              make_sdt(pr_extra=repeating_item_pr(), content=
                       para(make_sdt('sub_date', None,
                                     pr_extra=date_pr('date', 'yyyy.M'),
                                     content=run('x')))))))
    item = make_sdt(pr_extra=repeating_item_pr(), content=content)
    return build_doc(make_sdt('work_experiences', tag,
                              pr_extra=repeating_pr(), content=item))


def _work_schema(tag: str = 'h @sort=start_date:desc', *,
                 date_format: str = 'yyyy.M', store: str = 'date',
                 nested: bool = False) -> Schema:
    return derive_schema(BytesIO(
        _work_template(tag, date_format=date_format, store=store,
                       nested=nested)))


# --- derive: directive parsing ---

def test_tag_sort_directive_becomes_a_spec_and_leaves_the_remark():
    repeat = _work_schema('工作经历：逐条保留 @sort=start_date:desc').children[0]
    assert repeat.remark == '工作经历：逐条保留'
    assert repeat.sort.key == 'start_date'
    assert repeat.sort.desc is True and repeat.sort.kind == 'date'
    # the JSON schema never carries the directive
    assert '@sort' not in repeat.to_json_schema()['description']


def test_order_defaults_to_asc_and_kind_to_alphanumeric():
    repeat = _work_schema('notes @sort=company').children[0]
    assert repeat.sort.key == 'company'
    assert repeat.sort.desc is False and repeat.sort.kind == 'alphanumeric'


def test_datetime_format_infers_datetime_kind():
    repeat = _work_schema('h @sort=start_date:desc',
                          date_format='yyyy-MM-dd').children[0]
    assert repeat.sort.kind == 'date'  # a date picker stays a date sort


def test_unknown_sort_key_fails_at_derive():
    with pytest.raises(DeriveError, match='not a field'):
        _work_schema('h @sort=salary:desc')


def test_sort_key_walking_through_a_repeat_fails_at_derive():
    with pytest.raises(DeriveError, match='walks through a repeat'):
        _work_schema('h @sort=sub.sub_date:desc', nested=True)


def test_mangled_directive_fails_at_derive():
    with pytest.raises(DeriveError, match='unparseable @sort'):
        _work_schema('h @sort=start_date:descending')


# --- sorting: the three kinds, tails, stability ---

def job(date, name='c'):
    return {'company': name, 'start_date': date}


def test_date_sort_desc_newest_first():
    data = {'work_experiences': [job('2019.4'), job('2023.5'), job('2022-05')]}
    apply_sorts(_work_schema(), data)
    assert [j['start_date'] for j in data['work_experiences']] == \
        ['2023.5', '2022-05', '2019.4']


def test_apply_sorts_reports_position_to_original_index():
    data = {'work_experiences': [job('2019.4', 'A'), job('2023.5', 'B'),
                                 job('2022-05', 'C')]}
    perms = apply_sorts(_work_schema(), data)
    assert perms == {'work_experiences': [1, 2, 0]}
    # identity sort is a permutation too — hosts realign by it blindly
    assert apply_sorts(_work_schema(),
                       {'work_experiences': [job('2023.5', 'B')]}) == \
        {'work_experiences': [0]}
    assert apply_sorts(_work_schema(), {}) == {}


def test_duplicate_repeat_nodes_keep_the_first_permutation():
    # a template duplicate makes derive emit one repeat twice: the
    # second walk sees the now-sorted array and must not overwrite the
    # pre-sort permutation with an identity
    from copy import deepcopy
    schema = _work_schema()
    schema.children.append(deepcopy(schema.children[0]))
    data = {'work_experiences': [job('2019.4'), job('2023.5')]}
    assert apply_sorts(schema, data) == {'work_experiences': [1, 0]}


def test_date_sort_asc_oldest_first():
    data = {'work_experiences': [job('2023-05'), job('2019.4')]}
    apply_sorts(_work_schema('h @sort=start_date'), data)
    assert [j['start_date'] for j in data['work_experiences']] == \
        ['2019.4', '2023-05']


def test_unparsable_dates_sort_to_the_tail_either_way():
    data = {'work_experiences': [job('2021-04'), job(''), job('2020-01')]}
    apply_sorts(_work_schema(), data)
    assert [j['start_date'] for j in data['work_experiences']] == \
        ['2021-04', '2020-01', '']
    data = {'work_experiences': [job(''), job('2020-01')]}
    apply_sorts(_work_schema('h @sort=start_date'), data)
    assert [j['start_date'] for j in data['work_experiences']] == \
        ['2020-01', '']


def test_date_sort_is_stable_for_equal_keys():
    a, b = job('2021-04', 'A'), job('2021-04', 'B')
    data = {'work_experiences': [a, b]}
    apply_sorts(_work_schema(), data)
    assert data['work_experiences'] == [a, b]


def test_datetime_sort_orders_clock_time():
    schema = _work_schema('h @sort=start_date:desc',
                          date_format='yyyy-MM-dd', store='dateTime')
    assert schema.children[0].sort.kind == 'datetime'
    data = {'work_experiences': [
        {'company': 'c', 'start_date': '2023-05-01 09:00'},
        {'company': 'c', 'start_date': '2023-05-01 09:01'},
        {'company': 'c', 'start_date': '2023-04-30 23:59'}]}
    apply_sorts(schema, data)
    assert [j['start_date'] for j in data['work_experiences']] == \
        ['2023-05-01 09:01', '2023-05-01 09:00', '2023-04-30 23:59']


def test_alphanumeric_sort_is_natural():
    data = {'work_experiences': [{'company': 'Level 10'},
                                 {'company': 'level 2'},
                                 {'company': 'Acme'}]}
    apply_sorts(_work_schema('h @sort=company'), data)
    assert [j['company'] for j in data['work_experiences']] == \
        ['Acme', 'level 2', 'Level 10']


def test_nested_repeat_sorts_each_parent_item():
    schema = _work_schema(nested=True)
    data = {'work_experiences': [
        {'company': 'c1', 'start_date': '2020-01',
         'sub': [{'sub_date': '2019-04'}, {'sub_date': '2021-04'}]},
        {'company': 'c2', 'start_date': '2018-01',
         'sub': [{'sub_date': '2022-01'}, {'sub_date': '2021-01'}]}]}
    apply_sorts(schema, data)
    assert [j['company'] for j in data['work_experiences']] == ['c1', 'c2']
    assert [s['sub_date'] for s in data['work_experiences'][0]['sub']] == \
        ['2021-04', '2019-04']
    assert [s['sub_date'] for s in data['work_experiences'][1]['sub']] == \
        ['2022-01', '2021-01']


def test_nested_repeat_sorts_but_reports_no_perms():
    schema = _work_schema('h', nested=True)  # only the sub repeat sorts
    data = {'work_experiences': [
        {'company': 'c1', 'start_date': '2020-01',
         'sub': [{'sub_date': '2019-04'}, {'sub_date': '2021-04'}]}]}
    assert apply_sorts(schema, data) == {}
    assert [s['sub_date'] for s in data['work_experiences'][0]['sub']] == \
        ['2021-04', '2019-04']


def test_section_nested_key_addresses_through_sections():
    section = make_sdt('info', None, pr_extra=docpartlist_pr(), content=
                       para(make_sdt('start_date', None,
                                     pr_extra=date_pr('date', 'yyyy.M'),
                                     content=run('x'))))
    item = make_sdt(pr_extra=repeating_item_pr(), content=[section])
    template = build_doc(make_sdt('work_experiences',
                                  'h @sort=info.start_date:desc',
                                  pr_extra=repeating_pr(), content=item))
    schema = derive_schema(BytesIO(template))
    data = {'work_experiences': [{'info': {'start_date': '2019-04'}},
                                 {'info': {'start_date': '2023-05'}}]}
    apply_sorts(schema, data)
    assert [j['info']['start_date'] for j in data['work_experiences']] == \
        ['2023-05', '2019-04']


def test_repeat_nested_under_a_section_sorts_too():
    inner = make_sdt('start_date', None,
                     pr_extra=date_pr('date', 'yyyy.M'), content=run('x'))
    repeat = make_sdt('work_experiences', 'h @sort=start_date:desc',
                      pr_extra=repeating_pr(), content=
                      make_sdt(pr_extra=repeating_item_pr(), content=
                               para(inner)))
    section = make_sdt('history', None, pr_extra=docpartlist_pr(),
                       content=[repeat])
    template = build_doc(section)
    schema = derive_schema(BytesIO(template))
    data = {'history': {'work_experiences': [
        {'start_date': '2019.4'}, {'start_date': '2023.5'}]}}
    apply_sorts(schema, data)
    assert [j['start_date'] for j in data['history']['work_experiences']] == \
        ['2023.5', '2019.4']


def test_render_applies_sort_directives():
    template = _work_template('h @sort=start_date:desc')
    data = {'work_experiences': [
        {'company': 'older', 'start_date': '2019.4'},
        {'company': 'newer', 'start_date': '2023.5'}]}
    result = render(BytesIO(template), data)
    text = full_text(reopen(result))
    assert text.index('newer') < text.index('older')
    # the sort is in place: the caller's dict and the document agree
    assert [j['company'] for j in data['work_experiences']] == \
        ['newer', 'older']
    assert result.sort == {'work_experiences': [1, 0]}
