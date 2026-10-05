"""Parsing --since values and reducing Teams HTML to text."""
from datetime import datetime, timedelta, timezone

import pytest

from teams_reader.render import html_to_text
from teams_reader.timeutil import parse_since, parse_time

NOW = datetime(2026, 10, 5, 12, 0, tzinfo=timezone.utc)


@pytest.mark.parametrize('value, expected', [
    ('30m', NOW - timedelta(minutes=30)),
    ('12h', NOW - timedelta(hours=12)),
    ('7d', NOW - timedelta(days=7)),
    ('2W', NOW - timedelta(weeks=2)),
    ('2026-10-01', datetime(2026, 10, 1, tzinfo=timezone.utc)),
    ('2026-10-01T09:30', datetime(2026, 10, 1, 9, 30, tzinfo=timezone.utc)),
    ('2026-10-01T09:30:00Z', datetime(2026, 10, 1, 9, 30, tzinfo=timezone.utc)),
    ('2026-10-01T11:30:00+02:00', datetime(2026, 10, 1, 9, 30, tzinfo=timezone.utc)),
])
def test_since_values(value, expected):
    assert parse_since(value, now=NOW) == expected


@pytest.mark.parametrize('value', ['yesterday', '7', 'd7', '2026-13-01'])
def test_invalid_since_values(value):
    with pytest.raises(ValueError):
        parse_since(value, now=NOW)


@pytest.mark.parametrize('stamp', ['2026-10-02T13:23:27.03Z', '2026-10-02T13:23:27.0300000Z'])
def test_graph_timestamps_with_any_fraction_length(stamp):
    assert parse_time(stamp) == datetime(2026, 10, 2, 13, 23, 27, 30000, tzinfo=timezone.utc)


@pytest.mark.parametrize('html, text', [
    ('<p>Hello&nbsp;<at id="0">Sam</at>, see&amp;go</p>', 'Hello @Sam, see&go'),
    ('<p>one</p>\\\n<p>two<br>three</p>', 'one\ntwo\nthree'),
    ('<ul><li>alpha</li><li>beta</li></ul>', '- alpha\n- beta'),
    ('<p>Run <code>make test</code></p>', 'Run `make test`'),
    ('<a href="https://example.com/x">docs</a>', 'docs (https://example.com/x)'),
    ('<a href="https://example.com/x">https://example.com/x</a>', 'https://example.com/x'),
    ('<p>Nice <emoji id="1" alt="👍" title="Like"><img src="x"></emoji></p>', 'Nice 👍'),
    ('<p>see</p><img src="https://x/y.png" alt="image">', 'see\n[image]'),
    ('<attachment id="123"></attachment><p>reply</p>', 'reply'),
    ('<table><tr><th>a</th><th>b</th></tr><tr><td>1</td><td>2</td></tr></table>',
     '| a | b\n| 1 | 2'),
])
def test_html_to_text(html, text):
    assert html_to_text(html) == text


def test_plain_text_bodies_are_kept():
    assert html_to_text('a <b> c', 'text') == 'a <b> c'
