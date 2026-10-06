"""Reading chats, channels and replies through the CLI."""
import json

from conftest import ago, chat, msg, system_event

CHAT = '19:abc123@thread.v2'
TEAM = '00000000-0000-0000-0000-00000000aaaa'
CHANNEL = '19:general@thread.tacv2'


def history(count: int) -> list[dict]:
    """`count` chat messages one hour apart, newest first (Graph's order)."""
    return [msg(f'm{i}', ago(hours=i), f'<p>message {i}</p>') for i in range(count)]


def test_limit_returns_newest_messages_without_fetching_whole_history(graph, run):
    graph.add(f'/chats/{CHAT}/messages', history(120))

    code, out, _ = run('messages', CHAT, '--limit', '3')

    assert code == 0
    assert [m['id'] for m in json.loads(out)] == ['m0', 'm1', 'm2']
    assert len(graph.urls) == 1


def test_top_is_an_alias_and_zero_means_everything(graph, run):
    graph.add(f'/chats/{CHAT}/messages', history(120))

    assert len(json.loads(run('messages', CHAT, '--top', '7')[1])) == 7
    assert len(json.loads(run('messages', CHAT, '--limit', '0')[1])) == 120


def test_since_relative_keeps_only_recent_messages(graph, run):
    graph.add(f'/chats/{CHAT}/messages', history(120))

    code, out, _ = run('messages', CHAT, '--since', '1d', '--limit', '0')

    assert code == 0
    assert [m['id'] for m in json.loads(out)] == [f'm{i}' for i in range(24)]
    assert len(graph.urls) == 1  # paging stops once messages are older than --since


def test_since_accepts_iso_dates(graph, run):
    old = msg('old', '2026-01-31T18:59:59Z', 'old')  # local midnight is 19:00Z
    new = msg('new', '2026-01-31T19:00:00.5Z', 'new')
    graph.add(f'/chats/{CHAT}/messages', [new, old])

    _, out, _ = run('messages', CHAT, '--since', '2026-02-01')

    assert [m['id'] for m in json.loads(out)] == ['new']


def test_invalid_since_is_a_usage_error(graph, run):
    code, _, err = run('messages', CHAT, '--since', 'last tuesday')
    assert code == 2
    assert 'invalid --since' in err
    assert graph.calls == []


def test_chats_since_filters_by_last_message(graph, run):
    graph.add('/me/chats', [chat('c1', ago(hours=2), topic='Launch'),
                            chat('c2', ago(days=3)), chat('c3', ago(days=9))])

    _, out, _ = run('chats', '--since', '7d')

    assert [c['id'] for c in json.loads(out)] == ['c1', 'c2']


def test_markdown_chats_name_untitled_chats_by_members(graph, run):
    graph.add('/me/chats', [chat('c1', ago(hours=2), topic='Launch plan'),
                            chat('c2', ago(hours=3), members=('Priya Shah', 'Alex Doe'),
                                 preview='see you <b>soon</b>')])

    _, out, _ = run('chats', '--format', 'md')

    assert '**Launch plan** (group)' in out
    assert '**Priya Shah, Alex Doe** (oneOnOne) — Priya Shah: see you soon' in out
    assert 'id: c2' in out


def test_markdown_messages_are_plain_text_oldest_first_without_system_events(graph, run):
    graph.add(f'/chats/{CHAT}/messages', [
        msg('3', '2026-10-02T09:05:00Z', '<p>Thanks <at id="0">Sam Lee</at>&nbsp;&amp; team</p>'
            '\\\n<p>Second&nbsp;line</p>', attachments=[
                {'contentType': 'reference', 'name': 'plan.pdf'}]),
        {**msg('2', '2026-10-02T09:04:00Z', 'gone'), 'deletedDateTime': '2026-10-02T09:06:00Z'},
        system_event('1', '2026-10-02T09:03:00Z'),
        msg('0', '2026-10-01T17:00:00Z', '<div>Draft is <a href="https://example.com/d">here</a>'
            '</div>', sender='Sam Lee'),
    ])

    code, out, _ = run('-f', 'md', 'messages', CHAT)

    assert code == 0
    assert out == (
        '## 2026-10-01\n'
        '22:00 PKT Sam Lee: Draft is here (https://example.com/d)\n'
        '## 2026-10-02\n'
        '14:05 PKT Alex Doe: Thanks @Sam Lee & team\n'
        '  Second line\n'
        '  [file: plan.pdf]\n')


def test_json_output_keeps_system_events(graph, run):
    graph.add(f'/chats/{CHAT}/messages', [msg('1', ago(hours=1), 'hi'),
                                          system_event('0', ago(hours=2))])
    _, out, _ = run('messages', CHAT)
    assert [m['id'] for m in json.loads(out)] == ['1', '0']


def test_markdown_limit_counts_real_messages_not_system_events(graph, run):
    graph.add(f'/chats/{CHAT}/messages', [system_event('s', ago(minutes=1)),
                                          msg('a', ago(hours=1), 'one'),
                                          msg('b', ago(hours=2), 'two')])
    _, out, _ = run('messages', CHAT, '-n', '2', '-f', 'md')
    assert 'one' in out and 'two' in out


def test_channel_threads_since_counts_recent_replies(graph, run):
    graph.add(f'/teams/{TEAM}/channels/{CHANNEL}/messages', [
        # Graph orders threads by latest activity, so an old post with a new
        # reply comes first.
        {**msg('p-old', ago(days=30), 'old post'),
         'replies': [msg('r1', ago(hours=1), 'fresh reply', sender='Sam Lee')]},
        msg('p-new', ago(days=2), 'new post'),
        msg('p-stale', ago(days=20), 'stale post'),
    ])

    _, out, _ = run('posts', TEAM, CHANNEL, '--since', '7d')
    posts = json.loads(out)
    assert [p['id'] for p in posts] == ['p-old', 'p-new']
    assert all('replies' not in p for p in posts)

    _, out, _ = run('posts', TEAM, CHANNEL, '--since', '7d', '--replies', '-f', 'md')
    assert 'Replies (1):' in out
    assert 'Sam Lee: fresh reply' in out
    assert 'stale post' not in out


def test_teams_channels_and_replies(graph, run):
    graph.add('/me/joinedTeams', [{'id': TEAM, 'displayName': 'Platform', 'description': 'Infra'}])
    graph.add(f'/teams/{TEAM}/channels', [{'id': CHANNEL, 'displayName': 'General',
                                           'membershipType': 'standard'}])
    graph.add(f'/teams/{TEAM}/channels/{CHANNEL}/messages/p1/replies',
              [msg('r2', ago(hours=1), 'second'), msg('r1', ago(hours=2), 'first')])

    assert run('-f', 'md', 'teams')[1] == f'- **Platform** — Infra\n  id: {TEAM}\n'
    assert run('channels', TEAM)[1].count(CHANNEL) == 1
    _, out, _ = run('replies', TEAM, CHANNEL, 'p1', '-f', 'md')
    assert out.index('first') < out.index('second')


def test_status_reports_the_connection(graph, run):
    code, out, _ = run('status')
    assert code == 0
    assert json.loads(out)['connectedAs'] == 'alex@contoso.example'


def test_cards_with_backslash_line_endings_show_their_text(graph, run):
    card = ('{\r\\\n  "type": "AdaptiveCard",\r\\\n  "body": [\r\\\n'
            '    {"type": "TextBlock", "text": "PDTS decision"},\r\\\n'
            '    {"type": "TextBlock", "text": "Approve batch 4"}\r\\\n  ]\r\\\n}')
    graph.add(f'/chats/{CHAT}/messages', [msg('1', '2026-10-02T09:00:00Z', '', attachments=[
        {'contentType': 'application/vnd.microsoft.card.adaptive', 'content': card}])])

    _, out, _ = run('-f', 'md', 'messages', CHAT)

    assert '[card: PDTS decision / Approve batch 4]' in out
