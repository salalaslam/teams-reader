"""Searching messages: Microsoft Search by default, client-side scan with --scan."""
import json

from conftest import ago, chat, msg, search_hit

TEAM = '00000000-0000-0000-0000-00000000aaaa'
CHANNEL = '19:general@thread.tacv2'


def named_sources(graph):
    graph.add_object('/chats/c1', chat('c1', ago(hours=1), topic='Launch'))
    graph.add_object('/chats/c2', chat('c2', ago(days=2), members=('Priya Shah', 'Alex Doe')))
    graph.add_object(f'/teams/{TEAM}/channels/{CHANNEL}', {'id': CHANNEL, 'displayName': 'General'})
    graph.add('/me/joinedTeams', [{'id': TEAM, 'displayName': 'Platform'}])


def test_search_uses_microsoft_search_and_names_each_source(graph, run):
    named_sources(graph)
    graph.search_hits([
        search_hit('h1', ago(hours=1), 'The <c0>budget</c0> is approved<ddd/>', 'Sam Lee',
                   chat_id='c1'),
        search_hit('h2', ago(hours=5), '<c0>Budget</c0> numbers attached', 'Priya Shah',
                   team_id=TEAM, channel_id=CHANNEL),
        search_hit('h3', ago(days=3), 'old <c0>budget</c0> chat', chat_id='c2'),
    ])

    code, out, err = run('search', 'budget')

    assert code == 0, err
    assert graph.search_queries == ['budget']
    hits = json.loads(out)
    assert [h['messageId'] for h in hits] == ['h1', 'h2', 'h3']
    assert {k: hits[0][k] for k in ('source', 'chatId', 'chatName', 'from', 'text')} == {
        'source': 'chat', 'chatId': 'c1', 'chatName': 'Launch', 'from': 'Sam Lee',
        'text': 'The budget is approved…'}
    assert (hits[1]['teamName'], hits[1]['channelName']) == ('Platform', 'General')
    assert hits[2]['chatName'] == 'Priya Shah, Alex Doe'


def test_search_treats_a_chat_hit_with_a_teamless_channel_identity_as_a_chat(graph, run):
    # Microsoft Search fills channelIdentity.channelId with the chat ID on chat hits.
    named_sources(graph)
    graph.search_hits([search_hit('h1', ago(hours=1), 'budget', chat_id='c1', channel_id='c1')])

    code, out, err = run('search', 'budget')

    assert code == 0, err
    hit, = json.loads(out)
    assert (hit['source'], hit['chatId'], hit['chatName']) == ('chat', 'c1', 'Launch')


def test_search_since_is_sent_to_microsoft_search_and_enforced(graph, run):
    named_sources(graph)
    graph.search_hits([search_hit('new', ago(hours=1), 'budget', chat_id='c1'),
                       search_hit('old', ago(days=3), 'budget', chat_id='c2')])

    _, out, _ = run('search', 'budget', '--since', '2026-01-15')
    # Local midnight on the 15th is 19:00Z on the 14th; the query allows a day of slack.
    assert graph.search_queries == ['budget sent>=2026-01-13']

    _, out, _ = run('search', 'budget', '--since', '1d')
    assert [h['messageId'] for h in json.loads(out)] == ['new']


def test_search_limit_and_channel_opt_out(graph, run):
    named_sources(graph)
    graph.search_hits([search_hit('h1', ago(hours=1), 'x', team_id=TEAM, channel_id=CHANNEL),
                       search_hit('h2', ago(hours=2), 'x', chat_id='c1'),
                       search_hit('h3', ago(hours=3), 'x', chat_id='c2')])

    _, out, _ = run('search', 'x', '--no-channels')
    assert [h['messageId'] for h in json.loads(out)] == ['h2', 'h3']

    _, out, _ = run('search', 'x', '--limit', '1')
    assert [h['messageId'] for h in json.loads(out)] == ['h1']


def test_search_still_returns_results_when_a_name_lookup_fails(graph, run):
    graph.search_hits([search_hit('h1', ago(hours=1), 'budget', chat_id='c-gone')])
    code, out, err = run('search', 'budget', '-f', 'md')
    assert code == 0
    assert 'chat **c-gone** · Alex Doe: budget\n  chat: c-gone' in out
    assert 'warning: skipped chat c-gone' in err


def recent_workspace(graph):
    graph.add('/me/chats', [chat('c1', ago(hours=1), topic='Launch'),
                            chat('c2', ago(days=2), members=('Priya Shah', 'Alex Doe')),
                            chat('c-old', ago(days=40), topic='Archive')])
    graph.add('/chats/c1/messages', [
        msg('m3', ago(hours=1), '<p>Agreed</p>', attachments=[{
            'contentType': 'messageReference',
            'content': json.dumps({'messagePreview': 'Budget is approved',
                                   'messageSender': {'user': {'displayName': 'Sam Lee'}}})}]),
        msg('m2', ago(hours=2), '<p>The <b>BUDGET</b> is approved</p>', sender='Sam Lee'),
        msg('m1', ago(days=10), '<p>Old budget thread</p>'),
    ])
    graph.add('/chats/c2/messages', [msg('m4', ago(days=2), 'no match here')])
    graph.add('/chats/c-old/messages', [msg('m5', ago(days=40), 'budget from long ago')])
    graph.add('/me/joinedTeams', [{'id': TEAM, 'displayName': 'Platform'}])
    graph.add(f'/teams/{TEAM}/channels', [{'id': CHANNEL, 'displayName': 'General'}])
    graph.add(f'/teams/{TEAM}/channels/{CHANNEL}/messages', [
        {**msg('p1', ago(days=20), 'Quarterly planning'),
         'replies': [msg('r1', ago(hours=3), 'xBudgetx numbers', sender='Priya Shah')]},
    ])


def test_scan_matches_substrings_in_recent_chats_and_channel_replies(graph, run):
    recent_workspace(graph)

    code, out, err = run('search', 'budget', '--scan')

    assert code == 0, err
    assert graph.search_queries == []
    hits = json.loads(out)
    # m3 only quotes the budget message; m1 and m5 are older than the default 7d.
    assert [h['messageId'] for h in hits] == ['m2', 'r1']
    assert {k: hits[0][k] for k in ('source', 'chatId', 'chatName', 'from', 'text')} == {
        'source': 'chat', 'chatId': 'c1', 'chatName': 'Launch', 'from': 'Sam Lee',
        'text': 'The BUDGET is approved'}
    assert (hits[1]['teamName'], hits[1]['channelName'], hits[1]['threadId']) == (
        'Platform', 'General', 'p1')


def test_scan_respects_since_limit_and_channel_opt_out(graph, run):
    recent_workspace(graph)

    _, out, _ = run('search', 'budget', '--scan', '--since', '60d', '--no-channels')
    assert [h['messageId'] for h in json.loads(out)] == ['m2', 'm1', 'm5']

    _, out, _ = run('search', 'budget', '--scan', '--limit', '1')
    assert [h['messageId'] for h in json.loads(out)] == ['m2']


def test_scan_markdown_points_to_the_source(graph, run):
    recent_workspace(graph)
    _, out, _ = run('search', 'budget', '--scan', '-f', 'md')
    assert 'chat **Launch** · Sam Lee: The BUDGET is approved\n  chat: c1' in out
    assert '**Platform / General** · Priya Shah: xBudgetx numbers' in out
    assert 'thread: p1' in out


def test_scan_skips_a_failing_source_with_a_warning(graph, run):
    recent_workspace(graph)
    graph.fail('/chats/c2/messages', 'Forbidden')

    code, out, err = run('search', 'budget', '--scan')

    assert code == 0
    assert len(json.loads(out)) == 2
    assert 'warning: skipped chat Priya Shah, Alex Doe: Forbidden' in err
