"""Client-side search across chats and channels."""
import json

from conftest import ago, chat, msg

TEAM = '00000000-0000-0000-0000-00000000aaaa'
CHANNEL = '19:general@thread.tacv2'


def workspace(graph):
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
         'replies': [msg('r1', ago(hours=3), 'Budget numbers attached', sender='Priya Shah')]},
    ])


def test_search_finds_case_insensitive_matches_in_chats_and_channel_replies(graph, run):
    workspace(graph)

    code, out, err = run('search', 'budget')

    assert code == 0, err
    hits = json.loads(out)
    # m3 only quotes the budget message, m1 and m5 are older than the default 7d.
    assert [h['messageId'] for h in hits] == ['m2', 'r1']
    assert hits[0] | {'createdDateTime': None, 'webUrl': None} == {
        'source': 'chat', 'chatId': 'c1', 'chatName': 'Launch', 'messageId': 'm2',
        'from': 'Sam Lee', 'text': 'The BUDGET is approved',
        'createdDateTime': None, 'webUrl': None}
    assert hits[1]['source'] == 'channel'
    assert (hits[1]['teamName'], hits[1]['channelName'], hits[1]['threadId']) == (
        'Platform', 'General', 'p1')


def test_search_respects_since_limit_and_channel_opt_out(graph, run):
    workspace(graph)

    _, out, _ = run('search', 'budget', '--since', '60d', '--no-channels')
    assert [h['messageId'] for h in json.loads(out)] == ['m2', 'm1', 'm5']

    _, out, _ = run('search', 'budget', '--limit', '1')
    assert [h['messageId'] for h in json.loads(out)] == ['m2']


def test_search_markdown_points_to_the_source(graph, run):
    workspace(graph)
    _, out, _ = run('search', 'budget', '-f', 'md')
    assert 'chat **Launch** · Sam Lee: The BUDGET is approved\n  chat: c1' in out
    assert '**Platform / General** · Priya Shah: Budget numbers attached' in out
    assert 'thread: p1' in out


def test_search_skips_a_failing_source_with_a_warning(graph, run):
    workspace(graph)
    graph.fail('/chats/c2/messages', 'Forbidden')

    code, out, err = run('search', 'budget')

    assert code == 0
    assert len(json.loads(out)) == 2
    assert 'warning: skipped chat Priya Shah, Alex Doe: Forbidden' in err
