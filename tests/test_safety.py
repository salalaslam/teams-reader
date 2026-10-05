"""Read-only guarantees, setup and error handling."""
import json

import pytest

from conftest import ago, chat, msg

TEAM = '00000000-0000-0000-0000-00000000aaaa'
CHANNEL = '19:general@thread.tacv2'


def test_every_read_command_only_issues_get_requests(graph, run):
    graph.add('/me/chats', [chat('c1', ago(hours=1))])
    graph.add('/chats/c1/messages', [msg('m1', ago(hours=1), 'hello')])
    graph.add('/me/joinedTeams', [{'id': TEAM, 'displayName': 'Platform'}])
    graph.add(f'/teams/{TEAM}/channels', [{'id': CHANNEL, 'displayName': 'General'}])
    graph.add(f'/teams/{TEAM}/channels/{CHANNEL}/messages', [msg('p1', ago(hours=1), 'post')])
    graph.add(f'/teams/{TEAM}/channels/{CHANNEL}/messages/p1/replies', [])

    for argv in (['status'], ['chats'], ['messages', 'c1'], ['teams'], ['channels', TEAM],
                 ['posts', TEAM, CHANNEL, '--replies'], ['replies', TEAM, CHANNEL, 'p1'],
                 ['search', 'hello']):
        for fmt in ('json', 'md'):
            assert run(*argv, '--format', fmt)[0] == 0, argv

    for call in graph.calls:
        assert call[0] in ('status', 'request'), call
        if call[0] == 'request':
            assert call[call.index('--method') + 1] == 'get'
            assert call[call.index('--url') + 1].startswith('https://graph.microsoft.com/v1.0/')
            assert '--body' not in call


def test_ids_cannot_change_the_request_path(graph, run):
    hostile = '19:x/../../me/sendMail?'
    run('messages', hostile)
    (url,) = graph.urls
    path = url.split('?')[0]
    assert path == ('https://graph.microsoft.com/v1.0/chats/'
                    '19%3Ax%2F..%2F..%2Fme%2FsendMail%3F/messages')


def test_m365_errors_are_reported_without_a_traceback(graph, run):
    graph.fail('/me/chats', 'Log in to Microsoft 365 first')
    code, out, err = run('chats')
    assert code == 1
    assert out == ''
    assert err == 'teams-reader: Log in to Microsoft 365 first\n'


def test_missing_m365_explains_how_to_install(monkeypatch, tmp_path, run):
    monkeypatch.delenv('TEAMS_READER_M365', raising=False)
    monkeypatch.setenv('PATH', str(tmp_path))
    monkeypatch.setenv('HOME', str(tmp_path))
    code, _, err = run('chats')
    assert code == 1
    assert 'npm install -g @pnp/cli-microsoft365' in err


def test_login_uses_device_code_with_the_configured_app(graph, run, monkeypatch, tmp_path):
    config = tmp_path / 'config' / 'teams-reader'
    config.mkdir(parents=True)
    (config / 'account.json').write_text(json.dumps({'appId': 'app-guid', 'tenant': 'tenant-guid'}))
    monkeypatch.setenv('XDG_CONFIG_HOME', str(tmp_path / 'config'))

    assert run('login')[0] == 0
    (call,) = graph.calls
    assert call[0] == 'login'
    assert call[call.index('--authType') + 1] == 'deviceCode'
    assert call[call.index('--appId') + 1] == 'app-guid'
    assert call[call.index('--tenant') + 1] == 'tenant-guid'


def test_login_flags_override_config(graph, run, monkeypatch, tmp_path):
    monkeypatch.setenv('XDG_CONFIG_HOME', str(tmp_path / 'none'))
    monkeypatch.setenv('HOME', str(tmp_path))
    monkeypatch.setenv('TEAMS_READER_TENANT', 'env-tenant')
    assert run('login', '--app-id', 'flag-app')[0] == 0
    (call,) = graph.calls
    assert call[call.index('--appId') + 1] == 'flag-app'
    assert call[call.index('--tenant') + 1] == 'env-tenant'


def test_login_without_configuration_explains_setup(graph, run, monkeypatch, tmp_path):
    monkeypatch.setenv('XDG_CONFIG_HOME', str(tmp_path / 'none'))
    monkeypatch.setenv('HOME', str(tmp_path))
    monkeypatch.delenv('TEAMS_READER_APP_ID', raising=False)
    monkeypatch.delenv('TEAMS_READER_TENANT', raising=False)
    code, _, err = run('login')
    assert code == 2
    assert 'TEAMS_READER_APP_ID' in err
    assert graph.calls == []


@pytest.mark.parametrize('argv', [['messages', 'c1', '--limit', '-1'],
                                  ['chats', '--limit', 'ten'],
                                  ['search', '  ']])
def test_bad_arguments_are_usage_errors(graph, run, argv):
    assert run(*argv)[0] == 2
    assert graph.calls == []
