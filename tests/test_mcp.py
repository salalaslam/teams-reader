"""The MCP server (optional `mcp` extra)."""
import asyncio
import json
import os
import sys

import pytest

pytest.importorskip('mcp')

from mcp import ClientSession  # noqa: E402
from mcp.client.stdio import StdioServerParameters, stdio_client  # noqa: E402

from conftest import ago, chat, msg  # noqa: E402

CHAT = '19:abc123@thread.v2'


async def with_session(fn):
    params = StdioServerParameters(command=sys.executable, args=['-m', 'teams_reader', 'mcp'],
                                   env=dict(os.environ))
    async with stdio_client(params) as (read, write):
        async with ClientSession(read, write) as session:
            await session.initialize()
            return await fn(session)


def text_of(result) -> str:
    assert not result.is_error, result.content
    return ''.join(block.text for block in result.content)


def test_every_tool_is_declared_read_only_and_none_can_write(graph):
    tools = asyncio.run(with_session(lambda s: s.list_tools())).tools

    names = {t.name for t in tools}
    assert {'list_chats', 'read_chat', 'search_messages', 'read_channel'} <= names
    for tool in tools:
        assert tool.annotations.read_only_hint is True, tool.name
        assert tool.annotations.destructive_hint is False, tool.name
        assert not any(word in tool.name for word in
                       ('send', 'post_', 'reply_to', 'delete', 'update', 'login', 'logout'))


def test_tools_return_compact_markdown_by_default(graph):
    graph.add('/me/chats', [chat(CHAT, ago(hours=1), topic='Launch')])
    graph.add(f'/chats/{CHAT}/messages', [msg('m1', ago(hours=1), '<p>Ship <b>it</b></p>')])

    async def calls(session):
        chats = await session.call_tool('list_chats', {'since': '1d'})
        messages = await session.call_tool('read_chat', {'chat_id': CHAT})
        raw = await session.call_tool('read_chat', {'chat_id': CHAT, 'format': 'json'})
        return chats, messages, raw

    chats, messages, raw = asyncio.run(with_session(calls))
    assert '**Launch** (group)' in text_of(chats)
    assert 'Alex Doe: Ship it' in text_of(messages)
    assert json.loads(text_of(raw))[0]['id'] == 'm1'


def test_tool_errors_are_reported_to_the_client(graph):
    graph.fail('/me/chats', 'Log in to Microsoft 365 first')
    result = asyncio.run(with_session(lambda s: s.call_tool('list_chats', {})))
    assert result.is_error
    assert 'Log in to Microsoft 365 first' in ''.join(b.text for b in result.content)
