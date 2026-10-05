"""Compact Markdown rendering of Graph objects, for LLM context windows.

HTML is reduced to plain text, mentions become ``@Name``, attachments become
one-line placeholders, and system events are expected to be filtered out
before rendering (see ``reader.is_system_event``).
"""
from __future__ import annotations

import json
import re
from datetime import datetime
from html.parser import HTMLParser
from typing import Any, Callable

from .timeutil import item_time

_BLOCK = {'p', 'div', 'h1', 'h2', 'h3', 'h4', 'h5', 'h6', 'tr', 'ul', 'ol', 'table',
          'blockquote', 'pre', 'codeblock', 'hr'}
_SKIP = {'attachment', 'systemeventmessage', 'style', 'script', 'emoji'}


class _TextExtractor(HTMLParser):
    def __init__(self) -> None:
        super().__init__(convert_charrefs=True)
        self.parts: list[str] = []
        self.skip_depth = 0
        self.links: list[tuple[str, int]] = []

    def handle_starttag(self, tag: str, attrs: list[tuple[str, str | None]]) -> None:
        attr = dict(attrs)
        if tag in _SKIP:
            if tag == 'emoji' and attr.get('alt') and not self.skip_depth:
                self.parts.append(attr['alt'] or '')
            self.skip_depth += 1
            return
        if self.skip_depth:
            return
        if tag in _BLOCK:
            self.parts.append('\n')
        elif tag == 'br':
            self.parts.append('\n')
        elif tag == 'li':
            self.parts.append('\n- ')
        elif tag in ('td', 'th'):
            self.parts.append(' | ')
        elif tag == 'at':
            self.parts.append('@')
        elif tag == 'ddd':  # elision marker in Microsoft Search snippets
            self.parts.append('…')
        elif tag == 'code':
            self.parts.append('`')
        elif tag == 'img':
            if 'emoji' in (attr.get('itemtype') or '').lower() and attr.get('alt'):
                self.parts.append(attr['alt'] or '')
            else:
                self.parts.append('[image]')
        elif tag == 'a':
            self.links.append((attr.get('href') or '', len(self.parts)))

    def handle_startendtag(self, tag: str, attrs: list[tuple[str, str | None]]) -> None:
        if tag in _SKIP:  # self-closing: nothing to skip
            if tag == 'emoji' and dict(attrs).get('alt') and not self.skip_depth:
                self.parts.append(dict(attrs)['alt'] or '')
            return
        self.handle_starttag(tag, attrs)

    def handle_endtag(self, tag: str) -> None:
        if tag in _SKIP:
            self.skip_depth = max(0, self.skip_depth - 1)
            return
        if self.skip_depth:
            return
        if tag in _BLOCK:
            self.parts.append('\n')
        elif tag == 'code':
            self.parts.append('`')
        elif tag == 'a' and self.links:
            href, start = self.links.pop()
            text = ''.join(self.parts[start:]).strip()
            if href.startswith(('http://', 'https://', 'mailto:')) and href not in text:
                self.parts.append(f' ({href})')

    def handle_data(self, data: str) -> None:
        if not self.skip_depth:
            self.parts.append(re.sub(r'[ \t\r\n\f\xa0]+', ' ', data))


def html_to_text(content: str | None, content_type: str = 'html') -> str:
    """Plain text from a Teams message body, one non-empty line per block."""
    if not content:
        return ''
    if content_type != 'html':
        text = content
    else:
        parser = _TextExtractor()
        # Teams separates blocks with a literal backslash-newline in the raw
        # HTML source; neither is meaningful in HTML, so drop both.
        parser.feed(re.sub(r'\\\r?\n', '\n', content))
        parser.close()
        text = ''.join(parser.parts)
    lines = (re.sub(r' {2,}', ' ', line).strip() for line in text.splitlines())
    return '\n'.join(line for line in lines if line)


def message_text(message: dict, quotes: bool = True) -> str:
    """Body text plus attachment placeholders (optionally without quoted messages)."""
    body = message.get('body') or {}
    parts = [html_to_text(body.get('content'), body.get('contentType', 'html'))]
    subject = message.get('subject')
    if subject:
        parts.insert(0, f'Subject: {subject}')
    attachments = [(a.get('contentType') or '', _attachment(a))
                   for a in message.get('attachments') or []]
    quoted = [text for kind, text in attachments if kind.endswith('essageReference')]
    others = [text for kind, text in attachments if not kind.endswith('essageReference')]
    return '\n'.join(p for p in [*(quoted if quotes else []), *parts, *others] if p)


def _attachment(att: dict) -> str:
    kind = att.get('contentType') or ''
    content = _json(att.get('content'))
    if kind == 'reference':
        return f"[file: {att.get('name') or 'attachment'}]"
    if kind in ('messageReference', 'forwardedMessageReference'):
        sender = _name((content or {}).get('messageSender'))
        preview = _clip(html_to_text((content or {}).get('messagePreview'), 'text'), 120)
        label = 'forwarded from' if kind.startswith('forwarded') else 'replying to'
        return f'> {label} {sender}: {preview}'
    if kind == 'application/vnd.microsoft.card.adaptive':
        texts = [t for t in _card_texts(content) if t]
        return f"[card: {_clip(' / '.join(texts), 300)}]" if texts else '[card]'
    if kind == 'application/vnd.microsoft.card.audio':
        return '[audio]'
    return f"[attachment: {att.get('name') or kind or 'unknown'}]"


def _card_texts(node: Any) -> list[str]:
    if isinstance(node, dict):
        out = []
        if node.get('type') in ('TextBlock', 'TextRun') and isinstance(node.get('text'), str):
            out.append(html_to_text(node['text'], 'text'))
        for value in node.values():
            if isinstance(value, (dict, list)):
                out.extend(_card_texts(value))
        return out
    if isinstance(node, list):
        return [t for item in node for t in _card_texts(item)]
    return []


def _json(value: Any) -> Any:
    if isinstance(value, str):
        try:
            return json.loads(value)
        except json.JSONDecodeError:
            return None
    return value


def _clip(text: str, length: int) -> str:
    text = ' '.join(text.split())
    return text if len(text) <= length else text[:length - 1].rstrip() + '…'


def _name(identity: dict | None) -> str:
    identity = identity or {}
    for key in ('user', 'application', 'device'):
        name = (identity.get(key) or {}).get('displayName')
        if name:
            return name
    return 'unknown'


def sender(message: dict) -> str:
    return _name(message.get('from'))


def _stamp(when: datetime | None, fmt: str = '%Y-%m-%d %H:%MZ') -> str:
    return when.strftime(fmt) if when else '????-??-?? ??:??Z'


def _indent(text: str, prefix: str = '  ') -> str:
    return '\n'.join(prefix + line for line in text.splitlines())


def chat_name(chat: dict, limit: int = 4) -> str:
    if chat.get('topic'):
        return chat['topic']
    names = [m.get('displayName') for m in chat.get('members') or [] if m.get('displayName')]
    if not names:
        return '(untitled chat)'
    extra = f' +{len(names) - limit}' if len(names) > limit else ''
    return ', '.join(names[:limit]) + extra


def chats(items: list[dict]) -> str:
    lines = []
    for chat in items:
        preview = chat.get('lastMessagePreview') or {}
        when = _stamp(item_time(preview))
        line = f"- {when} **{chat_name(chat)}** ({chat.get('chatType', '?')})"
        text = html_to_text((preview.get('body') or {}).get('content'),
                            (preview.get('body') or {}).get('contentType', 'html'))
        if text:
            line += f' — {sender(preview)}: {_clip(text, 100)}'
        lines.append(f"{line}\n  id: {chat.get('id')}")
    return '\n'.join(lines) or '(no chats)'


def messages(items: list[dict], empty: str = '(no messages)') -> str:
    """Chat messages or replies, oldest first, grouped by UTC date."""
    out: list[str] = []
    day = None
    for msg in sorted(items, key=lambda m: m.get('createdDateTime') or ''):
        when = item_time(msg)
        this_day = _stamp(when, '%Y-%m-%d')
        if this_day != day:
            day = this_day
            out.append(f'## {day}')
        out.append(_message_block(msg, _stamp(when, '%H:%MZ')))
    return '\n'.join(out) or empty


def _message_block(msg: dict, when: str) -> str:
    text = message_text(msg) or '(empty)'
    first, _, rest = text.partition('\n')
    block = f'{when} {sender(msg)}: {first}'
    return block + ('\n' + _indent(rest) if rest else '')


def posts(items: list[dict]) -> str:
    out = []
    for post in items:
        out.append(f"### {_stamp(item_time(post))} · {sender(post)} · id {post.get('id')}")
        out.append(message_text(post) or '(empty)')
        replies = post.get('replies')
        if replies:
            out.append(f'Replies ({len(replies)}):')
            for reply in sorted(replies, key=lambda r: r.get('createdDateTime') or ''):
                first, _, rest = _message_block(reply, _stamp(item_time(reply))).partition('\n')
                out.append('  - ' + first + ('\n' + _indent(rest, '    ') if rest else ''))
    return '\n'.join(out) or '(no posts)'


def teams(items: list[dict]) -> str:
    return '\n'.join(_named(t) for t in items) or '(no teams)'


def channels(items: list[dict]) -> str:
    return '\n'.join(_named(c, c.get('membershipType')) for c in items) or '(no channels)'


def _named(item: dict, kind: str | None = None) -> str:
    line = f"- **{item.get('displayName') or '(unnamed)'}**"
    if kind and kind != 'standard':
        line += f' ({kind})'
    description = _clip(item.get('description') or '', 100)
    if description and description != item.get('displayName'):
        line += f' — {description}'
    return f"{line}\n  id: {item.get('id')}"


def status(info: Any) -> str:
    if not isinstance(info, dict):
        return str(info)
    return '\n'.join(f'- {k}: {v}' for k, v in info.items())


def search_results(items: list[dict]) -> str:
    out = []
    for hit in items:
        when = _stamp(item_time(hit))
        if hit.get('source') == 'chat':
            where = f"chat **{hit.get('chatName') or hit.get('chatId')}**"
            ref = f"chat: {hit.get('chatId')}"
        else:
            where = f"**{hit.get('teamName') or '?'} / {hit.get('channelName') or '?'}**"
            ref = f"team: {hit.get('teamId')} channel: {hit.get('channelId')} "
            ref += (f"thread: {hit['threadId']}" if hit.get('threadId')
                    else f"message: {hit.get('messageId')}")
        out.append(f"- {when} {where} · {hit.get('from')}: {_clip(hit.get('text') or '', 400)}"
                   f"\n  {ref}")
    return '\n'.join(out) or '(no matches)'


RENDERERS: dict[str, Callable[[Any], str]] = {
    'status': status, 'chats': chats, 'messages': messages, 'teams': teams,
    'channels': channels, 'posts': posts,
    'replies': lambda items: messages(items, '(no replies)'), 'search': search_results,
}
