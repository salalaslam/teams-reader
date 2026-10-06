"""Parsing for --since values and Graph timestamps."""
from __future__ import annotations

import re
from datetime import datetime, timedelta, timezone

_RELATIVE = re.compile(r'^\s*(\d+)\s*([mhdw])\s*$', re.IGNORECASE)
_UNITS = {'m': 'minutes', 'h': 'hours', 'd': 'days', 'w': 'weeks'}
_FRACTION = re.compile(r'\.(\d+)')
_DURATION = re.compile(r'^\s*(\d+(?:\.\d+)?)\s*([smhd])\s*$', re.IGNORECASE)
_DURATION_UNITS = {'s': 'seconds', 'm': 'minutes', 'h': 'hours', 'd': 'days'}


def parse_since(value: str | None, now: datetime | None = None) -> datetime | None:
    """Parse ``7d``/``12h``/``30m``/``2w`` or an ISO date/date-time into UTC.

    ISO values without a timezone are taken as local time.
    """
    if value is None or value == '':
        return None
    now = now or datetime.now(timezone.utc)
    match = _RELATIVE.match(value)
    if match:
        amount, unit = int(match.group(1)), match.group(2).lower()
        return now - timedelta(**{_UNITS[unit]: amount})
    try:
        parsed = parse_time(value.strip())
    except ValueError:
        raise ValueError(
            f'invalid --since value {value!r}: use e.g. 7d, 12h, 30m, 2w, '
            '2026-10-01 or 2026-10-01T09:00:00Z') from None
    return parsed


def parse_duration(value: str) -> timedelta:
    """Parse ``90s``/``30m``/``8h``/``2d`` into a timedelta."""
    match = _DURATION.match(value)
    if not match:
        raise ValueError(f'invalid duration {value!r}: use e.g. 90s, 30m, 8h or 2d')
    amount, unit = float(match.group(1)), match.group(2).lower()
    return timedelta(**{_DURATION_UNITS[unit]: amount})


def parse_time(value: str) -> datetime:
    """Parse an ISO 8601 timestamp as Graph returns it (any fraction length)."""
    text = value.strip()
    if text.endswith(('Z', 'z')):
        text = text[:-1] + '+00:00'
    # Python < 3.11 only accepts 3 or 6 fractional digits; Graph emits 1-7.
    text = _FRACTION.sub(lambda m: '.' + (m.group(1) + '000000')[:6], text, count=1)
    # astimezone() reads a naive value as local time; Graph always sends Z.
    return datetime.fromisoformat(text).astimezone(timezone.utc)


def item_time(item: dict | None, key: str = 'createdDateTime') -> datetime | None:
    raw = (item or {}).get(key)
    if not raw:
        return None
    try:
        return parse_time(raw)
    except ValueError:
        return None
