# teams-reader

**Read-only Microsoft Teams access for AI agents: a CLI and MCP server that reads your chats and channels with your own delegated permissions.**

[![test](https://github.com/salalaslam/teams-reader/actions/workflows/test.yml/badge.svg)](https://github.com/salalaslam/teams-reader/actions/workflows/test.yml)

teams-reader lists your chats, reads messages, channel threads and replies, and searches
your Teams history. It outputs JSON or compact Markdown sized for an LLM context window.
It is a small Python layer over [CLI for Microsoft 365](https://pnp.github.io/cli-microsoft365/)
(`m365`), which handles sign-in and calls Microsoft Graph. You can use it from a shell, from any
agent that can run commands, or as a stdio MCP server.

## Why

Agents are useful for catching up on Teams ("what did I miss in the launch chat?", "find the
thread where we agreed the budget"). The existing options have costs that teams-reader avoids:

- **Microsoft's Work IQ Teams MCP server** is tied to Microsoft 365 Copilot licensing (or
  Work IQ usage-based billing). Its tools include posting, editing and deleting messages,
  and it has no read-only mode.
- **Anthropic's Microsoft 365 connector for Claude** needs a Global Administrator to consent
  for the whole tenant before anyone can use it.
- **teams-reader** is read-only by construction and runs with *delegated* permissions: it can
  read only what you can already read, as you, from your own machine. No bot, webhook, public
  endpoint or AI subscription is involved.

It still needs a Microsoft Entra app registration in your tenant, and some of its permissions
need an administrator's approval. See [Tenant requirements](#tenant-requirements) before you
start.

## Install

Install CLI for Microsoft 365 (needs Node.js LTS) and teams-reader (needs Python 3.11 or later):

```sh
npm install -g @pnp/cli-microsoft365

# Run without installing
uvx --from git+https://github.com/salalaslam/teams-reader teams-reader --help

# Or install it
pipx install git+https://github.com/salalaslam/teams-reader
uv tool install git+https://github.com/salalaslam/teams-reader
```

The MCP server needs the optional `mcp` extra:

```sh
pipx install 'teams-reader[mcp] @ git+https://github.com/salalaslam/teams-reader'
```

teams-reader looks for `m365` on `PATH`. To use another copy, set `TEAMS_READER_M365` to its
path.

## Set up

1. **Register an app** in the [Microsoft Entra admin center](https://entra.microsoft.com) →
   App registrations → New registration. Choose *Accounts in this organizational directory
   only* (single tenant). No redirect URI is needed.
2. Under **Authentication**, set **Allow public client flows** to *Yes*. This enables
   device-code sign-in, so no client secret is needed.
3. Under **API permissions**, add the [Microsoft Graph delegated permissions](#permissions)
   below, then get them consented (see [Tenant requirements](#tenant-requirements)).
4. Save the app's client ID and your tenant ID in `~/.config/teams-reader/account.json`:

   ```json
   { "appId": "00000000-0000-0000-0000-000000000000", "tenant": "11111111-1111-1111-1111-111111111111" }
   ```

   Alternatively, pass `--app-id` and `--tenant` to `login`, or set `TEAMS_READER_APP_ID` and
   `TEAMS_READER_TENANT`. These IDs are identifiers, not secrets.
5. Sign in:

   ```sh
   teams-reader login    # prints a microsoft.com/devicelogin code
   teams-reader status
   ```

`m365` stores the session in its token cache in your home directory. teams-reader never
sees your password. Run `teams-reader login` again when Microsoft expires the session, and
`teams-reader logout` to end it.

## Usage

```text
teams-reader chats                                   list chats, most recently active first
teams-reader messages CHAT_ID                        read a chat, newest first
teams-reader teams                                   list your teams
teams-reader channels TEAM_ID                        list a team's channels
teams-reader posts TEAM_ID CHANNEL_ID [--replies]    read channel threads, most recently active first
teams-reader replies TEAM_ID CHANNEL_ID MESSAGE_ID   read replies to a channel post
teams-reader search TEXT [--scan]                    search chat and channel messages
teams-reader sync                                    archive new chat messages locally (run on a timer)
teams-reader recent [--project P] [--chat C]         read the local archive, without calling Graph
teams-reader wait CHAT... | --project P [--from N]   wait until someone else posts
teams-reader mcp                                     run the MCP server on stdio
teams-reader status | login | logout
```

All read commands accept these options:

- `--format json|md` (`-f`). `json` (the default) returns Microsoft Graph objects as Graph
  sends them. `md` returns compact text: HTML is stripped, mentions become `@Name`, cards
  keep their text, files and quoted replies become one-line placeholders, and system events
  (members added, calls, renames) and deleted messages are dropped. On real chats it came
  out 3 to 16 times smaller than the JSON.
- `--since WHEN` takes a relative time (`30m`, `12h`, `7d`, `2w`) or an ISO date or
  date-time (`2026-10-01`, `2026-10-01T09:00:00Z`). A date or date-time without a
  timezone is local time. Output times are local too, labelled with the zone; set `TZ`
  (for example `TZ=UTC`) to see another one.
- `--limit N` (aliases `--top`, `-n`) caps the results. The default is 50 (25 for search);
  `0` means no limit. Results are fetched page by page, and paging stops once the limit or
  the `--since` cut-off is reached, so a small limit is fast even on a long chat.

The examples below use made-up data, shown with `TZ=UTC`.

```console
$ teams-reader chats --since 2d -f md
- 2026-10-02 13:48 UTC **Launch plan** (group) — Priya Shah: Final checklist is in the channel, please review by Friday
  id: 19:3f2a9c1e5b7d4e0f8a6b2c4d1e9f7a3b@thread.v2
- 2026-10-02 11:38 UTC **Sam Lee, Alex Doe** (oneOnOne) — Sam Lee: sounds good
  id: 19:0b6c...@unq.gbl.spaces
```

```console
$ teams-reader messages '19:3f2a9c1e5b7d4e0f8a6b2c4d1e9f7a3b@thread.v2' -n 4 -f md
## 2026-10-02
12:41 UTC Alex Doe: Can we move the demo to Thursday?
12:45 UTC Priya Shah: Thursday works. @Sam Lee can you update the invite?
12:52 UTC Sam Lee: > replying to Priya Shah: Thursday works. @Sam Lee can you update the invite?
  Done, and I attached the run sheet
  [file: demo-run-sheet.docx]
13:48 UTC Priya Shah: Final checklist is in the channel, please review by Friday
```

```console
$ teams-reader search budget --since 30d -f md
- 2026-09-30 16:05 UTC **Finance / Planning** · Sam Lee: Q4 budget is approved, see the updated sheet...
  team: 6f1d... channel: 19:a1b2...@thread.tacv2 message: 1727712300000
- 2026-09-24 09:12 UTC chat **Launch plan** · Alex Doe: ...need sign-off on the launch budget before...
  chat: 19:3f2a9c1e5b7d4e0f8a6b2c4d1e9f7a3b@thread.v2
```

```console
$ teams-reader messages '19:3f2a9c1e5b7d4e0f8a6b2c4d1e9f7a3b@thread.v2' -n 1
[
  {
    "id": "1727873280000",
    "messageType": "message",
    "createdDateTime": "2026-10-02T13:48:00.000Z",
    "from": { "user": { "displayName": "Priya Shah", "id": "..." } },
    "body": { "contentType": "html", "content": "<p>Final checklist is in the channel, please review by Friday</p>" },
    "attachments": [],
    ...
  }
]
```

### Search

By default, `search` uses Microsoft Search (`m365 search --scopes chatMessage`). It searches
your whole chat and channel history in one call and returns a snippet per message. Queries
use [KQL](https://learn.microsoft.com/sharepoint/dev/general-development/keyword-query-language-kql-syntax-reference):
words and prefixes (`deploy*`), `"exact phrases"`, `AND`/`OR`.

`search --scan` reads recent messages directly and matches `TEXT` as a case-insensitive
substring of the full text. Use it for fragments Microsoft Search won't match, such as part
of a word or an ID. It covers the 25 most recently active chats (`--chats N`) and every
channel of every team you belong to, within `--since` (default `7d`). It is slower: each
chat or channel is a separate request.

### Local archive, notifications and waiting

`teams-reader sync` lists your chats by latest message and fetches only the ones that
changed since its last run, appending new messages to one JSON Lines file per chat in
`~/.local/state/teams-reader` (or `$XDG_STATE_HOME/teams-reader`), readable only by you.
The first run archives the last 7 days (`--since` changes that). After that, a run with
nothing new is one Graph request. Run it on a timer with the systemd units in
[contrib/systemd](contrib/systemd):

```sh
cp contrib/systemd/teams-reader-sync.* ~/.config/systemd/user/
systemctl --user daemon-reload && systemctl --user enable --now teams-reader-sync.timer
```

`teams-reader recent` reads the archive with no network call, newest first, and groups the
Markdown output by chat. It warns on stderr when the last sync is over 15 minutes old or
failed. Filter with `--since`, `--chat` (an ID or a Teams link), `--project` and `--from`.

`teams-reader wait` polls chats live (every 60 seconds by default) and exits as soon as
someone other than you posts, printing the new messages. It exits with status 3 after
`--timeout` (default `8h`). An agent can run it in the background and pick the
conversation up when it returns. `--since` counts messages already sent after a given time.

Commands that take a chat ID also accept a Teams link to the chat or to one of its messages.

Optional `~/.config/teams-reader/projects.toml` names groups of chats and controls
notifications. After the first sync, `sync` can publish to an
[ntfy](https://ntfy.sh) topic for direct messages, @mentions of you, chosen senders, and
every message in chosen projects' chats:

```toml
[notify]
ntfy = "http://127.0.0.1:8090/teams"  # server URL and topic; omit to disable
people = ["Priya Shah"]               # always notify for these senders
direct = true                         # direct messages (default true)
mentions = true                       # @mentions of you (default true)

[projects.launch]
chats = ["19:3f2a9c1e5b7d4e0f8a6b2c4d1e9f7a3b@thread.v2"]
notify = true                         # every message in these chats
```

Notifications contain message text, so use a self-hosted ntfy server or a topic only you
know. teams-reader ignores keys it doesn't use, so the same file can also list each
project's repos and people for agents. A failing sync (for example an expired sign-in)
sends one notification and keeps failing quietly until it recovers.

## MCP server

`teams-reader mcp` runs a stdio MCP server. Its tools are `list_chats`, `read_chat`,
`list_teams`, `list_channels`, `read_channel`, `read_thread_replies`, `search_messages`
and `teams_status`. Every tool is annotated `readOnlyHint: true`. There are no tools for
sending, editing, deleting, signing in or signing out. Tools return Markdown by default
and accept `format: "json"`. Sign in once with `teams-reader login` before starting the
server.

**Claude Code**

```sh
claude mcp add teams-reader -- uvx --from 'teams-reader[mcp] @ git+https://github.com/salalaslam/teams-reader' teams-reader mcp
```

**Codex**

```sh
codex mcp add teams-reader -- uvx --from 'teams-reader[mcp] @ git+https://github.com/salalaslam/teams-reader' teams-reader mcp
```

or in `~/.codex/config.toml`:

```toml
[mcp_servers.teams-reader]
command = "uvx"
args = ["--from", "teams-reader[mcp] @ git+https://github.com/salalaslam/teams-reader", "teams-reader", "mcp"]
```

If you installed with `pipx` or `uv tool`, use `teams-reader mcp` as the command instead.

## Tenant requirements

### You need your own Entra app

CLI for Microsoft 365 used to sign in through a shared multi-tenant app, the *PnP Management
Shell*. PnP retired it and
[deleted it on 9 September 2024](https://pnp.github.io/blog/post/changes-pnp-management-shell-registration/).
Since then, `m365` must sign in with an app registered in your own tenant. A single-tenant
app is the normal choice, and that is what teams-reader expects.

### Permissions

All permissions are Microsoft Graph **delegated** permissions. No application permissions,
client secrets or write scopes are involved.

| Permission | Used for | Admin consent required? |
| --- | --- | --- |
| `Chat.Read` | `chats`, `messages`, chat results in `search` (`GET /me/chats`, `GET /chats/{id}/messages`) | No, but see below |
| `ChannelMessage.Read.All` | `posts`, `replies`, channel results in `search` (`GET /teams/{id}/channels/{id}/messages[/{id}/replies]`) | **Yes, always** |
| `Team.ReadBasic.All` | `teams`, team names (`GET /me/joinedTeams`) | No |
| `Channel.ReadBasic.All` | `channels`, channel names (`GET /teams/{id}/channels`) | No |
| `User.Read` | `sync` and `wait`, to tell your own messages and @mentions apart (`GET /me`) | No |

`offline_access` (to refresh the session) is requested automatically at sign-in and needs no
admin consent. `User.Read` is only needed by `sync` and `wait`; new app registrations
include it by default.

Do not add `ChatMessage.Send`, any `*.ReadWrite*` permission, or any application permission.
The Graph permissions are the real security boundary, because `m365` itself can also write.
teams-reader only ever runs `m365 request --method get`, `m365 search` and `m365 status`, and
it builds Graph URLs from fixed paths with every ID percent-encoded.

### If you are not an administrator

Whether you can grant these permissions yourself depends on your tenant's
[user consent settings](https://learn.microsoft.com/entra/identity/enterprise-apps/configure-user-consent):

- **Registering the app.** By default, members can register applications. Many organisations
  turn this off. In that case an administrator has to create the app for you.
- **`ChannelMessage.Read.All`** always needs an administrator. A Global Administrator,
  Application Administrator or Cloud Application Administrator can grant it.
- **`Chat.Read`** doesn't require admin consent in principle. However, the
  [Microsoft-managed default policy](https://learn.microsoft.com/entra/identity/enterprise-apps/manage-app-consent-policies#microsoft-recommended-current-settings)
  ("Let Microsoft manage your consent settings", the default for new tenants) excludes
  `Chat.Read` from user consent. Users can consent to it themselves only if their tenant
  allows user consent for all apps, or allows apps registered in the organisation and
  classifies `Chat.Read` as low impact.
- `m365` requests every permission configured on the app at once (the `.default` scope). If
  any of them needs admin consent, you can't sign in at all until an administrator grants
  it. If your tenant does allow user consent, you can make an app with only `Chat.Read`,
  `Team.ReadBasic.All` and `Channel.ReadBasic.All`. You consent to it yourself at the first
  `login`. Chats, teams and channel lists work. Reading channel posts returns an access
  error.

Otherwise, ask an administrator to grant consent for the app. If your tenant has the
[admin consent workflow](https://learn.microsoft.com/entra/identity/enterprise-apps/configure-admin-consent-workflow)
enabled, you can request approval from the sign-in screen. [admin-approval.md](admin-approval.md)
is a request template you can send them. An administrator can also restrict the app to
specific users (Enterprise applications → the app → Properties → *Assignment required*).

## Notes and limits

- Every Graph call starts an `m365` process, which takes roughly 2 to 5 seconds. `search` is
  a single call plus name lookups. `search --scan` makes one call per chat and channel,
  four at a time.
- Graph orders channel threads by their latest activity. `posts --since` therefore keeps a
  thread if the post or any reply is new enough.
- In `md` output, `--limit` counts real messages; system events don't use up the limit.
- `chats` lists at most 25 members per chat (a Graph limit).
- `m365` keeps its token cache in your home directory, readable only by you. teams-reader
  creates files with a `077` umask. Any process running as your OS user can use the
  session.
- Message content is sensitive. Only send output to tools and models you are allowed to
  share it with.

## Upgrading from the `teams-read` script

Earlier versions were a single `teams-read` script in `~/.local/share/teams-reader`. The
package still installs a `teams-read` command. If `m365` isn't on `PATH`, it falls back to
`~/.local/share/teams-reader/node_modules/.bin/m365`, and it reads `account.json` from that
directory when `~/.config/teams-reader/account.json` doesn't exist. Defaults have changed:
list commands now return at most 50 items unless you pass `--limit 0`.

## Development

```sh
uv run --extra mcp pytest
uv build
```

The tests run the CLI and MCP server against a fake `m365` executable that serves canned
Graph responses, so no tenant is needed.

## License

MIT
