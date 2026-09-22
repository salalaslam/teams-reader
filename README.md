# teams-reader

A small read-only command interface for Microsoft Teams, built on
[PnP CLI for Microsoft 365](https://pnp.github.io/cli-microsoft365/).
Use it from a terminal or an agent with shell access to list existing chats,
read chat messages, and read team channel posts and replies. Output is JSON.
No AI model, bot gateway, or public webhook is required.

## Requirements

- Linux or macOS with Python 3, Node.js, and npm. The original installation was
  verified with Node.js 24 and CLI for Microsoft 365 11.11.0.
- A Microsoft 365 work account with access to Teams.
- A dedicated Microsoft Entra application with the delegated permissions below.
  Your organization may require administrator approval.

## Configure Microsoft Entra

1. Register an application in your organization's Entra tenant. Prefer a
   single-tenant application for personal or internal use.
2. Under **Authentication**, enable **Allow public client flows** for device-code
   sign-in. No client secret or redirect URI is needed for this flow.
3. Add these Microsoft Graph **delegated** permissions:

   | Permission | Purpose |
   | --- | --- |
   | `Chat.Read` | Read the signed-in user's chats and messages |
   | `ChannelMessage.Read.All` | Read channel messages as the signed-in user |
   | `Channel.ReadBasic.All` | List channel names and descriptions |
   | `Team.ReadBasic.All` | List team names and descriptions |
   | `User.Read` | Identify the signed-in user |
   | `offline_access` | Refresh the authorized session |

4. Have an administrator approve the application if required. See the
   [approval request template](admin-approval.md).
5. Record the application client ID and directory tenant ID for local setup.

Do not grant message send/edit/delete scopes or application permissions.
Delegated access remains limited by the signed-in user's access and tenant
policies. A single-tenant app is not automatically restricted to a single user;
an administrator can configure enterprise application assignment if needed.

## Install

From a clone of this repository:

```sh
mkdir -p "$HOME/.local/share/teams-reader" "$HOME/.local/bin"
npm install --prefix "$HOME/.local/share/teams-reader" \
  --save-exact @pnp/cli-microsoft365@11.11.0
install -m 700 teams-read "$HOME/.local/share/teams-reader/teams-read"
ln -sfn "$HOME/.local/share/teams-reader/teams-read" "$HOME/.local/bin/teams-read"
```

For a new installation, copy `account.example.json` to
`~/.local/share/teams-reader/account.json` and replace both placeholders with your
own IDs. Do not overwrite an existing configured file when upgrading.

```sh
chmod 600 "$HOME/.local/share/teams-reader/account.json"
export PATH="$PATH:$HOME/.local/bin"
teams-read login
```

Open the Microsoft device sign-in URL printed by the command, enter its code,
and sign in with your work account. The OAuth session is stored by Microsoft 365
CLI; the wrapper does not request or store your account password.

Add `~/.local/bin` to the PATH of any shell or agent that will run the tool, or
invoke `~/.local/bin/teams-read` by its full path.

## Usage

```sh
teams-read status
teams-read chats
teams-read messages 'CHAT_ID'
teams-read teams
teams-read channels 'TEAM_ID'
teams-read posts 'TEAM_ID' 'CHANNEL_ID'
teams-read replies 'TEAM_ID' 'CHANNEL_ID' 'MESSAGE_ID'
teams-read logout
```

Use IDs from list output in subsequent commands. List commands may return a
large history; this wrapper currently exposes no date or result-limit flags.
Run `teams-read --help` for command descriptions.

## Sessions and privacy

- The wrapper exposes read commands plus login, status, and logout. Microsoft
  Graph permissions are the actual security boundary; the underlying CLI also
  implements write commands. Use a dedicated app with read-only scopes.
- The wrapper uses the CLI's active connection. Login names its connection
  `teams-reader`; switching connections using the underlying CLI can change the
  account used by subsequent reads. Check `teams-read status` before use.
- Microsoft 365 CLI uses its standard cache in the user's home directory. The
  wrapper sets a restrictive creation mask (`077`) for newly created files.
  Other processes running as the same OS user can use the saved session.
- Session expiry or revoked access may require `teams-read login` again.
  `teams-read logout` logs out the active CLI connection.
- Keep `account.json`, authentication caches, and exported messages out of Git.
  Client and tenant IDs are not passwords, but this repository uses placeholders
  to avoid publishing deployment identifiers. `.gitignore` excludes common local
  configuration/cache files and an optional `exports/` directory.
- This repository contains no account configuration, tokens, or message data.
  Only share message output with systems authorized to receive it.

## Validation

The original deployment successfully read one-to-one, group, and meeting chats,
listed teams and channels, and read channel posts and replies. Those checks were
performed against a private account; no account identifiers or message contents
are included here. Access in another organization depends on its own permissions.
