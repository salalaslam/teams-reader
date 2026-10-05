# Administrator approval request template

Replace the placeholders before forwarding this request. Keep completed copies
outside the repository or in a `*.local.md` file.

> Please grant admin consent for the **Teams Reader** application so I can read
> my existing Teams chats and channel messages from a terminal or AI agent
> running under my own account. It is read-only and uses delegated permissions,
> so it can only see what I can already see in Teams.
>
> Application (client) ID: `YOUR_ENTRA_APPLICATION_CLIENT_ID`
>
> Directory (tenant) ID: `YOUR_ENTRA_TENANT_ID`
>
> Requested user: `YOUR_WORK_ACCOUNT`

In **Entra admin center → App registrations → the application → API
permissions**, review the permissions and select **Grant admin consent**. A
Global Administrator, Application Administrator or Cloud Application
Administrator can do this.

The requested Microsoft Graph permissions are all **delegated**:

| Permission | Why |
| --- | --- |
| `Chat.Read` | Read the signed-in user's chats and chat messages |
| `ChannelMessage.Read.All` | Read channel posts and replies in the user's teams (requires admin consent) |
| `Team.ReadBasic.All` | List the teams the user belongs to |
| `Channel.ReadBasic.All` | List channel names in those teams |

`offline_access` is requested automatically at sign-in so the session can be
refreshed. `User.Read` may also be listed; new registrations include it by
default.

The client is [teams-reader](https://github.com/salalaslam/teams-reader), which
signs in through CLI for Microsoft 365 using the device-code flow. There are no
application permissions, client secrets, redirect URIs, public webhooks, or
message send/edit/delete permissions. Access remains subject to the signed-in
user's own Teams membership and tenant policies.

To limit the app to the requesting user, open **Enterprise applications → the
application → Properties**, set **Assignment required** to *Yes*, and assign the
user under **Users and groups**. Approving this app does not require relaxing
the tenant's user consent settings.
