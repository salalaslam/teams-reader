# Administrator approval request template

Replace the placeholders before forwarding this request. Keep completed copies
outside the repository or in a `*.local.md` file.

> Please approve the **Teams Reader** application so I can read my existing Teams
> conversations from a terminal or agent running under my OS account.
>
> Application client ID: `YOUR_ENTRA_APPLICATION_CLIENT_ID`
>
> Directory tenant ID: `YOUR_ENTRA_TENANT_ID`
>
> Requested user: `YOUR_WORK_ACCOUNT`

In **Entra → App registrations → the application → API permissions**, review
and grant consent according to your organization's policy.

The requested Microsoft Graph permissions are all **delegated**:

- `Chat.Read`
- `ChannelMessage.Read.All`
- `Channel.ReadBasic.All`
- `Team.ReadBasic.All`
- `User.Read`
- `offline_access`

The client uses interactive OAuth device-code sign-in with CLI for Microsoft
365. There are no application permissions, client secrets, public webhooks, or
message send/edit/delete scopes. Access remains subject to the signed-in user's
permissions and tenant policies. Restrict enterprise application assignment to
the intended user if required. Approve this app without disabling tenant consent
restrictions. The user will complete sign-in after approval.
