# Creating channels (`create_channel`)

An opt-in MCP tool that creates a standard channel in a team through Microsoft
Graph. Standard channels only, and only in the teams the operator lists.

## Setup

1. Set `TEAMS_CHANNEL_CREATE_TEAMS` to the team's Microsoft 365 group id(s),
   comma-separated. Each is a GUID, the same id Graph uses in `/teams/{team-id}`.
   Without it the tool is off.
2. Build the manifest with `scripts/make-teams-manifest.sh --rsc-channel-create`
   and install or upgrade the app in the team, so the team owner consents to
   `Channel.Create.Group`.

## Arguments

- `team_id`: one of the listed teams; optional when only one is listed.
- `name`: trimmed, 1-50 characters (counted as UTF-16 code units, so an emoji
  counts twice), no control characters. Characters that Teams itself refuses in
  channel names are not checked here; Graph refuses them and the error comes
  back to Claude.
- `description`: optional, at most 1024 characters after trimming; tab and line
  breaks are allowed, other control characters are not.

The tool uses the same app credentials as the Planner tool (a client-credentials
token for graph.microsoft.com). The call is made with the app's permission, not
a person's.

Not measured against a live Teams: the request and the permission follow the
Microsoft Graph and Teams documentation; private and shared channels are not
supported here.
