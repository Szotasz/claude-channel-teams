# Planner tasks (`create_planner_task`)

The `create_planner_task` MCP tool creates a Microsoft Planner task through
Microsoft Graph, using the bot's own app-only token (client credentials). It
can set a title, bucket, assignees, due date, priority, description, a link
back to the Teams thread, and a checklist. It returns `{"task_id"}`.

## Turning it on

The tool is **off** unless `TEAMS_PLANNER_PLAN_IDS` lists plan ids
(comma-separated, e.g. `TEAMS_PLANNER_PLAN_IDS=planA,planB`). Each id must match
`[A-Za-z0-9_-]{1,64}`. If any entry is invalid, the plugin writes one line to
stderr and the tool stays off. With exactly one configured plan, `plan_id` may
be omitted; with several it is required. A plan not on the list is refused.

## Permission

The bot's app registration needs the Graph **application** permission
`Tasks.ReadWrite.All`, with admin consent. This permission covers **every plan
in the tenant**, which is why the plan allow-list exists: the tool only writes
to the plans you list here.

## Finding a plan id

Call `GET /groups/{group-id}/planner/plans` in Graph: each plan in the answer
has its `id`.

## Arguments and limits

- `title` (required), `bucket` (bucket id or exact name; an unknown bucket
  creates nothing).
- `assignee_ids`: AAD object ids, at most 20. The `aadObjectId` of an inbound
  Teams message sender works.
- `due_date`: `YYYY-MM-DD`. It is stored as 12:00 UTC, so the calendar day is
  the same from UTC-11 to UTC+11.
- `priority`: `urgent`, `important`, `medium` or `low`.
- `description`, `link` (https; appended as `Thread: <link>`), `checklist`
  (at most 20 items, kept in the given order).

If the task is created but setting its description or checklist fails, the
error says so and includes the `task_id`.

## Security notes

The client secret goes only to the token endpoint. Tokens and secrets are never
logged or put in an error, errors from Graph carry only the HTTP status and
error code, and no request follows a redirect or a link outside
`https://graph.microsoft.com/v1.0/`.

The token helper duplicates the one in the SharePoint PR (#13) and can be
merged with it later.
