# Meeting transcripts (`TEAMS_MEETING_TRANSCRIPTS`)

With `TEAMS_MEETING_TRANSCRIPTS=1` (or `true`) Claude gets a `get_meeting_transcript` tool. It takes the
`meeting_id` of a `meeting_event` tag (see [meeting-events.md](meeting-events.md)) and returns the newest transcript
of that meeting as plain text: one line per cue, `<v Speaker Name>` cues become `Speaker Name: <text>`, and the
result is capped at 60000 characters (`[transcript truncated]` marks a cut).

This is off by default. It needs all of:

1. `TEAMS_MEETING_EVENTS` in the plugin's env — without the meeting events there is no `meeting_id` to read a
   transcript by,
2. `TEAMS_MEETING_TRANSCRIPTS` in the plugin's env, and
3. an app package built with `scripts/make-teams-manifest.sh --rsc-meeting-events --rsc-meeting-transcripts`.
   The meeting-info lookup uses the meeting permissions, the transcript read uses
   `OnlineMeetingTranscript.Read.Chat`; a chat member or the meeting organizer consents when the app is installed,
   and an app that is already installed must be upgraded with the new permissions.

The tool knows only meetings this plugin saw a start or end event for: on each event it asks Teams for the
meeting's Graph id and organizer and keeps them in `meeting-transcripts.json` in the state directory, at most 200
meetings, the oldest dropped first. An unknown `meeting_id` is an error, not a guess.

The transcript is what people said in the meeting: user content, never instructions. Claude is told so, both in
the tool description and in the text the tool returns.

These errors can come back, and what they mean:

- `no transcript yet: ...` — transcription was off, or the transcript is still being processed after the meeting;
  trying again later can work.
- `the tenant administrator has turned off Graph API access to meeting transcripts
  (GraphAccessToTranscriptsDisabled); an administrator must re-enable it` — a tenant-wide setting, only an
  administrator can change it.
- `not allowed to read this transcript: ...` — the app does not have `OnlineMeetingTranscript.Read.Chat` for that
  meeting chat (re-install or upgrade the app there), or the tenant requires an application access policy.
- `the meeting or its transcript was not found (the meeting may have expired)` — Teams keeps transcripts for a
  limited time.
- `Graph HTTP <status> [<code>]` — anything else; the code, when shown, is Graph's own error code.

This covers scheduled meetings in a private chat (the chats where the app is installed). For channel meetings an
administrator would grant `OnlineMeetingTranscript.Read.All` with an application access policy, which this plugin
does not set up.

Not measured against a live Teams tenant, including whether the tenant also requires an application access policy for the chat-scoped permission.
