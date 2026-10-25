# Meeting events (`TEAMS_MEETING_EVENTS`)

With `TEAMS_MEETING_EVENTS=1` (or `true`) the plugin delivers the start and end of a
scheduled or channel meeting in a chat or team where the app is installed. Claude is
told about it, nothing else: the transcript of the meeting is not part of this (that
is a separate change).

What arrives in the session is the text `[Meeting started] <title>` /
`[Meeting ended] <title>` (no title, no suffix), and the meta keys `meeting_event`
(`started` / `ended`), `meeting_title`, `meeting_type`, `meeting_id`, `meeting_time`,
plus the usual `conversation_id` and `tenant_id`. There is no sender to allowlist —
`aad_object_id` is empty — so the event is delivered for the conversation it names and
is never a request: it carries no instruction, and the meeting title is user content,
cleaned like every other meta value (no quotes, angle brackets, control characters or
bidi controls, bounded length).

This is off by default. It needs both:

1. `TEAMS_MEETING_EVENTS` in the plugin's env, and
2. an app package built with `scripts/make-teams-manifest.sh --rsc-meeting-events`,
   which adds the RSC permissions `OnlineMeeting.ReadBasic.Chat` and
   `ChannelMeeting.ReadBasic.Group` (application) with the `webApplicationInfo` block.
   A chat member or the meeting organizer consents when the app is installed; an app
   that is already installed must be upgraded with the new permissions.

`meeting_id` is `channelData.meeting.id` — what the Microsoft documentation says to
use, not the id in the activity's `value` and not the conversation id. The join URL
Teams also sends is not passed on: a join link is not needed to act on the event.

Not measured against a live Teams tenant: the activity shape follows the botbuilder
package types and the Microsoft Learn page on meeting apps APIs.
