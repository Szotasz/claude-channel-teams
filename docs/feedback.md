# Feedback buttons (`TEAMS_FEEDBACK`)

Set `TEAMS_FEEDBACK=1` (or `true`) and every reply gets Teams' thumbs up / down buttons. A click from an allowlisted
person reaches Claude as a channel event marked `feedback="like"` or `feedback="dislike"`, with `message_id` set to the
rated message and the person's optional comment as the text (at most 2000 characters).

The plugin answers every click (Teams expects it) but delivers only clicks from allowlisted people in the configured
tenant. A feedback click never starts a pairing and is never treated as a permission reply.

The plugin does not store feedback; what to do with it is up to the agent.

Not measured against a live Teams: the wire shape follows the Microsoft Teams SDK (addFeedback, message/submitAction)
and documentation.
