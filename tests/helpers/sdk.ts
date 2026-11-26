/**
 * Test-side stand-ins for the botbuilder classes the tests used to build turns with (TEAMSSDK1126).
 *
 * `BotAdapter`: the tests subclass it with a recording sendActivities / updateActivity / deleteActivity, exactly as
 * they subclassed botbuilder's; the Agents SDK's TurnContext calls only those on the adapter.
 * `TurnContext`: the Agents SDK's. It calls Activity instance methods on its activity (getConversationReference when
 * it sends), so the test's plain object gets the Activity prototype, WITHOUT Activity.fromObject's validation: the
 * handler's own guards for malformed activities are under test here, and the Graph backfill path hands the handler
 * activities that never went through the SDK.
 */
import { Activity } from '@microsoft/agents-activity'
import { TurnContext as AgentsTurnContext } from '@microsoft/agents-hosting'

export type { Activity as ActivityType, Attachment, ConversationReference } from '../../src/teams/sdk'
export { TeamsInfo } from '../../src/teams/sdk'
export type Activity = import('@microsoft/agents-activity').Activity

export abstract class BotAdapter {
  abstract sendActivities(context: unknown, activities: unknown[]): Promise<unknown>
  abstract updateActivity(context: unknown, activity: unknown): Promise<unknown>
  abstract deleteActivity(context: unknown, reference: unknown): Promise<unknown>
  abstract continueConversation(...args: unknown[]): Promise<unknown>
}

export class TurnContext extends AgentsTurnContext {
  constructor(adapter: unknown, activity: unknown) {
    super(adapter as never, activity instanceof Activity ? activity : Object.setPrototypeOf(activity as object, Activity.prototype))
  }
}
