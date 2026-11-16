/**
 * `create_channel` implementation: the opt-in tool that creates a standard
 * channel in a team through Microsoft Graph with the bot's app-only token.
 * Standard channels only (no private or shared channels).
 *
 * Permission: the resource-specific consent (RSC) permission
 * Channel.Create.Group, declared in the app manifest
 * (scripts/make-teams-manifest.sh --rsc-channel-create) and consented by the
 * team owner. The tool only writes to the teams the operator lists in
 * TEAMS_CHANNEL_CREATE_TEAMS (docs/create-channel.md).
 */
import { graph } from './planner.js'

export const MAX_CHANNEL_NAME = 50
export const MAX_CHANNEL_DESCRIPTION = 1024

export interface CreateChannelInput {
  teamId: string
  name: string
  description?: string
}

/** Checks the tool arguments; throws on the first problem. No network. */
export function validateCreateChannelArgs(
  args: Record<string, unknown>,
  allowedTeams: ReadonlySet<string>,
): CreateChannelInput {
  if (allowedTeams.size === 0) throw new Error('create_channel is off (TEAMS_CHANNEL_CREATE_TEAMS is not set)')
  let teamId: string
  const rawTeam = args.team_id
  if (rawTeam === undefined || rawTeam === null) {
    if (allowedTeams.size !== 1) throw new Error('team_id is required when more than one team is configured')
    teamId = [...allowedTeams][0]!
  } else if (typeof rawTeam === 'string' && allowedTeams.has(rawTeam.toLowerCase())) {
    teamId = rawTeam.toLowerCase()
  } else {
    throw new Error('team not allowed')
  }
  const rawName = args.name
  if (typeof rawName !== 'string' || rawName.trim().length === 0) throw new Error('name is required')
  const name = rawName.trim()
  if (name.length > MAX_CHANNEL_NAME) throw new Error(`name can be at most ${MAX_CHANNEL_NAME} characters`)
  if (/[\u0000-\u001f\u007f-\u009f]/.test(name)) throw new Error('name must not contain control characters')
  const rawDesc = args.description
  let description: string | undefined
  if (rawDesc !== undefined && rawDesc !== null) {
    if (typeof rawDesc !== 'string') throw new Error('description must be a string')
    const desc = rawDesc.trim()
    if (desc.length > MAX_CHANNEL_DESCRIPTION) {
      throw new Error(`description can be at most ${MAX_CHANNEL_DESCRIPTION} characters`)
    }
    // A description may span lines: tab, line feed and carriage return are allowed, every other control character is not.
    if (/[\u0000-\u0008\u000b\u000c\u000e-\u001f\u007f-\u009f]/.test(desc)) {
      throw new Error('description must not contain control characters other than tab and line breaks')
    }
    if (desc.length > 0) description = desc
  }
  return { teamId, name, ...(description ? { description } : {}) }
}

export async function createChannel(
  input: CreateChannelInput,
  deps: { getToken: () => Promise<string>; fetch?: typeof fetch },
): Promise<{ channelId: string; webUrl?: string }> {
  const json = await graph(
    deps.getToken,
    deps.fetch,
    'POST',
    `https://graph.microsoft.com/v1.0/teams/${encodeURIComponent(input.teamId)}/channels`,
    {
      displayName: input.name,
      ...(input.description ? { description: input.description } : {}),
      membershipType: 'standard',
    },
  )
  const obj = typeof json === 'object' && json !== null && !Array.isArray(json) ? (json as Record<string, unknown>) : {}
  const id = obj.id
  if (typeof id !== 'string' || id.length === 0) throw new Error('Graph returned no channel id')
  const webUrl = obj.webUrl
  return { channelId: id, ...(typeof webUrl === 'string' && webUrl.startsWith('https://') ? { webUrl } : {}) }
}

export const CHANNEL_TOOL = {
  name: 'create_channel',
  description:
    "Creates a standard channel in one of the teams the operator allowed (TEAMS_CHANNEL_CREATE_TEAMS). Give team_id when more than one team is allowed. name: at most 50 characters. Returns the new channel's id (channel_id) and, when Teams gives one, its link (web_url).",
  inputSchema: {
    type: 'object',
    properties: {
      team_id: {
        type: 'string',
        description: 'The team (Microsoft 365 group id). Optional when only one team is allowed.',
      },
      name: { type: 'string', maxLength: 50, description: 'The channel name.' },
      description: { type: 'string', maxLength: 1024, description: 'Optional channel description.' },
    },
    required: ['name'],
  },
}

export function channelTools(teams: ReadonlySet<string>): (typeof CHANNEL_TOOL)[] {
  return teams.size > 0 ? [CHANNEL_TOOL] : []
}
