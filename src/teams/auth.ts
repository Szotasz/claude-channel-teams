/**
 * The adapter the listener hands requests to (Microsoft 365 Agents SDK, TEAMSSDK1126).
 *
 * Security shape:
 *   - Inbound tokens are NOT checked here. The plugin's own channel-token verifier (channelToken.ts, TEAMSEMUTOKEN1118)
 *     runs in handleMessagesPost before the adapter sees anything, and admits only what the Bot Framework channel
 *     service signs for this bot. The SDK's own `authorizeJWT` middleware is deliberately not used: it admits an
 *     anonymous request when no client id is configured outside production, and it accepts Entra issuers the verifier
 *     refuses on purpose. The verified claims are handed to `process` as `request.user` (adapter.ts).
 *   - The adapter still checks the token's `serviceurl` claim against the activity, and the outbound service URL host
 *     (CloudAdapter.validateServiceUrl / ensureOutboundServiceUrlAllowed).
 *   - Outbound tokens: client credentials of the bot's Entra app in the pinned tenant (SingleTenant only, design F.1).
 *
 * The configuration uses the SDK's names (clientId, clientSecret, tenantId); the mapping from the plugin's config is
 * here so the rest of the code keeps its own names.
 */

import { AuthType, CloudAdapter, type AuthConfiguration } from '@microsoft/agents-hosting'
import type { Config } from '../config.js'

export function agentsAuthConfig(config: Config): AuthConfiguration {
  if (!config.appId || !config.appPassword || !config.tenantId) {
    // The SDK would fall back to an anonymous identity without a client id; never start that way.
    throw new Error('teams channel: the bot app id, password and tenant id are all required')
  }
  return {
    clientId: config.appId,
    clientSecret: config.appPassword,
    tenantId: config.tenantId,
    authType: AuthType.ClientSecret,
  }
}

/** Build the `CloudAdapter` for SingleTenant operation. */
export function createCloudAdapter(config: Config): CloudAdapter {
  const adapter = new CloudAdapter(agentsAuthConfig(config))
  // The SDK's default onTurnError posts "The agent encountered an error or bug." into the user's chat. botbuilder had
  // no default: the error went back to the caller (adapter.process -> handleMessagesPost logs it and answers 500).
  // Keep that: nothing the plugin did not write ever appears in a chat.
  adapter.onTurnError = async (_context, error) => {
    throw error
  }
  return adapter
}
