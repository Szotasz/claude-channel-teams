/**
 * Environment / config loader for the Teams channel plugin.
 *
 * Phase 2: real implementation.
 *
 * Responsibilities:
 *   1. Resolve the state directory:
 *      - explicit env (`TEAMS_PLUGIN_STATE_DIR`, or legacy `TEAMS_STATE_DIR`),
 *      - then XDG_DATA_HOME (`$XDG_DATA_HOME/claude/channels/teams`),
 *      - then `~/.local/share/claude/channels/teams` on Linux,
 *      - then `~/.claude/channels/teams` as the fallback / Telegram-aligned default.
 *   2. Load `<stateDir>/.env` into process.env, real env winning.
 *   3. `chmod 0600` the .env so credentials are owner-only (no-op on Windows).
 *   4. Validate required keys.
 *   5. Return a frozen Config object.
 *
 * Failure to load a required key writes a helpful diagnostic to stderr and
 * the caller is expected to exit non-zero. We don't call process.exit here so
 * the function stays testable.
 */

import { chmodSync, mkdirSync, readFileSync } from 'fs'
import { homedir } from 'os'
import { join } from 'path'

/** A comma-separated id list from an env var: trimmed, empty entries dropped. */
export function parseIdList(raw: string | undefined): ReadonlySet<string> {
  return new Set((raw ?? '').split(',').map(s => s.trim()).filter(s => s.length > 0))
}

/**
 * Planner plan ids from an env var (comma-separated, like parseIdList). Every
 * entry must look like a Graph plan id; one bad entry makes the whole list
 * invalid (null) so a typo cannot silently widen or shrink the allow-list.
 */
export function parsePlannerPlanIds(raw: string | undefined): ReadonlySet<string> | null {
  const ids = parseIdList(raw)
  for (const id of ids) {
    if (!/^[A-Za-z0-9_-]{1,64}$/.test(id)) return null
  }
  return ids
}

/** True only when TEAMS_TARGETED_MESSAGES is '1' or 'true' (case-insensitive, trimmed). */
export function targetedMessagesEnabled(env: NodeJS.ProcessEnv): boolean {
  const v = env.TEAMS_TARGETED_MESSAGES?.trim().toLowerCase()
  return v === '1' || v === 'true'
}

export interface Config {
  /**
   * Shared conversations (channel or group chat) in which an allowlisted
   * sender's message is delivered without an @mention of the bot. Entries are
   * conversation ids without any ';messageid=...' suffix; '*' means every
   * shared conversation. Empty (the default) keeps the mention requirement
   * everywhere. From TEAMS_UNMENTIONED_CONVERSATIONS (comma-separated).
   */
  unmentionedConversations?: ReadonlySet<string>
  /** Persistent state dir — credentials, allowlist, future inbox. Mode 0700. */
  stateDir: string
  /** Path to `<stateDir>/allowlist.json`. */
  allowlistFile: string
  /** Path to `<stateDir>/pending.json` — pending pairings (Phase 3). */
  pendingFile: string
  /** Bot app (client) ID GUID. */
  appId: string
  /** Bot app client secret. */
  appPassword: string
  /** Always SingleTenant in v1 (design decision F.1). */
  appType: 'SingleTenant'
  /** Tenant GUID — required because we ship single-tenant only. */
  tenantId: string
  /** Address to bind the local HTTP listener (loopback by default). */
  bindHost: string
  /**
   * Use Teams targeted messages for the permission prompt and the pairing code
   * in shared conversations. Off by default; env `TEAMS_TARGETED_MESSAGES=1`.
   */
  targetedMessages: boolean
  /** Port for the local HTTP listener. */
  port: number
  /**
   * Where downloaded Teams attachments are saved. Configurable via the
   * `RECEIVED_FILES_DIR` env var. Defaults to `<stateDir>/files/received-files`
   * (under the sendable root so received attachments can be re-sent).
   */
  receivedFilesDir: string
  /**
   * Where outbound files are staged before being downloaded by Teams.
   * Wiped on plugin start so orphans from a previous boot don't accumulate.
   * Configurable via `OUTBOX_DIR`. Must be a child of `sendableFilesRoot`.
   */
  outboxDir: string
  /** TTL for outbound download tokens, in seconds. Default 1800 (30 min). */
  outboxTtlSeconds: number
  /**
   * Root directory under which `send_file` is allowed to source files from.
   * A path argument outside this root is refused — defence against a
   * prompt-injected Claude being talked into exfiltrating system files.
   * Configurable via `SENDABLE_FILES_ROOT`. Default `<stateDir>/files` (a SUBDIR,
   * so the state-dir secrets .env/access.json stay outside the send boundary).
   */
  sendableFilesRoot: string
  /**
   * Public base URL the `send_file` download links are built on, e.g.
   * `https://<host>.<tailnet>.ts.net/files`. The token gets appended.
   *
   * Derived from THIS install's own public endpoint so the link reaches the
   * plugin's `/files/<token>` route through the same tunnel Azure already
   * uses: `TEAMS_FILES_BASE_URL` when set, else the origin of
   * `TEAMS_BOT_ENDPOINT_URL` (which setup-azure-bot.sh writes) + `/files`.
   * `null` when neither yields an https URL; `send_file` then refuses
   * instead of minting a token behind a link that cannot work.
   */
  filesBaseUrl: string | null
  /**
   * Planner plans the `create_planner_task` tool may write to. Empty (the
   * default) keeps the tool off. From TEAMS_PLANNER_PLAN_IDS (comma-separated);
   * see docs/planner.md.
   */
  plannerPlanIds: ReadonlySet<string>
}

/**
 * Resolve where state lives. We prefer the explicit env var; otherwise fall
 * back to the same path the official Telegram plugin uses (~/.claude/channels/<name>)
 * so operators familiar with that layout get it for free.
 *
 * Documented in docs/installation.md.
 */
export function resolveStateDir(env: NodeJS.ProcessEnv = process.env): string {
  if (env.TEAMS_PLUGIN_STATE_DIR) return env.TEAMS_PLUGIN_STATE_DIR
  // Legacy alias, mirrors design.md naming.
  if (env.TEAMS_STATE_DIR) return env.TEAMS_STATE_DIR
  // Anthropic-published channel plugins (Telegram, Discord, iMessage) all
  // resolve to ~/.claude/channels/<name>. Stay consistent.
  return join(homedir(), '.claude', 'channels', 'teams')
}

/**
 * Read `<stateDir>/.env` and merge it into process.env. Real env wins so
 * operators can override per-launch.
 *
 * Security: the `.env` is forced to mode 0600 on every load. A loose mode
 * would mean any other user on the host can read the bot's client secret —
 * a credential equivalent to bot-impersonation.
 */
function loadEnvFile(envFile: string): void {
  let raw: string
  try {
    // Force owner-only perms before reading.
    // No-op on Windows (would need ACLs); harmless if file doesn't exist (we
    // skip via the catch below).
    chmodSync(envFile, 0o600)
    raw = readFileSync(envFile, 'utf8')
  } catch (err) {
    // Missing file is fine — operator may have set everything via real env.
    if ((err as NodeJS.ErrnoException).code === 'ENOENT') return
    process.stderr.write(
      `teams channel: failed to read ${envFile}: ${(err as Error).message}\n`,
    )
    return
  }
  for (const line of raw.split('\n')) {
    // Empty lines and `# comments` are ignored; KEY=VALUE only. We don't
    // unquote — operators should write raw values.
    const m = line.match(/^([A-Z_][A-Z0-9_]*)=(.*)$/)
    if (!m) continue
    const [, key, value] = m
    if (key && value !== undefined && process.env[key] === undefined) {
      process.env[key] = value
    }
  }
}

/**
 * Resolve the public base URL for outbound file downloads (see
 * `Config.filesBaseUrl`). Never falls back to a fixed host: a link to any
 * host other than this install's own would send the download token (and,
 * through Microsoft's link prefetch, the request itself) to a third party.
 * Returns `null` and writes one stderr line when nothing usable is set.
 */
export function resolveFilesBaseUrl(env: NodeJS.ProcessEnv): string | null {
  const explicit = env.TEAMS_FILES_BASE_URL?.trim()
  if (explicit) {
    const url = parseHttpsUrl(explicit)
    if (url) return `${url.origin}${url.pathname.replace(/\/+$/, '')}`
    process.stderr.write(
      `teams channel: TEAMS_FILES_BASE_URL is not an https URL (got ${explicit}); send_file disabled\n`,
    )
    return null
  }
  const endpoint = env.TEAMS_BOT_ENDPOINT_URL?.trim()
  if (endpoint) {
    const url = parseHttpsUrl(endpoint)
    if (url) return `${url.origin}/files`
    process.stderr.write(
      `teams channel: TEAMS_BOT_ENDPOINT_URL is not an https URL (got ${endpoint}); send_file disabled\n`,
    )
    return null
  }
  process.stderr.write(
    'teams channel: no TEAMS_BOT_ENDPOINT_URL or TEAMS_FILES_BASE_URL set; send_file disabled\n',
  )
  return null
}

function parseHttpsUrl(raw: string): URL | null {
  try {
    const url = new URL(raw)
    return url.protocol === 'https:' && url.hostname ? url : null
  } catch {
    return null
  }
}

function required(env: NodeJS.ProcessEnv, key: string): string {
  const v = env[key]
  if (!v || v.trim() === '') {
    throw new Error(`teams channel: ${key} is required (set in env or <stateDir>/.env)`)
  }
  return v.trim()
}

/**
 * Load and validate the plugin config. Throws on any required-key miss.
 *
 * The function is intentionally side-effect-light: it makes the state dir if
 * absent (mode 0700 — only the operator should be able to read pending state
 * and conversation references) and chmods .env, but does not start any
 * servers or write the allowlist. Callers do that.
 */
export function loadConfig(env: NodeJS.ProcessEnv = process.env): Config {
  const stateDir = resolveStateDir(env)
  // Make state dir owner-only — same hardening Telegram does.
  mkdirSync(stateDir, { recursive: true, mode: 0o700 })

  const envFile = join(stateDir, '.env')
  loadEnvFile(envFile)
  // Re-read after the file load — process.env is now fully populated.
  env = process.env

  const appId = required(env, 'TEAMS_BOT_APP_ID')
  const appPassword = required(env, 'TEAMS_BOT_APP_PASSWORD')
  const tenantId = required(env, 'TEAMS_BOT_TENANT_ID')

  const bindHost = env.TEAMS_PLUGIN_BIND_HOST?.trim() || '127.0.0.1'
  const portStr = env.TEAMS_PLUGIN_PORT?.trim() || '3979'
  const port = Number.parseInt(portStr, 10)
  if (!Number.isFinite(port) || port < 1 || port > 65535) {
    throw new Error(`teams channel: TEAMS_PLUGIN_PORT must be 1..65535 (got ${portStr})`)
  }

  // File dirs default UNDER the state dir, not the upstream
  // '/home/ccuser/workspace/...' hardcode (the cand1 author's path) -- that
  // crashes (ENOTSUP mkdir /home/ccuser) on every host where the user is not
  // 'ccuser' (i.e. all of them). The sendable root is a SUBDIR ('files'), so the
  // secrets in the state-dir root (.env, access.json, pending.json) stay OUTSIDE
  // the send boundary and cannot be exfiltrated; received + outbox live under it
  // so they remain re-sendable.
  const filesRoot = join(stateDir, 'files')
  const sendableFilesRoot =
    env.SENDABLE_FILES_ROOT?.trim() || filesRoot
  const receivedFilesDir =
    env.RECEIVED_FILES_DIR?.trim() || join(filesRoot, 'received-files')
  const outboxDir =
    env.OUTBOX_DIR?.trim() || join(filesRoot, 'outbox')

  const outboxTtlStr = env.OUTBOX_TTL_SECONDS?.trim() || '1800'
  const outboxTtlSeconds = Number.parseInt(outboxTtlStr, 10)
  if (!Number.isFinite(outboxTtlSeconds) || outboxTtlSeconds < 1) {
    throw new Error(
      `teams channel: OUTBOX_TTL_SECONDS must be a positive integer (got ${outboxTtlStr})`,
    )
  }

  let plannerPlanIds = parsePlannerPlanIds(env.TEAMS_PLANNER_PLAN_IDS)
  if (plannerPlanIds === null) {
    process.stderr.write(
      'teams channel: TEAMS_PLANNER_PLAN_IDS has an invalid plan id; create_planner_task stays off\n',
    )
    plannerPlanIds = new Set()
  }

  return Object.freeze({
    unmentionedConversations: parseIdList(env.TEAMS_UNMENTIONED_CONVERSATIONS),
    stateDir,
    allowlistFile: join(stateDir, 'allowlist.json'),
    pendingFile: join(stateDir, 'pending.json'),
    appId,
    appPassword,
    appType: 'SingleTenant' as const,
    tenantId,
    bindHost,
    targetedMessages: targetedMessagesEnabled(env),
    port,
    receivedFilesDir,
    outboxDir,
    outboxTtlSeconds,
    sendableFilesRoot,
    filesBaseUrl: resolveFilesBaseUrl(env),
    plannerPlanIds,
  })
}
