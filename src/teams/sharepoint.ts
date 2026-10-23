/**
 * Uploads a file into a folder of one SharePoint site with an app-only Microsoft Graph token of the bot's own
 * app registration (client credentials). The app needs the Graph application permission Sites.Selected, admin consent, and a
 * write grant on that one site (docs/sharepoint.md). The client secret goes only to the Microsoft identity platform token
 * endpoint; no token or secret is ever logged or put in an error message. No request follows a redirect (redirect:
 * 'error'), so neither the secret nor the file can be sent on to another host.
 */

export interface SharePointDeps {
  tenantId: string
  appId: string
  appPassword: string
  siteId: string
  folder: string
  fetch?: typeof fetch
  now?: () => number
}

export interface SharePointUpload {
  webUrl: string
  name: string
}

export const SIMPLE_UPLOAD_MAX = 4 * 1024 * 1024
export const CHUNK_BYTES = 327680 * 12

export function createSharePointUploader(deps: SharePointDeps): {
  upload(filename: string, content: Buffer): Promise<SharePointUpload>
} {
  const rawFetch = deps.fetch ?? globalThis.fetch
  // A fetch error can carry the request URL in its message (measured: Bun's UnexpectedRedirect does), and the upload
  // URL is pre-authenticated. So a failed request is rethrown with a fixed message and only the error code.
  const doFetch = async (url: string, init: RequestInit): Promise<Response> => {
    try {
      return await rawFetch(url, init)
    } catch (err) {
      const code = (err as { code?: unknown } | null)?.code
      const safe = typeof code === 'string' && /^[A-Za-z_]{1,40}$/.test(code) ? code : 'network error'
      throw new Error(`SharePoint request failed (${safe})`)
    }
  }
  const now = deps.now ?? Date.now
  let cached: { token: string; expiresAt: number } | null = null

  async function getToken(): Promise<string> {
    if (cached && now() < cached.expiresAt) return cached.token
    const res = await doFetch(
      `https://login.microsoftonline.com/${encodeURIComponent(deps.tenantId)}/oauth2/v2.0/token`,
      {
        method: 'POST',
        redirect: 'error',
        headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
        body: new URLSearchParams({
          client_id: deps.appId,
          client_secret: deps.appPassword,
          scope: 'https://graph.microsoft.com/.default',
          grant_type: 'client_credentials',
        }),
      },
    )
    let json: any = null
    try {
      json = await res.json()
    } catch {}
    if (
      res.status === 200 &&
      json &&
      typeof json.access_token === 'string' &&
      typeof json.expires_in === 'number'
    ) {
      cached = {
        token: json.access_token,
        expiresAt: now() + (json.expires_in - 60) * 1000,
      }
      return cached.token
    }
    const e = json?.error
    const err = typeof e === 'string' && /^[a-z_]{1,64}$/.test(e) ? e : ''
    throw new Error(`SharePoint token request failed (HTTP ${res.status}${err ? ', ' + err : ''})`)
  }

  function parseItem(json: any): SharePointUpload {
    if (!json || typeof json.webUrl !== 'string' || typeof json.name !== 'string') {
      throw new Error('SharePoint upload returned no webUrl')
    }
    return { webUrl: json.webUrl, name: json.name }
  }

  async function readJson(res: Response): Promise<any> {
    try {
      return await res.json()
    } catch {
      return null
    }
  }

  async function upload(filename: string, content: Buffer): Promise<SharePointUpload> {
    const token = await getToken()
    // No path separators, and never a '.' or '..' segment.
    const flat = filename.replace(/[\/\\]/g, '_')
    const safeName = flat === '.' || flat === '..' || flat === '' ? '_' : flat
    const itemPath = `${deps.folder}/${safeName}`
      .split('/')
      .map(encodeURIComponent)
      .join('/')
    const base = `https://graph.microsoft.com/v1.0/sites/${deps.siteId}/drive/root:/${itemPath}:`

    if (content.length <= SIMPLE_UPLOAD_MAX) {
      const res = await doFetch(`${base}/content?@microsoft.graph.conflictBehavior=rename`, {
        method: 'PUT',
        redirect: 'error',
        headers: {
          Authorization: `Bearer ${token}`,
          'Content-Type': 'application/octet-stream',
        },
        body: content as unknown as BodyInit,
      })
      if (res.status !== 200 && res.status !== 201) {
        throw new Error(`SharePoint upload failed (HTTP ${res.status})`)
      }
      return parseItem(await readJson(res))
    }

    const sess = await doFetch(`${base}/createUploadSession`, {
      method: 'POST',
      redirect: 'error',
      headers: {
        Authorization: `Bearer ${token}`,
        'Content-Type': 'application/json',
      },
      body: '{"item":{"@microsoft.graph.conflictBehavior":"rename"}}',
    })
    if (sess.status !== 200) {
      throw new Error(`SharePoint upload failed (HTTP ${sess.status})`)
    }
    const sessJson = await readJson(sess)
    const uploadUrl = sessJson?.uploadUrl
    let okUrl = false
    if (typeof uploadUrl === 'string' && uploadUrl.startsWith('https://')) {
      try {
        okUrl = new URL(uploadUrl).hostname.toLowerCase().endsWith('.sharepoint.com')
      } catch {}
    }
    if (!okUrl) throw new Error('SharePoint returned an unexpected upload URL')

    const total = content.length
    for (let start = 0; start < total; start += CHUNK_BYTES) {
      const end = Math.min(start + CHUNK_BYTES, total) - 1
      const chunk = content.subarray(start, end + 1)
      const res = await doFetch(uploadUrl, {
        method: 'PUT',
        redirect: 'error',
        headers: {
          'Content-Length': String(chunk.length),
          'Content-Range': `bytes ${start}-${end}/${total}`,
        },
        body: chunk as unknown as BodyInit,
      })
      const last = end === total - 1
      if (last && (res.status === 200 || res.status === 201)) {
        return parseItem(await readJson(res))
      }
      if (res.status !== 202 || last) {
        throw new Error(`SharePoint upload failed (HTTP ${res.status})`)
      }
    }
    throw new Error('SharePoint upload failed (HTTP 0)')
  }

  return { upload }
}
