# SharePoint target for send_file (opt-in)

## 1. What it does

When you configure a SharePoint site, `send_file` in a shared conversation (team channel or group chat) uploads the file into one folder of that one site and posts a card that links the SharePoint item. 1:1 chats keep the short-lived download link served by this plugin. If the SharePoint upload fails for any reason, the plugin falls back to the download link. It is off by default.

## 2. Microsoft Entra

In the bot's app registration, add the Microsoft Graph **application** permission `Sites.Selected` and grant admin consent (needs a Global Administrator or Privileged Role Administrator). `Sites.Selected` alone gives access to no site.

## 3. Site grant

An administrator gives the app write access to the one site with Microsoft Graph:

```
POST https://graph.microsoft.com/v1.0/sites/{site-id}/permissions
{"roles":["write"],"grantedToIdentities":[{"application":{"id":"<bot app id>","displayName":"<bot name>"}}]}
```

The caller needs `Sites.FullControl.All` (for example in Graph Explorer). Microsoft's "Create permission" page for sites is the reference for the exact body.

## 4. Find the site id

```
GET https://graph.microsoft.com/v1.0/sites/{hostname}:/sites/{site-path}
```

Use the `id` field of the answer (`hostname,site-guid,web-guid`).

## 5. Plugin config

In the plugin's env file:

```
TEAMS_SHAREPOINT_SITE_ID=<id>
TEAMS_SHAREPOINT_FOLDER=Teams bot files   # optional
```

Restart the session. The bot's existing `TEAMS_BOT_APP_ID`, `TEAMS_BOT_APP_PASSWORD` and tenant are used for the token.

## 6. What is sent where

- The token request goes to `login.microsoftonline.com` with the bot's client secret.
- Uploads go to `graph.microsoft.com` and, for large files, to the SharePoint upload URL Graph returns (only `*.sharepoint.com` is accepted).
- Existing names are kept: a new upload with the same name is renamed by SharePoint.

## 7. Limits

Files are limited by the plugin's 50 MB `send_file` cap. The card links the item, which is readable by people with access to the site.
