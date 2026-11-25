/* Which Microsoft tenant, app and list the Site Log uses. None of these are secrets: the ids are public by design
   (a public client app, the same one as the KEI task board), and the list is protected by the signed-in person's
   own SharePoint permissions on the KEI Tasks site. */
var CONFIG = {
  "tenantId": "ff85c429-ed5b-458f-aca2-478fd1f538e1",
  "clientId": "d6fb3bfc-b421-4d95-b7b3-f3cce4dad884",
  "siteId": "kootenayenv.sharepoint.com,d6190ecd-1fad-4ef0-9fcc-eebca2ff6a06,6f9302de-fdc0-42da-a9dd-a48c907c14a5",
  "siteUrl": "https://kootenayenv.sharepoint.com/sites/KEITasks",
  "siteLogList": "2eb1edb6-be58-4a20-8297-c77c8b120a10"
};
