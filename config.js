/* Which Microsoft tenant, app and lists this board uses.
   None of these are secrets: the client id and tenant id are public by design
   (this is a public client app), and the lists themselves are protected by the
   signed-in person's own SharePoint permissions. */
var CONFIG = {
  "tenantId": "ff85c429-ed5b-458f-aca2-478fd1f538e1",
  "clientId": "d6fb3bfc-b421-4d95-b7b3-f3cce4dad884",
  "siteId": "kootenayenv.sharepoint.com,d6190ecd-1fad-4ef0-9fcc-eebca2ff6a06,6f9302de-fdc0-42da-a9dd-a48c907c14a5",
  "siteUrl": "https://kootenayenv.sharepoint.com/sites/KEITasks",
  "lists": {
    "KEI Tasks": "95cc11de-141a-4b5f-93d6-76f1452ffff1",
    "Personal Tasks": "eddf252e-38a6-4227-8570-1607b62a6fde",
    "Board Settings": "deb7edd9-5c89-4eda-8be9-9ea792eeb460",
    "Activity": "04abf394-8097-425f-9179-e3c85b076032"
  }
};
