/* KEI Voice - which Microsoft tenant and app it signs in with, and where uploads go.
   None of these are secrets: the tenant and client ids are public by design (the
   same public-client app registration as the KEI board), and the folder is in the
   signed-in person's own OneDrive. The PC side (KEI Recordings/_tools/phone_inbox.py)
   reads that folder - change both together. */
var CONFIG = {
  tenantId: 'ff85c429-ed5b-458f-aca2-478fd1f538e1',
  clientId: 'd6fb3bfc-b421-4d95-b7b3-f3cce4dad884',
  folder: 'KEI Recordings/phone_inbox'
};
