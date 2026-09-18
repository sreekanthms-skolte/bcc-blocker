# Outlook BCC Blocker

## Purpose
Blocks users from sending a message when they manually add any Bcc recipient.

The existing Exchange Online transport rule that automatically adds the compliance mailbox as Bcc remains unchanged.

## Important
Replace every occurrence of:

https://sreekanthms-skolte.github.io/bcc-blocker

with your real public HTTPS host before deployment.

Example:

https://mailtools.yourcompany.com

The following URLs must then be reachable publicly over HTTPS:

/bcc-blocker/commands.html
/bcc-blocker/launchevent.js
/bcc-blocker/taskpane.html
/bcc-blocker/assets/icon-16.png
/bcc-blocker/assets/icon-32.png
/bcc-blocker/assets/icon-80.png

## Logic

1. User clicks Send.
2. Outlook runs OnMessageSend.
3. Add-in reads the compose-time Bcc list.
4. If Bcc exists: send is blocked.
5. If Bcc is empty: send is allowed.
6. Exchange Online transport rule can then add the compliance Bcc on the server side.

## Deployment

1. Host these files on a public HTTPS web server.
2. Edit manifest.xml and replace YOUR-DOMAIN.example.
3. In Microsoft 365 Admin Center:
   Settings -> Integrated apps -> Add-ins
4. Deploy the custom Outlook add-in using manifest.xml.
5. Test on one user first.
6. After testing, deploy to the required users/groups.

## Expected test

Normal mail:
To = customer@example.com
Cc = blank
Bcc = blank
=> Send succeeds.

Manual Bcc:
To = customer@example.com
Bcc = someone@example.com
=> Send is blocked.

Compliance rule:
No user Bcc entered
=> Outlook allows the message
=> Exchange transport rule adds compliance mailbox as Bcc
=> compliance mailbox receives the copy.

## Note
This package uses the Outlook Smart Alerts / OnMessageSend feature with SendMode=Block.
Microsoft documents Business Basic as eligible for centralized Office Add-in deployment,
and documents Bcc recipient access via the Outlook compose API.
