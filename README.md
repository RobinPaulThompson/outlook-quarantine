# outlook-quarantine
This lightweight Outlook web add‑in adds a ribbon button that opens Microsoft Defender portal → Quarantine (https://security.microsoft.com/quarantine) directly from Outlook (Win, Mac, Web). It uses Office.js command buttons and a simple function to launch a secure pop‑up (Office dialog) to the Defender portal.

✅ Works in: New Outlook for Windows, classic Outlook for Windows (2021/365), Outlook on the web, and Outlook for Mac (new UI), as long as add‑ins are allowed and the user has portal access.

🔐 Sign-in & permissions: The add‑in does not handle auth itself; the Defender portal will prompt and enforce the user’s existing Microsoft Entra ID (Azure AD) auth and RBAC. You can optionally restrict who sees the command via admin deployment.
