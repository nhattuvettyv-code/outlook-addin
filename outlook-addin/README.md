# Outlook Smart Alert Add-in — Customer Name Validator

Validates that the customer name in an email subject line (format: `[Customer Name]`)
matches the domains/addresses of every recipient before the message is sent.
Mismatches are hard-blocked with a clear error message.

---

## File structure

```
outlook-addin/
├── manifest.xml          ← Office Add-in manifest
├── src/
│   ├── commands.html     ← Hidden page loaded by Outlook to run the handler
│   ├── commands.js       ← Smart Alert logic (OnMessageSend handler)
│   └── mapping.json      ← Customer name → allowed domains/addresses
└── assets/
    ├── icon-16.png
    ├── icon-32.png
    └── icon-80.png       ← Add placeholder PNGs or replace with real icons
```

---

## Editing the mapping

Open `src/mapping.json` and add entries following this schema:

```json
{
  "Customer Name": {
    "domains": ["customer.com", "customer.org"],
    "addresses": ["specific-contact@otherdomain.com"]
  }
}
```

- **domains** — any recipient whose email ends in one of these domains is allowed.
- **addresses** — specific email addresses that are allowed regardless of domain.
- Matching is **case-insensitive**.

---

## Hosting (required before sideloading)

The add-in files must be served over **HTTPS**. Options:

### Option A — GitHub Pages (free, easiest)
1. Push this folder to a public GitHub repo.
2. Enable GitHub Pages (Settings → Pages → `main` branch / root).
3. Your base URL will be `https://<username>.github.io/<repo>/`.

### Option B — Azure Static Web Apps
1. `az staticwebapp create ...` and deploy the folder.
2. Use the provided `*.azurestaticapps.net` URL.

### Option C — Local dev with ngrok (for testing only)
```bash
npx http-server . -p 3000 --cors   # from inside outlook-addin/
ngrok http 3000                    # exposes https://xxxx.ngrok.io
```

---

## Replace placeholder URLs

Once you have a hosting URL, replace every occurrence of `https://YOUR-HOST` in
`manifest.xml` with your actual URL, e.g.:

```
https://YOUR-HOST/src/commands.html
→ https://myorg.github.io/outlook-addin/src/commands.html
```

---

## Sideloading (Outlook desktop — Windows)

1. Open Outlook desktop.
2. Go to **File → Manage Add-ins** (opens the web Add-ins centre).
3. Click **My add-ins → Add a custom add-in → Add from file**.
4. Select `manifest.xml`.
5. Confirm the warning prompt.

The add-in is now active for your account. Smart Alerts fire automatically on send.

---

## Sideloading (Outlook on the web)

1. Go to **Settings (⚙) → View all Outlook settings → Mail → Customize actions**.
2. Or navigate to: `https://outlook.office.com/mail/inclientstore`
3. Upload `manifest.xml`.

---

## Deploying org-wide (Microsoft 365 admin)

1. M365 Admin Centre → **Settings → Integrated apps → Upload custom app**.
2. Upload `manifest.xml` and assign to users/groups.

---

## How it works

1. User clicks **Send**.
2. Outlook fires `OnMessageSend` — the add-in intercepts it **before** delivery.
3. The handler:
   - Reads the subject and extracts text between `[` and `]`.
   - Fetches `mapping.json` from the hosted URL.
   - Looks up the customer name (case-insensitive).
   - Checks every recipient in To / CC / BCC against allowed domains and addresses.
4. If all pass → `event.completed({ allowEvent: true })` — email sends normally.
5. If any fail → `event.completed({ allowEvent: false, errorMessage: "…" })` — send is blocked and the user sees the error dialog.

---

## Supported Outlook versions

| Platform | Minimum version |
|----------|----------------|
| Outlook on Windows | 2206 (Build 15330.20196) |
| Outlook on Mac | 16.65 |
| Outlook on the web | All current versions |
| New Outlook for Windows | Supported |

Requirement set: **Mailbox 1.12** (Smart Alerts).

---

## Troubleshooting

| Symptom | Fix |
|---------|-----|
| Add-in doesn't fire on send | Check manifest `<Set Name="Mailbox" MinVersion="1.12"/>` and that Outlook is updated |
| `mapping.json` fetch fails | Ensure the file is served with `Content-Type: application/json` and CORS allows the add-in origin |
| "Unknown customer" error for a valid name | Check for extra spaces or special characters in the subject brackets; mapping lookup is case-insensitive but exact otherwise |
| Icons missing | Place 16×16, 32×32, 80×80 PNG files in `assets/` and update manifest URLs |
