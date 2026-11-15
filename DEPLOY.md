# Free trial run: GitHub Codespaces

The `.devcontainer` folder makes Codespaces install Chromium, Node.js and the Python packages automatically.

1. Upload everything in this `cloud-deploy` folder to your GitHub repo, **including the `.devcontainer` folder**.
   On GitHub, use **Add file > Upload files** and drag the whole folder in. The repo root should then
   contain `.devcontainer/devcontainer.json`.
2. *(Recommended)* Add your secrets before you start. On GitHub, go to your profile picture >
   **Settings > Codespaces > Secrets > New secret**, and allow the `Lab-management-agent` repo for each:
   `ADMIN_KEY` (a password you make up), `GMAIL_APP_PASSWORD`, `GEMINI_API_KEY` (optional).
3. Open the repo, then **Code > Codespaces > Create codespace on main**.
   The first build takes about 3–5 minutes.
4. In the codespace terminal, run:
   ```bash
   python agent.py
   ```
   A QR code prints in the terminal. Scan it with WhatsApp > **Linked devices > Link a device**.
5. To load your data, drag `lab_inventory.xlsx` from your PC into the `data` folder in the file explorer
   on the left, then restart (`Ctrl+C`, then `python agent.py` again).
   You can also open the **Ports** tab, click the globe icon on port 5000, and add `/admin?key=<ADMIN_KEY>` to the URL.
6. Send test messages from WhatsApp.

Notes:
- Port 5000 is **private** by default: only you, signed in to GitHub, can open it. Keep it that way.
- The codespace stops after about 30 minutes of inactivity and the bot goes offline. Reopen it and run step 4 again.
  The `data` folder (Excel file and WhatsApp login) survives stops, but is deleted if you **delete** the codespace.
- The `data` folder is excluded from git (`.gitignore`), so your Excel file and login are never pushed to the public repo.
- Free personal accounts get roughly 120 core-hours a month. Stop the codespace when you're done
  (**Code > Codespaces > ... > Stop codespace**).

---

# Deploying the Lab Agent to Railway (always on)

One container runs both `agent.py` (Flask) and `bridge.js` (WhatsApp). The `Dockerfile` installs
Python, Node.js and Chromium. Railway costs about US$5/month (Hobby plan) and supports the
persistent volume this app needs.

## 1. Put these files in your GitHub repo

Upload everything in this `cloud-deploy` folder to https://github.com/devinleeyh-dot/Lab-management-agent
(GitHub > **Add file > Upload files**), replacing the old `agent.py`, `bridge.js`, `package.json`
and `requirements.txt`:

`agent.py`, `bridge.js`, `package.json`, `requirements.txt`, `Dockerfile`, `.dockerignore`, `.gitignore`

Do **not** upload `.env`, `lab_inventory.xlsx`, `.wwebjs_auth` or `.wwebjs_cache`. The repo is public.
The Excel file goes up later through the private admin page.

## 2. Create the Railway service

1. Go to railway.com, sign in with GitHub, then **New Project > Deploy from GitHub repo**, and pick `Lab-management-agent`.
2. Railway detects the `Dockerfile` automatically.
3. **Add a volume:** right-click the service, choose **Attach volume**, and set the mount path to `/data`.
   It stores the Excel database and the WhatsApp login. Without it, both are wiped on every redeploy.
4. **Variables** tab: add

   | Variable | Value |
   |---|---|
   | `ADMIN_KEY` | a long random password you make up (protects the admin page) |
   | `GMAIL_APP_PASSWORD` | your 16-character Gmail app password |
   | `GEMINI_API_KEY` | your Gemini key (optional, only for free-form sentences) |
   | `OWNER_PHONE` | `6589033626` (change if the owner changes) |
   | `GMAIL_USER` | the Gmail address that sends PO emails |

5. **Settings > Networking > Generate Domain** gives you a URL such as `https://lab-agent-production.up.railway.app`.

## 3. Link WhatsApp (one time)

Open `https://<your-domain>/admin?key=<ADMIN_KEY>` and scan the QR code with
WhatsApp > **Linked devices > Link a device**. The status changes to **ready**.
The login is saved in `/data`, so restarts and redeploys don't need a new scan.

## 4. Load your existing data

On the same admin page, use **Upload & replace** to upload your local `lab_inventory.xlsx`
(it contains your vendors, requesters and stock). Use **Download** at any time to back it up or
edit it in Excel, then upload it again. This replaces the old start-up menu.

## Addresses

- `/` : health check (shows the WhatsApp status)
- `/admin?key=...` : QR code, Excel download and upload
- `/incoming`, `/users`, `/unknown` : internal only. Requests from the internet get 403.

## Known issue: first-time WhatsApp login

In testing on 2026-11-15, whatsapp-web.js (1.34.7 and 1.34.8) failed to show a QR code for a
**brand-new** login on Windows ("Execution context was destroyed"). The existing local login
was unaffected. If the Railway logs show the same error repeating and `/admin` never shows a QR:
- Check the whatsapp-web.js GitHub issues for a fix, bump the version in `package.json`, and redeploy.
- Or set `WA_WEB_VERSION` to pin a WhatsApp Web build (see the comment in `bridge.js`).
