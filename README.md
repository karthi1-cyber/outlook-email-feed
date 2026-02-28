# 📧 Email Feed — Outlook Web Add-in

Select multiple emails → view them as a stacked feed → reply one by one.

---

## Setup (5 minutes)

### Step 1 — Create the GitHub repo

1. Go to **https://github.com/new**
2. Name the repo **`outlook-email-feed`** (this exact name matters)
3. Set it to **Public**
4. Click **Create repository**

### Step 2 — Upload the files

1. On your new repo page, click **"uploading an existing file"** (or the **Add file → Upload files** button)
2. Drag and drop **all** files and folders from this project into the upload area:
   - `manifest.xml`
   - `taskpane.html`
   - `taskpane.css`
   - `taskpane.js`
   - `assets/` folder (with all the icon PNGs)
3. Click **Commit changes**

### Step 3 — Enable GitHub Pages

1. In your repo, go to **Settings** → **Pages** (left sidebar)
2. Under **Source**, select **Deploy from a branch**
3. Set branch to **`main`** and folder to **`/ (root)`**
4. Click **Save**
5. Wait 1–2 minutes, then visit: `https://YOUR_GITHUB_USERNAME.github.io/outlook-email-feed/taskpane.html`
6. You should see the Email Feed interface load (it will say "No emails loaded" — that's correct)

### Step 4 — Update the manifest

1. Open `manifest.xml` in any text editor
2. **Find and replace** all instances of `YOUR_GITHUB_USERNAME` with your actual GitHub username
   - There are about 12 instances to replace
   - Example: `YOUR_GITHUB_USERNAME` → `karthi-pfg`
3. Save the file
4. Re-upload the updated `manifest.xml` to your GitHub repo (Add file → Upload files → overwrite)

### Step 5 — Sideload into Outlook

1. Go to **https://outlook.office.com** and sign in
2. Click any email to open it
3. Go to **https://aka.ms/olksideload** (or find Add-ins via the ribbon **"…"** menu → **Get Add-ins**)
4. Scroll to **"Custom add-ins"** at the bottom
5. Click **"Add from file…"**
6. Select your updated `manifest.xml`
7. Click **Install**

### Step 6 — Use it

1. In your Outlook inbox, **select multiple emails** (hold `Ctrl` / `Cmd` and click)
2. Click **"Open Feed"** in the ribbon (may be under the **"…"** overflow menu)
3. The task pane opens → click **"Load Selected Emails"**
4. Reply, Reply All, or Forward from each card
5. Hit **Send** (or `Cmd/Ctrl + Enter`)

---

## Troubleshooting

| Problem | Fix |
|---|---|
| Task pane is blank | Visit your GitHub Pages URL directly in the browser to confirm it loads |
| "Open Feed" button missing | Refresh Outlook. New add-ins can take a minute to appear |
| Multi-select not working | Must use Outlook on the web or New Outlook (not classic desktop) |
| "Failed to get token" | Your M365 account may need admin consent for REST API access |
| GitHub Pages 404 | Check the repo name is exactly `outlook-email-feed` and Pages is enabled on `main` branch |

---

## What it does

- Loads selected emails via the Outlook REST API
- Displays them as a scrollable card feed in the task pane
- Each card shows sender, subject, timestamp, and body preview
- Inline Reply / Reply All / Forward composer on each card
- Green "Replied" badge after sending, auto-scrolls to next email
- Pinnable task pane stays open as you navigate
