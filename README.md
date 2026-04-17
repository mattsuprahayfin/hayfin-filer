[README.md](https://github.com/user-attachments/files/26831329/README.md)
# Hayfin Inbox Filer — Outlook Add-in

An Outlook sidebar add-in that reads your inbox, groups emails into conversations, gets AI folder suggestions from Claude, and moves them with one click.

---

## What it does

- Opens as a sidebar panel in Outlook (web or desktop)
- Fetches your inbox and groups emails by conversation thread
- Calls the Anthropic API to suggest the right folder for each thread
- Lets you confirm, skip, or override each suggestion
- Moves emails directly via Exchange Web Services (no copy — the actual message moves)
- Learns from your corrections over time (stored locally)

---

## Files

```
hayfin-filer/
├── taskpane.html        — the sidebar UI
├── manifest.xml         — tells Outlook where to find the add-in
├── src/
│   ├── app.js           — main application logic
│   └── folders.js       — your Hayfin folder structure + IDs
├── assets/
│   ├── icon-16.png      — required icons (generate or provide your own)
│   ├── icon-32.png
│   ├── icon-64.png
│   ├── icon-80.png
│   └── icon-128.png
└── generate_icons.py    — optional: generates placeholder icons
```

---

## Deployment (sideloading — no IT required)

### Step 1: Host on GitHub Pages (free, ~5 minutes)

1. Create a free GitHub account at github.com if you don't have one
2. Create a new **public** repository called `hayfin-filer`
3. Upload all files in this folder (drag and drop onto the GitHub web UI works fine)
4. Go to **Settings → Pages** in the repository
5. Set Source to **Deploy from a branch**, select `main`, folder `/` (root)
6. Click Save. After ~60 seconds, your add-in will be live at:
   `https://YOUR-USERNAME.github.io/hayfin-filer/`

### Step 2: Update the manifest with your URL

Open `manifest.xml` and replace every occurrence of:
```
YOUR-GITHUB-USERNAME
```
with your actual GitHub username. There are 5 places to update.

Re-upload the updated `manifest.xml` to GitHub.

### Step 3: Generate icons

```bash
pip install pillow
python3 generate_icons.py
```

Upload the generated `assets/` folder to GitHub.

### Step 4: Sideload into Outlook

**In Outlook on the web (recommended for first test):**
1. Open https://outlook.office.com
2. Click the gear icon (Settings) → **View all Outlook settings**
3. Go to **Mail → Customise actions** or navigate to **Add-ins**
   *(exact path varies slightly — search for "manage add-ins" in settings)*
4. Click **My add-ins** → **Add a custom add-in** → **Add from file**
5. Upload your `manifest.xml`
6. The "File Inbox" button will appear in the toolbar when reading any email

**In Outlook desktop (Windows):**
1. Open Outlook
2. Go to **File → Manage Add-ins** (or **Get Add-ins** in the ribbon)
3. This opens the Office Add-ins web page
4. Click **My add-ins** → **+ Add custom add-in** → **Add from file**
5. Select `manifest.xml`

**In Outlook desktop (Mac):**
1. Open Outlook
2. Go to **Tools → Add-ins…**
3. Click **+** → choose `manifest.xml`

---

## First use

1. Click the **File Inbox** button in the Outlook toolbar while reading any email
2. The sidebar opens — you'll be prompted for your Anthropic API key
3. Get a key at https://console.anthropic.com → API Keys (free tier available)
4. Paste the key (starts with `sk-ant-`) and click Save
5. Press **Load inbox** to fetch your inbox
6. Review suggestions, click **Move ✓** to file each thread

Your API key is stored in browser localStorage — it never leaves your browser except to go directly to Anthropic.

---

## Updating folders

If you add or rename Outlook folders, update `src/folders.js`. You can use Claude to fetch fresh folder IDs — just say "refresh the folder list" in a conversation with the MS365 connector active.

---

## Data & privacy notes

- Email subjects and sender addresses are sent to the Anthropic API for suggestions
- Full email bodies are never sent
- Your folder structure is stored locally in the add-in files
- Your API key is stored in browser localStorage (not in any file or server)
- All email moves happen directly between your browser and your Exchange server

---

## Troubleshooting

| Problem | Fix |
|---|---|
| "Move failed" error | The add-in needs `ReadWriteMailbox` permission in the manifest — check it's there |
| Suggestions don't load | Check your API key is valid at console.anthropic.com |
| Add-in doesn't appear | Make sure the manifest URL matches your GitHub Pages URL exactly |
| Icons don't load | Run `generate_icons.py` or add your own PNGs to the `assets/` folder |
| Outlook shows "this app can't be loaded" | The manifest XML must be valid — re-download from GitHub raw view |

---

## Costs

The add-in uses Claude claude-sonnet-4-20250514. At typical usage (50 emails/day):
- ~1,500 input tokens per suggestion request
- ~50 emails = ~75,000 tokens/day input
- At $3 per million input tokens = **~$0.22/day** or ~£4/month
