# CED Platform — GitHub Pages Deployment Guide

## What changed from the Claude artifact version

| Feature | Claude Artifact | GitHub Pages |
|---------|----------------|--------------|
| Data storage | `window.storage` (Claude API) | `localStorage` (browser built-in) |
| AI extraction | Auto-authenticated | Requires your Anthropic API key |
| Hosting | Claude servers | Your GitHub Pages |

Everything else (login, all 17 pages, analytics, exports) works identically.

---

## Prerequisites

- A GitHub account (free)
- Node.js installed on your computer (download from https://nodejs.org — choose LTS version)
- Git installed (download from https://git-scm.com)

---

## Step-by-step deployment

### Step 1: Create a GitHub repository

1. Go to https://github.com/new
2. Repository name: `ced-platform` (or any name you prefer)
3. Set to **Public** (required for free GitHub Pages)
4. Do NOT check "Add a README" — we'll push our own files
5. Click **Create repository**
6. Keep this page open — you'll need the URL

### Step 2: Set up the project on your computer

Open a terminal (Command Prompt on Windows, Terminal on Mac) and run these commands one by one:

```bash
# 1. Create project folder and navigate to it
mkdir ced-platform
cd ced-platform

# 2. Initialize git
git init
```

### Step 3: Copy the project files

Copy ALL the files from the downloaded zip into your `ced-platform` folder. The structure should look like:

```
ced-platform/
├── index.html
├── package.json
├── vite.config.js
└── src/
    ├── main.jsx
    └── App.jsx
```

### Step 4: Update the repo name in vite.config.js

Open `vite.config.js` and change the `base` line to match your GitHub repo name:

```js
base: '/ced-platform/',  // ← Must match your repo name exactly
```

If your repo is named `my-university-app`, change it to:
```js
base: '/my-university-app/',
```

### Step 5: Install dependencies and build

```bash
# Install all packages (this takes 1-2 minutes)
npm install

# Build the production version
npm run build
```

You should see a `dist/` folder created.

### Step 6: Push to GitHub

```bash
# Add all files
git add .

# Commit
git commit -m "Initial commit - CED Platform"

# Connect to your GitHub repo (replace with YOUR repo URL)
git remote add origin https://github.com/YOUR_USERNAME/ced-platform.git

# Push
git branch -M main
git push -u origin main
```

### Step 7: Deploy to GitHub Pages

```bash
npm run deploy
```

This builds and pushes to a `gh-pages` branch automatically.

### Step 8: Enable GitHub Pages

1. Go to your repo on GitHub
2. Click **Settings** → **Pages** (in the left sidebar)
3. Under "Source", select **Deploy from a branch**
4. Branch: select `gh-pages` and `/ (root)`
5. Click **Save**
6. Wait 2-3 minutes

### Step 9: Access your site

Your site will be live at:
```
https://YOUR_USERNAME.github.io/ced-platform/
```

---

## Setting up the API key (for AI extraction features)

The AI document extraction (camera/file upload) requires an Anthropic API key:

1. Go to https://console.anthropic.com/settings/keys
2. Create a new API key
3. Open your CED platform → Settings → **Anthropic API Key** section
4. Paste your key (starts with `sk-ant-...`)
5. The key is stored in your browser only — it never leaves your device

**Note:** Without an API key, all manual features work perfectly. Only the AI extraction buttons (📄 and 📷) need the key.

---

## Updating the site

When you make changes to the code:

```bash
# Build and deploy in one command
npm run deploy
```

The site updates in 1-2 minutes.

---

## Important notes

### Data storage
- All data is stored in your browser's `localStorage`
- Data persists across sessions on the SAME browser
- Different browsers/devices have separate data
- Clearing browser data will erase all platform data
- To back up data: use the Backup & Export page in the app

### Security
- The API key is stored in localStorage (your browser only)
- Login passwords are SHA-256 hashed
- No data is sent to any server except Anthropic (for AI extraction only)

### Storage limits
- localStorage has a 5-10MB limit per domain (browser dependent)
- If you approach the limit, use the Archive feature in Backup settings
- Excel/PDF exports are not affected by storage limits

---

## Troubleshooting

| Problem | Solution |
|---------|----------|
| Blank page after deploy | Check `base` in vite.config.js matches repo name |
| 404 on page refresh | Normal for single-page apps on GitHub Pages — navigate from the root URL |
| AI extraction fails | Check API key in Settings |
| `npm install` fails | Make sure Node.js is installed: run `node --version` |
| `npm run deploy` fails | Make sure you've pushed to GitHub first with `git push` |
| Data disappeared | Check if you switched browsers or cleared browser data |
