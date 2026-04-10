# Production Sync — Style Ledger Diff Engine

Auto-pulls 3 factory style ledgers from Dropbox every hour, compares against your master template, and shows what's new, what changed, and what's missing.

## Deploy to Render (free tier works)

### 1. Push to GitHub
```bash
git init
git add .
git commit -m "production sync"
git remote add origin https://github.com/YOUR_USERNAME/production-sync.git
git push -u origin main
```

### 2. Create Render Web Service
1. Go to [render.com](https://render.com) → New → Web Service
2. Connect your GitHub repo
3. Settings:
   - **Build Command:** `npm install`
   - **Start Command:** `npm start`
   - **Instance Type:** Free (or Starter for always-on)

### 3. Set Environment Variables
In Render dashboard → Environment → Add these:

| Key | Value |
|-----|-------|
| `DROPBOX_APP_KEY` | Your Dropbox app key |
| `DROPBOX_APP_SECRET` | Your Dropbox app secret |
| `DROPBOX_REFRESH_TOKEN` | Your Dropbox refresh token |
| `SYNC_MINUTES` | `60` |
| `PATH_FY1` | `/Production_Style_Ledger/FY1_Production_Style_Ledger.xlsx` |
| `PATH_NB` | `/Production_Style_Ledger/NB_Production_Style_Ledger.xlsx` |
| `PATH_PC` | `/Production_Style_Ledger/PC_Production_Style_Ledger.xlsx` |
| `PATH_DAVID` | `/Versa Share Files/David - Dropbox/Style Ledger/Style Ledger Template for David.xlsx` |

### 4. Done
Render deploys automatically. Your app syncs on startup and every 60 minutes.

> **Note:** Render free tier spins down after 15 min of inactivity. Use Starter ($7/mo) for always-on. Alternatively, use a cron service like UptimeRobot to ping your URL every 14 min to keep it alive.

## Local Development
```bash
cp .env.example .env
# Edit .env with your Dropbox token
npm install
npm start
# Open http://localhost:3000
```

## How It Works
- **New Styles:** Factory styles not in your ledger with ETD in the last 60 days or future. Old shipped items are ignored.
- **Updates:** Matched styles where Ship Units, ETD, Brand, PO Name, or Production# differ.
- **Missing:** Styles in your ledger but not found in any factory file.
- **Export:** Download a .xlsx report of all changes via the Export button or `/api/export`.
