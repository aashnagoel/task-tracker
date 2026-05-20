# Task Tracker - Complete Setup Instructions

## You Have GitHub Ready ✓

- GitHub account: aashnagoel
- Repository: https://github.com/aashnagoel/task-tracker
- GitHub Pages enabled
- Repository cloned to: ~/Desktop/task-tracker/

## Step 1: Configure Git (One-Time)

Open Terminal and run:

```bash
git config --global user.name "aashnagoel"
git config --global user.email "aashnagoel1999@gmail.com"
```

Verify it worked:
```bash
git config --global user.email
# Should show: aashnagoel1999@gmail.com
```

## Step 2: Move Files to Your Repository

Download these 5 files from outputs and put them in `~/Desktop/task-tracker/`:

1. `task_tracker_FINAL.py`
2. `slack_users_fetcher_FIXED.py`
3. `github_push.py`
4. `README.md`
5. `slack_users_tsip_contributors.json` (from your original task_tracker_data folder)

Also move your Excel files:
- `week_2026_05_24.xlsx`
- Any other weekly Excel files

Your folder should look like:
```
~/Desktop/task-tracker/
├── task_tracker_FINAL.py
├── slack_users_fetcher_FIXED.py
├── github_push.py
├── README.md
├── slack_users_tsip_contributors.json
├── week_2026_05_24.xlsx
└── .git/ (hidden folder, already there)
```

## Step 3: Install Python Dependencies

```bash
pip3 install openpyxl requests
```

## Step 4: Test It Works

```bash
cd ~/Desktop/task-tracker
python3 task_tracker_FINAL.py
```

Should output:
```
=== Task Tracker Processor ===
✓ Loaded 179 Slack users
✓ Found Excel: week_2026_05_24.xlsx
✓ Claim Sheet: 14 taskers
✓ Model Decomp: 8 taskers
✓ Generating dashboards...
✓ Generated dashboard.html
✓ Generated approval_ui.html
✓ COMPLETE!
```

If you see this, you're ready!

## Daily Workflow (Every Morning)

### 9:15 AM: Download Excel

Export from your task sheet as `.xlsx` and save to:
```
~/Desktop/task-tracker/
```

Name it with the week: `week_2026_05_24.xlsx` or similar

### 9:20 AM: Process Data

```bash
cd ~/Desktop/task-tracker
python3 task_tracker_FINAL.py
```

This creates:
- `dashboard.html` - Tabs for Claim Sheet, Decomp, Historic
- `approval_ui.html` - Review + send Slack DMs

### 9:25 AM: Review Dashboard

Open the generated files in your browser:
- `dashboard.html` - See activity overview
- `approval_ui.html` - Review silent people

### 9:30 AM: Push to GitHub

```bash
python3 github_push.py
```

This commits + pushes to GitHub. Takes 5-10 seconds.

Your dashboard is now live at:
```
https://aashnagoel.github.io/task-tracker/dashboard.html
```

**Share this URL with your team!** They can view the latest dashboard anytime.

## What Each Tab Shows

### Claim Sheet Activity
- Day-by-day task counts for each person
- Who's active vs silent (48+ hours inactive)
- Total tasks completed

### Decomp Progress
- Day-by-day decomp work by person
- Who's active vs silent on decomp
- Total decomp items

### Historic Data
- Coming soon: accumulate data across weeks
- Compare trends over time

## Slack DM Sending (Optional)

If you want to send check-in messages:

1. Open `approval_ui.html` in browser
2. See list of silent people
3. For each person:
   - **Green box** = Username found automatically, ready to send
   - **Yellow box** = Multiple matches, pick the right one
   - **Red box** = Not found, type username manually
4. Click checkboxes for who to message
5. Click "Send Messages via Slack"
6. Confirms will be sent to Slack as DMs

**Note:** Currently shows a confirmation popup. Backend Slack sending integration coming next.

## Troubleshooting

### Python command not found
Try:
```bash
python3.9 task_tracker_FINAL.py
# or
python task_tracker_FINAL.py
```

### Module not found (openpyxl, requests)
Run:
```bash
pip3 install openpyxl requests --upgrade
```

### Excel file not found
- Make sure file is in `~/Desktop/task-tracker/`
- Make sure filename ends with `.xlsx`
- Try running just one recent file

### Git push fails
Make sure you configured git:
```bash
git config --global user.name "aashnagoel"
git config --global user.email "aashnagoel1999@gmail.com"
```

### Dashboard not updating on GitHub
- Wait 1-2 minutes (GitHub Pages caches)
- Hard refresh browser: Cmd+Shift+R (Mac) or Ctrl+Shift+R (Windows)

## Your GitHub Pages URL

**Dashboard:**
```
https://aashnagoel.github.io/task-tracker/dashboard.html
```

**Approval UI:**
```
https://aashnagoel.github.io/task-tracker/approval_ui.html
```

**GitHub Repository:**
```
https://github.com/aashnagoel/task-tracker
```

## You're Ready! 🚀

Follow the daily workflow above each morning. The system is fully automated.

Questions? Check the README.md or look at the Python scripts.
