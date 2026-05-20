# Task Tracker Dashboard

Automated task tracking and Slack check-in system for data labeling team.

## Features

- **Claim Sheet Activity** - Track daily task completions
- **Decomp Progress** - Monitor model decomposition work
- **Silent Detection** - Automatically flag people inactive for 48+ hours
- **Slack DMs** - Send personalized check-in messages
- **GitHub Pages** - Share dashboard with team via URL
- **Multi-user Support** - Track 100+ contributors across multiple task types

## Dashboard URL

**Share this link with your team:**
```
https://aashnagoel.github.io/task-tracker/dashboard.html
```

Updated daily after processing your Excel files.

## Daily Workflow

### 1. Download Excel

Export your task sheet as `.xlsx` and save to this folder:
```
~/Desktop/task-tracker/
```

### 2. Process Data

Open Terminal and run:
```bash
cd ~/Desktop/task-tracker
python3 task_tracker_FINAL.py
```

This generates:
- `dashboard.html` - Activity overview (Claim Sheet, Decomp, Historic tabs)
- `approval_ui.html` - Silent people with message previews

### 3. Send Slack Messages (Optional)

Open `approval_ui.html` in your browser:
- Review people inactive for 48+ hours
- Confirm or edit Slack usernames
- Preview personalized messages
- Click "Send Messages via Slack"

### 4. Push to GitHub

Automatically updates the public dashboard:
```bash
python3 github_push.py
```

Team sees updates instantly at the dashboard URL.

## File Structure

```
~/Desktop/task-tracker/
├── task_tracker_FINAL.py              # Main processor
├── slack_users_fetcher_FIXED.py       # Fetch Slack user list
├── github_push.py                      # Auto-push to GitHub
├── slack_users_tsip_contributors.json  # Slack user cache
├── dashboard.html                      # Generated daily
├── approval_ui.html                    # Generated daily
├── week_2026_05_24.xlsx               # Your Excel file
└── README.md                           # This file
```

## Setup (One-Time)

### Prerequisites

- Python 3.6+
- Git installed
- GitHub account with this repository
- Slack bot token

### Initial Setup

1. Clone this repository to your laptop:
   ```bash
   git clone https://github.com/aashnagoel/task-tracker.git
   cd task-tracker
   ```

2. Install Python dependencies:
   ```bash
   pip3 install openpyxl requests
   ```

3. Fetch Slack users once:
   ```bash
   python3 slack_users_fetcher_FIXED.py
   ```

4. Configure git:
   ```bash
   git config --global user.name "aashnagoel"
   git config --global user.email "aashnagoel1999@gmail.com"
   ```

Done! Now use the daily workflow above.

## How It Works

1. **Excel Processing** - Reads Claim Sheet + Model Decomp
2. **Silent Detection** - Identifies people with no activity in 48+ hours
3. **Slack Matching** - Uses fuzzy matching to find Slack usernames
4. **Dashboard Generation** - Creates HTML files for viewing
5. **GitHub Push** - Auto-uploads to make shareable

## Message Template

Check-in message for inactive people:
```
Hey {FirstName}, I noticed you haven't submitted any tasks in the past couple of days. If you're facing any blockers or have questions, let me know and I'm happy to help!
```

## Slack Features

- ✅ Fuzzy username matching (auto-finds users)
- ✅ Manual username override (if matching fails)
- ✅ Message preview before sending
- ✅ Prevents duplicate messages (5-day cooldown)
- ✅ First name personalization

## Troubleshooting

### Excel not found
Make sure your `.xlsx` file is in `~/Desktop/task-tracker/`

### Slack users cache error
Run: `python3 slack_users_fetcher_FIXED.py` to refresh

### Git push fails
Check git is configured:
```bash
git config --global user.email
git config --global user.name
```

### Dashboard not updating on GitHub
Wait 1-2 minutes for GitHub Pages to refresh after pushing.

## Version History

- **v1.0** - Initial release with Claim Sheet + Decomp tracking, Slack DMs, GitHub Pages
