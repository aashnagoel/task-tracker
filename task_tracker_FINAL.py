#!/usr/bin/env python3
"""
Task Tracker - FINAL VERSION with Persistent Notes on Dashboard
Reads Claim Sheet + Model Decomp, generates dashboards with tabs
Detects silent people, shows editable notes column on dashboard
"""

import json
from datetime import datetime, timedelta
from pathlib import Path
import openpyxl
from collections import defaultdict
from difflib import SequenceMatcher

# Notes storage file
NOTES_FILE = Path.cwd() / "tasker_notes.json"

def load_notes():
    """Load existing notes about taskers"""
    if NOTES_FILE.exists():
        with open(NOTES_FILE) as f:
            return json.load(f)
    return {}

def load_slack_users():
    """Load Slack users from cache"""
    cache_file = Path.cwd() / "slack_users_tsip_contributors.json"
    if not cache_file.exists():
        print(f"ERROR: slack_users_tsip_contributors.json not found")
        return []
    with open(cache_file) as f:
        return json.load(f).get("users", [])

def find_latest_excel():
    """Find most recent Excel file"""
    excel_files = list(Path.cwd().glob("*.xlsx"))
    if not excel_files:
        print(f"ERROR: No Excel files found")
        return None
    return max(excel_files, key=lambda p: p.stat().st_mtime)

def extract_first_name(full_name):
    """Extract first name from 'First Last' format"""
    if not full_name:
        return ""
    parts = str(full_name).strip().split()
    return parts[0] if parts else ""

def read_claim_sheet(excel_path):
    """Read Claim Sheet"""
    wb = openpyxl.load_workbook(excel_path)
    sheet = wb["Claim Sheet"]
    
    headers = {}
    for col_num, cell in enumerate(sheet[1], 1):
        if cell.value:
            headers[cell.value] = col_num
    
    tasker_col = headers.get("Claimed By")
    date_col = headers.get("Date fixed")
    completion_col = headers.get("Task Completed")
    
    if not all([tasker_col, date_col, completion_col]):
        return None, []
    
    tasker_by_date = defaultdict(lambda: defaultdict(int))
    all_taskers = set()
    
    for row_num in range(2, sheet.max_row + 1):
        tasker_name = sheet.cell(row_num, tasker_col).value
        task_date = sheet.cell(row_num, date_col).value
        is_completed = sheet.cell(row_num, completion_col).value
        
        if not all([tasker_name, task_date, is_completed]):
            continue
        
        tasker_name = str(tasker_name).strip()
        all_taskers.add(tasker_name)
        
        if isinstance(task_date, datetime):
            task_date = task_date.date()
        elif isinstance(task_date, str):
            try:
                task_date = datetime.strptime(task_date, "%Y-%m-%d").date()
            except:
                continue
        
        tasker_by_date[tasker_name][task_date] += 1
    
    return tasker_by_date, sorted(all_taskers)

def read_decomp_sheet(excel_path):
    """Read Model Decomp Sheet"""
    wb = openpyxl.load_workbook(excel_path)
    if "Model Decomp" not in wb.sheetnames:
        return None, []
    
    sheet = wb["Model Decomp"]
    
    headers = {}
    for col_num, cell in enumerate(sheet[1], 1):
        if cell.value:
            headers[cell.value] = col_num
    
    tasker_col = headers.get("Claimed By")
    prompt_date_col = headers.get("Date Prompt Com")
    completion_date_col = headers.get("Date Completed")
    
    if not tasker_col:
        return None, []
    
    tasker_by_date = defaultdict(lambda: defaultdict(int))
    all_taskers = set()
    
    for row_num in range(2, sheet.max_row + 1):
        tasker_name = sheet.cell(row_num, tasker_col).value
        
        if not tasker_name:
            continue
        
        tasker_name = str(tasker_name).strip()
        all_taskers.add(tasker_name)
        
        latest_date = None
        
        if prompt_date_col:
            prompt_date = sheet.cell(row_num, prompt_date_col).value
            if prompt_date:
                if isinstance(prompt_date, datetime):
                    latest_date = prompt_date.date()
                elif isinstance(prompt_date, str):
                    try:
                        latest_date = datetime.strptime(prompt_date, "%Y-%m-%d").date()
                    except:
                        pass
        
        if completion_date_col:
            completion_date = sheet.cell(row_num, completion_date_col).value
            if completion_date:
                if isinstance(completion_date, datetime):
                    comp_date = completion_date.date()
                elif isinstance(completion_date, str):
                    try:
                        comp_date = datetime.strptime(completion_date, "%Y-%m-%d").date()
                    except:
                        comp_date = None
                else:
                    comp_date = None
                
                if comp_date and (latest_date is None or comp_date > latest_date):
                    latest_date = comp_date
        
        if latest_date:
            tasker_by_date[tasker_name][latest_date] += 1
    
    return tasker_by_date, sorted(all_taskers)

def get_silent_taskers(tasker_by_date, all_taskers):
    """Find silent people (48+ hours)"""
    silent_threshold = timedelta(hours=48)
    check_time = datetime.now()
    silent_list = []
    
    for tasker in all_taskers:
        dates = tasker_by_date.get(tasker, {})
        if not dates:
            continue
        
        last_task_date = max(dates.keys())
        time_since_task = check_time - datetime.combine(last_task_date, datetime.min.time())
        
        if time_since_task > silent_threshold:
            first_name = extract_first_name(tasker)
            silent_list.append({
                "name": tasker,
                "first_name": first_name,
                "last_task_date": last_task_date.isoformat(),
                "hours_silent": int(time_since_task.total_seconds() / 3600),
                "days_silent": round(time_since_task.total_seconds() / 86400, 1)
            })
    
    return sorted(silent_list, key=lambda x: x["hours_silent"], reverse=True)

def generate_dashboards(claim_data, claim_taskers, decomp_data, decomp_taskers, slack_users):
    """Generate dashboard with editable notes column"""
    
    notes = load_notes()
    
    # Get all dates from both sheets
    all_dates = set()
    for dates_dict in claim_data.values():
        all_dates.update(dates_dict.keys())
    for dates_dict in decomp_data.values():
        all_dates.update(dates_dict.keys())
    
    if not all_dates:
        print("No task data found")
        return
    
    sorted_dates = sorted(all_dates)
    date_headers = [d.strftime("%a %m/%d") for d in sorted_dates]
    
    # Get silent taskers
    claim_silent = get_silent_taskers(claim_data, claim_taskers)
    decomp_silent = get_silent_taskers(decomp_data, decomp_taskers)
    
    # === DASHBOARD with NOTES COLUMN ===
    html_dashboard = f"""<!DOCTYPE html>
<html>
<head>
    <title>Task Tracker Dashboard</title>
    <style>
        * {{ margin: 0; padding: 0; box-sizing: border-box; }}
        body {{ font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif; background: #f5f5f5; padding: 20px; }}
        .container {{ max-width: 1600px; margin: 0 auto; background: white; padding: 32px; border-radius: 8px; box-shadow: 0 2px 8px rgba(0,0,0,0.08); }}
        h1 {{ font-size: 28px; font-weight: 600; color: #1a1a1a; margin-bottom: 8px; }}
        .meta {{ color: #666; font-size: 14px; margin-bottom: 24px; }}
        .tabs {{ display: flex; border-bottom: 2px solid #e0e0e0; margin-bottom: 24px; gap: 0; }}
        .tab-button {{ padding: 12px 20px; background: none; border: none; cursor: pointer; font-size: 15px; font-weight: 500; color: #666; border-bottom: 3px solid transparent; margin-bottom: -2px; }}
        .tab-button.active {{ color: #1a1a1a; border-bottom-color: #3b82f6; }}
        .tab-button:hover {{ color: #1a1a1a; }}
        .tab-content {{ display: none; }}
        .tab-content.active {{ display: block; }}
        .stats {{ display: grid; grid-template-columns: repeat(auto-fit, minmax(150px, 1fr)); gap: 16px; margin-bottom: 24px; }}
        .stat-card {{ background: #f9f9f9; padding: 16px; border-radius: 6px; border-left: 4px solid #3b82f6; }}
        .stat-label {{ color: #666; font-size: 12px; text-transform: uppercase; margin-bottom: 6px; }}
        .stat-value {{ font-size: 24px; font-weight: 600; color: #1a1a1a; }}
        table {{ width: 100%; border-collapse: collapse; }}
        th, td {{ padding: 12px; text-align: left; border-bottom: 1px solid #e0e0e0; word-break: break-word; }}
        th {{ background: #f9f9f9; font-weight: 600; color: #1a1a1a; }}
        th.date {{ text-align: center; width: 70px; }}
        td.date {{ text-align: center; font-weight: 500; color: #0066cc; }}
        tr:hover {{ background: #fafafa; }}
        .silent {{ background: #fffbea; }}
        .active {{ background: #f0fdf4; }}
        .badge {{ display: inline-block; padding: 4px 8px; border-radius: 4px; font-size: 12px; font-weight: 500; }}
        .badge.warning {{ background: #fcd34d; color: #7c2d12; }}
        .badge.success {{ background: #bbf7d0; color: #166534; }}
        .notes-input {{ width: 100%; padding: 6px; border: 1px solid #ddd; border-radius: 3px; font-size: 12px; }}
        .notes-input:focus {{ outline: none; border-color: #3b82f6; background: #eff6ff; }}
        .notes-cell {{ min-width: 150px; }}
        .save-indicator {{ font-size: 11px; color: #10b981; margin-top: 2px; }}
    </style>
</head>
<body>
    <div class="container">
        <h1>Task Tracker Dashboard</h1>
        <div class="meta">Generated: {datetime.now().strftime('%A, %B %d, %Y at %I:%M %p')} EST</div>
        
        <div class="tabs">
            <button class="tab-button active" onclick="openTab(event, 'claim')">Claim Sheet Activity</button>
            <button class="tab-button" onclick="openTab(event, 'decomp')">Decomp Progress</button>
            <button class="tab-button" onclick="openTab(event, 'historic')">Historic Data</button>
        </div>
        
        <!-- CLAIM SHEET TAB -->
        <div id="claim" class="tab-content active">
            <div class="stats">
                <div class="stat-card">
                    <div class="stat-label">Total Taskers</div>
                    <div class="stat-value">{len(claim_taskers)}</div>
                </div>
                <div class="stat-card">
                    <div class="stat-label">Active</div>
                    <div class="stat-value">{len(claim_taskers) - len(claim_silent)}</div>
                </div>
                <div class="stat-card">
                    <div class="stat-label">Silent (48+hrs)</div>
                    <div class="stat-value" style="color: #dc2626;">{len(claim_silent)}</div>
                </div>
                <div class="stat-card">
                    <div class="stat-label">Total Tasks</div>
                    <div class="stat-value">{sum(sum(dates.values()) for dates in claim_data.values())}</div>
                </div>
            </div>
            
            <table>
                <thead>
                    <tr>
                        <th>Tasker Name</th>
"""
    
    for date_header in date_headers:
        html_dashboard += f"                        <th class='date'>{date_header}</th>\n"
    
    html_dashboard += """                        <th class='date'>Total</th>
                        <th>Status</th>
                        <th class="notes-cell">Notes</th>
                    </tr>
                </thead>
                <tbody>
"""
    
    silent_names = {s["name"] for s in claim_silent}
    
    for tasker in sorted(claim_taskers):
        row_class = "silent" if tasker in silent_names else "active"
        tasker_notes = notes.get(tasker, "")
        
        html_dashboard += f"                    <tr class='{row_class}'>\n"
        html_dashboard += f"                        <td>{tasker}</td>\n"
        
        daily_counts = []
        for date in sorted_dates:
            count = claim_data[tasker].get(date, 0)
            daily_counts.append(count)
            html_dashboard += f"                        <td class='date'>{count if count > 0 else '-'}</td>\n"
        
        total = sum(daily_counts)
        html_dashboard += f"                        <td class='date'>{total}</td>\n"
        
        if tasker in silent_names:
            html_dashboard += f"                        <td><span class='badge warning'>⚠ Silent</span></td>\n"
        else:
            html_dashboard += f"                        <td><span class='badge success'>✓ Active</span></td>\n"
        
        # Notes input
        html_dashboard += f"""                        <td class="notes-cell">
                            <input type="text" class="notes-input" data-tasker="{tasker}" value="{tasker_notes}" placeholder="Add note..." onchange="saveNote(this)">
                            <div class="save-indicator" style="display:none;">✓ Saved</div>
                        </td>
"""
        html_dashboard += "                    </tr>\n"
    
    html_dashboard += """                </tbody>
            </table>
        </div>
        
        <!-- DECOMP TAB -->
        <div id="decomp" class="tab-content">
            <div class="stats">
                <div class="stat-card">
                    <div class="stat-label">Total Decomp Tasks</div>
                    <div class="stat-value">""" + str(len(decomp_taskers)) + """</div>
                </div>
                <div class="stat-card">
                    <div class="stat-label">Active</div>
                    <div class="stat-value">""" + str(len(decomp_taskers) - len(decomp_silent)) + """</div>
                </div>
                <div class="stat-card">
                    <div class="stat-label">Silent (48+hrs)</div>
                    <div class="stat-value" style="color: #dc2626;">""" + str(len(decomp_silent)) + """</div>
                </div>
                <div class="stat-card">
                    <div class="stat-label">Total Decomp Items</div>
                    <div class="stat-value">""" + str(sum(sum(dates.values()) for dates in decomp_data.values())) + """</div>
                </div>
            </div>
            
            <table>
                <thead>
                    <tr>
                        <th>Tasker Name</th>
"""
    
    for date_header in date_headers:
        html_dashboard += f"                        <th class='date'>{date_header}</th>\n"
    
    html_dashboard += """                        <th class='date'>Total</th>
                        <th>Status</th>
                        <th class="notes-cell">Notes</th>
                    </tr>
                </thead>
                <tbody>
"""
    
    silent_decomp_names = {s["name"] for s in decomp_silent}
    
    for tasker in sorted(decomp_taskers):
        row_class = "silent" if tasker in silent_decomp_names else "active"
        tasker_notes = notes.get(tasker, "")
        
        html_dashboard += f"                    <tr class='{row_class}'>\n"
        html_dashboard += f"                        <td>{tasker}</td>\n"
        
        daily_counts = []
        for date in sorted_dates:
            count = decomp_data[tasker].get(date, 0)
            daily_counts.append(count)
            html_dashboard += f"                        <td class='date'>{count if count > 0 else '-'}</td>\n"
        
        total = sum(daily_counts)
        html_dashboard += f"                        <td class='date'>{total}</td>\n"
        
        if tasker in silent_decomp_names:
            html_dashboard += f"                        <td><span class='badge warning'>⚠ Silent</span></td>\n"
        else:
            html_dashboard += f"                        <td><span class='badge success'>✓ Active</span></td>\n"
        
        # Notes input
        html_dashboard += f"""                        <td class="notes-cell">
                            <input type="text" class="notes-input" data-tasker="{tasker}" value="{tasker_notes}" placeholder="Add note..." onchange="saveNote(this)">
                            <div class="save-indicator" style="display:none;">✓ Saved</div>
                        </td>
"""
        html_dashboard += "                    </tr>\n"
    
    html_dashboard += """                </tbody>
            </table>
        </div>
        
        <!-- HISTORIC DATA TAB -->
        <div id="historic" class="tab-content">
            <p style="color: #666; padding: 20px; text-align: center;">Historic data will appear here as you track more weeks.</p>
        </div>
    </div>
    
    <script>
        function openTab(evt, tabName) {
            var i, tabcontent, tabbuttons;
            tabcontent = document.getElementsByClassName("tab-content");
            for (i = 0; i < tabcontent.length; i++) {
                tabcontent[i].classList.remove("active");
            }
            tabbuttons = document.getElementsByClassName("tab-button");
            for (i = 0; i < tabbuttons.length; i++) {
                tabbuttons[i].classList.remove("active");
            }
            document.getElementById(tabName).classList.add("active");
            evt.currentTarget.classList.add("active");
        }
        
        function saveNote(element) {
            const tasker = element.getAttribute('data-tasker');
            const noteText = element.value;
            const indicator = element.parentElement.querySelector('.save-indicator');
            
            // Save to localStorage (persists in browser)
            const notes = JSON.parse(localStorage.getItem('taskerNotes') || '{{}}');
            notes[tasker] = noteText;
            localStorage.setItem('taskerNotes', JSON.stringify(notes));
            
            // Show saved indicator
            indicator.style.display = 'block';
            setTimeout(() => {
                indicator.style.display = 'none';
            }, 2000);
        }
        
        // Load notes from localStorage on page load
        window.onload = function() {
            const notes = JSON.parse(localStorage.getItem('taskerNotes') || '{{}}');
            document.querySelectorAll('.notes-input').forEach(input => {
                const tasker = input.getAttribute('data-tasker');
                if (notes[tasker]) {
                    input.value = notes[tasker];
                }
            });
        };
    </script>
</body>
</html>"""
    
    # Write dashboard
    dashboard_path = Path.cwd() / "dashboard.html"
    with open(dashboard_path, 'w') as f:
        f.write(html_dashboard)
    print(f"✓ Generated dashboard.html with notes column")

def main():
    print(f"\n=== Task Tracker Processor ===")
    print(f"Working directory: {Path.cwd()}")
    
    # Load Slack users
    slack_users = load_slack_users()
    if not slack_users:
        print("ERROR: Could not load Slack users")
        return
    print(f"✓ Loaded {len(slack_users)} Slack users")
    
    # Find Excel
    excel_path = find_latest_excel()
    if not excel_path:
        return
    print(f"✓ Found Excel: {excel_path.name}")
    
    # Read Claim Sheet
    claim_data, claim_taskers = read_claim_sheet(excel_path)
    if claim_data is None:
        return
    print(f"✓ Claim Sheet: {len(claim_taskers)} taskers")
    
    # Read Decomp Sheet
    decomp_data, decomp_taskers = read_decomp_sheet(excel_path)
    if decomp_data is None:
        decomp_data = {}
        decomp_taskers = []
    print(f"✓ Model Decomp: {len(decomp_taskers)} taskers")
    
    # Generate dashboards
    print(f"\n✓ Generating dashboards...")
    generate_dashboards(claim_data, claim_taskers, decomp_data, decomp_taskers, slack_users)
    
    print(f"\n✓ COMPLETE!")
    print(f"Dashboard: https://aashnagoel.github.io/task-tracker/dashboard.html")

if __name__ == "__main__":
    main()
