#!/usr/bin/env python3
"""
Slack Message Server
Sends DMs via Slack API and saves notes about silent taskers
Run this in background: python3 slack_server.py
"""

import json
import os
from pathlib import Path
from flask import Flask, request, jsonify
import requests
from datetime import datetime

app = Flask(__name__)

NOTES_FILE = Path.cwd() / "silent_tasker_notes.json"

def load_notes():
    """Load existing notes"""
    if NOTES_FILE.exists():
        with open(NOTES_FILE) as f:
            return json.load(f)
    return {}

def save_notes(notes):
    """Save notes to file"""
    with open(NOTES_FILE, 'w') as f:
        json.dump(notes, f, indent=2)

def get_slack_user_id(username, bot_token):
    """Get Slack user ID from username"""
    url = "https://slack.com/api/users.lookupByEmail"
    
    # Try email first
    headers = {"Authorization": f"Bearer {bot_token}"}
    response = requests.post(url, headers=headers, json={"email": f"{username}@slack.com"})
    data = response.json()
    
    if data.get("ok"):
        return data["user"]["id"]
    
    # Try username search
    url = "https://slack.com/api/users.list"
    response = requests.get(url, headers=headers)
    data = response.json()
    
    if data.get("ok"):
        for user in data["members"]:
            if user.get("name", "").lower() == username.lower():
                return user["id"]
    
    return None

def send_slack_dm(user_id, message, bot_token):
    """Send DM via Slack API"""
    url = "https://slack.com/api/conversations.open"
    headers = {"Authorization": f"Bearer {bot_token}"}
    
    # Open DM conversation
    response = requests.post(url, headers=headers, json={"users": user_id})
    data = response.json()
    
    if not data.get("ok"):
        return False, data.get("error", "Failed to open DM")
    
    channel_id = data["channel"]["id"]
    
    # Send message
    url = "https://slack.com/api/chat.postMessage"
    response = requests.post(
        url,
        headers=headers,
        json={"channel": channel_id, "text": message}
    )
    data = response.json()
    
    return data.get("ok"), data.get("error", "Message sent" if data.get("ok") else "Failed")

@app.route('/send_slack_messages', methods=['POST'])
def send_slack_messages():
    """Handle sending Slack messages"""
    try:
        data = request.get_json()
        people = data.get("people", [])
        
        # Get bot token from environment
        bot_token = os.getenv("SLACK_BOT_TOKEN")
        if not bot_token:
            return jsonify({"error": "SLACK_BOT_TOKEN not set"}), 500
        
        notes = load_notes()
        results = []
        
        for person in people:
            name = person.get("name")
            reason = person.get("reason", "")
            
            # Extract first name
            first_name = name.split()[0] if name else name
            
            # Create message
            message = f"Hey {first_name}, I noticed you haven't submitted any tasks in the past couple of days. If you're facing any blockers or have questions, let me know and I'm happy to help!"
            
            # Get Slack user ID (try multiple formats)
            user_id = None
            for attempt_name in [name.lower().replace(" ", ""), first_name.lower()]:
                user_id = get_slack_user_id(attempt_name, bot_token)
                if user_id:
                    break
            
            if not user_id:
                results.append({"name": name, "success": False, "error": "Could not find Slack user"})
                continue
            
            # Send DM
            success, error = send_slack_dm(user_id, message, bot_token)
            
            if success:
                # Save notes
                notes[name] = {
                    "reason": reason,
                    "last_outreach": datetime.now().strftime("%Y-%m-%d %H:%M"),
                    "message_sent": True
                }
                results.append({"name": name, "success": True})
            else:
                results.append({"name": name, "success": False, "error": error})
        
        # Save all notes
        save_notes(notes)
        
        return jsonify({
            "success": True,
            "results": results,
            "sent": sum(1 for r in results if r.get("success"))
        })
    
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route('/health', methods=['GET'])
def health():
    """Health check"""
    return jsonify({"status": "ok"})

if __name__ == "__main__":
    print("\n=== Slack Message Server ===")
    print("Starting Flask server for Slack message sending...")
    print("Make sure SLACK_BOT_TOKEN is set: export SLACK_BOT_TOKEN='xoxb-...'")
    app.run(port=5000, debug=False)
