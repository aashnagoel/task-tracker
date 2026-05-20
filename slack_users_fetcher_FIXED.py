#!/usr/bin/env python3
"""
Slack Users Fetcher - SIMPLIFIED VERSION
"""

import json
from pathlib import Path
import requests
import sys

def main():
    bot_token = "xoxb-10479801607904-11164856158566-gSb4JEMo1m7jERALkyydLNLj"
    workspace_name = "TSIP Contributors"
    
    # Get current directory (wherever you run this from)
    current_dir = Path.cwd()
    
    print(f"Current directory: {current_dir}")
    print(f"Fetching users from Slack workspace: {workspace_name}")
    
    # Fetch from Slack
    url = "https://slack.com/api/users.list"
    headers = {
        "Authorization": f"Bearer {bot_token}",
        "Content-Type": "application/json"
    }
    
    try:
        response = requests.get(url, headers=headers)
        data = response.json()
        
        if not data.get("ok"):
            print(f"ERROR: Slack API error: {data.get('error')}")
            return
        
        users = data.get("members", [])
        
        # Filter and process
        processed_users = []
        for user in users:
            if user.get("is_bot") or user.get("deleted"):
                continue
            
            processed_users.append({
                "id": user.get("id"),
                "username": user.get("name"),
                "real_name": user.get("real_name"),
                "display_name": user.get("profile", {}).get("display_name"),
            })
        
        processed_users.sort(key=lambda x: x["username"])
        
        # Save to file
        cache_file = current_dir / "slack_users_tsip_contributors.json"
        
        with open(cache_file, 'w') as f:
            json.dump({
                "workspace": workspace_name,
                "user_count": len(processed_users),
                "users": processed_users
            }, f, indent=2)
        
        print(f"\n✓ Successfully fetched {len(processed_users)} users from Slack")
        print(f"✓ Saved to: {cache_file}")
        print(f"\nFirst 10 users:")
        for user in processed_users[:10]:
            print(f"  - {user['username']:25} ({user['real_name']})")
        if len(processed_users) > 10:
            print(f"  ... and {len(processed_users) - 10} more")
        
    except Exception as e:
        print(f"ERROR: {e}")
        return

if __name__ == "__main__":
    main()
