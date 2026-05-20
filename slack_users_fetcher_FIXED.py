#!/usr/bin/env python3
"""
Slack Users Fetcher - SECURE VERSION
Fetches list of Slack users and caches them locally
Token comes from environment variable, NOT hardcoded
"""

import json
import os
from pathlib import Path

def fetch_slack_users(bot_token):
    """Fetch Slack users using bot token"""
    try:
        import requests
    except ImportError:
        print("ERROR: requests library not found. Run: pip3 install requests")
        return []
    
    url = "https://slack.com/api/users.list"
    headers = {"Authorization": f"Bearer {bot_token}"}
    
    try:
        response = requests.get(url, headers=headers, timeout=10)
        data = response.json()
        
        if not data.get("ok"):
            print(f"ERROR: {data.get('error', 'Unknown error')}")
            return []
        
        users = data.get("members", [])
        print(f"✓ Fetched {len(users)} Slack users")
        return users
    
    except Exception as e:
        print(f"ERROR: {e}")
        return []

def save_users_cache(users):
    """Save users to local cache file"""
    cache_file = Path.cwd() / "slack_users_tsip_contributors.json"
    cache_data = {"users": users, "count": len(users)}
    
    with open(cache_file, 'w') as f:
        json.dump(cache_data, f, indent=2)
    
    print(f"✓ Saved cache to {cache_file.name}")

def main():
    print("\n=== Slack Users Fetcher ===")
    
    # Get token from environment variable (SECURE - not hardcoded)
    bot_token = os.getenv("SLACK_BOT_TOKEN")
    
    if not bot_token:
        print("ERROR: SLACK_BOT_TOKEN environment variable not set")
        print("\nTo use this script:")
        print("1. Set your token: export SLACK_BOT_TOKEN='xoxb-...'")
        print("2. Run: python3 slack_users_fetcher_FIXED.py")
        return
    
    # Fetch and cache users
    users = fetch_slack_users(bot_token)
    if users:
        save_users_cache(users)
        print(f"\n✓ COMPLETE!")
    else:
        print("ERROR: Could not fetch users")

if __name__ == "__main__":
    main()
