#!/bin/bash
cd ~/Desktop/task-tracker
git add -A
git commit -m "Update dashboards - $(date)"
git push https://aashnagoel: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@github.com/aashnagoel/task-tracker.git main
echo "✓ Pushed to GitHub!"
