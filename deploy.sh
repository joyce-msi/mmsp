#!/bin/bash
# Sync from Jira, update public data, and deploy to GitHub Pages

set -e

echo "Syncing from Jira..."
npm run sync

echo "Copying data to public folder..."
cp data/*.json public/data/

echo "Committing and pushing..."
git add data/ public/data/
git commit -m "Update Jira data $(date +%Y-%m-%d)" || git commit --allow-empty -m "Redeploy $(date +%Y-%m-%d)"
git push

echo "Done! Site will update in ~1 minute at https://joyce-msi.github.io/mmsp/"
