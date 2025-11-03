#!/bin/bash
# Deploy script that injects timestamp before pushing to Apps Script

set -e

# Generate timestamp
TIMESTAMP=$(date -u +"%Y-%m-%d %H:%M:%S UTC")
echo "Injecting deployment timestamp: $TIMESTAMP"

# Update the DEPLOYED_AT constant in Code.js
sed -i.bak "s/DEPLOYED_AT: '.*'/DEPLOYED_AT: '$TIMESTAMP'/" apps-script/Code.js

# Push to Apps Script
echo "Pushing to Apps Script..."
cd apps-script && clasp push

# Restore the backup (keep git clean)
cd ..
mv apps-script/Code.js.bak apps-script/Code.js

echo "Deployment complete! Deployed at: $TIMESTAMP"
echo "Remember to RELOAD the spreadsheet to get the new version!"
