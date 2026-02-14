#!/bin/bash
set -e

SCRIPTS_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
TEMPLATE_PATH="$SCRIPTS_DIR/../templates/0-origin"
ORI_TEMPLATE_PATH="$SCRIPTS_DIR/template"
mkdir -p "$ORI_TEMPLATE_PATH"

echo "Installing dependencies..."

cd "$TEMPLATE_PATH"

[[ -f "0-origin.zip" ]] && unzip 0-origin.zip > /dev/null 2>&1
if [[ -d "0-origin" ]]; then
  TEMPLATE_PATH="$TEMPLATE_PATH/0-origin"
fi

cd "$TEMPLATE_PATH"
npm install

# copy all files, especially node_modules, to origin template path
# PATH should contain "/app/.kimi/skills/webapp-building/scripts/template/node_modules/.bin" etc.
cp -r "$TEMPLATE_PATH"/* "$ORI_TEMPLATE_PATH"/

echo ""
echo "Setup complete: $TEMPLATE_PATH"
echo ""
