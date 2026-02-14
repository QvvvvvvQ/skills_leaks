#!/bin/bash
set -e

# ─────────────────────────────────────────────────────────────────────────────
# Configuration
# ─────────────

# 这是最早上线的默认模板，用户如果没指定模板，则使用这个模板
ORIGIN_TEMPLATE_NAME="0-origin"

SCRIPTS_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"

# 这个目录里面有 node_modules 的所有东西，是在 prepare 脚本里准备好的。运行时需要 copy 到 PROJECT_PATH
ORI_TEMPLATE_PATH="$SCRIPTS_DIR/template"

REPO_ROOT="$(cd "$SCRIPTS_DIR/.." && pwd)"
PROJECT_NAME="$1"
TEMPLATE_NAME="${2:-$ORIGIN_TEMPLATE_NAME}"
[[ "$OSTYPE" == "darwin"* ]] && SED_INPLACE="sed -i ''" || SED_INPLACE="sed -i"

TEMPLATE_PATH="$REPO_ROOT/templates/$TEMPLATE_NAME"

# ─────────────────────────────────────────────────────────────────────────────
# Validation
# ─────────────────────────────────────────────────────────────────────────────

! command -v npm &>/dev/null && { echo "Error: npm not found"; exit 1; }
[[ -z "$1" ]] && { echo "Usage: $0 <project-name>"; exit 1; }
[[ ! -d "$TEMPLATE_PATH" ]] && { echo "Error: template not found: $TEMPLATE_NAME"; exit 1; }
TEMPLATE_ZIP="$TEMPLATE_PATH/$TEMPLATE_NAME.zip"
INFO_MD="$TEMPLATE_PATH/info.md"
if [[ ! -f "$TEMPLATE_ZIP" ]]; then
    echo "Error: Template '$TEMPLATE_NAME' not found at: $TEMPLATE_ZIP"
    echo "Available templates:"
    ls "$REPO_ROOT/templates" 2>/dev/null | sed 's/^/  - /' || true
    exit 1
fi

PROJECT_PATH=${PROJECT_PATH:-"/mnt/okcomputer/output/app"}
TEMP_PATH=${TEMP_PATH:-"/tmp/temp-webapp"}

# ─────────────────────────────────────────────────────────────────────────────
# Project Creation
# ─────────────────────────────────────────────────────────────────────────────

echo "Creating project: $PROJECT_PATH"
mkdir -p "$PROJECT_PATH"
[[ -d "$TEMP_PATH" ]] && rm -rf "$TEMP_PATH"
mkdir -p "$TEMP_PATH"

# ─────────────────────────────────────────────────────────────────────────────
# Dependencies
# ─────────────────────────────────────────────────────────────────────────────

echo "Installing dependencies..."

# cp -r "$TEMPLATE_PATH"/* "$TEMP_PATH"/
unzip "$TEMPLATE_ZIP" -d "$TEMP_PATH" > /dev/null 2>&1
# 如果解压的目录带着模板名，则使用带着模板名的目录
if [[ -d "$TEMP_PATH/$TEMPLATE_NAME" ]]; then
    TEMP_PATH="$TEMP_PATH/$TEMPLATE_NAME"
fi

cp "$INFO_MD" "$TEMP_PATH"/info.md

ESCAPED_REPLACE=$(printf '%s\n' "$PROJECT_NAME" | sed 's/[\/&]/\\&/g')
$SED_INPLACE 's/<title>.*<\/title>/<title>'"$ESCAPED_REPLACE"'<\/title>/' "$TEMP_PATH"/index.html

# 这里需要保证没有 template_name 那一层；如果是 0-origin，则理论上会有 node_modules 那一层
cp -r "$TEMP_PATH"/* "$PROJECT_PATH"/

# 把 node_modules 复制到目标目录
cp -r "$ORI_TEMPLATE_PATH"/node_modules "$PROJECT_PATH"/

# 对于用户指定的非默认模板来说，copy 到目标目录之后，需要安装依赖
if [[ "$TEMPLATE_NAME" != "$ORIGIN_TEMPLATE_NAME" ]]; then
  cd "$PROJECT_PATH"
  npm install
fi

# ─────────────────────────────────────────────────────────────────────────────
# Done
# ─────────────────────────────────────────────────────────────────────────────
if [[ "$TEMPLATE_NAME" != "$ORIGIN_TEMPLATE_NAME" ]]; then
  echo ""
  echo "✓ Template '$TEMPLATE_NAME' extracted successfully to: $PROJECT_PATH"
  echo ""
  echo "═══════════════════════════════════════════════════════════════════════════"
  echo "TEMPLATE INFO"
  echo "═══════════════════════════════════════════════════════════════════════════"
  echo ""
  # Output info.md for the AI agent to read
  if [[ -f "$INFO_MD" ]]; then
    cat "$INFO_MD"
  else
    echo "⚠ No info.md found for this template."
  fi
  echo ""
  echo "═══════════════════════════════════════════════════════════════════════════"
else
  # 默认模板就直接输出info.md
  [[ -f "$INFO_MD" ]] && cat "$INFO_MD"
fi