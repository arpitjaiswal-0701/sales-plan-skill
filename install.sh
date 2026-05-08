#!/bin/bash
# ============================================================================
# /sales plan — Adobe Lite DX Business Plan Skill Installer
# ============================================================================
set -e

RED='\033[0;31m'
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
BLUE='\033[0;34m'
CYAN='\033[0;36m'
NC='\033[0m'

echo ""
echo -e "${BLUE}╔══════════════════════════════════════════════════════════════╗${NC}"
echo -e "${BLUE}║                                                              ║${NC}"
echo -e "${BLUE}║${NC}   ${CYAN}/sales plan — Adobe Lite DX Business Plan Skill${NC}           ${BLUE}║${NC}"
echo -e "${BLUE}║${NC}   ${GREEN}1 skill · 1 Python script · Auto-populated PPT${NC}            ${BLUE}║${NC}"
echo -e "${BLUE}║                                                              ║${NC}"
echo -e "${BLUE}╚══════════════════════════════════════════════════════════════╝${NC}"
echo ""

# ---------------------------------------------------------------------------
# Locate source files (works for both local clone and curl | bash)
# ---------------------------------------------------------------------------
GITHUB_REPO="arpitjaiswal-0701/sales-plan-skill"
TEMP_DIR=""

if [ -n "${BASH_SOURCE[0]:-}" ] && [ "${BASH_SOURCE[0]}" != "bash" ]; then
    SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
    if [ -f "$SCRIPT_DIR/install.sh" ] && [ -d "$SCRIPT_DIR/skills" ]; then
        SOURCE_DIR="$SCRIPT_DIR"
        echo -e "${GREEN}Installing from local directory:${NC} $SOURCE_DIR"
    fi
fi

if [ -z "${SOURCE_DIR:-}" ]; then
    echo -e "${YELLOW}Cloning from GitHub...${NC}"
    TEMP_DIR=$(mktemp -d)
    if command -v git &>/dev/null; then
        git clone --depth 1 "https://github.com/$GITHUB_REPO.git" "$TEMP_DIR/repo" 2>/dev/null
        SOURCE_DIR="$TEMP_DIR/repo"
        echo -e "${GREEN}Cloned successfully.${NC}"
    else
        echo -e "${RED}Error: git is required. Install git and try again.${NC}"
        exit 1
    fi
fi

# ---------------------------------------------------------------------------
# Prerequisites check
# ---------------------------------------------------------------------------
echo -e "${BLUE}Checking prerequisites...${NC}"

if command -v claude &>/dev/null; then
    echo -e "  ${GREEN}✓${NC} Claude Code found"
else
    echo -e "  ${YELLOW}⚠${NC}  Claude Code CLI not found — skills will be installed but won't run without it"
    echo -e "      Get it at: https://claude.ai/download"
fi

PYTHON_CMD=""
for cmd in python3 python py; do
    if command -v "$cmd" &>/dev/null && "$cmd" -c "import sys; assert sys.version_info >= (3,8)" 2>/dev/null; then
        PYTHON_CMD="$cmd"
        echo -e "  ${GREEN}✓${NC} Python found: $cmd ($(${cmd} --version 2>&1))"
        break
    fi
done

if [ -z "$PYTHON_CMD" ]; then
    echo -e "  ${RED}✗${NC}  Python 3.8+ not found — required for PPT generation"
    echo -e "      Install from https://www.python.org/downloads/"
    exit 1
fi

# Check python-pptx
if "$PYTHON_CMD" -c "import pptx" 2>/dev/null; then
    echo -e "  ${GREEN}✓${NC} python-pptx installed"
else
    echo -e "  ${YELLOW}⚠${NC}  python-pptx not installed — installing now..."
    "$PYTHON_CMD" -m pip install python-pptx lxml -q
    echo -e "  ${GREEN}✓${NC} python-pptx installed"
fi

# ---------------------------------------------------------------------------
# Prompt: PPT template path
# ---------------------------------------------------------------------------
echo ""
echo -e "${BLUE}Configuration${NC}"
echo ""

DEFAULT_TEMPLATE_MAC="$HOME/Library/CloudStorage/OneDrive-Adobe/Lite DX Business Plan Template - Arpit.pptx"
DEFAULT_TEMPLATE_WIN="$USERPROFILE/OneDrive - Adobe/Lite DX Business Plan Template - Arpit.pptx"

echo -e "  Enter the full path to your ${CYAN}Lite DX Business Plan Template .pptx${NC} file."
echo -e "  (This is your personal copy synced from OneDrive.)"
echo ""
if [ -f "$DEFAULT_TEMPLATE_MAC" ]; then
    SUGGESTED="$DEFAULT_TEMPLATE_MAC"
elif [ -f "$DEFAULT_TEMPLATE_WIN" ]; then
    SUGGESTED="$DEFAULT_TEMPLATE_WIN"
else
    SUGGESTED=""
fi

if [ -n "$SUGGESTED" ]; then
    echo -e "  Detected: ${CYAN}${SUGGESTED}${NC}"
    read -r -p "  Use this path? [Y/n]: " USE_DETECTED
    if [[ "$USE_DETECTED" =~ ^[Nn]$ ]]; then
        read -r -p "  Template path: " TEMPLATE_PATH
    else
        TEMPLATE_PATH="$SUGGESTED"
    fi
else
    read -r -p "  Template path: " TEMPLATE_PATH
fi

if [ ! -f "$TEMPLATE_PATH" ]; then
    echo -e "  ${RED}✗${NC}  File not found: $TEMPLATE_PATH"
    echo -e "  Make sure OneDrive is synced and the path is correct, then re-run the installer."
    exit 1
fi
echo -e "  ${GREEN}✓${NC} Template found"

# ---------------------------------------------------------------------------
# Prompt: deals root directory
# ---------------------------------------------------------------------------
DEFAULT_DEALS="$HOME/Desktop/claude-workspace/deals"
echo ""
echo -e "  Enter your ${CYAN}deals root directory${NC} (where account folders live)."
echo -e "  Default: ${CYAN}${DEFAULT_DEALS}${NC}"
read -r -p "  [Press Enter to accept default]: " DEALS_ROOT
DEALS_ROOT="${DEALS_ROOT:-$DEFAULT_DEALS}"
mkdir -p "$DEALS_ROOT"
echo -e "  ${GREEN}✓${NC} Deals root: $DEALS_ROOT"

# ---------------------------------------------------------------------------
# Install skill and script
# ---------------------------------------------------------------------------
SKILLS_DIR="$HOME/.claude/skills"
mkdir -p "$SKILLS_DIR/sales-plan"

SCRIPT_DEST="$SKILLS_DIR/sales-plan/sales_plan_ppt.py"

echo ""
echo -e "${BLUE}Installing...${NC}"

# Copy and patch SKILL.md
sed \
    -e "s|__DEALS_ROOT__|${DEALS_ROOT}|g" \
    -e "s|__PPT_SCRIPT__|${SCRIPT_DEST}|g" \
    -e "s|__TEMPLATE_PATH__|${TEMPLATE_PATH}|g" \
    -e "s|__PYTHON_CMD__|${PYTHON_CMD}|g" \
    "$SOURCE_DIR/skills/sales-plan/SKILL.md" > "$SKILLS_DIR/sales-plan/SKILL.md"
echo -e "  ${GREEN}✓${NC} sales-plan skill installed"

# Copy and patch Python script
sed \
    -e "s|__TEMPLATE_PATH__|${TEMPLATE_PATH}|g" \
    "$SOURCE_DIR/tools/sales_plan_ppt.py" > "$SCRIPT_DEST"
chmod +x "$SCRIPT_DEST"
echo -e "  ${GREEN}✓${NC} sales_plan_ppt.py installed"

# ---------------------------------------------------------------------------
# Dependency check for the ai-sales-team-claude sub-skills
# ---------------------------------------------------------------------------
echo ""
echo -e "${BLUE}Checking required sub-skills...${NC}"
MISSING_SKILLS=()
for skill in sales-prospect sales-qualify sales-contacts sales-competitors sales-prep sales-outreach; do
    if [ -f "$SKILLS_DIR/$skill/SKILL.md" ]; then
        echo -e "  ${GREEN}✓${NC} $skill"
    else
        echo -e "  ${RED}✗${NC}  $skill — NOT FOUND"
        MISSING_SKILLS+=("$skill")
    fi
done

if [ ${#MISSING_SKILLS[@]} -gt 0 ]; then
    echo ""
    echo -e "  ${YELLOW}⚠  Missing sub-skills. Install the AI Sales Team first:${NC}"
    echo -e "  ${CYAN}curl -fsSL https://raw.githubusercontent.com/zubair-trabzada/ai-sales-team-claude/main/install.sh | bash${NC}"
fi

# ---------------------------------------------------------------------------
# Cleanup
# ---------------------------------------------------------------------------
if [ -n "$TEMP_DIR" ] && [ -d "$TEMP_DIR" ]; then
    rm -rf "$TEMP_DIR"
fi

# ---------------------------------------------------------------------------
# Summary
# ---------------------------------------------------------------------------
echo ""
echo -e "${GREEN}╔══════════════════════════════════════════════════════════════╗${NC}"
echo -e "${GREEN}║  Installation Complete!                                      ║${NC}"
echo -e "${GREEN}╚══════════════════════════════════════════════════════════════╝${NC}"
echo ""
echo -e "  ${CYAN}Skill:${NC}     $SKILLS_DIR/sales-plan/SKILL.md"
echo -e "  ${CYAN}Script:${NC}    $SCRIPT_DEST"
echo -e "  ${CYAN}Template:${NC}  $TEMPLATE_PATH"
echo -e "  ${CYAN}Deals:${NC}     $DEALS_ROOT"
echo ""
echo -e "${BLUE}Quick start:${NC}"
echo ""
echo -e "  Open Claude Code and run:"
echo -e "  ${CYAN}/sales plan https://www.example.com${NC}"
echo ""
echo -e "  With deal context:"
echo -e "  ${CYAN}/sales plan https://www.example.com --arr=150000 --renewal=Q3 --stage=POC --close-date=2026-09-30${NC}"
echo ""
