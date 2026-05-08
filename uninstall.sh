#!/bin/bash
# ============================================================================
# /sales plan — Uninstaller
# ============================================================================

GREEN='\033[0;32m'
YELLOW='\033[1;33m'
NC='\033[0m'

SKILLS_DIR="$HOME/.claude/skills"

echo ""
echo "Uninstalling /sales plan skill..."
echo ""

if [ -d "$SKILLS_DIR/sales-plan" ]; then
    rm -rf "$SKILLS_DIR/sales-plan"
    echo -e "  ${GREEN}✓${NC} Removed $SKILLS_DIR/sales-plan"
else
    echo -e "  ${YELLOW}⚠${NC}  Not found: $SKILLS_DIR/sales-plan — nothing to remove"
fi

echo ""
echo "Done. Your deal folders and generated PPTs are untouched."
echo ""
