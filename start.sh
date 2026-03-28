#!/bin/bash
# ============================================================
# Tsifulator.ai — One-command startup
# Double-click this file or run: bash start.sh
# ============================================================

ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
BACKEND="$ROOT/backend"
ADDIN="$ROOT/excel-addin"
CERTS="$HOME/.office-addin-dev-certs"

# Colors
GREEN='\033[0;32m'
PURPLE='\033[0;35m'
RED='\033[0;31m'
NC='\033[0m'

echo ""
echo -e "${PURPLE}⚡ Tsifulator.ai — Starting up...${NC}"
echo "────────────────────────────────────"

# Kill any previous add-in server
pkill -f "webpack serve" 2>/dev/null
sleep 1

# 1. Check cloud backend is alive
echo -e "  ${GREEN}▶ Checking cloud backend...${NC}"
RAILWAY_URL="https://focused-solace-production-6839.up.railway.app"
if curl -s "$RAILWAY_URL/" > /dev/null 2>&1; then
  echo -e "  ${GREEN}✅ Backend live       → $RAILWAY_URL${NC}"
else
  echo -e "  ${RED}❌ Backend unreachable — check railway.app${NC}"
fi

# 2. Start Excel add-in server (still local — serves the sidebar UI)
echo -e "  ${GREEN}▶ Starting Excel add-in server...${NC}"
cd "$ADDIN" && \
  npm start -- --no-open \
  > "$ROOT/logs/addin.log" 2>&1 &

sleep 6
if curl -sk https://localhost:3000/ > /dev/null 2>&1; then
  echo -e "  ${GREEN}✅ Excel add-in ready → https://localhost:3000${NC}"
else
  echo -e "  ${GREEN}⏳ Excel add-in starting (may take 10s)...${NC}"
fi

echo ""
echo "────────────────────────────────────"
echo -e "${PURPLE}⚡ Tsifulator.ai is running!${NC}"
echo ""
echo "  Excel:    Open Excel → Home → Add-ins → Tsifulator.ai"
echo "  RStudio:  Addins → Tsifulator.ai"
echo "  Terminal: python3 terminal-client/tsifulator.py"
echo "  Gmail:    python3 gmail-client/gmail_setup.py  (first time only)"
echo ""
echo "  Logs:    $ROOT/logs/"
echo "  Stop:    Press Ctrl+C"
echo "────────────────────────────────────"
echo ""

# Keep running and show live backend logs
tail -f "$ROOT/logs/backend.log"
