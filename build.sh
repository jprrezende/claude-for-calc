#!/bin/bash
# Gera dist/ClaudeCalc.oxt
set -e
cd "$(dirname "$0")/src"
find . -name "__pycache__" -type d -prune -exec rm -rf {} + 2>/dev/null || true
mkdir -p ../dist && rm -f ../dist/ClaudeCalc.oxt
zip -qr -X ../dist/ClaudeCalc.oxt META-INF description description.xml images python Addons.xcu Factories.xcu ProtocolHandler.xcu Sidebar.xcu
echo "built: $(cd .. && pwd)/dist/ClaudeCalc.oxt ($(du -h ../dist/ClaudeCalc.oxt | cut -f1))"
