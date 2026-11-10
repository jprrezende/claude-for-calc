# Claude for Calc (LibreOffice extension)

*Leia em português: [README.pt-BR.md](README.pt-BR.md)*

A chat sidebar inside LibreOffice Calc, in the spirit of "Claude for Excel". Ask questions about the open
spreadsheet and let Claude read cells, explain formulas, write or fix formulas, format, sort, and build charts
and pivot tables. Every change is a regular LibreOffice action and can be undone with **Ctrl+Z**.

> Community project, not an official Anthropic product.

## Installation (Windows 11 + LibreOffice 26.2)
1. In Calc: **Tools ▸ Extension Manager ▸ Add…** and pick `ClaudeCalc.oxt` (build it with `./build.sh`, see below).
2. Restart LibreOffice.
3. Open a spreadsheet and use the **Claude ▸ Open Claude panel** menu (or the orange icon in the right sidebar).

## How authentication works
The extension **never reads or stores credentials**. It runs the **Claude Code** CLI (`claude -p`) that is already
installed on your computer and uses the sign-in you already have. The executable is detected in this order:
path set in Settings → `claude` on the PATH → `%USERPROFILE%\.local\bin\claude.exe` → `%APPDATA%\npm\claude.cmd`
→ the Claude Code bundled with the Claude Desktop app (`%APPDATA%\Claude\claude-code\<version>\claude.exe`,
including the MSIX variant under `%LOCALAPPDATA%\Packages\Claude_*\LocalCache\Roaming\...`).
The Claude Desktop app's own executable (`WindowsApps\Claude.exe`) is deliberately ignored.

- If the panel says **"Not signed in"**, click **Sign in**: it opens Claude's own sign-in flow in your browser.
- Alternative without a subscription sign-in: set the `ANTHROPIC_API_KEY` environment variable (a Console key).
- Note: how `claude -p` usage is billed on subscription plans has changed over time. Check Anthropic's help article
  "Use the Claude Agent SDK with your Claude plan" for the current rules.

## What Claude can do (tools)
Read: `get_workbook_info`, `read_range`, `find_cells`, `find_errors`, `get_format`.
Change: `write_range` (values and Excel-style formulas), `clear_range`, `format_range`, `copy_range`,
`insert_delete`, `resize`, `sort_range`, `manage_sheet`, `freeze_panes`, `add_comment`, `define_name`,
`create_chart`, `create_pivot_table`. Show: `select_range` (selects the cell on screen).
Claude only sees these tools: **no terminal, no files, no internet**.

Examples: "Explain how we got the total in F20" · "Find the cells with errors and fix them" ·
"Build a summary by region on another sheet, using formulas" · "Format the header and freeze the first row" ·
"Create a column chart of sales by month".

## Security and privacy
- The tools are served by a local server on `127.0.0.1` (random port), protected by a per-panel token and a Host
  header check. Nothing is exposed to the network.
- The content of cells that Claude **reads** is sent to Anthropic, as with Claude for Excel.
- **Settings ▸ "Ask for confirmation before every change"** shows a Yes/No box before each write.

## Known limitations
- Plain-text chat (no clickable cell links; Claude uses `select_range` to point at cells).
- No conditional formatting, data validation, or macros yet.
- Newer Excel functions only work if LibreOffice supports them (on `#NAME?`, Claude switches to an equivalent).
- Reads are limited to 5000 cells per call (adjustable in Settings); Claude reads in chunks.
- Tested on Linux with LibreOffice 24.2 and Claude Code 2.1.322 against a mock API; not yet validated on
  Windows 11 with LibreOffice 26.2 or with a live model.

## Troubleshooting
- Log: `%APPDATA%\ClaudeCalc\claudecalc.log`. Settings: `%APPDATA%\ClaudeCalc\settings.json`.
- "Claude Code not found": Settings ▸ enter the path to `claude.exe` ▸ Detect.
- Sign-in errors: click Sign in, or run `claude auth login` in a terminal.
- On Windows, Claude Code may fail to start without Git for Windows; install it if you see a git-bash error.

## Development
- `./build.sh` builds `dist/ClaudeCalc.oxt` from `src/`.
- `tests/test_tools.py` exercises the spreadsheet tools against a real LibreOffice
  (`soffice --headless --accept="socket,host=127.0.0.1,port=2002;urp;"`).
- Layout: `src/python/claude_calc.py` (UNO component) and `src/python/claudecalc/` (`tools.py`, `panel.py`,
  `mcp_server.py`, `claude_cli.py`, `executor.py`, …).
