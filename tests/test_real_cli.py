import sys, json, uuid, threading, time
sys.path.insert(0, "/home/claude/work/ext/python")
from claudecalc import mcp_server, claude_cli, tools, settings, prompts
class Fake(mcp_server.Session):
    def list_tools(self): return tools.list_tools()
    def call_tool(self, n, a):
        print("   >> MCP tool call:", n, json.dumps(a)[:100])
        if n == "get_workbook_info": return json.dumps({"sheets":[{"name":"S","used_range":"A1:C3"}]}), False
        return json.dumps({"ok": True}), False
tok, url = mcp_server.HUB.register(Fake())
cfg = settings.load(); cfg["claude_path"] = ""
cli = claude_cli.find_claude(cfg, force=True); print("CLI:", cli)
print("auth:", claude_cli.auth_status(cli["path"]).get("logged_in"))
ev = []; done = threading.Event()
def on(k, d):
    ev.append((k, d))
    print("EVENT", k, str(d)[:140])
    if k == "done": done.set()
t = claude_cli.Turn(cli, cfg, url, tok, str(uuid.uuid4()), False, "Some A1:A3 em D1", prompts.build_system(""), on)
t.start(); done.wait(120)
print("kinds:", [e[0] for e in ev])
