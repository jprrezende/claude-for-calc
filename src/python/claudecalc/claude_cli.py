# -*- coding: utf-8 -*-
"""Ponte com o Claude Code (CLI `claude`), que já traz a autenticação do usuário.

A extensão NUNCA lê nem copia credenciais: ela só executa `claude -p` na máquina do
usuário, que usa o login que o próprio Claude já tem (ou a ANTHROPIC_API_KEY do
ambiente). O CLI recebe as ferramentas do Calc por MCP (servidor local em 127.0.0.1).
"""
import glob
import json
import os
import re
import subprocess
import threading
import uuid

from . import settings as S

CREATE_NO_WINDOW = 0x08000000
CREATE_NEW_CONSOLE = 0x00000010

_WIN = os.name == "nt"
_cache = {"key": None, "found": None, "help": {}}


# --------------------------------------------------------------------------
# Localização do executável
# --------------------------------------------------------------------------

def _looks_like_desktop_app(path):
    """Evita executar o aplicativo desktop (que abriria uma janela)."""
    p = path.lower().replace("/", "\\")
    return "\\windowsapps\\" in p or "\\anthropicclaude\\" in p


def _version_key(p):
    m = re.search(r"claude-code[\\/]([0-9][0-9.]*)", p.replace("\\", "/"))
    if not m:
        return ()
    return tuple(int(x) for x in m.group(1).split(".") if x.isdigit())


def _candidates(configured):
    seen = set()

    def add(p, origin):
        if p and p not in seen and os.path.isfile(p):
            seen.add(p)
            return [(p, origin)]
        return []

    out = []
    if configured:
        out += add(configured, "configurado")
    names = ["claude.exe", "claude.cmd", "claude"] if _WIN else ["claude"]
    for d in os.environ.get("PATH", "").split(os.pathsep):
        if not d:
            continue
        for n in names:
            out += add(os.path.join(d.strip('"'), n), "PATH")
    home = os.path.expanduser("~")
    appdata = os.environ.get("APPDATA", "")
    local = os.environ.get("LOCALAPPDATA", "")
    out += add(os.path.join(home, ".local", "bin", "claude.exe" if _WIN else "claude"), "instalação nativa")
    if appdata:
        out += add(os.path.join(appdata, "npm", "claude.cmd"), "npm")
    # Claude Code que acompanha o app Claude Desktop
    pats = []
    if appdata:
        pats.append(os.path.join(appdata, "Claude", "claude-code", "*", "claude.exe"))
    if local:
        pats.append(os.path.join(local, "Packages", "Claude_*", "LocalCache", "Roaming", "Claude", "claude-code", "*", "claude.exe"))
    # Layouts mais novos do app: ...\Claude-Data\<id>\claude-code\<versão>\<hash>\claude.exe
    for base in (local, appdata):
        if base:
            for sub in ("Claude-Data", "Claude"):
                pats.append(os.path.join(base, sub, "*", "claude-code", "*", "*", "claude.exe"))
                pats.append(os.path.join(base, sub, "claude-code", "*", "*", "claude.exe"))
    if appdata:
        pats.append(os.path.join(appdata, "Claude", "claude-code", "*", "*", "claude.exe"))
    if local:
        pats.append(os.path.join(local, "Packages", "Claude_*", "LocalCache", "*", "Claude*", "claude-code", "*", "*", "claude.exe"))
        pats.append(os.path.join(local, "Packages", "Claude_*", "LocalCache", "*", "Claude*", "*", "claude-code", "*", "*", "claude.exe"))
    if not _WIN:
        pats.append(os.path.join(home, ".config", "Claude", "claude-code", "*", "claude"))
        pats.append(os.path.join(home, "Library", "Application Support", "Claude", "claude-code", "*", "claude"))
    found = []
    for pat in pats:
        found += glob.glob(pat)
    for p in sorted(found, key=_version_key, reverse=True):
        out += add(p, "app Claude Desktop")
    return out


def _popen_flags():
    return CREATE_NO_WINDOW if _WIN else 0


def run_quiet(args, timeout=30, input_text=None):
    try:
        p = subprocess.run(args, input=(input_text.encode("utf-8") if input_text else None),
                           stdout=subprocess.PIPE, stderr=subprocess.PIPE, timeout=timeout,
                           creationflags=_popen_flags(), cwd=S.workspace_dir(), stdin=subprocess.DEVNULL if input_text is None else None)
        return p.returncode, p.stdout.decode("utf-8", "replace"), p.stderr.decode("utf-8", "replace")
    except Exception as e:  # noqa
        return -1, "", str(e)


def probe(path):
    """Versão do Claude Code em `path`, ou None se não for o CLI."""
    if _looks_like_desktop_app(path):
        return None
    code, out, err = run_quiet([path, "--version"], timeout=25)
    txt = (out + " " + err).strip()
    if code == 0 and "claude code" in txt.lower():
        return txt.splitlines()[0].strip()
    return None


def find_claude(cfg, force=False):
    """Devolve dict(path, version, origin) ou None."""
    key = cfg.get("claude_path", "")
    if not force and _cache["key"] == key and _cache["found"]:
        return _cache["found"]
    for p, origin in _candidates(key):
        v = probe(p)
        if v:
            res = {"path": p, "version": v, "origin": origin}
            _cache["key"], _cache["found"] = key, res
            return res
    _cache["key"], _cache["found"] = key, None
    return None


def supported_flags(path):
    h = _cache["help"].get(path)
    if h is None:
        code, out, err = run_quiet([path, "--help"], timeout=30)
        h = out + err
        _cache["help"][path] = h
    return h


def auth_status(path):
    code, out, err = run_quiet([path, "auth", "status"], timeout=40)
    d = {}
    try:
        d = json.loads(out)
    except Exception:
        pass
    logged = bool(d.get("loggedIn")) if d else ("logged in" in out.lower() and "not" not in out.lower())
    email = d.get("email") or (d.get("account") or {}).get("email") or d.get("emailAddress")
    return {"logged_in": logged, "method": d.get("authMethod"), "email": email, "raw": d, "text": (out + err).strip()}


def start_login(path):
    """Abre o fluxo de login do próprio Claude (navegador). Não toca em credenciais."""
    try:
        if _WIN:
            subprocess.Popen([path, "auth", "login"], creationflags=CREATE_NEW_CONSOLE, cwd=S.workspace_dir())
        else:
            subprocess.Popen([path, "auth", "login"], cwd=S.workspace_dir(),
                             stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, stdin=subprocess.DEVNULL)
        return True
    except Exception:
        return False


# --------------------------------------------------------------------------
# Execução de uma pergunta
# --------------------------------------------------------------------------

def _kill_tree(proc):
    try:
        if _WIN:
            subprocess.run(["taskkill", "/PID", str(proc.pid), "/T", "/F"], creationflags=CREATE_NO_WINDOW,
                           stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, timeout=10)
        else:
            proc.kill()
    except Exception:
        try:
            proc.kill()
        except Exception:
            pass


class Turn:
    """Executa `claude -p` em uma thread e entrega eventos via callback.

    on_event(kind, data):
      init      {session_id, mcp: {...}}
      text      str                  (trecho de resposta, em streaming)
      tool_use  {name, input}
      tool_err  str
      result    {is_error, text, cost, duration_ms, session_id}
      error     str
      done      None                  (sempre o último)
    """

    def __init__(self, cli, cfg, mcp_url, mcp_token, session_id, resuming, prompt, system_prompt, on_event):
        self.cli, self.cfg = cli, cfg
        self.mcp_url, self.mcp_token = mcp_url, mcp_token
        self.session_id, self.resuming = session_id, resuming
        self.prompt, self.system_prompt = prompt, system_prompt
        self.on_event = on_event
        self.proc = None
        self.inline_prompt = False
        self._stopped = False
        self._thread = None
        self._tmp = []

    def start(self):
        self._thread = threading.Thread(target=self._run, name="claudecalc-turn", daemon=True)
        self._thread.start()

    def stop(self):
        self._stopped = True
        if self.proc is not None and self.proc.poll() is None:
            _kill_tree(self.proc)

    # ------------------------------------------------------------------
    def _write_tmp(self, suffix, text):
        p = os.path.join(S.workspace_dir(), "turn-%s%s" % (uuid.uuid4().hex[:10], suffix))
        with open(p, "w", encoding="utf-8") as f:
            f.write(text)
        self._tmp.append(p)
        return p

    def _build_args(self):
        path = self.cli["path"]
        h = supported_flags(path)

        def has(flag):
            return flag in h

        mcp_cfg = {"mcpServers": {"calc": {"type": "http", "url": self.mcp_url,
                                           "headers": {"Authorization": "Bearer " + self.mcp_token}}}}
        cfg_file = self._write_tmp(".mcp.json", json.dumps(mcp_cfg))
        sp_file = self._write_tmp(".system.txt", self.system_prompt)
        args = [path, "-p", "--output-format", "stream-json", "--verbose", "--include-partial-messages",
                "--mcp-config", cfg_file, "--strict-mcp-config"]
        if not self.inline_prompt:
            args += ["--append-system-prompt-file", sp_file]
        else:
            sp = self.system_prompt
            if path.lower().endswith((".cmd", ".bat")):  # evita metacaracteres do cmd.exe
                sp = re.sub(r'[&|<>^%"\r\n]', " ", sp)
            args += ["--append-system-prompt", sp]
        if has("--tools"):
            args += ["--tools", ""]
        if has("--permission-mode"):
            args += ["--permission-mode", "dontAsk"]
        if has("--disable-slash-commands"):
            args.append("--disable-slash-commands")
        if has("--no-chrome"):
            args.append("--no-chrome")
        model = (self.cfg.get("model") or "").strip()
        if model:
            args += ["--model", model]
        if self.resuming:
            args += ["--resume", self.session_id]
        else:
            args += ["--session-id", self.session_id]
        args += ["--allowedTools", "mcp__calc"]
        return args

    def _run(self):
        try:
            if not self._run_once() and not self._stopped:
                # CLI antigo sem --append-system-prompt-file: tenta de novo com o texto inline
                self.inline_prompt = True
                self._run_once(final=True)
        finally:
            for p in self._tmp:
                try:
                    os.remove(p)
                except Exception:
                    pass
            self.on_event("done", None)

    def _run_once(self, final=False):
        """Devolve False se precisa repetir com prompt inline."""
        ev = self.on_event
        streamed = False
        got_result = False
        stderr_buf = []
        try:
            args = self._build_args()
            env = dict(os.environ)
            env.pop("CLAUDECODE", None)
            self.proc = subprocess.Popen(args, stdin=subprocess.PIPE, stdout=subprocess.PIPE, stderr=subprocess.PIPE,
                                         cwd=S.workspace_dir(), env=env, creationflags=_popen_flags())

            def drain_err():
                try:
                    for ln in self.proc.stderr:
                        stderr_buf.append(ln.decode("utf-8", "replace"))
                        del stderr_buf[:-80]
                except Exception:
                    pass

            threading.Thread(target=drain_err, daemon=True).start()
            try:
                self.proc.stdin.write(self.prompt.encode("utf-8"))
                self.proc.stdin.close()
            except Exception:
                pass

            timeout_s = float(self.cfg.get("timeout_minutes", 15)) * 60
            timer = threading.Timer(timeout_s, self.stop)
            timer.daemon = True
            timer.start()
            try:
                for raw in self.proc.stdout:
                    line = raw.decode("utf-8", "replace").strip()
                    if not line or not line.startswith("{"):
                        continue
                    try:
                        m = json.loads(line)
                    except Exception:
                        continue
                    t = m.get("type")
                    if t == "system" and m.get("subtype") == "init":
                        servers = {s.get("name"): s.get("status") for s in (m.get("mcp_servers") or [])}
                        ev("init", {"session_id": m.get("session_id"), "mcp": servers})
                    elif t == "stream_event":
                        e = m.get("event") or {}
                        et = e.get("type")
                        if et == "message_start":
                            streamed = False
                        elif et == "content_block_delta":
                            d = e.get("delta") or {}
                            if d.get("type") == "text_delta" and d.get("text"):
                                streamed = True
                                ev("text", d["text"])
                        elif et == "content_block_stop" and streamed:
                            pass
                    elif t == "assistant":
                        for b in (m.get("message") or {}).get("content") or []:
                            if b.get("type") == "tool_use":
                                ev("tool_use", {"name": b.get("name", ""), "input": b.get("input") or {}})
                            elif b.get("type") == "text" and not streamed and b.get("text"):
                                ev("text", b["text"])
                        streamed = False
                    elif t == "user":
                        c = (m.get("message") or {}).get("content")
                        if isinstance(c, list):
                            for b in c:
                                if b.get("type") == "tool_result" and b.get("is_error"):
                                    cont = b.get("content")
                                    if isinstance(cont, list):
                                        cont = " ".join(x.get("text", "") for x in cont if isinstance(x, dict))
                                    ev("tool_err", str(cont)[:400])
                    elif t == "result":
                        got_result = True
                        ev("result", {"is_error": bool(m.get("is_error")), "text": m.get("result") or "",
                                      "cost": m.get("total_cost_usd"), "duration_ms": m.get("duration_ms"),
                                      "session_id": m.get("session_id"), "subtype": m.get("subtype")})
            finally:
                timer.cancel()
            rc = self.proc.wait()
            if not got_result and not self._stopped:
                tail = "".join(stderr_buf).strip()[-900:]
                if (not final and "system-prompt-file" in tail and "unknown" in tail.lower()):
                    return False
                ev("error", tail or ("O Claude terminou sem resposta (código %s)." % rc))
        except FileNotFoundError:
            ev("error", "Não encontrei o executável do Claude Code.")
        except Exception as e:  # noqa
            ev("error", "%s: %s" % (e.__class__.__name__, e))
        return True
