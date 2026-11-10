# -*- coding: utf-8 -*-
"""Servidor MCP mínimo (JSON-RPC sobre HTTP, 127.0.0.1) para o Claude Code usar as ferramentas do Calc."""
import json
import secrets
import threading
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer

SUPPORTED = ["2025-11-25", "2025-06-18", "2025-03-26", "2024-11-05"]
SERVER_INFO = {"name": "claude-calc", "version": "1.0.0"}


class Session:
    """Uma sessão = um painel/documento. Implementada pelo painel."""
    def list_tools(self):
        raise NotImplementedError

    def call_tool(self, name, args):
        """Devolve (texto, is_error)."""
        raise NotImplementedError


class McpHub:
    def __init__(self):
        self._sessions = {}
        self._lock = threading.Lock()
        self._httpd = None
        self.port = None

    # ------------------------------------------------------------------
    def ensure_started(self):
        with self._lock:
            if self._httpd is not None:
                return
            hub = self

            class Handler(BaseHTTPRequestHandler):
                protocol_version = "HTTP/1.1"

                def log_message(self, *a):  # silencioso
                    pass

                # -- helpers
                def _send(self, code, body=None, ctype="application/json"):
                    data = b"" if body is None else (body if isinstance(body, bytes) else json.dumps(body, ensure_ascii=False).encode("utf-8"))
                    self.send_response(code)
                    if data:
                        self.send_header("Content-Type", ctype)
                    self.send_header("Content-Length", str(len(data)))
                    self.end_headers()
                    if data:
                        self.wfile.write(data)

                def _session(self):
                    host = (self.headers.get("Host") or "").split(":")[0]
                    if host not in ("127.0.0.1", "localhost"):
                        return None
                    auth = self.headers.get("Authorization", "")
                    if not auth.startswith("Bearer "):
                        return None
                    return hub._sessions.get(auth[7:].strip())

                # -- verbos
                def do_GET(self):
                    self._send(405)

                def do_DELETE(self):
                    self._send(200)

                def do_POST(self):
                    sess = self._session()
                    if sess is None:
                        self._send(401, {"error": "unauthorized"})
                        return
                    if self.path.split("?")[0] != "/mcp":
                        self._send(404)
                        return
                    try:
                        n = int(self.headers.get("Content-Length") or 0)
                        raw = self.rfile.read(n) if n else b""
                        msg = json.loads(raw.decode("utf-8"))
                    except Exception:
                        self._send(400, {"jsonrpc": "2.0", "id": None,
                                         "error": {"code": -32700, "message": "parse error"}})
                        return
                    batch = isinstance(msg, list)
                    replies = []
                    for m in (msg if batch else [msg]):
                        r = hub._handle(sess, m)
                        if r is not None:
                            replies.append(r)
                    if not replies:
                        self._send(202)
                    else:
                        self._send(200, replies if batch else replies[0])

            self._httpd = ThreadingHTTPServer(("127.0.0.1", 0), Handler)
            self._httpd.daemon_threads = True
            self.port = self._httpd.server_address[1]
            t = threading.Thread(target=self._httpd.serve_forever, name="claudecalc-mcp", daemon=True)
            t.start()

    def stop(self):
        with self._lock:
            if self._httpd is not None:
                try:
                    self._httpd.shutdown()
                    self._httpd.server_close()
                except Exception:
                    pass
                self._httpd = None
                self.port = None

    # ------------------------------------------------------------------
    def register(self, session):
        self.ensure_started()
        token = secrets.token_urlsafe(32)
        with self._lock:
            self._sessions[token] = session
        return token, "http://127.0.0.1:%d/mcp" % self.port

    def unregister(self, token):
        with self._lock:
            self._sessions.pop(token, None)

    # ------------------------------------------------------------------
    def _handle(self, sess, m):
        if not isinstance(m, dict):
            return {"jsonrpc": "2.0", "id": None, "error": {"code": -32600, "message": "invalid request"}}
        method = m.get("method")
        mid = m.get("id")
        params = m.get("params") or {}
        if mid is None:  # notificação
            return None

        def ok(result):
            return {"jsonrpc": "2.0", "id": mid, "result": result}

        def err(code, msg):
            return {"jsonrpc": "2.0", "id": mid, "error": {"code": code, "message": msg}}

        try:
            if method == "initialize":
                req = params.get("protocolVersion")
                ver = req if req in SUPPORTED else SUPPORTED[0]
                return ok({"protocolVersion": ver, "capabilities": {"tools": {"listChanged": False}},
                           "serverInfo": SERVER_INFO,
                           "instructions": "Ferramentas para ler e editar a planilha aberta no LibreOffice Calc."})
            if method == "ping":
                return ok({})
            if method == "tools/list":
                return ok({"tools": sess.list_tools()})
            if method == "tools/call":
                name = params.get("name")
                args = params.get("arguments") or {}
                text, is_err = sess.call_tool(name, args)
                return ok({"content": [{"type": "text", "text": text}], "isError": bool(is_err)})
            if method in ("resources/list", "prompts/list", "resources/templates/list"):
                key = method.split("/")[0]
                return ok({key: []})
            return err(-32601, "method not found: %s" % method)
        except Exception as e:  # noqa
            return ok({"content": [{"type": "text", "text": json.dumps({"error": "%s: %s" % (e.__class__.__name__, e)}, ensure_ascii=False)}],
                       "isError": True})


HUB = McpHub()
