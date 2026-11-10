# -*- coding: utf-8 -*-
"""Painel lateral (sidebar) do Calc: chat com o Claude."""
import json
import re
import threading
import time
import uuid

import uno
import unohelper
from com.sun.star.awt import XActionListener, XKeyListener, XWindowListener
from com.sun.star.lang import XComponent
from com.sun.star.ui import LayoutSize, XSidebarPanel, XToolPanel, XUIElement

from . import claude_cli, dialogs, i18n, mcp_server, prompts, tools
from . import settings as S
from .executor import MainThread
from .util import log, log_exc

POSSIZE = 15
PANELS = []  # painéis vivos (para atualizar status depois de mudar configurações)


# --------------------------------------------------------------------------
# Sessão MCP deste painel: recebe chamadas de ferramentas vindas do Claude
# --------------------------------------------------------------------------

class _PanelSession(mcp_server.Session):
    def __init__(self, panel):
        self.panel = panel

    def list_tools(self):
        return tools.list_tools()

    def call_tool(self, name, args):
        return self.panel.exec_tool(name, args)


class _Btn(unohelper.Base, XActionListener):
    def __init__(self, fn):
        self.fn = fn

    def actionPerformed(self, ev):
        try:
            self.fn()
        except Exception:
            log_exc("botão")

    def disposing(self, src):
        pass


class _Resize(unohelper.Base, XWindowListener):
    def __init__(self, fn):
        self.fn = fn

    def windowResized(self, ev):
        try:
            self.fn(ev.Width, ev.Height)
        except Exception:
            log_exc("resize")

    def windowMoved(self, ev):
        pass

    def windowShown(self, ev):
        pass

    def windowHidden(self, ev):
        pass

    def disposing(self, src):
        pass


class _Keys(unohelper.Base, XKeyListener):
    """Ctrl+Enter envia."""

    def __init__(self, fn):
        self.fn = fn

    def keyPressed(self, ev):
        try:
            if ev.KeyCode == 1280 and (ev.Modifiers & 2):  # RETURN + CTRL(MOD1)
                self.fn()
        except Exception:
            log_exc("tecla")

    def keyReleased(self, ev):
        pass

    def disposing(self, src):
        pass


class _LineStream:
    """Acumula o texto em streaming e entrega pedaços 'limpos' (sem marcações de Markdown)."""

    def __init__(self, sink):
        self.sink = sink
        self.buf = ""

    def feed(self, s):
        self.buf += s
        if "\n" in self.buf:
            head, _, tail = self.buf.rpartition("\n")
            self.sink(prompts.clean_markdown(head) + "\n")
            self.buf = tail
        if len(self.buf) > 90:
            cut = self.buf.rfind(" ", 0, 90)
            seg = self.buf[:cut + 1] if cut > 20 else None
            if seg is not None and seg.count("**") % 2 == 0 and seg.count("`") % 2 == 0:
                self.sink(seg)
                self.buf = self.buf[cut + 1:]

    def flush(self):
        if self.buf:
            self.sink(prompts.clean_markdown(self.buf))
            self.buf = ""


# --------------------------------------------------------------------------
# O painel
# --------------------------------------------------------------------------

class ClaudePanel(unohelper.Base, XUIElement, XToolPanel, XSidebarPanel, XComponent):
    def __init__(self, uno_ctx, resource_url, frame, parent):
        self.ctx = uno_ctx
        self.smgr = uno_ctx.ServiceManager
        self.resource_url = resource_url
        self.frame = frame
        self.parent = parent
        self.cfg = S.load()
        self.T = i18n.table(i18n.ui_lang(uno_ctx))
        self.main = MainThread(uno_ctx)
        self.session = _PanelSession(self)
        self.token, self.mcp_url = mcp_server.HUB.register(self.session)
        self.cli = None
        self.auth = None
        self.turn = None
        self.turn_doc = None
        self.busy = False
        self.session_id = str(uuid.uuid4())
        self.resuming = False
        self._pending = []
        self._plock = threading.Lock()
        self._flush_sched = False
        self._disposed = False
        self._listeners = []
        self._stream = _LineStream(self.emit)
        self._t0 = 0.0
        self._at_bol = True      # último texto emitido terminou em quebra de linha?
        self._after_tool = False  # a última coisa emitida foi uma linha de ferramenta?
        self._build_ui()
        PANELS.append(self)
        self.emit(self.T["welcome"] + "\n")
        self.refresh_auth()

    # ------------------------------------------------------------------ UNO
    def getFrame(self):
        return self.frame

    def getResourceURL(self):
        return self.resource_url

    def getType(self):
        return uno.getConstantByName("com.sun.star.ui.UIElementType.TOOLPANEL")

    def getRealInterface(self):
        return self

    def getWindow(self):
        return self.container

    def createAccessible(self, parent):
        return None

    def getHeightForWidth(self, width):
        return LayoutSize(260, -1, 520)

    def getMinimalWidth(self):
        return 220

    def addEventListener(self, l):
        self._listeners.append(l)

    def removeEventListener(self, l):
        if l in self._listeners:
            self._listeners.remove(l)

    def dispose(self):
        if self._disposed:
            return
        self._disposed = True
        try:
            if self.turn is not None:
                self.turn.stop()
        except Exception:
            pass
        mcp_server.HUB.unregister(self.token)
        if self in PANELS:
            PANELS.remove(self)
        for l in list(self._listeners):
            try:
                l.disposing(None)
            except Exception:
                pass
        try:
            self.container.dispose()
        except Exception:
            pass

    # ------------------------------------------------------------------ UI
    def _build_ui(self):
        T = self.T
        toolkit = self.smgr.createInstanceWithContext("com.sun.star.awt.Toolkit", self.ctx)
        self.cmodel = self.smgr.createInstanceWithContext("com.sun.star.awt.UnoControlContainerModel", self.ctx)
        self.container = self.smgr.createInstanceWithContext("com.sun.star.awt.UnoControlContainer", self.ctx)
        self.container.setModel(self.cmodel)
        self.container.createPeer(toolkit, self.parent)
        self._toolkit = toolkit

        def add(name, kind, **props):
            m = self.smgr.createInstanceWithContext("com.sun.star.awt.UnoControl%sModel" % kind, self.ctx)
            for k, v in props.items():
                m.setPropertyValue(k, v)
            ctrl = self.smgr.createInstanceWithContext("com.sun.star.awt.UnoControl%s" % kind, self.ctx)
            ctrl.setModel(m)
            self.container.addControl(name, ctrl)
            return m

        self.m_status = add("status", "FixedText", Label=T["checking"], MultiLine=True)
        self.m_new = add("btnNew", "Button", Label=T["new_chat"])
        self.m_set = add("btnSet", "Button", Label=T["settings"])
        self.m_login = add("btnLogin", "Button", Label=T["login"])
        self.m_log = add("log", "Edit", MultiLine=True, ReadOnly=True, VScroll=True, HScroll=False, Border=1)
        self.m_in = add("input", "Edit", MultiLine=True, VScroll=True, HScroll=False, Border=1)
        self.m_send = add("btnSend", "Button", Label=T["send"], DefaultButton=False)
        self.m_stop = add("btnStop", "Button", Label=T["stop"], Enabled=False)
        self.m_foot = add("foot", "FixedText", Label=T["ready"])

        c = self.container.getControl
        c("btnNew").addActionListener(_Btn(self.new_chat))
        c("btnSet").addActionListener(_Btn(self.open_settings))
        c("btnLogin").addActionListener(_Btn(self.login))
        c("btnSend").addActionListener(_Btn(self.send))
        c("btnStop").addActionListener(_Btn(self.stop))
        c("input").addKeyListener(_Keys(self.send))
        self.container.addWindowListener(_Resize(self.layout))
        self.container.setVisible(True)
        try:
            r = self.parent.getPosSize()
            self.container.setPosSize(0, 0, r.Width or 300, r.Height or 500, POSSIZE)
            self.layout(r.Width or 300, r.Height or 500)
        except Exception:
            self.layout(300, 500)

    def layout(self, w, h):
        if self._disposed or w <= 0 or h <= 0:
            return
        c = self.container.getControl
        pad = 4
        iw = max(w - 2 * pad, 50)
        y = pad
        c("status").setPosSize(pad, y, iw, 30, POSSIZE)
        y += 32
        bw = (iw - 2 * pad) // 3
        for i, n in enumerate(("btnNew", "btnSet", "btnLogin")):
            c(n).setPosSize(pad + i * (bw + pad), y, bw, 24, POSSIZE)
        y += 28
        foot_h, send_h, in_h = 18, 26, 70
        log_h = max(h - y - (in_h + send_h + foot_h + 4 * pad), 40)
        c("log").setPosSize(pad, y, iw, log_h, POSSIZE)
        y += log_h + pad
        c("input").setPosSize(pad, y, iw, in_h, POSSIZE)
        y += in_h + pad
        half = (iw - pad) // 2
        c("btnSend").setPosSize(pad, y, half, send_h, POSSIZE)
        c("btnStop").setPosSize(pad + half + pad, y, half, send_h, POSSIZE)
        y += send_h + pad
        c("foot").setPosSize(pad, y, iw, foot_h, POSSIZE)
        self._scroll_end()

    # ------------------------------------------------------------------ texto
    def _append(self, text):
        if self._disposed:
            return
        m = self.m_log
        m.Text = m.Text + text
        self._scroll_end()

    def _scroll_end(self):
        """Mostra o fim do chat. Vai ao topo primeiro, para só rolar o mínimo necessário."""
        try:
            ctrl = self.container.getControl("log")
            n = len(self.m_log.Text)
            sel = uno.createUnoStruct("com.sun.star.awt.Selection")
            sel.Min = sel.Max = 0
            ctrl.setSelection(sel)
            sel2 = uno.createUnoStruct("com.sun.star.awt.Selection")
            sel2.Min = sel2.Max = n
            ctrl.setSelection(sel2)
        except Exception:
            pass

    def emit(self, text):
        """Pode ser chamado de qualquer thread."""
        if not text:
            return
        with self._plock:
            self._pending.append(text)
            self._at_bol = text.endswith("\n")
            if self._flush_sched:
                return
            self._flush_sched = True
        self.main.post(self._flush)

    def _flush(self):
        with self._plock:
            txt = "".join(self._pending)
            self._pending = []
            self._flush_sched = False
        if txt:
            self._append(txt)

    def set_status(self, text):
        def f():
            if not self._disposed:
                self.m_status.Label = text
        self.main.post(f)

    def set_foot(self, text):
        def f():
            if not self._disposed:
                self.m_foot.Label = text
        self.main.post(f)

    def set_busy(self, busy):
        def f():
            if self._disposed:
                return
            self.busy = busy
            self.m_send.Enabled = not busy
            self.m_stop.Enabled = busy
        self.main.post(f)

    # ------------------------------------------------------------------ login / status
    def refresh_auth(self):
        self.set_status(self.T["checking"])

        def work():
            try:
                self.cfg = S.load()
                self.cli = claude_cli.find_claude(self.cfg, force=True)
                if not self.cli:
                    self.auth = None
                    self.set_status(self.T["not_found"])
                    return
                self.auth = claude_cli.auth_status(self.cli["path"])
                if self.auth["logged_in"]:
                    who = self.auth.get("email")
                    self.set_status(self.T["logged_as"] % who if who else self.T["logged"])
                else:
                    self.set_status(self.T["not_logged"])
            except Exception:
                log_exc("refresh_auth")
                self.set_status(self.T["not_found"])

        threading.Thread(target=work, daemon=True).start()

    def login(self):
        if not self.cli:
            self.refresh_auth()
            return
        if claude_cli.start_login(self.cli["path"]):
            self.set_foot(self.T["login_started"])

            def poll():
                for _ in range(40):
                    time.sleep(5)
                    if self._disposed:
                        return
                    try:
                        st = claude_cli.auth_status(self.cli["path"])
                    except Exception:
                        continue
                    if st["logged_in"]:
                        self.auth = st
                        who = st.get("email")
                        self.set_status(self.T["logged_as"] % who if who else self.T["logged"])
                        self.set_foot(self.T["ready"])
                        return

            threading.Thread(target=poll, daemon=True).start()
        else:
            self.set_foot(self.T["login_failed"])

    def open_settings(self):
        try:
            peer = self.frame.getContainerWindow()
            new = dialogs.settings_dialog(self.ctx, peer, self.T, self.cfg)
            if new is not None:
                S.save(new)
                self.cfg = new
                for p in list(PANELS):
                    p.cfg = S.load()
                    p.refresh_auth()
        except Exception:
            log_exc("open_settings")

    def new_chat(self):
        if self.busy:
            self.stop()
        self.session_id = str(uuid.uuid4())
        self.resuming = False
        self.m_log.Text = ""
        self.emit(self.T["welcome"] + "\n")

    # ------------------------------------------------------------------ enviar
    def _doc(self):
        try:
            return self.frame.getController().getModel()
        except Exception:
            return None

    def stop(self):
        if self.turn is not None:
            self.turn.stop()
            self.emit("\n" + self.T["stopped"] + "\n")

    def send(self):
        if self.busy:
            self.set_foot(self.T["busy"])
            return
        text = self.m_in.Text.strip()
        if not text:
            return
        doc = self._doc()
        if doc is None or not doc.supportsService("com.sun.star.sheet.SpreadsheetDocument"):
            self.set_foot(self.T["no_doc"])
            return
        self.cfg = S.load()
        if not self.cli:
            self.cli = claude_cli.find_claude(self.cfg, force=True)
        if not self.cli:
            self.set_status(self.T["not_found"])
            return
        self.turn_doc = doc
        # contexto da planilha (lido aqui, na thread principal)
        try:
            cctx = tools.CalcContext(self.ctx, doc, self.cfg)
            info = tools.get_workbook_info(cctx, {})
            sheets = ", ".join("%s (%s)" % (s["name"], s["used_range"] or "vazia") for s in info["sheets"])
            header = "[Contexto: arquivo \"%s\" | planilha ativa: %s | seleção: %s | planilhas: %s]\n\n" % (
                info["file"], info.get("active_sheet"), info.get("selection") or "nenhuma", sheets)
        except Exception:
            log_exc("contexto")
            header = ""
        prompt = header + text
        self.m_in.Text = ""
        self.emit("%s: %s\n\n%s: " % (self.T["you"], text, self.T["claude"]))
        self.set_busy(True)
        self.set_foot(self.T["working"])
        self._t0 = time.time()
        self._stream = _LineStream(self.emit)
        self._after_tool = False
        if not self.resuming:
            self.session_id = str(uuid.uuid4())
        self._ok_turn = False
        self.turn = claude_cli.Turn(self.cli, self.cfg, self.mcp_url, self.token, self.session_id, self.resuming,
                                    prompt, prompts.build_system(self.cfg.get("extra_instructions")), self._on_event)
        self.turn.start()

    # worker thread -----------------------------------------------------
    def _on_event(self, kind, d):
        try:
            if kind == "init":
                if d["mcp"].get("calc") not in (None, "connected"):
                    self.emit("\n⚠ MCP: %s\n" % d["mcp"])
            elif kind == "text":
                if self._after_tool:
                    self.emit("\n")
                    self._after_tool = False
                self._stream.feed(d)
                self._text_open = True
            elif kind == "tool_use":
                self._stream.flush()
                label = prompts.summarize_tool(d["name"], d["input"])
                self.emit(("" if self._at_bol else "\n") + "  · %s\n" % label)
                self.set_foot(self.T["using"] % label)
                self._after_tool = True
            elif kind == "tool_err":
                msg = d
                try:
                    j = json.loads(d)
                    if isinstance(j, dict) and j.get("error"):
                        msg = j["error"]
                except Exception:
                    pass
                self.emit("  ⚠ %s\n" % msg)
            elif kind == "result":
                self._stream.flush()
                if d.get("is_error"):
                    msg = d.get("text") or "erro"
                    self.emit("\n⚠ %s\n" % msg)
                    if re.search(r"log ?in|authenticat|api key|unauthorized|credential", msg, re.I):
                        self.emit("%s\n" % self.T["auth_hint"])
                else:
                    self._ok_turn = True
                    self.resuming = True
            elif kind == "error":
                self._stream.flush()
                self.emit("\n⚠ %s\n" % d)
            elif kind == "done":
                self._stream.flush()
                self.emit("\n\n")
                self.set_busy(False)
                self.set_foot(self.T["finished"] % (time.time() - self._t0))
                self.turn = None
        except Exception:
            log_exc("on_event")

    # ferramentas (chamadas pelo servidor MCP, thread HTTP) --------------------
    def exec_tool(self, name, args):
        short = name

        def run():
            doc = self.turn_doc or self._doc()
            if doc is None:
                return ('{"error": "Nenhuma planilha aberta."}', True)
            cctx = tools.CalcContext(self.ctx, doc, self.cfg)
            if self.cfg.get("confirm_writes") and tools.tool_is_write(short):
                if not self._confirm(prompts.summarize_tool(short, args)):
                    return ('{"error": "%s"}' % self.T["denied"], True)
            return tools.run_tool(cctx, short, args)

        return self.main.call(run, timeout=180)

    def _confirm(self, summary):
        T = self.T
        MBT = uno.Enum("com.sun.star.awt.MessageBoxType", "QUERYBOX")
        yes_no = uno.getConstantByName("com.sun.star.awt.MessageBoxButtons.BUTTONS_YES_NO")
        box = self._toolkit.createMessageBox(self.frame.getContainerWindow(), MBT, yes_no, T["confirm_title"],
                                             T["confirm_body"] % summary)
        return box.execute() == uno.getConstantByName("com.sun.star.awt.MessageBoxResults.YES")
