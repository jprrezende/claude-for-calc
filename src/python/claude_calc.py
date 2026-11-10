# -*- coding: utf-8 -*-
"""Claude para Calc – componente UNO (fábrica do painel lateral + comandos de menu)."""
import os
import sys

import uno
import unohelper
from com.sun.star.frame import XDispatch, XDispatchProvider
from com.sun.star.lang import XInitialization, XServiceInfo
from com.sun.star.ui import XUIElementFactory

_HERE = os.path.dirname(os.path.abspath(__file__))
if _HERE not in sys.path:
    sys.path.insert(0, _HERE)

IMPL_FACTORY = "com.claudecalc.sidebar.PanelFactory"
IMPL_DISPATCH = "com.claudecalc.Dispatch"
PROTOCOL = "com.claudecalc.dispatch:"


def _log_exc(where):
    try:
        from claudecalc.util import log_exc
        log_exc(where)
    except Exception:
        pass


class PanelFactory(unohelper.Base, XUIElementFactory, XServiceInfo):
    def __init__(self, ctx):
        self.ctx = ctx

    def createUIElement(self, resource_url, args):
        try:
            from claudecalc.panel import ClaudePanel
            d = {a.Name: a.Value for a in args}
            return ClaudePanel(self.ctx, resource_url, d.get("Frame"), d.get("ParentWindow"))
        except Exception:
            _log_exc("createUIElement")
            raise

    def getImplementationName(self):
        return IMPL_FACTORY

    def supportsService(self, name):
        return name in self.getSupportedServiceNames()

    def getSupportedServiceNames(self):
        return ("com.sun.star.ui.UIElementFactory",)


class Dispatch(unohelper.Base, XServiceInfo, XDispatchProvider, XDispatch, XInitialization):
    def __init__(self, ctx):
        self.ctx = ctx
        self.frame = None

    # XInitialization
    def initialize(self, args):
        if args:
            self.frame = args[0]

    # XDispatchProvider
    def queryDispatch(self, url, target, flags):
        if url.Protocol == PROTOCOL:
            return self
        return None

    def queryDispatches(self, requests):
        return tuple(self.queryDispatch(r.FeatureURL, r.FrameName, r.SearchFlags) for r in requests)

    # XDispatch
    def dispatch(self, url, args):
        cmd = url.Path
        try:
            if cmd == "OpenPanel":
                self._open_panel()
            elif cmd == "Settings":
                self._settings()
            elif cmd == "Login":
                self._login()
        except Exception:
            _log_exc("dispatch " + cmd)

    def addStatusListener(self, listener, url):
        pass

    def removeStatusListener(self, listener, url):
        pass

    def _open_panel(self):
        smgr = self.ctx.ServiceManager
        helper = smgr.createInstanceWithContext("com.sun.star.frame.DispatchHelper", self.ctx)
        helper.executeDispatch(self.frame, ".uno:SidebarDeck.ClaudeCalcDeck", "", 0, ())

    def _find_panel(self):
        from claudecalc import panel
        for p in panel.PANELS:
            try:
                if p.frame == self.frame or p._doc() == self.frame.getController().getModel():
                    return p
            except Exception:
                continue
        return panel.PANELS[0] if panel.PANELS else None

    def _settings(self):
        p = self._find_panel()
        if p is None:
            self._open_panel()
            return
        p.open_settings()

    def _login(self):
        p = self._find_panel()
        if p is None:
            self._open_panel()
            return
        p.login()

    # XServiceInfo
    def getImplementationName(self):
        return IMPL_DISPATCH

    def supportsService(self, name):
        return name in self.getSupportedServiceNames()

    def getSupportedServiceNames(self):
        return ("com.sun.star.frame.ProtocolHandler",)


g_ImplementationHelper = unohelper.ImplementationHelper()
g_ImplementationHelper.addImplementation(PanelFactory, IMPL_FACTORY, ("com.sun.star.ui.UIElementFactory",))
g_ImplementationHelper.addImplementation(Dispatch, IMPL_DISPATCH, ("com.sun.star.frame.ProtocolHandler",))
