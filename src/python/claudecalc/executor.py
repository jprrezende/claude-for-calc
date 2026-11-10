# -*- coding: utf-8 -*-
"""Executa funções na thread principal do LibreOffice (via AsyncCallback)."""
import threading

import unohelper
from com.sun.star.awt import XCallback


class _Cb(unohelper.Base, XCallback):
    def __init__(self, fn, done):
        self.fn = fn
        self.done = done

    def notify(self, data):
        try:
            self.fn()
        finally:
            self.done(self)


class MainThread:
    """Deve ser criado na thread principal (ex.: ao montar o painel)."""

    def __init__(self, uno_ctx):
        self.uno_ctx = uno_ctx
        self.main_ident = threading.get_ident()
        self._svc = uno_ctx.ServiceManager.createInstanceWithContext("com.sun.star.awt.AsyncCallback", uno_ctx)
        self._alive = set()
        self._lock = threading.Lock()

    def _done(self, cb):
        with self._lock:
            self._alive.discard(cb)

    def post(self, fn):
        """Agenda `fn` na thread principal e retorna imediatamente."""
        if threading.get_ident() == self.main_ident:
            # já estamos na principal; ainda assim adia para não reentrar em callbacks de UI
            pass
        cb = _Cb(fn, self._done)
        with self._lock:
            self._alive.add(cb)
        self._svc.addCallback(cb, None)

    def call(self, fn, timeout=120.0):
        """Executa `fn` na thread principal e devolve o resultado (bloqueia quem chama)."""
        if threading.get_ident() == self.main_ident:
            return fn()
        ev = threading.Event()
        box = {}

        def run():
            try:
                box["r"] = fn()
            except BaseException as e:  # noqa
                box["e"] = e
            finally:
                ev.set()

        self.post(run)
        if not ev.wait(timeout):
            raise TimeoutError("O LibreOffice não respondeu em %ds (há algum diálogo aberto?)." % int(timeout))
        if "e" in box:
            raise box["e"]
        return box.get("r")
