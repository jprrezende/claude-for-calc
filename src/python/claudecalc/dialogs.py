# -*- coding: utf-8 -*-
"""Diálogo de configurações (montado por código, sem arquivo XDL)."""
import uno
import unohelper
from com.sun.star.awt import XActionListener

from . import claude_cli
from .util import log_exc


class _Detect(unohelper.Base, XActionListener):
    def __init__(self, fn):
        self.fn = fn

    def actionPerformed(self, ev):
        try:
            self.fn()
        except Exception:
            log_exc("detect")

    def disposing(self, src):
        pass


def settings_dialog(uno_ctx, parent_peer, T, cfg):
    """Mostra o diálogo. Devolve o novo dict de configurações, ou None se cancelado."""
    smgr = uno_ctx.ServiceManager
    dm = smgr.createInstanceWithContext("com.sun.star.awt.UnoControlDialogModel", uno_ctx)
    dm.Width, dm.Height = 270, 196
    dm.Title = T["settings_title"]

    def add(name, kind, x, y, w, h, **props):
        m = dm.createInstance("com.sun.star.awt.UnoControl%sModel" % kind)
        m.PositionX, m.PositionY, m.Width, m.Height = x, y, w, h
        for k, v in props.items():
            m.setPropertyValue(k, v)
        dm.insertByName(name, m)
        return m

    add("l1", "FixedText", 8, 8, 250, 9, Label=T["s_path"])
    m_path = add("path", "Edit", 8, 19, 190, 12, Text=cfg.get("claude_path", ""))
    add("btnDetect", "Button", 202, 18, 60, 14, Label=T["s_detect"])
    m_found = add("found", "FixedText", 8, 34, 254, 9, Label="")
    add("l2", "FixedText", 8, 50, 250, 9, Label=T["s_model"])
    m_model = add("model", "Edit", 8, 61, 120, 12, Text=cfg.get("model", ""))
    m_conf = add("confirm", "CheckBox", 8, 80, 254, 10, Label=T["s_confirm"], State=1 if cfg.get("confirm_writes") else 0)
    add("l3", "FixedText", 8, 97, 150, 9, Label=T["s_maxcells"])
    m_max = add("maxc", "NumericField", 160, 95, 60, 12, Value=float(cfg.get("max_read_cells", 5000)),
                ValueMin=100.0, ValueMax=200000.0, DecimalAccuracy=0, Spin=False)
    add("l4", "FixedText", 8, 114, 250, 9, Label=T["s_extra"])
    m_extra = add("extra", "Edit", 8, 125, 254, 42, Text=cfg.get("extra_instructions", ""), MultiLine=True,
                  VScroll=True, HScroll=False)
    add("ok", "Button", 150, 175, 55, 14, Label=T["s_ok"], PushButtonType=1, DefaultButton=True)
    add("cancel", "Button", 209, 175, 53, 14, Label=T["s_cancel"], PushButtonType=2)

    dlg = smgr.createInstanceWithContext("com.sun.star.awt.UnoControlDialog", uno_ctx)
    dlg.setModel(dm)

    def detect():
        tmp = dict(cfg)
        tmp["claude_path"] = m_path.Text.strip()
        res = claude_cli.find_claude(tmp, force=True)
        m_found.Label = (T["s_found"] % ("%s – %s" % (res["path"], res["version"]))) if res else T["s_notfound"]

    dlg.getControl("btnDetect").addActionListener(_Detect(detect))
    toolkit = smgr.createInstanceWithContext("com.sun.star.awt.Toolkit", uno_ctx)
    dlg.createPeer(toolkit, parent_peer)
    try:
        detect()
    except Exception:
        log_exc("detect-initial")
    rc = dlg.execute()
    out = None
    if rc == 1:
        out = dict(cfg)
        out["claude_path"] = m_path.Text.strip()
        out["model"] = m_model.Text.strip()
        out["confirm_writes"] = bool(m_conf.State)
        out["max_read_cells"] = int(m_max.Value)
        out["extra_instructions"] = m_extra.Text
    dlg.dispose()
    return out
