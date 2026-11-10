# -*- coding: utf-8 -*-
import os
import time
import traceback

from . import settings as S


def log(msg):
    try:
        p = os.path.join(S.data_dir(), "claudecalc.log")
        if os.path.exists(p) and os.path.getsize(p) > 300000:
            os.replace(p, p + ".old")
        with open(p, "a", encoding="utf-8") as f:
            f.write("%s %s\n" % (time.strftime("%Y-%m-%d %H:%M:%S"), msg))
    except Exception:
        pass


def log_exc(where):
    log("ERRO em %s:\n%s" % (where, traceback.format_exc()))
