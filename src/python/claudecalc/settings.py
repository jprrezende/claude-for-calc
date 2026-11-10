# -*- coding: utf-8 -*-
"""Configurações persistentes (JSON em %APPDATA%\\ClaudeCalc no Windows)."""
import json
import os

APP = "ClaudeCalc"

DEFAULTS = {
    "claude_path": "",          # vazio = detectar automaticamente
    "model": "",                # vazio = padrão do Claude Code (ou alias: sonnet, opus, haiku)
    "confirm_writes": False,    # pedir confirmação antes de cada alteração
    "max_read_cells": 5000,     # limite de células por leitura
    "extra_instructions": "",   # instruções extras anexadas ao prompt de sistema
    "timeout_minutes": 15,      # tempo máximo por pergunta
}


def data_dir():
    if os.name == "nt":
        base = os.environ.get("APPDATA") or os.path.expanduser("~")
    else:
        base = os.environ.get("XDG_CONFIG_HOME") or os.path.expanduser("~/.config")
    d = os.path.join(base, APP)
    os.makedirs(d, exist_ok=True)
    return d


def workspace_dir():
    d = os.path.join(data_dir(), "workspace")
    os.makedirs(d, exist_ok=True)
    return d


def _path():
    return os.path.join(data_dir(), "settings.json")


def load():
    s = dict(DEFAULTS)
    try:
        with open(_path(), "r", encoding="utf-8") as f:
            s.update({k: v for k, v in json.load(f).items() if k in DEFAULTS})
    except Exception:
        pass
    return s


def save(s):
    clean = {k: s.get(k, DEFAULTS[k]) for k in DEFAULTS}
    tmp = _path() + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(clean, f, ensure_ascii=False, indent=2)
    os.replace(tmp, _path())
