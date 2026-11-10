# -*- coding: utf-8 -*-
"""Textos da interface (pt-BR e en)."""

_PT = {
    "send": "Enviar", "stop": "Parar", "new_chat": "Nova conversa", "settings": "Config.", "login": "Entrar",
    "checking": "Verificando o Claude…",
    "not_found": "Claude Code não encontrado. Abra Config. e informe o caminho.",
    "not_logged": "Não conectado. Clique em Entrar.",
    "logged": "Conectado", "logged_as": "Conectado como %s",
    "ready": "Pronto. Ctrl+Enter envia.",
    "working": "Trabalhando…", "using": "Usando: %s", "finished": "Concluído em %.1f s",
    "stopped": "Interrompido.", "you": "Você", "claude": "Claude",
    "busy": "Aguarde a resposta atual terminar.",
    "login_started": "Abri o login do Claude no navegador. Depois de entrar, volte aqui.",
    "login_failed": "Não consegui iniciar o login. Rode 'claude auth login' em um terminal.",
    "no_doc": "Abra uma planilha do Calc para usar o Claude.",
    "auth_hint": "Parece um problema de login. Clique em Entrar (ou rode 'claude auth login').",
    "confirm_title": "Claude quer alterar a planilha",
    "confirm_body": "Permitir esta ação?\n\n%s",
    "denied": "O usuário recusou esta alteração.",
    "welcome": "Olá! Pergunte sobre a planilha aberta ou peça alterações (fórmulas, formatação, gráficos…). "
               "Tudo pode ser desfeito com Ctrl+Z.\n",
    "settings_title": "Claude para Calc – Configurações",
    "s_path": "Caminho do Claude Code (vazio = detectar):", "s_detect": "Detectar",
    "s_model": "Modelo (vazio = padrão; ex.: sonnet, opus):",
    "s_confirm": "Pedir confirmação antes de cada alteração na planilha",
    "s_maxcells": "Máximo de células por leitura:",
    "s_extra": "Instruções extras para o Claude:",
    "s_ok": "Salvar", "s_cancel": "Cancelar",
    "s_found": "Encontrado: %s", "s_notfound": "Não encontrado.",
}

_EN = {
    "send": "Send", "stop": "Stop", "new_chat": "New chat", "settings": "Settings", "login": "Sign in",
    "checking": "Checking Claude…",
    "not_found": "Claude Code not found. Open Settings and set its path.",
    "not_logged": "Not signed in. Click Sign in.",
    "logged": "Connected", "logged_as": "Connected as %s",
    "ready": "Ready. Ctrl+Enter sends.",
    "working": "Working…", "using": "Using: %s", "finished": "Done in %.1f s",
    "stopped": "Stopped.", "you": "You", "claude": "Claude",
    "busy": "Wait for the current answer to finish.",
    "login_started": "Claude sign-in opened in your browser. Come back here afterwards.",
    "login_failed": "Could not start sign-in. Run 'claude auth login' in a terminal.",
    "no_doc": "Open a Calc spreadsheet to use Claude.",
    "auth_hint": "Looks like a sign-in problem. Click Sign in (or run 'claude auth login').",
    "confirm_title": "Claude wants to change the spreadsheet",
    "confirm_body": "Allow this action?\n\n%s",
    "denied": "The user declined this change.",
    "welcome": "Hi! Ask about the open spreadsheet or request changes (formulas, formatting, charts…). "
               "Everything can be undone with Ctrl+Z.\n",
    "settings_title": "Claude for Calc – Settings",
    "s_path": "Claude Code path (empty = auto-detect):", "s_detect": "Detect",
    "s_model": "Model (empty = default; e.g. sonnet, opus):",
    "s_confirm": "Ask for confirmation before every spreadsheet change",
    "s_maxcells": "Max cells per read:",
    "s_extra": "Extra instructions for Claude:",
    "s_ok": "Save", "s_cancel": "Cancel",
    "s_found": "Found: %s", "s_notfound": "Not found.",
}


def ui_lang(uno_ctx):
    """'pt' ou 'en' conforme o idioma da interface do LibreOffice."""
    try:
        prov = uno_ctx.ServiceManager.createInstanceWithContext("com.sun.star.configuration.ConfigurationProvider", uno_ctx)
        import uno
        pv = uno.createUnoStruct("com.sun.star.beans.PropertyValue")
        pv.Name, pv.Value = "nodepath", "/org.openoffice.Setup/L10N"
        acc = prov.createInstanceWithArguments("com.sun.star.configuration.ConfigurationAccess", (pv,))
        loc = acc.getByName("ooLocale") or ""
        return "pt" if loc.lower().startswith("pt") else "en"
    except Exception:
        return "pt"


def table(lang):
    return _PT if lang == "pt" else _EN
