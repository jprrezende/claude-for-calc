# -*- coding: utf-8 -*-
"""Prompt de sistema e utilitários de texto."""
import re

SYSTEM = """\
Você é o Claude, integrado ao LibreOffice Calc como assistente de planilhas (equivalente ao "Claude para Excel").
Você trabalha SOMENTE na planilha aberta do usuário, usando as ferramentas mcp__calc__*. Você não tem acesso a
arquivos, internet nem terminal.

Princípios
- Responda no idioma do usuário, de forma curta e direta. Cite células como Planilha!B2 ou B2:D9.
- Antes de afirmar algo sobre os dados, leia-os (get_workbook_info, read_range, find_cells). Nunca invente conteúdo.
- Para explicar um valor, siga as fórmulas: leia as células referenciadas e descreva a cadeia de cálculo.
- Prefira fórmulas a números digitados; mantenha premissas/entradas separadas dos cálculos; não sobrescreva
  dados do usuário sem ser pedido. Não use valores fixos dentro de fórmulas quando houver uma célula de premissa.
- Fórmulas: sintaxe Excel com funções em inglês, argumentos separados por vírgula e referências como
  Planilha2!A1 ou 'Minha Aba'!A1. Texto comum é gravado como texto; só strings iniciadas por "=" viram fórmula.
  Datas são números seriais: formate com number_format (ex.: dd/mm/yyyy).
- Depois de escrever fórmulas, confira `formula_errors` e corrija. Se uma função devolver #NAME?, use outra
  equivalente suportada pelo LibreOffice.
- Intervalos grandes: leia em partes. Para tarefas amplas, faça o plano em poucas linhas e execute em lotes.
- Use select_range para mostrar ao usuário uma célula ou intervalo importante (por exemplo, a origem de um erro).
- Cada alteração pode ser desfeita pelo usuário com Ctrl+Z. Em mudanças grandes ou destrutivas, resuma o que
  vai fazer antes de fazer.
- Ao terminar, diga em 1–3 frases o que mudou e onde. Não repita tabelas inteiras que já estão na planilha.
"""


def build_system(extra):
    extra = (extra or "").strip()
    return SYSTEM + ("\nInstruções adicionais do usuário:\n" + extra + "\n" if extra else "")


_TOOL_LABELS = {
    "get_workbook_info": ("lendo a estrutura da planilha", None),
    "read_range": ("lendo", "{ref}"),
    "find_cells": ("procurando", "“{query}”"),
    "find_errors": ("procurando erros de fórmula", None),
    "get_format": ("lendo formatação de", "{ref}"),
    "write_range": ("escrevendo em", "{ref}"),
    "clear_range": ("limpando", "{ref}"),
    "format_range": ("formatando", "{ref}"),
    "copy_range": ("copiando", "{src} → {dst}"),
    "insert_delete": ("{action} {dim}", "{start}"),
    "resize": ("ajustando tamanho de", "{ref}"),
    "sort_range": ("ordenando", "{ref}"),
    "manage_sheet": ("planilha: {action}", "{name}"),
    "freeze_panes": ("congelando painéis", None),
    "add_comment": ("comentando", "{ref}"),
    "define_name": ("definindo nome", "{name}"),
    "select_range": ("selecionando", "{ref}"),
    "create_chart": ("criando gráfico de", "{ref}"),
    "create_pivot_table": ("criando tabela dinâmica de", "{ref}"),
}


def summarize_tool(name, inp):
    n = name.replace("mcp__calc__", "")
    inp = inp or {}
    ref = inp.get("range") or inp.get("start_cell") or inp.get("data_range") or inp.get("cell") or inp.get("source_range") or ""
    if inp.get("sheet") and ref and "!" not in ref:
        ref = "%s!%s" % (inp["sheet"], ref)
    fields = {"ref": ref, "query": inp.get("query", ""), "src": inp.get("source", ""), "dst": inp.get("destination", ""),
              "action": inp.get("action", ""), "dim": {"rows": "linhas", "columns": "colunas"}.get(inp.get("dimension"), ""),
              "start": inp.get("start", ""), "name": inp.get("name", "")}
    verb, obj = _TOOL_LABELS.get(n, (n, None))
    try:
        s = verb.format(**fields)
        if obj:
            s += " " + obj.format(**fields)
    except Exception:
        s = n
    if n == "write_range":
        v = inp.get("values")
        if isinstance(v, list) and v:
            rows = len(v)
            cols = max((len(r) for r in v if isinstance(r, list)), default=1)
            s += " (%d×%d)" % (rows, cols)
    return s.strip()


def clean_markdown(text):
    """Limpa marcações do Markdown para exibir em um campo de texto simples."""
    text = re.sub(r"\*\*(.+?)\*\*", r"\1", text)
    text = re.sub(r"(?<!\w)__(.+?)__(?!\w)", r"\1", text)
    text = re.sub(r"`([^`\n]+)`", r"\1", text)
    text = re.sub(r"^\s{0,3}#{1,6}\s*", "", text, flags=re.M)
    text = re.sub(r"^(\s*)[-*]\s+", lambda m: m.group(1) + "• ", text, flags=re.M)
    return text
