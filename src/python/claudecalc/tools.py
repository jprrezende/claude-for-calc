# -*- coding: utf-8 -*-
"""Ferramentas de planilha que o Claude pode chamar (expostas via MCP).

Tudo aqui roda na thread principal do LibreOffice (ver executor.py) e usa
apenas a API UNO do Calc. Fórmulas são aceitas em sintaxe Excel (vírgulas,
Planilha1!A1) e convertidas para o Calc através do FormulaParser.
"""
import json
import re

import uno


class ToolError(Exception):
    """Erro "esperado" – a mensagem volta para o Claude para ele se corrigir."""


TOOLS = {}


def _obj(props, required=None):
    return {"type": "object", "properties": props, "required": required or [],
            "additionalProperties": False}


def _p(t, d, **kw):
    r = {"type": t, "description": d}
    r.update(kw)
    return r


SHEET_P = _p("string", "Nome da planilha (aba). Se omitido, usa a planilha ativa.")


def tool(name, description, props=None, required=None, writes=False):
    def deco(fn):
        TOOLS[name] = {
            "name": name,
            "description": description,
            "inputSchema": _obj(props or {}, required),
            "fn": fn,
            "writes": writes,
        }
        return fn
    return deco


# --------------------------------------------------------------------------
# Utilidades de endereço
# --------------------------------------------------------------------------

def col_to_letters(c):
    s = ""
    c += 1
    while c:
        c, rem = divmod(c - 1, 26)
        s = chr(65 + rem) + s
    return s


def letters_to_col(s):
    n = 0
    for ch in s.upper():
        n = n * 26 + ord(ch) - 64
    return n - 1


def cell_a1(col, row):
    return "%s%d" % (col_to_letters(col), row + 1)


def range_a1(a):
    """CellRangeAddress -> 'A1:C3' (ou 'A1' para célula única)."""
    s = cell_a1(a.StartColumn, a.StartRow)
    if a.StartColumn == a.EndColumn and a.StartRow == a.EndRow:
        return s
    return s + ":" + cell_a1(a.EndColumn, a.EndRow)


def quote_sheet(name):
    if re.match(r"^[A-Za-z_][A-Za-z0-9_]*$", name):
        return name
    return "'" + name.replace("'", "''") + "'"


_CELL_RE = re.compile(r"^\$?([A-Za-z]{1,3})\$?(\d+)$")


def parse_cell(ref):
    m = _CELL_RE.match((ref or "").strip())
    if not m:
        raise ToolError("Endereço de célula inválido: %r (use algo como B2)" % ref)
    return letters_to_col(m.group(1)), int(m.group(2)) - 1


def split_ref(ref):
    """'Planilha 1'!A1:B2 | Plan1!A1 | $Plan1.A1 | A1:B2 -> (planilha|None, 'A1:B2')."""
    ref = (ref or "").strip()
    m = re.match(r"^(?:'((?:[^']|'')+)'|([^'!]+))!(.+)$", ref)
    if m:
        name = (m.group(1) or m.group(2)).replace("''", "'")
        return name.strip(), m.group(3).strip()
    m = re.match(r"^\$?(?:'((?:[^']|'')+)'|([^.'$:]+))\.(\$?[A-Za-z]+\$?\d*(?::\$?[A-Za-z]+\$?\d*)?)$", ref)
    if m:
        name = (m.group(1) or m.group(2)).replace("''", "'")
        return name.strip(), m.group(3).strip()
    return None, ref


COLORS = {
    "black": 0x000000, "white": 0xFFFFFF, "red": 0xFF0000, "green": 0x00A000,
    "blue": 0x0000FF, "yellow": 0xFFFF00, "orange": 0xFF8000, "gray": 0x808080,
    "grey": 0x808080, "lightgray": 0xD9D9D9, "lightgrey": 0xD9D9D9,
    "purple": 0x800080, "pink": 0xFFC0CB, "brown": 0x8B4513, "lightblue": 0xADD8E6,
    "lightgreen": 0x90EE90, "darkblue": 0x00008B, "darkgreen": 0x006400,
}


def parse_color(v):
    if isinstance(v, int):
        return v
    s = str(v).strip().lower().replace(" ", "")
    if s in COLORS:
        return COLORS[s]
    s = s.lstrip("#")
    if re.match(r"^[0-9a-f]{6}$", s):
        return int(s, 16)
    raise ToolError("Cor inválida: %r (use #RRGGBB ou um nome como red, blue)" % v)


# --------------------------------------------------------------------------
# Contexto
# --------------------------------------------------------------------------

class CalcContext:
    def __init__(self, uno_ctx, doc, settings=None):
        self.uno_ctx = uno_ctx
        self.smgr = uno_ctx.ServiceManager
        self.doc = doc
        self.settings = settings or {}
        self._parser = None

    # ---- documento / vista
    @property
    def controller(self):
        return self.doc.getCurrentController()

    def active_sheet(self):
        c = self.controller
        if c is not None:
            try:
                return c.getActiveSheet()
            except Exception:
                pass
        return self.doc.Sheets.getByIndex(0)

    def get_sheet(self, name=None):
        if not name:
            return self.active_sheet()
        sheets = self.doc.Sheets
        if sheets.hasByName(name):
            return sheets.getByName(name)
        for n in sheets.ElementNames:
            if n.lower() == name.lower():
                return sheets.getByName(n)
        raise ToolError("A planilha %r não existe. Planilhas: %s" % (name, ", ".join(sheets.ElementNames)))

    def sheet_name(self, sheet):
        return sheet.Name

    def sheet_index(self, sheet):
        return sheet.RangeAddress.Sheet

    def max_cells(self):
        try:
            return int(self.settings.get("max_read_cells", 5000))
        except Exception:
            return 5000

    # ---- intervalos
    def resolve(self, sheet_name, ref, default_selection=False):
        """Devolve (planilha, intervalo). `ref` pode conter a planilha ('Plan1!A1')."""
        if (ref is None or str(ref).strip() == "") and default_selection:
            sel = self.selection_range()
            if sel is None:
                raise ToolError("Nenhuma célula selecionada. Informe o intervalo.")
            return self.doc.Sheets.getByIndex(sel.RangeAddress.Sheet), sel
        if ref is None or str(ref).strip() == "":
            raise ToolError("Informe o intervalo (ex.: A1:C10).")
        sn, rr = split_ref(str(ref))
        sheet = self.get_sheet(sn or sheet_name)
        try:
            rng = sheet.getCellRangeByName(rr.replace("$", ""))
        except Exception:
            rng = None
        if rng is None:
            raise ToolError("Intervalo inválido: %r" % ref)
        return sheet, rng

    def selection_range(self):
        c = self.controller
        if c is None:
            return None
        try:
            sel = c.getSelection()
        except Exception:
            return None
        if sel is None:
            return None
        try:
            if sel.supportsService("com.sun.star.sheet.SheetCellRange"):
                return sel
            if sel.supportsService("com.sun.star.sheet.SheetCellRanges"):
                addrs = sel.getRangeAddresses()
                if addrs:
                    a = addrs[0]
                    return self.doc.Sheets.getByIndex(a.Sheet).getCellRangeByPosition(
                        a.StartColumn, a.StartRow, a.EndColumn, a.EndRow)
        except Exception:
            return None
        return None

    def selection_text(self):
        sel = self.selection_range()
        if sel is None:
            return None
        a = sel.RangeAddress
        return "%s!%s" % (quote_sheet(self.doc.Sheets.getByIndex(a.Sheet).Name), range_a1(a))

    def used_area(self, sheet):
        """(c0, r0, c1, r1) da área usada, ou None se vazia."""
        cur = sheet.createCursor()
        cur.gotoEndOfUsedArea(False)
        a = cur.RangeAddress
        if a.EndColumn == 0 and a.EndRow == 0:
            if sheet.getCellByPosition(0, 0).Type.value == "EMPTY":
                return None
        return (0, 0, a.EndColumn, a.EndRow)

    # ---- fórmulas
    def parser(self):
        if self._parser is None:
            p = self.doc.createInstance("com.sun.star.sheet.FormulaParser")
            p.FormulaConvention = uno.getConstantByName("com.sun.star.sheet.AddressConvention.XL_A1")
            p.CompileEnglish = True
            self._parser = p
        return self._parser

    def set_formula(self, cell, text, sheet_idx, col, row):
        body = excel_to_api(text)[1:]
        addr = uno.createUnoStruct("com.sun.star.table.CellAddress")
        addr.Sheet, addr.Column, addr.Row = sheet_idx, col, row
        try:
            toks = self.parser().parseFormula(body, addr)
            cell.setTokens(toks)
        except Exception as e:  # noqa
            raise ToolError("Fórmula inválida em %s: %s (%s)" % (cell_a1(col, row), text, e))

    def formula_text(self, cell):
        """Fórmula de uma célula em sintaxe Excel (com '=')."""
        try:
            addr = uno.createUnoStruct("com.sun.star.table.CellAddress")
            ca = cell.CellAddress
            addr.Sheet, addr.Column, addr.Row = ca.Sheet, ca.Column, ca.Row
            txt = self.parser().printFormula(cell.getTokens(), addr)
            return "=" + api_to_excel(txt)
        except Exception:
            return cell.getFormula()


def excel_to_api(f):
    """Converte separadores Excel -> gramática API inglesa do Calc (fora de strings)."""
    out = []
    i, n, brace = 0, len(f), 0
    while i < n:
        ch = f[i]
        if ch == '"':
            j = i + 1
            while j < n:
                if f[j] == '"':
                    if j + 1 < n and f[j + 1] == '"':
                        j += 2
                        continue
                    break
                j += 1
            out.append(f[i:j + 1])
            i = j + 1
            continue
        if ch == "'":
            j = f.find("'", i + 1)
            while j != -1 and j + 1 < n and f[j + 1] == "'":
                j = f.find("'", j + 2)
            if j == -1:
                out.append(f[i:])
                break
            out.append(f[i:j + 1])
            i = j + 1
            continue
        if ch == "{":
            brace += 1
        elif ch == "}":
            brace -= 1
        if ch == ",":
            out.append(";")
        elif ch == ";" and brace > 0:
            out.append("|")
        else:
            out.append(ch)
        i += 1
    return "".join(out)


def api_to_excel(f):
    """Inverso de excel_to_api (para devolver fórmulas ao Claude)."""
    out = []
    i, n, brace = 0, len(f), 0
    while i < n:
        ch = f[i]
        if ch == '"':
            j = i + 1
            while j < n:
                if f[j] == '"':
                    if j + 1 < n and f[j + 1] == '"':
                        j += 2
                        continue
                    break
                j += 1
            out.append(f[i:j + 1])
            i = j + 1
            continue
        if ch == "'":
            j = f.find("'", i + 1)
            while j != -1 and j + 1 < n and f[j + 1] == "'":
                j = f.find("'", j + 2)
            if j == -1:
                out.append(f[i:])
                break
            out.append(f[i:j + 1])
            i = j + 1
            continue
        if ch == "{":
            brace += 1
        elif ch == "}":
            brace -= 1
        if ch == ";":
            out.append(",")
        elif ch == "|" and brace > 0:
            out.append(";")
        else:
            out.append(ch)
        i += 1
    return "".join(out)


class _Undo:
    def __init__(self, ctx, title):
        self.um = None
        try:
            self.um = ctx.doc.getUndoManager()
        except Exception:
            pass
        self.title = title
        self.entered = False

    def __enter__(self):
        if self.um is not None:
            try:
                self.um.enterUndoContext(self.title)
                self.entered = True
            except Exception:
                pass
        return self

    def __exit__(self, *a):
        if self.entered:
            try:
                self.um.leaveUndoContext()
            except Exception:
                pass
        return False


# --------------------------------------------------------------------------
# Leitura
# --------------------------------------------------------------------------

def _range_rect(rng):
    a = rng.RangeAddress
    return a.StartColumn, a.StartRow, a.EndColumn, a.EndRow


def _each_cell_in(addresses):
    for a in addresses:
        for r in range(a.StartRow, a.EndRow + 1):
            for c in range(a.StartColumn, a.EndColumn + 1):
                yield a.Sheet, c, r


@tool("get_workbook_info",
      "Visão geral da planilha aberta: arquivo, planilhas (com a área usada de cada uma), planilha ativa, "
      "seleção atual, nomes definidos, gráficos e tabelas dinâmicas. Chame isto primeiro para se orientar.")
def get_workbook_info(ctx, a):
    doc = ctx.doc
    info = {"file": doc.getTitle(), "url": doc.getURL() or None}
    try:
        info["active_sheet"] = ctx.active_sheet().Name
    except Exception:
        info["active_sheet"] = None
    info["selection"] = ctx.selection_text()
    sheets = []
    for i in range(doc.Sheets.Count):
        sh = doc.Sheets.getByIndex(i)
        ua = ctx.used_area(sh)
        d = {"index": i, "name": sh.Name, "visible": bool(sh.IsVisible),
             "used_range": None if ua is None else "A1:%s" % cell_a1(ua[2], ua[3])}
        try:
            ch = list(sh.Charts.ElementNames)
            if ch:
                d["charts"] = ch
        except Exception:
            pass
        try:
            pv = list(sh.DataPilotTables.ElementNames)
            if pv:
                d["pivot_tables"] = pv
        except Exception:
            pass
        if getattr(sh, "IsProtected", False) or (hasattr(sh, "isProtected") and sh.isProtected()):
            d["protected"] = True
        sheets.append(d)
    info["sheets"] = sheets
    names = []
    try:
        nr = doc.NamedRanges
        for n in nr.ElementNames:
            names.append({"name": n, "refers_to": nr.getByName(n).Content})
    except Exception:
        pass
    info["named_ranges"] = names
    info["formula_syntax_note"] = "Escreva fórmulas em sintaxe Excel (vírgulas, Planilha!A1)."
    return info


@tool("read_range",
      "Lê valores e fórmulas de um intervalo. Se `range` for omitido, lê a seleção atual. Aceita 'A1:F20', "
      "'A:C', 'Planilha2!B2:D9'. Devolve `values` (matriz, linha a linha), `formulas` (só células com fórmula, "
      "em sintaxe Excel) e `errors`. Intervalos grandes são limitados: leia em partes.",
      {"sheet": SHEET_P,
       "range": _p("string", "Intervalo em notação A1. Omita para usar a seleção atual."),
       "formatted": _p("boolean", "Se true, devolve também `display` com o texto formatado como aparece na tela "
                                  "(útil para datas, moedas e percentuais). Padrão: false.")})
def read_range(ctx, a):
    sheet, rng = ctx.resolve(a.get("sheet"), a.get("range"), default_selection=True)
    c0, r0, c1, r1 = _range_rect(rng)
    ua = ctx.used_area(sheet)
    requested = range_a1(rng.RangeAddress)
    if ua is None:
        return {"sheet": sheet.Name, "range": requested, "empty": True, "values": []}
    # recorta ao que é usado (permite 'A:C', '1:10', etc.)
    c1 = min(c1, ua[2])
    r1 = min(r1, ua[3])
    if c0 > c1 or r0 > r1:
        return {"sheet": sheet.Name, "range": requested, "empty": True, "values": [],
                "used_range": "A1:%s" % cell_a1(ua[2], ua[3])}
    ncols, nrows = c1 - c0 + 1, r1 - r0 + 1
    limit = ctx.max_cells()
    if ncols * nrows > limit:
        per = max(1, limit // ncols)
        raise ToolError(
            "Intervalo grande demais (%d células; limite %d). Leia em partes, por exemplo %s:%s (%d linhas por vez)."
            % (ncols * nrows, limit, cell_a1(c0, r0), cell_a1(c1, min(r1, r0 + per - 1)), per))
    sub = sheet.getCellRangeByPosition(c0, r0, c1, r1)
    data = sub.getDataArray()
    values = [[(None if v == "" else v) for v in row] for row in data]
    out = {"sheet": sheet.Name,
           "range": "%s:%s" % (cell_a1(c0, r0), cell_a1(c1, r1)) if (c0, r0) != (c1, r1) else cell_a1(c0, r0),
           "rows": nrows, "cols": ncols, "values": values}
    if (c1, r1) != _range_rect(rng)[2:]:
        out["clipped_to_used_area"] = True
    # fórmulas
    formulas = {}
    FORMULA = 16
    try:
        q = sub.queryContentCells(FORMULA)
        for sh, c, r in _each_cell_in(q.getRangeAddresses()):
            cell = sheet.getCellByPosition(c, r)
            formulas[cell_a1(c, r)] = ctx.formula_text(cell)
    except Exception:
        pass
    if formulas:
        out["formulas"] = formulas
    # erros
    errors = {}
    try:
        q = sub.queryFormulaCells(4)  # FormulaResult.ERROR
        for sh, c, r in _each_cell_in(q.getRangeAddresses()):
            cell = sheet.getCellByPosition(c, r)
            errors[cell_a1(c, r)] = cell.getString()
            values[r - r0][c - c0] = cell.getString()
    except Exception:
        pass
    if errors:
        out["errors"] = errors
    if a.get("formatted"):
        out["display"] = [[sub.getCellByPosition(c, r).getString() for c in range(ncols)] for r in range(nrows)]
    return out


@tool("find_cells",
      "Procura texto nas células (valores ou fórmulas). Devolve até 100 ocorrências com endereço, valor e fórmula.",
      {"query": _p("string", "Texto a procurar."),
       "sheet": _p("string", "Planilha onde procurar. Se omitido, procura em todas."),
       "in_formulas": _p("boolean", "Se true, procura dentro do texto das fórmulas (padrão: false = valores)."),
       "match_case": _p("boolean", "Diferenciar maiúsculas de minúsculas (padrão false)."),
       "whole_cell": _p("boolean", "A célula inteira deve ser igual ao texto (padrão false)."),
       "regex": _p("boolean", "Tratar `query` como expressão regular (padrão false).")},
      ["query"])
def find_cells(ctx, a):
    q = a["query"]
    sheets = [ctx.get_sheet(a["sheet"])] if a.get("sheet") else [ctx.doc.Sheets.getByIndex(i) for i in range(ctx.doc.Sheets.Count)]
    hits = []
    truncated = False
    for sh in sheets:
        sd = sh.createSearchDescriptor()
        sd.setSearchString(q)
        sd.SearchCaseSensitive = bool(a.get("match_case"))
        sd.SearchWords = bool(a.get("whole_cell"))
        sd.SearchRegularExpression = bool(a.get("regex"))
        try:
            sd.SearchType = 0 if a.get("in_formulas") else 1
        except Exception:
            pass
        found = sh.findAll(sd)
        if found is None:
            continue
        for _, c, r in _each_cell_in(found.getRangeAddresses()):
            if len(hits) >= 100:
                truncated = True
                break
            cell = sh.getCellByPosition(c, r)
            h = {"cell": "%s!%s" % (quote_sheet(sh.Name), cell_a1(c, r)), "text": cell.getString()}
            if cell.Type.value == "FORMULA":
                h["formula"] = ctx.formula_text(cell)
            hits.append(h)
        if truncated:
            break
    return {"count": len(hits), "truncated": truncated, "matches": hits}


@tool("find_errors",
      "Lista células com erro de fórmula (#DIV/0!, #REF!, #NAME?, Err:xxx…) com a fórmula de cada uma. "
      "Útil para depurar um modelo.",
      {"sheet": _p("string", "Planilha a verificar. Se omitido, verifica todas.")})
def find_errors(ctx, a):
    sheets = [ctx.get_sheet(a["sheet"])] if a.get("sheet") else [ctx.doc.Sheets.getByIndex(i) for i in range(ctx.doc.Sheets.Count)]
    res = []
    truncated = False
    for sh in sheets:
        ua = ctx.used_area(sh)
        if ua is None:
            continue
        whole = sh.getCellRangeByPosition(0, 0, ua[2], ua[3])
        q = whole.queryFormulaCells(4)
        for _, c, r in _each_cell_in(q.getRangeAddresses()):
            if len(res) >= 100:
                truncated = True
                break
            cell = sh.getCellByPosition(c, r)
            res.append({"cell": "%s!%s" % (quote_sheet(sh.Name), cell_a1(c, r)), "error": cell.getString(),
                        "formula": ctx.formula_text(cell)})
        if truncated:
            break
    return {"count": len(res), "truncated": truncated, "errors": res}


@tool("get_format",
      "Mostra a formatação de uma célula: fonte, cores, alinhamento, formato numérico, bordas, mesclagem.",
      {"sheet": SHEET_P, "cell": _p("string", "Célula, ex.: B2.")}, ["cell"])
def get_format(ctx, a):
    sheet, rng = ctx.resolve(a.get("sheet"), a["cell"])
    cell = rng.getCellByPosition(0, 0)
    fmt = ctx.doc.NumberFormats
    try:
        code = fmt.getByKey(cell.NumberFormat).FormatString
    except Exception:
        code = None
    return {
        "cell": range_a1(rng.RangeAddress),
        "font": cell.CharFontName, "size": cell.CharHeight,
        "bold": cell.CharWeight >= 150, "italic": cell.CharPosture.value != "NONE",
        "font_color": "#%06X" % (cell.CharColor & 0xFFFFFF) if cell.CharColor != -1 else "auto",
        "fill_color": None if cell.IsCellBackgroundTransparent else "#%06X" % (cell.CellBackColor & 0xFFFFFF),
        "number_format": code,
        "h_align": cell.HoriJustify.value, "wrap_text": bool(cell.IsTextWrapped),
        "merged": bool(rng.getCellByPosition(0, 0).IsMerged) if hasattr(cell, "IsMerged") else None,
    }


# --------------------------------------------------------------------------
# Escrita
# --------------------------------------------------------------------------

@tool("write_range",
      "Escreve valores e fórmulas a partir de uma célula inicial. `values` é uma matriz (lista de linhas). "
      "Números JSON viram números; texto é gravado como texto literal, EXCETO strings que começam com '=' "
      "que viram fórmulas (sintaxe Excel com funções em inglês: =SUM(A1:A3), =IF(A1>0,\"ok\",\"\"), "
      "=Plan2!B3*2, ='Minha Aba'!A1). Use null para deixar a célula como está e \"\" para esvaziá-la. "
      "A resposta lista `formula_errors` – corrija-os. O usuário pode desfazer com Ctrl+Z.",
      {"sheet": SHEET_P,
       "start_cell": _p("string", "Célula do canto superior esquerdo, ex.: B2."),
       "values": {"type": "array", "description": "Matriz de linhas, ex.: [[\"Item\",\"Preço\"],[\"A\",10]]",
                  "items": {"type": "array", "items": {}}}},
      ["start_cell", "values"], writes=True)
def write_range(ctx, a):
    sheet, rng = ctx.resolve(a.get("sheet"), a["start_cell"])
    c0, r0 = rng.RangeAddress.StartColumn, rng.RangeAddress.StartRow
    values = a["values"]
    if not isinstance(values, list):
        raise ToolError("`values` deve ser uma lista de linhas.")
    if values and not isinstance(values[0], list):
        values = [values]
    n = sum(len(r) for r in values)
    if n > 20000:
        raise ToolError("Muitas células de uma vez (%d). Escreva em blocos de até 20000." % n)
    idx = ctx.sheet_index(sheet)
    written = 0
    formula_cells = []
    with _Undo(ctx, "Claude: escrever dados"):
        for i, row in enumerate(values):
            for j, v in enumerate(row):
                if v is None:
                    continue
                col, r = c0 + j, r0 + i
                if col > 16383 or r > 1048575:
                    raise ToolError("A escrita passa do limite da planilha.")
                cell = sheet.getCellByPosition(col, r)
                if isinstance(v, bool):
                    cell.setValue(1 if v else 0)
                elif isinstance(v, (int, float)):
                    cell.setValue(v)
                elif isinstance(v, str):
                    if v.startswith("=") and len(v) > 1:
                        ctx.set_formula(cell, v, idx, col, r)
                        formula_cells.append((col, r))
                    elif v == "":
                        cell.setFormula("")
                    else:
                        cell.setString(v)
                else:
                    cell.setString(json.dumps(v, ensure_ascii=False))
                written += 1
    last_c = c0 + max((len(r) for r in values), default=1) - 1
    last_r = r0 + len(values) - 1
    out = {"written_cells": written,
           "range": "%s!%s:%s" % (quote_sheet(sheet.Name), cell_a1(c0, r0), cell_a1(last_c, last_r))}
    errs = {}
    for col, r in formula_cells:
        cell = sheet.getCellByPosition(col, r)
        e = cell.getError()
        if e:
            errs[cell_a1(col, r)] = "%s (%s)" % (cell.getString(), ctx.formula_text(cell))
    if errs:
        out["formula_errors"] = errs
    return out


@tool("clear_range", "Limpa conteúdo e/ou formatação de um intervalo.",
      {"sheet": SHEET_P, "range": _p("string", "Intervalo, ex.: A2:D50."),
       "what": _p("string", "O que limpar: 'contents' (valores e fórmulas, padrão), 'formats' ou 'all'.",
                  enum=["contents", "formats", "all"])},
      ["range"], writes=True)
def clear_range(ctx, a):
    sheet, rng = ctx.resolve(a.get("sheet"), a["range"])
    what = a.get("what", "contents")
    VALUE, DATETIME, STRING, ANNOTATION, FORMULA, HARDATTR, STYLES, OBJECTS, EDITATTR = 1, 2, 4, 8, 16, 32, 64, 128, 256
    flags = {"contents": VALUE | DATETIME | STRING | FORMULA,
             "formats": HARDATTR | STYLES | EDITATTR,
             "all": VALUE | DATETIME | STRING | FORMULA | HARDATTR | STYLES | EDITATTR | ANNOTATION}.get(what)
    if flags is None:
        raise ToolError("`what` deve ser contents, formats ou all.")
    with _Undo(ctx, "Claude: limpar"):
        rng.clearContents(flags)
    return {"cleared": range_a1(rng.RangeAddress), "what": what}


_NF_CACHE = {}


def _number_format_key(ctx, code):
    loc = uno.createUnoStruct("com.sun.star.lang.Locale")
    loc.Language, loc.Country, loc.Variant = "en", "US", ""
    nf = ctx.doc.NumberFormats
    key = nf.queryKey(code, loc, False)
    if key == -1:
        try:
            key = nf.addNew(code, loc)
        except Exception as e:  # noqa
            raise ToolError("Formato numérico inválido %r: %s" % (code, e))
    return key


def _border(color, style, width):
    b = uno.createUnoStruct("com.sun.star.table.BorderLine2")
    b.Color = color
    b.LineWidth = width
    b.OuterLineWidth = width
    b.LineStyle = style
    return b


@tool("format_range",
      "Aplica formatação a um intervalo: negrito, itálico, sublinhado, fonte, tamanho, cor da fonte, cor de "
      "preenchimento, alinhamento, quebra de texto, formato numérico (ex.: '#,##0.00', '0.0%', 'dd/mm/yyyy', "
      "'\"R$\" #,##0.00'), bordas e mesclagem. Só as propriedades informadas são alteradas.",
      {"sheet": SHEET_P, "range": _p("string", "Intervalo, ex.: A1:D1."),
       "bold": _p("boolean", "Negrito"), "italic": _p("boolean", "Itálico"),
       "underline": _p("boolean", "Sublinhado"), "strikethrough": _p("boolean", "Tachado"),
       "font_name": _p("string", "Nome da fonte"), "font_size": _p("number", "Tamanho em pontos"),
       "font_color": _p("string", "Cor do texto: #RRGGBB ou nome (red, blue…)"),
       "fill_color": _p("string", "Cor de fundo: #RRGGBB, nome, ou 'none' para remover"),
       "h_align": _p("string", "Alinhamento horizontal", enum=["left", "center", "right", "justify", "default"]),
       "v_align": _p("string", "Alinhamento vertical", enum=["top", "center", "bottom"]),
       "wrap_text": _p("boolean", "Quebrar texto na célula"),
       "number_format": _p("string", "Código de formato numérico estilo Excel/Calc (locale en-US)."),
       "borders": _p("string", "Bordas", enum=["all", "outline", "none", "bottom", "top"]),
       "border_style": _p("string", "Espessura das bordas", enum=["thin", "medium", "thick"]),
       "border_color": _p("string", "Cor das bordas (padrão preto)"),
       "merge": _p("boolean", "true para mesclar as células do intervalo, false para desfazer mesclagem")},
      ["range"], writes=True)
def format_range(ctx, a):
    sheet, rng = ctx.resolve(a.get("sheet"), a["range"])
    applied = []
    with _Undo(ctx, "Claude: formatar"):
        if "bold" in a:
            rng.CharWeight = 150.0 if a["bold"] else 100.0
            applied.append("bold")
        if "italic" in a:
            rng.CharPosture = uno.Enum("com.sun.star.awt.FontSlant", "ITALIC" if a["italic"] else "NONE")
            applied.append("italic")
        if "underline" in a:
            rng.CharUnderline = 1 if a["underline"] else 0
            applied.append("underline")
        if "strikethrough" in a:
            rng.CharStrikeout = 1 if a["strikethrough"] else 0
            applied.append("strikethrough")
        if a.get("font_name"):
            rng.CharFontName = a["font_name"]
            applied.append("font_name")
        if a.get("font_size"):
            rng.CharHeight = float(a["font_size"])
            applied.append("font_size")
        if a.get("font_color"):
            rng.CharColor = parse_color(a["font_color"])
            applied.append("font_color")
        if a.get("fill_color"):
            if str(a["fill_color"]).lower() in ("none", "transparent"):
                rng.IsCellBackgroundTransparent = True
            else:
                rng.CellBackColor = parse_color(a["fill_color"])
                rng.IsCellBackgroundTransparent = False
            applied.append("fill_color")
        if a.get("h_align"):
            m = {"left": "LEFT", "center": "CENTER", "right": "RIGHT", "justify": "BLOCK", "default": "STANDARD"}
            if a["h_align"] not in m:
                raise ToolError("h_align inválido")
            rng.HoriJustify = uno.Enum("com.sun.star.table.CellHoriJustify", m[a["h_align"]])
            applied.append("h_align")
        if a.get("v_align"):
            m = {"top": 1, "center": 2, "bottom": 3}
            if a["v_align"] not in m:
                raise ToolError("v_align inválido")
            rng.CellVertJustify2 = m[a["v_align"]]
            applied.append("v_align")
        if "wrap_text" in a:
            rng.IsTextWrapped = bool(a["wrap_text"])
            applied.append("wrap_text")
        if a.get("number_format"):
            rng.NumberFormat = _number_format_key(ctx, a["number_format"])
            applied.append("number_format")
        if a.get("borders"):
            width = {"thin": 26, "medium": 53, "thick": 88}.get(a.get("border_style", "thin"), 26)
            color = parse_color(a.get("border_color", "black"))
            tb = rng.TableBorder2
            kind = a["borders"]
            if kind == "none":
                nb = _border(0, 0, 0)
                for n in ("Top", "Bottom", "Left", "Right", "Horizontal", "Vertical"):
                    setattr(tb, n + "Line", nb)
                    setattr(tb, "Is%sLineValid" % n, True)
            else:
                bl = _border(color, 0, width)
                if kind in ("all", "outline"):
                    for n in ("Top", "Bottom", "Left", "Right"):
                        setattr(tb, n + "Line", bl)
                        setattr(tb, "Is%sLineValid" % n, True)
                if kind == "all":
                    for n in ("Horizontal", "Vertical"):
                        setattr(tb, n + "Line", bl)
                        setattr(tb, "Is%sLineValid" % n, True)
                if kind == "bottom":
                    tb.BottomLine = bl
                    tb.IsBottomLineValid = True
                if kind == "top":
                    tb.TopLine = bl
                    tb.IsTopLineValid = True
            rng.TableBorder2 = tb
            applied.append("borders")
        if "merge" in a:
            rng.merge(bool(a["merge"]))
            applied.append("merge")
    if not applied:
        raise ToolError("Nenhuma propriedade de formatação foi informada.")
    return {"formatted": "%s!%s" % (quote_sheet(sheet.Name), range_a1(rng.RangeAddress)), "applied": applied}


@tool("copy_range",
      "Copia (ou move) um intervalo para outro lugar, ajustando referências relativas das fórmulas.",
      {"sheet": SHEET_P, "source": _p("string", "Intervalo de origem, ex.: A1:C5 ou 'Plan2!A1:C5'."),
       "destination": _p("string", "Célula de destino (canto superior esquerdo), ex.: F1 ou Plan3!A1."),
       "move": _p("boolean", "true para mover em vez de copiar (padrão false).")},
      ["source", "destination"], writes=True)
def copy_range(ctx, a):
    ssheet, srng = ctx.resolve(a.get("sheet"), a["source"])
    dsn, dref = split_ref(a["destination"])
    dsheet = ctx.get_sheet(dsn or a.get("sheet") or (ssheet.Name if dsn is None else None))
    c, r = parse_cell(dref.split(":")[0])
    dest = uno.createUnoStruct("com.sun.star.table.CellAddress")
    dest.Sheet, dest.Column, dest.Row = ctx.sheet_index(dsheet), c, r
    with _Undo(ctx, "Claude: copiar"):
        if a.get("move"):
            dsheet.moveRange(dest, srng.RangeAddress)
        else:
            dsheet.copyRange(dest, srng.RangeAddress)
    return {"from": "%s!%s" % (quote_sheet(ssheet.Name), range_a1(srng.RangeAddress)),
            "to": "%s!%s" % (quote_sheet(dsheet.Name), cell_a1(c, r)), "moved": bool(a.get("move"))}


@tool("insert_delete",
      "Insere ou exclui linhas/colunas inteiras.",
      {"sheet": SHEET_P,
       "action": _p("string", "insert ou delete", enum=["insert", "delete"]),
       "dimension": _p("string", "rows ou columns", enum=["rows", "columns"]),
       "start": _p("string", "Primeira linha (número, ex.: '5') ou coluna (letra, ex.: 'C')."),
       "count": _p("integer", "Quantidade (padrão 1).")},
      ["action", "dimension", "start"], writes=True)
def insert_delete(ctx, a):
    sheet = ctx.get_sheet(a.get("sheet"))
    count = int(a.get("count", 1))
    if count < 1:
        raise ToolError("count deve ser >= 1")
    start = str(a["start"]).strip()
    if a["dimension"] == "rows":
        if not start.isdigit():
            raise ToolError("Para linhas, `start` é um número (1 = primeira linha).")
        idx, coll = int(start) - 1, sheet.Rows
    else:
        idx = letters_to_col(start) if start.isalpha() else int(start) - 1
        coll = sheet.Columns
    with _Undo(ctx, "Claude: %s %s" % (a["action"], a["dimension"])):
        if a["action"] == "insert":
            coll.insertByIndex(idx, count)
        else:
            coll.removeByIndex(idx, count)
    return {"ok": True, "action": a["action"], "dimension": a["dimension"], "start": start, "count": count}


@tool("resize",
      "Ajusta largura de colunas ou altura de linhas (em milímetros) ou ajusta automaticamente ao conteúdo.",
      {"sheet": SHEET_P,
       "range": _p("string", "Colunas ('A:D') ou linhas ('1:5') ou um intervalo que as contenha."),
       "dimension": _p("string", "columns ou rows", enum=["columns", "rows"]),
       "size_mm": _p("number", "Largura/altura em milímetros."),
       "autofit": _p("boolean", "Ajustar automaticamente ao conteúdo.")},
      ["range", "dimension"], writes=True)
def resize(ctx, a):
    sheet, rng = ctx.resolve(a.get("sheet"), a["range"])
    coll = rng.Columns if a["dimension"] == "columns" else rng.Rows
    with _Undo(ctx, "Claude: redimensionar"):
        if a.get("autofit"):
            if a["dimension"] == "columns":
                coll.OptimalWidth = True
            else:
                coll.OptimalHeight = True
        elif a.get("size_mm"):
            v = int(float(a["size_mm"]) * 100)
            if a["dimension"] == "columns":
                coll.Width = v
            else:
                coll.Height = v
        else:
            raise ToolError("Informe size_mm ou autofit=true.")
    return {"ok": True, "range": range_a1(rng.RangeAddress)}


@tool("sort_range",
      "Ordena um intervalo por uma ou mais colunas.",
      {"sheet": SHEET_P, "range": _p("string", "Intervalo a ordenar, ex.: A1:D100."),
       "keys": {"type": "array", "description": "Chaves de ordenação, em ordem de prioridade.",
                "items": {"type": "object", "properties": {
                    "column": {"type": "string", "description": "Letra da coluna, ex.: 'B'."},
                    "ascending": {"type": "boolean", "description": "true = A→Z / menor→maior (padrão)."}},
                    "required": ["column"]}},
       "has_header": _p("boolean", "A primeira linha é cabeçalho e não deve ser ordenada (padrão true).")},
      ["range", "keys"], writes=True)
def sort_range(ctx, a):
    sheet, rng = ctx.resolve(a.get("sheet"), a["range"])
    c0 = rng.RangeAddress.StartColumn
    fields = []
    for k in a["keys"]:
        f = uno.createUnoStruct("com.sun.star.table.TableSortField")
        col = letters_to_col(k["column"])
        if col < c0 or col > rng.RangeAddress.EndColumn:
            raise ToolError("A coluna %s está fora do intervalo." % k["column"])
        f.Field = col - c0
        f.IsAscending = bool(k.get("ascending", True))
        f.IsCaseSensitive = False
        f.FieldType = uno.Enum("com.sun.star.table.TableSortFieldType", "AUTOMATIC")
        fields.append(f)
    desc = rng.createSortDescriptor()
    pvs = []
    for p in desc:
        if p.Name == "SortFields":
            p.Value = uno.Any("[]com.sun.star.table.TableSortField", tuple(fields))
        elif p.Name == "ContainsHeader":
            p.Value = bool(a.get("has_header", True))
        pvs.append(p)
    with _Undo(ctx, "Claude: ordenar"):
        rng.sort(tuple(pvs))
    return {"sorted": range_a1(rng.RangeAddress), "keys": a["keys"]}


@tool("manage_sheet",
      "Gerencia planilhas (abas): add, rename, delete, duplicate, move, activate, tab_color, hide, show.",
      {"action": _p("string", "Ação", enum=["add", "rename", "delete", "duplicate", "move", "activate",
                                           "tab_color", "hide", "show"]),
       "name": _p("string", "Nome da planilha alvo (ou nome da nova planilha em 'add')."),
       "new_name": _p("string", "Novo nome (rename/duplicate)."),
       "position": _p("integer", "Posição, começando em 1 (add/move/duplicate)."),
       "color": _p("string", "Cor da aba (tab_color).")},
      ["action"], writes=True)
def manage_sheet(ctx, a):
    sheets = ctx.doc.Sheets
    act = a["action"]
    name = a.get("name")
    pos = a.get("position")
    with _Undo(ctx, "Claude: planilha"):
        if act == "add":
            if not name:
                raise ToolError("Informe `name`.")
            if sheets.hasByName(name):
                raise ToolError("Já existe uma planilha chamada %r." % name)
            sheets.insertNewByName(name, (pos - 1) if pos else sheets.Count)
        elif act == "rename":
            sh = ctx.get_sheet(name)
            if not a.get("new_name"):
                raise ToolError("Informe `new_name`.")
            sh.Name = a["new_name"]
        elif act == "delete":
            sh = ctx.get_sheet(name)
            if sheets.Count == 1:
                raise ToolError("Não é possível excluir a única planilha.")
            sheets.removeByName(sh.Name)
        elif act == "duplicate":
            sh = ctx.get_sheet(name)
            new = a.get("new_name") or (sh.Name + " (2)")
            sheets.copyByName(sh.Name, new, (pos - 1) if pos else sheets.Count)
        elif act == "move":
            sh = ctx.get_sheet(name)
            if not pos:
                raise ToolError("Informe `position`.")
            sheets.moveByName(sh.Name, pos - 1)
        elif act == "activate":
            sh = ctx.get_sheet(name)
            ctx.controller.setActiveSheet(sh)
        elif act == "tab_color":
            sh = ctx.get_sheet(name)
            sh.TabColor = parse_color(a.get("color", "white"))
        elif act in ("hide", "show"):
            sh = ctx.get_sheet(name)
            sh.IsVisible = (act == "show")
    return {"ok": True, "action": act, "sheets": list(sheets.ElementNames)}


@tool("freeze_panes", "Congela linhas e/ou colunas no topo/esquerda da planilha (0 para descongelar).",
      {"sheet": SHEET_P, "rows": _p("integer", "Quantidade de linhas a congelar no topo."),
       "columns": _p("integer", "Quantidade de colunas a congelar à esquerda.")}, writes=True)
def freeze_panes(ctx, a):
    sh = ctx.get_sheet(a.get("sheet"))
    ctrl = ctx.controller
    if ctrl is None:
        raise ToolError("Sem janela de documento ativa.")
    ctrl.setActiveSheet(sh)
    ctrl.freezeAtPosition(int(a.get("columns", 0)), int(a.get("rows", 0)))
    return {"ok": True, "rows": int(a.get("rows", 0)), "columns": int(a.get("columns", 0))}


@tool("add_comment", "Adiciona um comentário (nota) a uma célula.",
      {"sheet": SHEET_P, "cell": _p("string", "Célula, ex.: B2."), "text": _p("string", "Texto do comentário.")},
      ["cell", "text"], writes=True)
def add_comment(ctx, a):
    sheet, rng = ctx.resolve(a.get("sheet"), a["cell"])
    ad = rng.getCellByPosition(0, 0).CellAddress
    with _Undo(ctx, "Claude: comentário"):
        sheet.getAnnotations().insertNew(ad, a["text"])
    return {"ok": True, "cell": "%s!%s" % (quote_sheet(sheet.Name), cell_a1(ad.Column, ad.Row))}


@tool("define_name", "Cria (ou substitui) um nome definido que aponta para um intervalo.",
      {"name": _p("string", "Nome (sem espaços), ex.: Receita_Total."),
       "sheet": SHEET_P, "range": _p("string", "Intervalo, ex.: B2:B13.")},
      ["name", "range"], writes=True)
def define_name(ctx, a):
    sheet, rng = ctx.resolve(a.get("sheet"), a["range"])
    nr = ctx.doc.NamedRanges
    ad = uno.createUnoStruct("com.sun.star.table.CellAddress")
    ad.Sheet, ad.Column, ad.Row = ctx.sheet_index(sheet), 0, 0
    with _Undo(ctx, "Claude: nome definido"):
        if nr.hasByName(a["name"]):
            nr.removeByName(a["name"])
        nr.addNewByName(a["name"], rng.AbsoluteName, ad, 0)
    return {"ok": True, "name": a["name"], "refers_to": rng.AbsoluteName}


@tool("select_range",
      "Seleciona e mostra um intervalo na tela do usuário (equivale a 'citar' uma célula). "
      "Use depois de apontar algo importante, por exemplo uma célula com erro.",
      {"sheet": SHEET_P, "range": _p("string", "Intervalo, ex.: B2 ou A1:C3.")}, ["range"])
def select_range(ctx, a):
    sheet, rng = ctx.resolve(a.get("sheet"), a["range"])
    ctrl = ctx.controller
    if ctrl is None:
        raise ToolError("Sem janela de documento ativa.")
    ctrl.setActiveSheet(sheet)
    ctrl.select(rng)
    return {"selected": "%s!%s" % (quote_sheet(sheet.Name), range_a1(rng.RangeAddress))}


# --------------------------------------------------------------------------
# Gráficos e tabelas dinâmicas
# --------------------------------------------------------------------------

_CHART_TYPES = {
    "column": ("com.sun.star.chart.BarDiagram", {"Vertical": False}),
    "bar": ("com.sun.star.chart.BarDiagram", {"Vertical": True}),
    "line": ("com.sun.star.chart.LineDiagram", {}),
    "area": ("com.sun.star.chart.AreaDiagram", {}),
    "pie": ("com.sun.star.chart.PieDiagram", {}),
    "donut": ("com.sun.star.chart.DonutDiagram", {}),
    "scatter": ("com.sun.star.chart.XYDiagram", {}),
}


@tool("create_chart",
      "Cria um gráfico a partir de um intervalo de dados (primeira linha = rótulos das séries, primeira coluna "
      "= rótulos das categorias, por padrão).",
      {"sheet": SHEET_P, "data_range": _p("string", "Intervalo de dados incluindo cabeçalhos, ex.: A1:D13."),
       "chart_type": _p("string", "Tipo", enum=list(_CHART_TYPES)),
       "title": _p("string", "Título do gráfico."),
       "anchor_cell": _p("string", "Célula onde posicionar o canto superior esquerdo (padrão: ao lado dos dados)."),
       "width_mm": _p("number", "Largura em mm (padrão 140)."), "height_mm": _p("number", "Altura em mm (padrão 90)."),
       "first_row_labels": _p("boolean", "Primeira linha contém rótulos (padrão true)."),
       "first_col_labels": _p("boolean", "Primeira coluna contém rótulos (padrão true)."),
       "legend": _p("boolean", "Mostrar legenda (padrão true)."),
       "name": _p("string", "Nome interno do gráfico.")},
      ["data_range", "chart_type"], writes=True)
def create_chart(ctx, a):
    sheet, rng = ctx.resolve(a.get("sheet"), a["data_range"])
    if a["chart_type"] not in _CHART_TYPES:
        raise ToolError("chart_type deve ser um de: %s" % ", ".join(_CHART_TYPES))
    ra = rng.RangeAddress
    if a.get("anchor_cell"):
        c, r = parse_cell(split_ref(a["anchor_cell"])[1])
    else:
        c, r = ra.EndColumn + 2, ra.StartRow
    pos = sheet.getCellByPosition(c, r).Position
    rect = uno.createUnoStruct("com.sun.star.awt.Rectangle")
    rect.X, rect.Y = pos.X, pos.Y
    rect.Width = int(float(a.get("width_mm", 140)) * 100)
    rect.Height = int(float(a.get("height_mm", 90)) * 100)
    charts = sheet.getCharts()
    name = a.get("name") or "Claude_Grafico"
    base, k = name, 1
    while charts.hasByName(name):
        k += 1
        name = "%s_%d" % (base, k)
    with _Undo(ctx, "Claude: gráfico"):
        charts.addNewByName(name, rect, (ra,), bool(a.get("first_row_labels", True)),
                            bool(a.get("first_col_labels", True)))
        emb = charts.getByName(name).getEmbeddedObject()
        svc, props = _CHART_TYPES[a["chart_type"]]
        diagram = emb.createInstance(svc)
        emb.setDiagram(diagram)
        for k2, v in props.items():
            try:
                diagram.setPropertyValue(k2, v)
            except Exception:
                pass
        if a.get("title"):
            emb.HasMainTitle = True
            emb.Title.String = a["title"]
        emb.HasLegend = bool(a.get("legend", True))
    return {"ok": True, "chart": name, "sheet": sheet.Name, "anchored_at": cell_a1(c, r)}


_AGG = {"sum": "SUM", "count": "COUNT", "average": "AVERAGE", "max": "MAX", "min": "MIN",
        "product": "PRODUCT", "count_numbers": "COUNTNUMS", "stdev": "STDEV", "var": "VAR"}


@tool("create_pivot_table",
      "Cria uma tabela dinâmica a partir de um intervalo com cabeçalhos na primeira linha.",
      {"sheet": SHEET_P, "source_range": _p("string", "Dados de origem incluindo cabeçalhos, ex.: A1:E200."),
       "destination": _p("string", "Célula de destino (pode ser 'Resumo!A1')."),
       "row_fields": {"type": "array", "items": {"type": "string"}, "description": "Campos nas linhas (nomes do cabeçalho)."},
       "column_fields": {"type": "array", "items": {"type": "string"}, "description": "Campos nas colunas."},
       "data_fields": {"type": "array", "description": "Campos de valores.",
                       "items": {"type": "object", "properties": {
                           "field": {"type": "string"},
                           "function": {"type": "string", "enum": list(_AGG)}}, "required": ["field"]}},
       "name": _p("string", "Nome da tabela dinâmica.")},
      ["source_range", "destination", "data_fields"], writes=True)
def create_pivot_table(ctx, a):
    ssheet, srng = ctx.resolve(a.get("sheet"), a["source_range"])
    dsn, dref = split_ref(a["destination"])
    dsheet = ctx.get_sheet(dsn or a.get("sheet"))
    c, r = parse_cell(dref.split(":")[0])
    dts = dsheet.getDataPilotTables()
    name = a.get("name") or "Claude_Pivot"
    base, k = name, 1
    while dts.hasByName(name):
        k += 1
        name = "%s_%d" % (base, k)
    desc = dts.createDataPilotDescriptor()
    desc.setSourceRange(srng.RangeAddress)
    fields = desc.getDataPilotFields()
    avail = [fields.getByIndex(i).Name for i in range(fields.Count)] if hasattr(fields, "Count") else []

    def fld(n):
        try:
            return fields.getByName(n)
        except Exception:
            raise ToolError("Campo %r não existe no cabeçalho. Campos: %s" % (n, ", ".join(avail)))

    DPO = "com.sun.star.sheet.DataPilotFieldOrientation"
    with _Undo(ctx, "Claude: tabela dinâmica"):
        for n in a.get("row_fields", []):
            fld(n).Orientation = uno.Enum(DPO, "ROW")
        for n in a.get("column_fields", []):
            fld(n).Orientation = uno.Enum(DPO, "COLUMN")
        for d in a["data_fields"]:
            f = fld(d["field"])
            f.Orientation = uno.Enum(DPO, "DATA")
            f.Function = uno.Enum("com.sun.star.sheet.GeneralFunction", _AGG.get(d.get("function", "sum"), "SUM"))
        dest = uno.createUnoStruct("com.sun.star.table.CellAddress")
        dest.Sheet, dest.Column, dest.Row = ctx.sheet_index(dsheet), c, r
        dts.insertNewByName(name, dest, desc)
    return {"ok": True, "pivot_table": name, "at": "%s!%s" % (quote_sheet(dsheet.Name), cell_a1(c, r))}


# --------------------------------------------------------------------------
# Execução
# --------------------------------------------------------------------------

def list_tools():
    return [{"name": t["name"], "description": t["description"], "inputSchema": t["inputSchema"]}
            for t in TOOLS.values()]


def tool_is_write(name):
    t = TOOLS.get(name)
    return bool(t and t["writes"])


def run_tool(ctx, name, args):
    """Executa uma ferramenta. Devolve (texto_json, is_error)."""
    t = TOOLS.get(name)
    if t is None:
        return json.dumps({"error": "Ferramenta desconhecida: %s" % name}, ensure_ascii=False), True
    try:
        res = t["fn"](ctx, args or {})
        txt = json.dumps(res, ensure_ascii=False, default=str)
        if len(txt) > 120000:
            txt = txt[:120000] + '... [resposta truncada; peça um intervalo menor]'
        return txt, False
    except ToolError as e:
        return json.dumps({"error": str(e)}, ensure_ascii=False), True
    except Exception as e:  # noqa
        msg = getattr(e, "Message", None) or str(e) or e.__class__.__name__
        return json.dumps({"error": "Falha interna (%s): %s" % (e.__class__.__name__, msg)}, ensure_ascii=False), True
