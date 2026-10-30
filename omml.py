"""Build native Office Math (OMML) elements, not screenshots or plain LaTeX."""
from __future__ import annotations
from typing import Any
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

def _m(name: str, *, value: str | None = None):
    element = OxmlElement("m:"+name)
    if value is not None:
        element.set(qn("m:val"), value)
    return element

def _text(value: str):
    run=_m("r")
    text=_m("t")
    text.text=value
    # preserve leading/trailing spaces inside mathematical runs
    text.set("{http://www.w3.org/XML/1998/namespace}space","preserve")
    run.append(text)
    return run

def _wrap_expr(container: Any, expr: dict) -> None:
    container.extend(math_nodes(expr))

def math_nodes(expr: dict) -> list:
    typ=expr["type"]
    if typ=="text":
        return [_text(expr["value"])]
    if typ=="sequence":
        out=[]
        for item in expr["items"]:
            out.extend(math_nodes(item))
        return out
    if typ=="fraction":
        frac=_m("f")
        num=_m("num"); den=_m("den")
        _wrap_expr(num, expr["numerator"]); _wrap_expr(den,expr["denominator"])
        frac.extend([num,den])
        return [frac]
    if typ in ("power","subscript"):
        node=_m("sSup" if typ=="power" else "sSub")
        base=_m("e"); idx=_m("sup" if typ=="power" else "sub")
        _wrap_expr(base,expr["base"])
        _wrap_expr(idx,expr["exponent"] if typ=="power" else expr["index"])
        node.extend([base,idx])
        return [node]
    if typ=="matrix":
        delim=_m("d")
        props=_m("dPr")
        props.extend([_m("begChr",value="("),_m("endChr",value=")")])
        delim.append(props)
        container=_m("e")
        matrix=_m("m")
        rows=expr["rows"]
        matrix_props=_m("mPr")
        matrix_columns=_m("mcs")
        matrix_col=_m("mc")
        matrix_col_props=_m("mcPr")
        matrix_col_props.append(_m("count",value=str(len(rows[0]))))
        matrix_col.append(matrix_col_props); matrix_columns.append(matrix_col)
        matrix_props.append(matrix_columns); matrix.append(matrix_props)
        for row in rows:
            mr=_m("mr")
            for cell in row:
                e=_m("e"); _wrap_expr(e,cell); mr.append(e)
            matrix.append(mr)
        container.append(matrix)
        delim.append(container)
        return [delim]
    if typ=="limit":
        lim_low=_m("limLow")
        base=_m("e"); base.append(_text("lim"))
        lim=_m("lim"); lim.append(_text(expr["variable"]+"→"+expr["to"]))
        lim_low.extend([base,lim])
        return [lim_low,_text(" "),*math_nodes(expr["expression"])]
    raise ValueError(f"Unsupported math: {typ}")

def make_equation(expr:dict):
    para=_m("oMathPara")
    math=_m("oMath")
    math.extend(math_nodes(expr))
    para.append(math)
    return para
