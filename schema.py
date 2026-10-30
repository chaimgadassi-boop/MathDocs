"""Strict, minimal content contract. No network or AI calls happen here."""
from __future__ import annotations
import json
from pathlib import Path
from typing import Any

ALLOWED_BLOCKS = {"paragraph", "equation", "table", "note"}
ALLOWED_MATH = {"text", "sequence", "fraction", "power", "subscript", "matrix", "limit"}

class DocumentValidationError(ValueError):
    pass

def _require(cond: bool, msg: str) -> None:
    if not cond:
        raise DocumentValidationError(msg)

def validate_math(expr: dict[str, Any], context: str = "equation") -> None:
    _require(isinstance(expr, dict), f"{context}: must be an object")
    t = expr.get("type")
    _require(t in ALLOWED_MATH, f"{context}: unsupported math type {t!r}")
    if t == "text":
        _require(isinstance(expr.get("value"), str), f"{context}: text value must be string")
    elif t == "sequence":
        items = expr.get("items")
        _require(isinstance(items, list) and bool(items), f"{context}: sequence items missing")
        for i, item in enumerate(items):
            validate_math(item, f"{context}.items[{i}]")
    elif t == "fraction":
        for k in ("numerator", "denominator"):
            validate_math(expr.get(k), f"{context}.{k}")
    elif t in ("power", "subscript"):
        for k in ("base", "exponent" if t=="power" else "index"):
            validate_math(expr.get(k), f"{context}.{k}")
    elif t == "matrix":
        rows = expr.get("rows")
        _require(isinstance(rows, list) and len(rows)>0, f"{context}: matrix rows missing")
        _require(all(isinstance(row, list) for row in rows), f"{context}: invalid row")
        n = len(rows[0])
        _require(n>0 and all(len(row)==n for row in rows), f"{context}: non-rectangular matrix")
        _require(len(rows)<=12 and n<=12, f"{context}: max supported matrix size is 12x12")
        for ri,row in enumerate(rows):
            for ci,cell in enumerate(row):
                validate_math(cell, f"{context}.rows[{ri}][{ci}]")
    elif t == "limit":
        _require(isinstance(expr.get("variable"), str), f"{context}: missing limit variable")
        _require(isinstance(expr.get("to"), str), f"{context}: missing limit destination")
        validate_math(expr.get("expression"), f"{context}.expression")

def validate_document(data: dict[str, Any]) -> dict[str, Any]:
    _require(isinstance(data, dict), "Document must be a JSON object")
    _require(data.get("version") == "1", "Unsupported document schema version")
    _require(isinstance(data.get("title"), str) and bool(data["title"].strip()), "Missing title")
    sections = data.get("sections")
    _require(isinstance(sections, list) and bool(sections), "At least one section is required")
    _require(len(sections)<=100, "Too many sections")
    for si, sec in enumerate(sections):
        _require(isinstance(sec, dict) and isinstance(sec.get("heading"),str), f"Section {si}: invalid heading")
        blocks = sec.get("blocks")
        _require(isinstance(blocks, list), f"Section {si}: missing blocks")
        for bi, block in enumerate(blocks):
            c=f"sections[{si}].blocks[{bi}]"
            _require(isinstance(block, dict), f"{c}: must be object")
            kind=block.get("kind")
            _require(kind in ALLOWED_BLOCKS, f"{c}: unknown kind {kind!r}")
            if kind in ("paragraph", "note"):
                _require(isinstance(block.get("text"),str), f"{c}: text required")
            if kind == "equation":
                validate_math(block.get("expr"), c+".expr")
            if kind == "table":
                rows=block.get("rows")
                _require(isinstance(rows,list) and rows and all(isinstance(r,list) for r in rows),f"{c}: table rows missing")
                width=len(rows[0])
                _require(width>0 and all(len(r)==width for r in rows), f"{c}: table not rectangular")
                _require(all(all(isinstance(v,str) for v in r) for r in rows), f"{c}: table cells must be text")
    return data

def read_document(path: Path) -> dict[str, Any]:
    with open(path, encoding="utf-8") as f:
        return validate_document(json.load(f))
