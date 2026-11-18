#!/usr/bin/env python3
"""univer-office-snap -- offline bridge between XLSX files and Univer snapshots.

Univer's open-source repository deliberately keeps XLSX/Office file conversion
out of the core: the README states "Snapshot persistence is separate from XLSX
import and export, which uses the separately distributed file-conversion
capabilities" (dream-num/univer, README.md). This tool fills that gap with a
pure-stdlib, fully offline bridge:

  xlsx2snap  XLSX  -> Univer IWorkbookData snapshot (JSON)
  snap2xlsx  snapshot JSON -> XLSX
  diff       structural cell-level diff between two snapshots / XLSX files
  roundtrip  XLSX -> snapshot -> XLSX, then verify cell data survived

Exit codes: 0 = pass, 1 = drift / problems detected, 2 = usage or input error.

No network access is used anywhere. Only the Python standard library.
"""

from __future__ import annotations

import argparse
import json
import re
import sys
import zipfile
from xml.etree import ElementTree as ET

# ---------------------------------------------------------------------------
# Univer snapshot vocabulary (from dream-num/univer, packages/protocol and
# packages/core/src/sheets/typedef.ts). These numbers are part of the public
# snapshot contract and must not be renumbered.
# ---------------------------------------------------------------------------

CELL_UNKNOWN = 0
CELL_STRING = 1
CELL_NUMBER = 2
CELL_BOOLEAN = 3
CELL_FORCE_STRING = 4

TYPE_NAMES = {
    CELL_UNKNOWN: "unknown",
    CELL_STRING: "string",
    CELL_NUMBER: "number",
    CELL_BOOLEAN: "boolean",
    CELL_FORCE_STRING: "force_string",
}

NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships"
NS_CT = "http://schemas.openxmlformats.org/package/2006/content-types"

COL_RE = re.compile(r"^([A-Z]+)(\d*)$")


# ---------------------------------------------------------------------------
# Small helpers
# ---------------------------------------------------------------------------


class InputError(Exception):
    """Raised for malformed input; maps to exit code 2."""


def _local(tag: str) -> str:
    """Strip an XML namespace from a tag name."""
    return tag.split("}", 1)[1] if "}" in tag else tag


def _col_to_index(ref: str) -> int:
    """'A' -> 0, 'B' -> 1, 'AA' -> 26."""
    m = COL_RE.match(ref or "")
    if not m:
        raise InputError("bad cell reference: %r" % (ref,))
    letters = m.group(1)
    n = 0
    for ch in letters:
        n = n * 26 + (ord(ch) - ord("A") + 1)
    return n - 1


def _index_to_col(idx: int) -> str:
    """0 -> 'A', 25 -> 'Z', 26 -> 'AA'."""
    if idx < 0:
        raise InputError("negative column index: %d" % idx)
    out = ""
    n = idx + 1
    while n > 0:
        n, rem = divmod(n - 1, 26)
        out = chr(ord("A") + rem) + out
    return out


def _cell_ref(row: int, col: int) -> str:
    return "%s%d" % (_index_to_col(col), row + 1)


def _new_cell(value: dict, ctype: int, formula: str = "") -> dict:
    return {
        "v": value,
        "t": ctype,
        "s": "",
        "f": formula,
        "si": "",
        "ref": "",
        "xf": "",
    }


def _new_sheet(name: str, sheet_id: str, row_count=1000, column_count=26) -> dict:
    return {
        "name": name,
        "id": sheet_id,
        "rowCount": row_count,
        "columnCount": column_count,
        "cellData": {},
        "rowData": {},
        "columnData": {},
        "mergeData": [],
        "frozenRow": -1,
        "frozenCol": -1,
        "defaultRowHeight": 24,
        "defaultColumnWidth": 88,
        "showGridlines": 1,
        "zoomRatio": 1,
        "rightToLeft": 0,
        "hidden": 0,
        "tabColor": "",
    }


def _new_workbook(name: str = "UniverWorkbook") -> dict:
    return {
        "id": "univer-workbook",
        "name": name,
        "sheetOrder": [],
        "sheets": {},
        "creator": "",
        "rev": 1,
    }


# ---------------------------------------------------------------------------
# XLSX reading
# ---------------------------------------------------------------------------


def _read_zip_text(zf: zipfile.ZipFile, name: str) -> str:
    try:
        data = zf.read(name)
    except KeyError:
        raise InputError("missing part in xlsx: %s" % name)
    return data.decode("utf-8")


def _parse_shared_strings(zf: zipfile.ZipFile) -> list:
    if "xl/sharedStrings.xml" not in zf.namelist():
        return []
    root = ET.fromstring(_read_zip_text(zf, "xl/sharedStrings.xml"))
    out = []
    for si in root:
        if _local(si.tag) != "si":
            continue
        parts = []
        for node in si.iter():
            if _local(node.tag) == "t" and node.text:
                parts.append(node.text)
        out.append("".join(parts))
    return out


def _sheet_targets(zf: zipfile.ZipFile) -> list:
    """Return [(sheet_name, 'xl/worksheets/sheetN.xml')] in workbook order."""
    wb = ET.fromstring(_read_zip_text(zf, "xl/workbook.xml"))
    rels = {}
    if "xl/_rels/workbook.xml.rels" in zf.namelist():
        rroot = ET.fromstring(_read_zip_text(zf, "xl/_rels/workbook.xml.rels"))
        for rel in rroot:
            if _local(rel.tag) != "Relationship":
                continue
            rels[rel.get("Id")] = rel.get("Target", "")
    sheets = []
    for node in wb.iter():
        if _local(node.tag) != "sheet":
            continue
        name = node.get("name")
        rid = node.get("{%s}id" % NS_REL)
        target = rels.get(rid, "")
        if target.startswith("/"):
            target = target.lstrip("/")
        elif not target.startswith("xl/"):
            target = "xl/" + target.lstrip("./")
        sheets.append((name, target))
    if not sheets:
        raise InputError("xlsx contains no worksheets")
    return sheets


def _cell_value(ctype_attr: str, vtext: str, shared: list, cell_el) -> tuple:
    """Return (value_dict, univer_type)."""
    ctype_attr = (ctype_attr or "n").lower()
    if ctype_attr == "s":
        try:
            idx = int(vtext)
        except (TypeError, ValueError):
            raise InputError("shared-string index is not an integer: %r" % vtext)
        if idx < 0 or idx >= len(shared):
            raise InputError("shared-string index out of range: %d" % idx)
        return {"strV": shared[idx]}, CELL_STRING
    if ctype_attr == "str":
        return {"strV": vtext or ""}, CELL_STRING
    if ctype_attr == "inlinestr":
        parts = []
        for node in cell_el.iter():
            if _local(node.tag) == "t" and node.text:
                parts.append(node.text)
        return {"strV": "".join(parts)}, CELL_STRING
    if ctype_attr == "b":
        return {"boolV": vtext == "1" or vtext == "true"}, CELL_BOOLEAN
    if ctype_attr == "e":
        return {"strV": vtext or ""}, CELL_STRING
    # numeric (default)
    if vtext is None or vtext == "":
        return None, CELL_UNKNOWN
    try:
        num = float(vtext)
    except ValueError:
        # A non-numeric value with no type attribute: keep as string.
        return {"strV": vtext}, CELL_STRING
    if num == int(num) and abs(num) < 1e15:
        return {"numV": int(num)}, CELL_NUMBER
    return {"numV": num}, CELL_NUMBER


def read_xlsx(path: str) -> dict:
    """Parse an .xlsx file into a Univer IWorkbookData snapshot dict."""
    try:
        zf = zipfile.ZipFile(path)
    except (OSError, zipfile.BadZipFile) as exc:
        raise InputError("cannot open xlsx %s: %s" % (path, exc))
    with zf:
        names = set(zf.namelist())
        if "xl/workbook.xml" not in names:
            raise InputError("%s is not a valid xlsx (no xl/workbook.xml)" % path)
        shared = _parse_shared_strings(zf)
        wb = _new_workbook()
        targets = _sheet_targets(zf)
        for order, (name, target) in enumerate(targets, start=1):
            sheet_id = "sheet-%02d" % order
            sheet = _new_sheet(name, sheet_id)
            if target not in names:
                raise InputError("worksheet part missing: %s" % target)
            root = ET.fromstring(_read_zip_text(zf, target))
            for row_el in root.iter():
                if _local(row_el.tag) != "row":
                    continue
                row_idx = int(row_el.get("r", "1")) - 1
                if row_idx < 0:
                    continue
                for c_el in row_el:
                    if _local(c_el.tag) != "c":
                        continue
                    ref = c_el.get("r", "")
                    if ref:
                        try:
                            col_idx = _col_to_index(ref)
                        except InputError:
                            col_idx = None
                    else:
                        col_idx = None
                    if col_idx is None:
                        continue
                    formula = ""
                    vtext = None
                    for child in c_el:
                        tag = _local(child.tag)
                        if tag == "f" and child.text:
                            formula = child.text
                        elif tag == "v":
                            vtext = child.text
                    value, ctype = _cell_value(c_el.get("t"), vtext, shared, c_el)
                    if value is None:
                        continue
                    sheet["cellData"].setdefault(str(row_idx), {})[str(col_idx)] = (
                        _new_cell(value, ctype, formula)
                    )
            wb["sheets"][sheet_id] = sheet
            wb["sheetOrder"].append(sheet_id)
        wb["name"] = targets[0][0]
    return wb


# ---------------------------------------------------------------------------
# XLSX writing
# ---------------------------------------------------------------------------


def _xml_escape(text: str) -> str:
    return (
        text.replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
        .replace('"', "&quot;")
    )


def _cell_xml(row_idx: int, col_idx: int, cell: dict) -> str:
    ref = _cell_ref(row_idx, col_idx)
    ctype = cell.get("t", CELL_UNKNOWN)
    value = cell.get("v") or {}
    formula = cell.get("f") or ""
    if "strV" in value:
        inner = '<is><t xml:space="preserve">%s</t></is>' % _xml_escape(value["strV"])
        return '<c r="%s" t="inlineStr">%s</c>' % (ref, inner)
    if "boolV" in value:
        b = "1" if value["boolV"] else "0"
        body = "<f>%s</f>" % _xml_escape(formula) if formula else ""
        return '<c r="%s" t="b">%s<v>%s</v></c>' % (ref, body, b)
    if "numV" in value:
        body = "<f>%s</f>" % _xml_escape(formula) if formula else ""
        num = value["numV"]
        if isinstance(num, float) and num == int(num):
            num = int(num)
        return '<c r="%s">%s<v>%s</v></c>' % (ref, body, num)
    return '<c r="%s" t="inlineStr"><is><t></t></is></c>' % ref


def _sheet_xml(sheet: dict) -> str:
    rows = []
    for row_str, cols in sorted(sheet.get("cellData", {}).items(), key=lambda kv: int(kv[0])):
        row_idx = int(row_str)
        cells = []
        for col_str in sorted(cols, key=int):
            cells.append(_cell_xml(row_idx, int(col_str), cols[col_str]))
        if cells:
            rows.append('<row r="%d">%s</row>' % (row_idx + 1, "".join(cells)))
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<worksheet xmlns="%s"><sheetData>%s</sheetData></worksheet>'
        % (NS_MAIN, "".join(rows))
    )


def write_xlsx(snapshot: dict, path: str) -> int:
    """Write a Univer snapshot dict to an .xlsx file. Returns sheet count."""
    if not isinstance(snapshot, dict) or "sheets" not in snapshot:
        raise InputError("snapshot has no 'sheets' object")
    order = snapshot.get("sheetOrder") or []
    sheets = snapshot.get("sheets") or {}
    ids = [sid for sid in order if sid in sheets] or list(sheets.keys())
    if not ids:
        raise InputError("snapshot contains no sheets")

    ct_overrides = "".join(
        '<Override PartName="/xl/worksheets/sheet%d.xml" ContentType='
        '"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
        % (i + 1)
        for i in range(len(ids))
    )
    content_types = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Types xmlns="%s">'
        '<Default Extension="rels" ContentType='
        '"application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Override PartName="/xl/workbook.xml" ContentType='
        '"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
        "%s</Types>" % (NS_CT, ct_overrides)
    )
    root_rels = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="%s">'
        '<Relationship Id="rId1" Type='
        '"http://schemas.openxmlformats.org/officeDocument/2006/relationships/'
        'officeDocument" Target="xl/workbook.xml"/></Relationships>' % NS_PKG_REL
    )
    sheet_tags = []
    wb_rels = []
    for i, sid in enumerate(ids, start=1):
        sheet = sheets[sid]
        name = sheet.get("name") or sid
        rid = "rId%d" % i
        sheet_tags.append(
            '<sheet name="%s" sheetId="%d" r:id="%s"/>'
            % (_xml_escape(name), i, rid)
        )
        wb_rels.append(
            '<Relationship Id="%s" Type='
            '"http://schemas.openxmlformats.org/officeDocument/2006/relationships/'
            'worksheet" Target="worksheets/sheet%d.xml"/>' % (rid, i)
        )
    workbook = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<workbook xmlns="%s" xmlns:r="%s"><sheets>%s</sheets></workbook>'
        % (NS_MAIN, NS_REL, "".join(sheet_tags))
    )
    wb_rels_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="%s">%s</Relationships>' % (NS_PKG_REL, "".join(wb_rels))
    )

    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", content_types)
        zf.writestr("_rels/.rels", root_rels)
        zf.writestr("xl/workbook.xml", workbook)
        zf.writestr("xl/_rels/workbook.xml.rels", wb_rels_xml)
        for i, sid in enumerate(ids, start=1):
            zf.writestr("xl/worksheets/sheet%d.xml" % i, _sheet_xml(sheets[sid]))
    return len(ids)


# ---------------------------------------------------------------------------
# Comparison / diff
# ---------------------------------------------------------------------------


def _iter_cells(snapshot: dict):
    for sid, sheet in (snapshot.get("sheets") or {}).items():
        for row_str, cols in (sheet.get("cellData") or {}).items():
            for col_str, cell in cols.items():
                yield sid, sheet, int(row_str), int(col_str), cell


def _values_equal(a: dict, b: dict) -> bool:
    return a == b


def diff_snapshots(a: dict, b: dict) -> dict:
    """Cell-level structural diff between two snapshots."""
    a_ids = list(a.get("sheetOrder") or []) or list((a.get("sheets") or {}).keys())
    b_ids = list(b.get("sheetOrder") or []) or list((b.get("sheets") or {}).keys())
    a_cells = {(sid, r, c): cell for sid, _, r, c, cell in _iter_cells(a)}
    b_cells = {(sid, r, c): cell for sid, _, r, c, cell in _iter_cells(b)}

    added_sheets = [sid for sid in b_ids if sid not in set(a_ids)]
    removed_sheets = [sid for sid in a_ids if sid not in set(b_ids)]
    added_cells = [k for k in b_cells if k not in a_cells]
    removed_cells = [k for k in a_cells if k not in b_cells]
    changed = []
    for key in sorted(set(a_cells) & set(b_cells)):
        if not _values_equal(a_cells[key], b_cells[key]):
            changed.append(
                {"sheet": key[0], "cell": _cell_ref(key[1], key[2]),
                 "before": a_cells[key], "after": b_cells[key]}
            )
    result = {
        "sheets_added": added_sheets,
        "sheets_removed": removed_sheets,
        "cells_added": len(added_cells),
        "cells_removed": len(removed_cells),
        "cells_changed": len(changed),
        "drift": changed[:50],
        "max_drift": len(changed),
    }
    result["drift_count"] = len(changed)
    result["clean"] = (
        not added_sheets and not removed_sheets and not added_cells
        and not removed_cells and not changed
    )
    return result


# ---------------------------------------------------------------------------
# Commands
# ---------------------------------------------------------------------------


def _load_snapshot(path: str) -> dict:
    try:
        with open(path, "r", encoding="utf-8") as fh:
            data = json.load(fh)
    except (OSError, ValueError) as exc:
        raise InputError("cannot read snapshot %s: %s" % (path, exc))
    if not isinstance(data, dict) or "sheets" not in data:
        raise InputError("%s is not a Univer snapshot (no 'sheets')" % path)
    return data


def _load_any(path: str) -> dict:
    low = path.lower()
    if low.endswith(".xlsx"):
        return read_xlsx(path)
    if low.endswith(".json"):
        return _load_snapshot(path)
    raise InputError("unsupported input (use .xlsx or .json): %s" % path)


def cmd_xlsx2snap(args) -> int:
    snap = read_xlsx(args.input)
    payload = {"snapshot": snap}
    if args.json:
        print(json.dumps(payload, ensure_ascii=False, indent=2))
    else:
        print("sheets: %d" % len(snap["sheetOrder"]))
        for sid in snap["sheetOrder"]:
            sh = snap["sheets"][sid]
            n = sum(len(c) for c in sh["cellData"].values())
            print("  %s  %-20s cells=%d" % (sid, sh["name"], n))
        if args.out:
            with open(args.out, "w", encoding="utf-8") as fh:
                json.dump(snap, fh, ensure_ascii=False, indent=2)
            print("written: %s" % args.out)
    return 0


def cmd_snap2xlsx(args) -> int:
    snap = _load_snapshot(args.input)
    n = write_xlsx(snap, args.out)
    payload = {"out": args.out, "sheets_written": n}
    if args.json:
        print(json.dumps(payload, ensure_ascii=False, indent=2))
    else:
        print("wrote %s (%d sheets)" % (args.out, n))
    return 0


def cmd_diff(args) -> int:
    a = _load_any(args.a)
    b = _load_any(args.b)
    result = diff_snapshots(a, b)
    if args.json:
        print(json.dumps(result, ensure_ascii=False, indent=2))
    else:
        print("sheets_added=%d sheets_removed=%d cells_added=%d cells_removed=%d cells_changed=%d"
              % (len(result["sheets_added"]), len(result["sheets_removed"]),
                 result["cells_added"], result["cells_removed"], result["cells_changed"]))
        for d in result["drift"]:
            print("  %s!%s changed" % (d["sheet"], d["cell"]))
    return 0 if result["clean"] else 1


def cmd_roundtrip(args) -> int:
    src = read_xlsx(args.input)
    snap_a = src
    out_path = args.out or "/tmp/univer-office-snap-rt.xlsx"
    write_xlsx(snap_a, out_path)
    snap_b = read_xlsx(out_path)
    result = diff_snapshots(snap_a, snap_b)
    result["input"] = args.input
    result["roundtrip_file"] = out_path
    result["clean"] = result["clean"] and len(result["sheets_added"]) == 0
    if args.json:
        print(json.dumps(result, ensure_ascii=False, indent=2))
    else:
        print("roundtrip %s -> snapshot -> %s" % (args.input, out_path))
        print("sheets_added=%d sheets_removed=%d cells_added=%d cells_removed=%d cells_changed=%d"
              % (len(result["sheets_added"]), len(result["sheets_removed"]),
                 result["cells_added"], result["cells_removed"], result["cells_changed"]))
        if result["clean"]:
            print("OK: all cells preserved")
        else:
            print("DRIFT detected")
    return 0 if result["clean"] else 1


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------


def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        prog="univer-office-snap",
        description="Offline XLSX <-> Univer snapshot bridge (pure stdlib).",
    )
    sub = p.add_subparsers(dest="command", required=True)

    s = sub.add_parser("xlsx2snap", help="convert an XLSX file to a Univer snapshot")
    s.add_argument("input", help="input .xlsx path")
    s.add_argument("--out", default="", help="optional snapshot JSON output path")
    s.add_argument("--json", action="store_true", help="emit machine-readable JSON")
    s.set_defaults(func=cmd_xlsx2snap)

    s = sub.add_parser("snap2xlsx", help="convert a Univer snapshot to XLSX")
    s.add_argument("input", help="input snapshot .json path")
    s.add_argument("--out", required=True, help="output .xlsx path")
    s.add_argument("--json", action="store_true", help="emit machine-readable JSON")
    s.set_defaults(func=cmd_snap2xlsx)

    s = sub.add_parser("diff", help="diff two snapshots or XLSX files")
    s.add_argument("a", help="first .xlsx/.json")
    s.add_argument("b", help="second .xlsx/.json")
    s.add_argument("--json", action="store_true", help="emit machine-readable JSON")
    s.set_defaults(func=cmd_diff)

    s = sub.add_parser("roundtrip", help="XLSX -> snapshot -> XLSX and verify")
    s.add_argument("input", help="input .xlsx path")
    s.add_argument("--out", default="", help="optional output .xlsx path")
    s.add_argument("--json", action="store_true", help="emit machine-readable JSON")
    s.set_defaults(func=cmd_roundtrip)
    return p


def main(argv=None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)
    try:
        return args.func(args)
    except InputError as exc:
        print("error: %s" % exc, file=sys.stderr)
        return 2


if __name__ == "__main__":
    sys.exit(main())
