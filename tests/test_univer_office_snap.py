#!/usr/bin/env python3
"""Tests for univer-office-snap. Run: python3 -m unittest discover -s tests"""

import json
import os
import subprocess
import sys
import tempfile
import unittest
import zipfile

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)

import univer_office_snap as u  # noqa: E402

TOOL = os.path.join(ROOT, "univer_office_snap.py")
SAMPLE = os.path.join(ROOT, "examples", "in", "sample.xlsx")


def run_cli(*args, **kw):
    return subprocess.run(
        [sys.executable, TOOL] + list(args),
        capture_output=True, text=True, cwd=ROOT, **kw
    )


def make_shared_xlsx(path):
    """A minimal XLSX that uses sharedStrings (the t="s" path)."""
    ct = ('<?xml version="1.0" encoding="UTF-8"?>'
          '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
          '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
          '<Default Extension="xml" ContentType="application/xml"/>'
          '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
          '</Types>')
    rels = ('<?xml version="1.0" encoding="UTF-8"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
            '</Relationships>')
    wb = ('<?xml version="1.0" encoding="UTF-8"?>'
          '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
          'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
          '<sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets></workbook>')
    wbr = ('<?xml version="1.0" encoding="UTF-8"?>'
           '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
           '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>'
           '</Relationships>')
    ss = ('<?xml version="1.0" encoding="UTF-8"?>'
          '<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="2" uniqueCount="2">'
          '<si><t>alpha</t></si><si><t>beta</t></si></sst>')
    sh = ('<?xml version="1.0" encoding="UTF-8"?>'
          '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>'
          '<row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1" t="s"><v>1</v></c></row>'
          '</sheetData></worksheet>')
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", ct)
        zf.writestr("_rels/.rels", rels)
        zf.writestr("xl/workbook.xml", wb)
        zf.writestr("xl/_rels/workbook.xml.rels", wbr)
        zf.writestr("xl/sharedStrings.xml", ss)
        zf.writestr("xl/worksheets/sheet1.xml", sh)


class TestReadXlsx(unittest.TestCase):
    def setUp(self):
        self.snap = u.read_xlsx(SAMPLE)

    def test_sheet_order_and_names(self):
        self.assertEqual(self.snap["sheetOrder"], ["sheet-01", "sheet-02"])
        self.assertEqual(self.snap["sheets"]["sheet-01"]["name"], "Sales")
        self.assertEqual(self.snap["sheets"]["sheet-02"]["name"], "Notes")

    def test_string_cell(self):
        cell = self.snap["sheets"]["sheet-01"]["cellData"]["0"]["0"]
        self.assertEqual(cell["v"], {"strV": "Region"})
        self.assertEqual(cell["t"], u.CELL_STRING)

    def test_int_number_cell(self):
        cell = self.snap["sheets"]["sheet-01"]["cellData"]["1"]["1"]
        self.assertEqual(cell["v"], {"numV": 10})
        self.assertEqual(cell["t"], u.CELL_NUMBER)

    def test_float_number_cell(self):
        cell = self.snap["sheets"]["sheet-01"]["cellData"]["1"]["2"]
        self.assertEqual(cell["v"], {"numV": 100.5})

    def test_boolean_cell(self):
        cell = self.snap["sheets"]["sheet-02"]["cellData"]["1"]["1"]
        self.assertEqual(cell["v"], {"boolV": True})
        self.assertEqual(cell["t"], u.CELL_BOOLEAN)

    def test_formula_captured(self):
        cell = self.snap["sheets"]["sheet-01"]["cellData"]["3"]["1"]
        self.assertEqual(cell["f"], "SUM(B2:B3)")

    def test_str_type_result(self):
        cell = self.snap["sheets"]["sheet-02"]["cellData"]["2"]["2"]
        self.assertEqual(cell["v"], {"strV": "cached result"})

    def test_cell_counts(self):
        s1 = self.snap["sheets"]["sheet-01"]
        s2 = self.snap["sheets"]["sheet-02"]
        self.assertEqual(sum(len(c) for c in s1["cellData"].values()), 11)
        self.assertEqual(sum(len(c) for c in s2["cellData"].values()), 3)


class TestSharedStrings(unittest.TestCase):
    def test_shared_string_resolution(self):
        with tempfile.TemporaryDirectory() as td:
            path = os.path.join(td, "shared.xlsx")
            make_shared_xlsx(path)
            snap = u.read_xlsx(path)
            cells = snap["sheets"]["sheet-01"]["cellData"]["0"]
            self.assertEqual(cells["0"]["v"], {"strV": "alpha"})
            self.assertEqual(cells["1"]["v"], {"strV": "beta"})


class TestColumnHelpers(unittest.TestCase):
    def test_col_index_roundtrip(self):
        for idx in (0, 1, 25, 26, 51, 701, 702):
            self.assertEqual(u._col_to_index(u._index_to_col(idx)), idx)

    def test_cell_ref(self):
        self.assertEqual(u._cell_ref(0, 0), "A1")
        self.assertEqual(u._cell_ref(9, 27), "AB10")


class TestWriteAndRoundtrip(unittest.TestCase):
    def test_write_read_roundtrip(self):
        snap = u.read_xlsx(SAMPLE)
        with tempfile.TemporaryDirectory() as td:
            out = os.path.join(td, "rt.xlsx")
            n = u.write_xlsx(snap, out)
            self.assertEqual(n, 2)
            back = u.read_xlsx(out)
        d = u.diff_snapshots(snap, back)
        self.assertTrue(d["clean"], d)

    def test_diff_detects_change(self):
        a = u.read_xlsx(SAMPLE)
        b = json.loads(json.dumps(a))
        b["sheets"]["sheet-01"]["cellData"]["1"]["1"]["v"] = {"numV": 999}
        d = u.diff_snapshots(a, b)
        self.assertFalse(d["clean"])
        self.assertEqual(d["cells_changed"], 1)
        self.assertEqual(d["drift"][0]["cell"], "B2")

    def test_diff_detects_added_sheet(self):
        a = u.read_xlsx(SAMPLE)
        b = json.loads(json.dumps(a))
        b["sheets"]["sheet-99"] = u._new_sheet("Extra", "sheet-99")
        b["sheetOrder"].append("sheet-99")
        d = u.diff_snapshots(a, b)
        self.assertIn("sheet-99", d["sheets_added"])


class TestCli(unittest.TestCase):
    def test_json_output_parseable(self):
        r = run_cli("xlsx2snap", SAMPLE, "--json")
        self.assertEqual(r.returncode, 0, r.stderr)
        data = json.loads(r.stdout)
        self.assertIn("snapshot", data)
        self.assertEqual(data["snapshot"]["sheetOrder"], ["sheet-01", "sheet-02"])

    def test_roundtrip_cli_rc0(self):
        r = run_cli("roundtrip", SAMPLE, "--json")
        self.assertEqual(r.returncode, 0, r.stderr)
        self.assertTrue(json.loads(r.stdout)["clean"])

    def test_diff_cli_rc1_on_change(self):
        with tempfile.TemporaryDirectory() as td:
            a = os.path.join(td, "a.xlsx")
            b = os.path.join(td, "b.xlsx")
            snap = u.read_xlsx(SAMPLE)
            u.write_xlsx(snap, a)
            snap2 = json.loads(json.dumps(snap))
            snap2["sheets"]["sheet-01"]["cellData"]["0"]["0"]["v"] = {"strV": "Changed"}
            u.write_xlsx(snap2, b)
            r = run_cli("diff", a, b, "--json")
        self.assertEqual(r.returncode, 1, r.stderr)
        self.assertEqual(json.loads(r.stdout)["cells_changed"], 1)

    def test_diff_cli_rc0_identical(self):
        r = run_cli("diff", SAMPLE, SAMPLE, "--json")
        self.assertEqual(r.returncode, 0, r.stderr)
        self.assertTrue(json.loads(r.stdout)["clean"])

    def test_missing_file_rc2(self):
        r = run_cli("xlsx2snap", "/tmp/definitely-not-here-12345.xlsx", "--json")
        self.assertEqual(r.returncode, 2)

    def test_bad_snapshot_rc2(self):
        with tempfile.TemporaryDirectory() as td:
            bad = os.path.join(td, "bad.json")
            with open(bad, "w") as fh:
                fh.write("{not json")
            r = run_cli("snap2xlsx", bad, "--out", os.path.join(td, "x.xlsx"))
        self.assertEqual(r.returncode, 2)


if __name__ == "__main__":
    unittest.main()