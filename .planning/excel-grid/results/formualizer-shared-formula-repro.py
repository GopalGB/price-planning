#!/usr/bin/env python3
"""Minimal reproduction: Formualizer 0.11.1 corrupts shared-formula followers.

Reported against the pinned engine used by this repository. Run it with the
project interpreter:

    .venv/bin/python .planning/excel-grid/results/formualizer-shared-formula-repro.py

The script builds a tiny in-memory workbook, so it never reads or writes company
data. It prints COUNTS and formula text only for its own synthetic fixture.

WHAT HAPPENS
    Sheet S02 holds a shared family whose master is

        SUMIF(S01!$D$1:$D$3,">=2",S01!$E$1:$E$3)+S01!$A$1

    anchored at A2 and spanning A2:C2. Excel translates only the relative parts
    of a shared formula; the sheet qualifier and the absolute anchors stay put,
    so every follower repeats the master text unchanged.

    Formualizer 0.11.1 runs the same A1 offset over the whole token stream. The
    sheet qualifier "S01" parses as the cell address column S row 1, so it is
    shifted with the follower:

        B2  ->  SUMIF(T1!$D$1:$D$3,">=2",T1!$E$1:$E$3)+T1!$A$1
        C2  ->  SUMIF(U1!$D$1:$D$3,">=2",U1!$E$1:$E$3)+U1!$A$1

    Sheets T1 and U1 do not exist, so every follower evaluates to #REF!.

IMPACT
    A template that shares formulas across a sheet whose name parses as a cell
    address (S01, P1, D2, H1, ...) loads with silently wrong followers. On the
    synthetic fixture at seed 20261011 scale 0.0001 this is 47 of 150 formula
    cells.

WORKAROUND IN THIS REPOSITORY
    pp/engine/native_load.py rewrites the package before load so that every
    shared follower carries its own explicit formula. Native load then needs no
    per-cell formula registration and reports 0 mismatches.

REPORTING
    Exit code 0 means the defect is present, 1 means the engine stopped
    corrupting the qualifier and the workaround in pp/engine/native_load.py can
    be reconsidered.

    Formualizer is a third-party pinned dependency; upstream is
    https://github.com/janverschelde/formualizer (verify before citing).
"""

import io
import sys
import zipfile

import formualizer

CONTENT_TYPES = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package'
    '.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>'
    '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats'
    '-officedocument.spreadsheetml.sheet.main+xml"/>'
    '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd'
    '.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
    '<Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd'
    '.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>')
ROOT_RELS = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006'
    '/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>')
WORKBOOK = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
    '<sheets><sheet name="S01" sheetId="1" r:id="rId1"/>'
    '<sheet name="S02" sheetId="2" r:id="rId2"/></sheets></workbook>')
WORKBOOK_RELS = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006'
    '/relationships/worksheet" Target="worksheets/sheet1.xml"/>'
    '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006'
    '/relationships/worksheet" Target="worksheets/sheet2.xml"/></Relationships>')
# S01 holds the palette the master reads: A1 = 7, D1:D3 = 1, 2, 3, E1:E3 = 100.
SOURCE_SHEET = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>'
    '<row r="1"><c r="A1" t="n"><v>7</v></c><c r="D1" t="n"><v>1</v></c>'
    '<c r="E1" t="n"><v>100</v></c></row>'
    '<row r="2"><c r="D2" t="n"><v>2</v></c><c r="E2" t="n"><v>100</v></c></row>'
    '<row r="3"><c r="D3" t="n"><v>3</v></c><c r="E3" t="n"><v>100</v></c></row>'
    '</sheetData></worksheet>')
MASTER = 'SUMIF(S01!$D$1:$D$3,"&gt;=2",S01!$E$1:$E$3)+S01!$A$1'
TARGET_SHEET = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>'
    '<row r="2">'
    f'<c r="A2" t="n"><f t="shared" ref="A2:C2" si="1">{MASTER}</f><v>307</v></c>'
    '<c r="B2" t="n"><f t="shared" si="1"></f><v>307</v></c>'
    '<c r="C2" t="n"><f t="shared" si="1"/><v>307</v></c>'
    '</row></sheetData></worksheet>')
FOLLOWERS = ((2, 2), (2, 3))


def build():
    payload = io.BytesIO()
    with zipfile.ZipFile(payload, "w", zipfile.ZIP_DEFLATED) as archive:
        archive.writestr("[Content_Types].xml", CONTENT_TYPES)
        archive.writestr("_rels/.rels", ROOT_RELS)
        archive.writestr("xl/workbook.xml", WORKBOOK)
        archive.writestr("xl/_rels/workbook.xml.rels", WORKBOOK_RELS)
        archive.writestr("xl/worksheets/sheet1.xml", SOURCE_SHEET)
        archive.writestr("xl/worksheets/sheet2.xml", TARGET_SHEET)
    return payload.getvalue()


def main():
    import importlib.metadata

    workbook = formualizer.Workbook.from_bytes(build())
    expected = 'SUMIF(S01!$D$1:$D$3,">=2",S01!$E$1:$E$3)+S01!$A$1'
    print(f"formualizer {importlib.metadata.version('formualizer')}")
    print(f"shared follower cells: {len(FOLLOWERS)}")
    print(f"master  A2 = {workbook.get_formula('S02', 2, 1)}")
    corrupt = 0
    for row, column in FOLLOWERS:
        formula = workbook.get_formula("S02", row, column)
        value = workbook.evaluate_cell("S02", row, column)
        address = f"{chr(64 + column)}{row}"
        matches = formula == expected
        corrupt += not matches
        print(f"follower {address} = {formula}")
        print(f"           value = {value}")
        print(f"           equals master text: {matches}")
    print(f"corrupt follower count: {corrupt}/{len(FOLLOWERS)}")
    return 0 if corrupt else 1


if __name__ == "__main__":
    sys.exit(main())