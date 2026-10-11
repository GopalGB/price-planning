"""Create deterministic, streamed synthetic OOXML workbooks for engine scaling."""

import argparse
import json
import math
import os
import zipfile
from collections import Counter
from dataclasses import dataclass
from pathlib import Path
from xml.sax.saxutils import escape


BASE_SHEETS = ((1_700_000, 100_000), (760_000, 650_000), (467_000, 364_000))
REMAINING_CELLS = 2_273_000
REMAINING_FORMULAS = 386_000
ARRAY_FORMULAS = 664_000
ARRAY_SHEET_TWO = 363_000
SHARED_FORMULAS = 680_000
SHEET_COUNT = 22
HIDDEN_SHEET_COUNT = 9
ROW_WIDTH = 50
SHARED_GROUP_SIZE = 32
INPUT_VALUES = {1: 7, 2: 3, 4: 1, 5: 100, 54: 2, 55: 200, 104: 3, 105: 300}
FAMILIES = ("IFERROR", "INDEX_MATCH", "VLOOKUP", "XLOOKUP", "LET", "SWITCH", "SUMIF",
            "NUMBERVALUE", "CONCAT", "COLUMN")
WEIGHTED_FAMILIES = tuple(family for family, weight in zip(FAMILIES, (3, 3, 2, 2, 2, 2, 2, 1, 2, 2))
                          for _ in range(weight))


def _column(number):
    result = ""
    while number:
        number, remainder = divmod(number - 1, 26)
        result = chr(65 + remainder) + result
    return result


def _coordinate(index):
    row, column = divmod(index - 1, ROW_WIDTH)
    return f"{_column(column + 1)}{row + 1}"


def _scaled_split(total, scale, parts):
    quotient, remainder = divmod(round(total * scale), parts)
    return [quotient + (position < remainder) for position in range(parts)]


def _allocate(count, capacities, order):
    allocation = [0] * len(capacities)
    for index in order:
        taken = min(count, capacities[index])
        allocation[index] = taken
        capacities[index] -= taken
        count -= taken
    if count:
        raise ValueError("formula encoding allocation exceeds formula capacity")
    return allocation


def _encoding_plan(formulas, scale):
    capacities = list(formulas)
    normal = _allocate(sum(formulas) - round(ARRAY_FORMULAS * scale) - round(SHARED_FORMULAS * scale),
                       capacities, [0, 2, *range(3, SHEET_COUNT), 1])
    required_second = round(ARRAY_SHEET_TWO * scale)
    if required_second > capacities[1]:
        raise ValueError("sheet two cannot hold its scaled array allocation")
    capacities[1] -= required_second
    array = _allocate(round(ARRAY_FORMULAS * scale) - required_second, capacities,
                      [2, *range(3, SHEET_COUNT), 0, 1])
    array[1] += required_second
    shared = _allocate(round(SHARED_FORMULAS * scale), capacities,
                       [1, 2, *range(3, SHEET_COUNT), 0])
    return {"array": array, "normal": normal, "shared": shared}


def validate_scale(value):
    scale = float(value)
    if not math.isfinite(scale) or not 0 < scale <= 1:
        raise ValueError("scale must be finite and satisfy 0 < scale <= 1")
    cells, formulas = (round(count * scale) for count in BASE_SHEETS[0])
    if cells - formulas + 1 <= max(INPUT_VALUES):
        raise ValueError("scale is too small to keep S01 inputs outside formulas")
    return scale


def build_manifest(scale, seed):
    """Return the exact deterministic allocation without creating a workbook."""
    scale = validate_scale(scale)
    cells = [round(count * scale) for count, _ in BASE_SHEETS]
    formulas = [round(count * scale) for _, count in BASE_SHEETS]
    cells.extend(_scaled_split(REMAINING_CELLS, scale, SHEET_COUNT - len(BASE_SHEETS)))
    formulas.extend(_scaled_split(REMAINING_FORMULAS, scale, SHEET_COUNT - len(BASE_SHEETS)))
    if any(formula > cell for cell, formula in zip(cells, formulas)):
        raise ValueError("scale produces more formulas than cells")
    encoding_plan = _encoding_plan(formulas, scale)
    sheets = []
    for index, (cell_count, formula_count) in enumerate(zip(cells, formulas), 1):
        sheets.append({"name": f"S{index:02d}", "cells": cell_count, "formulas": formula_count,
                       "hidden": index > SHEET_COUNT - HIDDEN_SHEET_COUNT,
                       "formula_encodings": {
                           kind: encoding_plan[kind][index - 1] for kind in encoding_plan
                       }})
    return {
        "seed": seed,
        "scale": scale,
        "sheet_count": SHEET_COUNT,
        "hidden_sheet_count": HIDDEN_SHEET_COUNT,
        "cell_count": sum(cells),
        "formula_count": sum(formulas),
        "formula_encoding_counts": {kind: sum(values) for kind, values in encoding_plan.items()},
        "formula_encoding_plan": encoding_plan,
        "sheets": sheets,
        "ordinary_coordinate": "B1",
        "broad_coordinate": "A1",
        "ordinary_expected_changed_count": 1,
        "broad_expected_changed_count": sum(formulas),
        "function_use_counts": {},
    }


def _family(seed, formula_index):
    return WEIGHTED_FAMILIES[(seed + formula_index * 5) % len(WEIGHTED_FAMILIES)]


def _formula_text(family, column, ordinary):
    suffix = "+S01!$A$1" + ("+S01!$B$1" if ordinary else "")
    formulas = {
        "IFERROR": "IFERROR(1/0,100)",
        "INDEX_MATCH": "INDEX(S01!$E$1:$E$3,MATCH(S01!$D$1,S01!$D$1:$D$3,0))",
        "VLOOKUP": "VLOOKUP(S01!$D$2,S01!$D$1:$E$3,2,FALSE)",
        "XLOOKUP": "_xlfn.XLOOKUP(S01!$D$3,S01!$D$1:$D$3,S01!$E$1:$E$3)",
        "LET": "_xlfn.LET(_xlpm.v,S01!$E$1+S01!$A$1,_xlpm.v-S01!$A$1)",
        "SWITCH": "_xlfn.SWITCH(S01!$D$1,1,100,0)",
        "SUMIF": 'SUMIF(S01!$D$1:$D$3,">=2",S01!$E$1:$E$3)',
        "NUMBERVALUE": '_xlfn.NUMBERVALUE("123.5",".",",")',
        "CONCAT": "LEN(_xlfn.CONCAT(\"S\",COLUMN()))",
        "COLUMN": "COLUMN()",
    }
    return formulas[family] + suffix


def _formula_value(family, column, ordinary):
    return _formula_base(family, column) + 7 + (3 if ordinary else 0)


def _formula_base(family, column):
    return {"IFERROR": 100, "INDEX_MATCH": 100, "VLOOKUP": 200, "XLOOKUP": 300,
            "LET": 100, "SWITCH": 100, "SUMIF": 500, "NUMBERVALUE": 123.5,
            "CONCAT": len("S" + str(column)), "COLUMN": column}[family]


def expected_formula_values(manifest, broad_value, ordinary_value):
    """Yield analytical values in the same coordinate order as generated formulas."""
    formula_offset = 0
    for sheet_index, sheet in enumerate(manifest["sheets"]):
        yield from _sheet_expected_values(
            sheet_index, sheet, (manifest["seed"], formula_offset, broad_value, ordinary_value)
        )
        formula_offset += sheet["formulas"]


def _sheet_expected_values(sheet_index, sheet, state):
    seed, formula_offset, broad_value, ordinary_value = state
    formula_start = sheet["cells"] - sheet["formulas"] + 1
    array_end = formula_start + sheet["formula_encodings"]["normal"]
    array_end += sheet["formula_encodings"]["array"]
    shared_remaining, shared_family = 0, None
    for index in range(formula_start, sheet["cells"] + 1):
        row, col = divmod(index - 1, ROW_WIDTH)
        formula_index = formula_offset + index - formula_start
        ordinary = sheet_index == 0 and index == formula_start
        if index < array_end:
            family = _family(seed, formula_index)
        else:
            if shared_remaining == 0:
                family = shared_family = _family(seed, formula_index)
                row_end = min((row + 1) * ROW_WIDTH, sheet["cells"])
                shared_remaining = min(SHARED_GROUP_SIZE, row_end - index + 1)
            else:
                family = shared_family
            shared_remaining -= 1
        value = _formula_base(family, col + 1) + broad_value
        yield sheet_index, row, col, value + (ordinary_value if ordinary else 0)


def _cell_xml(reference, value, formula_spec=None):
    if formula_spec is None:
        return f'<c r="{reference}" t="n"><v>{value}</v></c>'
    formula, attributes = formula_spec
    return (f'<c r="{reference}" t="n"><f{attributes}>{escape(formula)}</f>'
            f'<v>{value}</v></c>')


@dataclass(frozen=True)
class SheetWriteContext:
    sheet: dict
    categories: dict
    seed: int
    formula_offset: int
    function_uses: Counter

    @property
    def formula_start(self):
        return self.sheet["cells"] - self.sheet["formulas"] + 1


@dataclass
class FormulaWriteState:
    formula_index: int
    shared_remaining: int = 0
    shared_family: str | None = None
    shared_id: int = 0


@dataclass(frozen=True)
class RowWriteContext:
    sheet: SheetWriteContext
    state: FormulaWriteState
    row_end: int


def _write_sheet(target, context):
    target.write(b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>')
    target.write(b'<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>')
    state = FormulaWriteState(context.formula_offset)
    for row_start in range(1, context.sheet["cells"] + 1, ROW_WIDTH):
        row_end = min(row_start + ROW_WIDTH - 1, context.sheet["cells"])
        row = row_start // ROW_WIDTH + 1
        row_context = RowWriteContext(context, state, row_end)
        entries = [_sheet_cell(index, row_context)
                   for index in range(row_start, row_end + 1)]
        target.write(f'<row r="{row}">'.encode())
        target.write("".join(entries).encode())
        target.write(b"</row>")
    target.write(b"</sheetData></worksheet>")


def _sheet_cell(index, row_context):
    context, state = row_context.sheet, row_context.state
    reference = _coordinate(index)
    if index < context.formula_start:
        value = INPUT_VALUES.get(index, (index * 17 + context.seed) % 997)
        return _cell_xml(reference, value)
    column = ((index - 1) % ROW_WIDTH) + 1
    ordinary = context.sheet["name"] == "S01" and index == context.formula_start
    normal_end = context.formula_start + context.categories["normal"]
    array_end = normal_end + context.categories["array"]
    if index < array_end:
        family = _family(context.seed, state.formula_index)
        attributes = f' t="array" ref="{reference}"' if index >= normal_end else ""
        formula = _formula_text(family, column, ordinary)
    else:
        family, formula, attributes = _shared_formula(index, row_context)
    value = _formula_value(family, column, ordinary)
    _count_functions(context.function_uses, _formula_text(family, column, ordinary))
    state.formula_index += 1
    return _cell_xml(reference, value, (formula, attributes))


def _shared_formula(index, row_context):
    context, state, row_end = row_context.sheet, row_context.state, row_context.row_end
    reference = _coordinate(index)
    if state.shared_remaining == 0:
        state.shared_family = _family(context.seed, state.formula_index)
        state.shared_remaining = min(SHARED_GROUP_SIZE, row_end - index + 1)
        state.shared_id += 1
        end = _coordinate(index + state.shared_remaining - 1)
        attributes = f' t="shared" si="{state.shared_id}" ref="{reference}:{end}"'
        formula = _formula_text(state.shared_family, ((index - 1) % ROW_WIDTH) + 1,
                                context.sheet["name"] == "S01" and index == context.formula_start)
    else:
        attributes, formula = f' t="shared" si="{state.shared_id}"', ""
    state.shared_remaining -= 1
    return state.shared_family, formula, attributes


def _count_functions(counter, formula):
    text = formula.upper()
    for function in ("IFERROR", "INDEX", "MATCH", "VLOOKUP", "XLOOKUP", "LET", "SWITCH", "SUMIF",
                     "NUMBERVALUE", "LEN", "CONCAT", "COLUMN"):
        counter[function] += text.count(f"{function}(")


def _zip_info(name):
    info = zipfile.ZipInfo(name, (1980, 1, 1, 0, 0, 0))
    info.compress_type = zipfile.ZIP_DEFLATED
    info.create_system = 0
    info.external_attr = 0
    return info


def _package_xml(manifest):
    sheet_tags = []
    for index, sheet in enumerate(manifest["sheets"], 1):
        state = ' state="hidden"' if sheet["hidden"] else ""
        sheet_tags.append(f'<sheet name="{sheet["name"]}" sheetId="{index}" r:id="rId{index}"{state}/>')
    sheets = "".join(sheet_tags)
    workbook = ("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
                "<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" "
                "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">"
                f"<sheets>{sheets}</sheets><definedNames>"
                "<definedName name=\"BroadInput\">S01!$A$1</definedName>"
                "<definedName name=\"OrdinaryInput\">S01!$B$1</definedName>"
                "<definedName name=\"LookupKeys\">S01!$D$1:$D$3</definedName>"
                "<definedName name=\"LookupValues\">S01!$E$1:$E$3</definedName>"
                "</definedNames></workbook>")
    content_types = ("<?xml version=\"1.0\" encoding=\"UTF-8\"?>"
                     "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">"
                     "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>"
                     "<Default Extension=\"xml\" ContentType=\"application/xml\"/>"
                     "<Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/>"
                     + "".join(f'<Override PartName="/xl/worksheets/sheet{i}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' for i in range(1, SHEET_COUNT + 1)) + "</Types>")
    root_rels = ("<?xml version=\"1.0\" encoding=\"UTF-8\"?>"
                 "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
                 "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"xl/workbook.xml\"/>"
                 "</Relationships>")
    workbook_rels = ("<?xml version=\"1.0\" encoding=\"UTF-8\"?>"
                     "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
                     + "".join(f'<Relationship Id="rId{i}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet{i}.xml"/>' for i in range(1, SHEET_COUNT + 1)) + "</Relationships>")
    return content_types, root_rels, workbook, workbook_rels


def generate(path, scale, seed):
    manifest = build_manifest(scale, seed)
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    temporary = path.with_suffix(path.suffix + ".tmp")
    function_uses = Counter()
    content_types, root_rels, workbook, workbook_rels = _package_xml(manifest)
    with zipfile.ZipFile(temporary, "w", zipfile.ZIP_DEFLATED, compresslevel=6, strict_timestamps=True) as archive:
        for name, content in (("[Content_Types].xml", content_types), ("_rels/.rels", root_rels),
                              ("xl/workbook.xml", workbook), ("xl/_rels/workbook.xml.rels", workbook_rels)):
            archive.writestr(_zip_info(name), content.encode())
        formula_offset = 0
        for number, sheet in enumerate(manifest["sheets"], 1):
            categories = {kind: manifest["formula_encoding_plan"][kind][number - 1]
                          for kind in manifest["formula_encoding_plan"]}
            with archive.open(_zip_info(f"xl/worksheets/sheet{number}.xml"), "w") as target:
                context = SheetWriteContext(sheet, categories, seed, formula_offset, function_uses)
                _write_sheet(target, context)
            formula_offset += sheet["formulas"]
        manifest["function_use_counts"] = dict(sorted(function_uses.items()))
        manifest.pop("formula_encoding_plan")
        archive.writestr(_zip_info("engine-scale-manifest.json"),
                         json.dumps(manifest, sort_keys=True, separators=(",", ":")).encode())
    os.replace(temporary, path)
    return manifest


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("path", type=Path)
    parser.add_argument("--scale", type=validate_scale, required=True)
    parser.add_argument("--seed", type=int, required=True)
    args = parser.parse_args()
    print(json.dumps(generate(args.path, args.scale, args.seed), sort_keys=True))


if __name__ == "__main__":
    main()
