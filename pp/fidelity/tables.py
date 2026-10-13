"""Bounded table-style projection for requested row blocks."""


def table_style_overrides(sheet, query):
    output = {}
    request_end = query.start + query.count - 1
    for table in sheet.get("tables", []):
        row1, col1, row2, col2 = table["range"]
        start, end = max(row1, query.start), min(row2, request_end)
        if start > end or not table.get("style"):
            continue
        for row in range(start, end + 1):
            for col in range(col1, col2 + 1):
                style = _table_cell_style(table, row, col)
                if style:
                    output[(row, col)] = _merge_style(output.get((row, col), {}), style)
    return output


def merge_style_overrides(base, top):
    output = dict(base)
    for coordinate, style in top.items():
        output[coordinate] = _merge_style(output.get(coordinate, {}), style)
    return output


def _merge_style(base, top):
    merged = {**base, **top}
    if "border" in base or "border" in top:
        merged["border"] = {**base.get("border", {}), **top.get("border", {})}
    return merged


def _table_cell_style(table, row, col):
    output = {}
    for name, region in _cell_elements(table, row, col):
        element = table["style"].get(name)
        if element:
            output = _merge_style(output, _cell_dxf(element["style"], row, col, region))
    return output


def _cell_elements(table, row, col):
    row1, col1, row2, col2 = table["range"]
    data1 = row1 + int(table["header_row"])
    data2 = row2 - int(table["totals_row"])
    yield "wholeTable", table["range"]
    if data1 <= row <= data2:
        if table["show_row_stripes"] and (stripe := _stripe(table, row, data1, data2, False)):
            yield stripe
        if table["show_column_stripes"] and (stripe := _stripe(table, col, col1, col2, True)):
            yield stripe
    if table["show_first_column"] and col == col1:
        yield "firstColumn", [row1, col1, row2, col1]
    if table["show_last_column"] and col == col2:
        yield "lastColumn", [row1, col2, row2, col2]
    if table["header_row"] and row == row1:
        yield "headerRow", [row1, col1, row1, col2]
        if col == col1:
            yield "firstHeaderCell", [row1, col1, row1, col1]
        if col == col2:
            yield "lastHeaderCell", [row1, col2, row1, col2]
    if table["totals_row"] and row == row2:
        yield "totalRow", [row2, col1, row2, col2]
        if col == col1:
            yield "firstTotalCell", [row2, col1, row2, col1]
        if col == col2:
            yield "lastTotalCell", [row2, col2, row2, col2]


def _stripe(table, value, start, end, columns):
    prefix = "Column" if columns else "Row"
    first = table["style"].get(f"first{prefix}Stripe")
    second = table["style"].get(f"second{prefix}Stripe")
    first_size = first["size"] if first else 1
    second_size = second["size"] if second else 1
    offset, cycle = value - start, first_size + second_size
    cycle_start = start + offset // cycle * cycle
    if offset % cycle < first_size:
        return f"first{prefix}Stripe", _band_region(table, cycle_start, first_size, end, columns)
    band_start = cycle_start + first_size
    return f"second{prefix}Stripe", _band_region(table, band_start, second_size, end, columns)


def _band_region(table, start, size, end, columns):
    row1, col1, row2, col2 = table["range"]
    if columns:
        return [row1, start, row2, min(end, start + size - 1)]
    return [start, col1, min(end, start + size - 1), col2]


def _cell_dxf(style, row, col, region):
    output = {key: value for key, value in style.items() if key != "border"}
    source = style.get("border", {})
    row1, col1, row2, col2 = region
    edges = {}
    for key, matches in (
        ("t", row == row1), ("r", col == col2),
        ("b", row == row2), ("l", col == col1),
    ):
        if matches and key in source:
            edges[key] = source[key]
    if col < col2 and "vertical" in source:
        edges["r"] = source["vertical"]
    if row < row2 and "horizontal" in source:
        edges["b"] = source["horizontal"]
    if edges:
        output["border"] = edges
    return output
