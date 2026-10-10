"""Write input values into a copy of the original .xlsm by surgical, streamed XML edits (SPEC 3.6).

Only the touched worksheet parts, workbook.xml (fullCalcOnLoad), its rels and [Content_Types].xml
(calcChain removed) change. Every other member, including xl/vbaProject.bin, is copied unchanged and in
the original order. Worksheet XML is streamed row by row, so a 650 MB sheet part never sits in memory.
"""
import io
import math
import os
import posixpath
import re
import shutil
import tempfile
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path
from xml.sax.saxutils import escape

CHUNK = 1 << 20
MAX_ROWS, MAX_COLS = 1_048_576, 16_384
MAIN_NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
REL_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
PKG_REL_NS = 'http://schemas.openxmlformats.org/package/2006/relationships'
ILLEGAL_XML = re.compile('[\x00-\x08\x0b\x0c\x0e-\x1f￾￿]')


def col_letters(c):
    s = ''
    c += 1
    while c:
        c, rem = divmod(c - 1, 26)
        s = chr(65 + rem) + s
    return s


def col_index(letters):
    n = 0
    for ch in letters:
        n = n * 26 + ord(ch) - 64
    return n - 1


def _check(cell):
    sheet, r, c, v = cell
    for n in (sheet, r, c):
        if not isinstance(n, int) or isinstance(n, bool) or n < 0:
            raise ValueError(f'bad coordinate in {cell!r}')
    if r >= MAX_ROWS or c >= MAX_COLS:
        raise ValueError(f'cell outside the Excel grid: {cell!r}')
    if isinstance(v, float) and not math.isfinite(v):
        raise ValueError(f'non-finite number in {cell!r}')
    if isinstance(v, str) and len(v) > 32767:
        raise ValueError(f'text longer than the Excel cell limit of 32767 characters at {cell[:3]!r}')
    if isinstance(v, str) and ILLEGAL_XML.search(v):
        raise ValueError(f'text holds a character XML cannot store: {cell!r}')
    if v is not None and not isinstance(v, (bool, int, float, str)):
        raise ValueError(f'unsupported value type {type(v).__name__} in {cell!r}')


def _sheet_parts(zin):
    """Worksheet part name per sheet, in workbook order."""
    wb = ET.fromstring(zin.read('xl/workbook.xml'))
    rels = ET.fromstring(zin.read('xl/_rels/workbook.xml.rels'))
    targets = {r.get('Id'): r.get('Target') for r in rels.iter(f'{{{PKG_REL_NS}}}Relationship')}
    parts = []
    for sheet in wb.iter(f'{{{MAIN_NS}}}sheet'):
        target = targets[sheet.get(f'{{{REL_NS}}}id')]
        parts.append(target.lstrip('/') if target.startswith('/') else posixpath.normpath('xl/' + target))
    return parts


def _cell_xml(p, ref, style, v):
    s = f' s="{style}"' if style is not None else ''
    if v is None:
        return f'<{p}c r="{ref}"{s}/>'
    if isinstance(v, bool):
        return f'<{p}c r="{ref}"{s} t="b"><{p}v>{int(v)}</{p}v></{p}c>'
    if isinstance(v, (int, float)):
        return f'<{p}c r="{ref}"{s}><{p}v>{v!r}</{p}v></{p}c>'
    text = escape(v, {'"': '&quot;'})
    return f'<{p}c r="{ref}"{s} t="inlineStr"><{p}is><{p}t xml:space="preserve">{text}</{p}t></{p}is></{p}c>'


def _attr(tag, name):
    m = re.search(rf'\s{name}="([^"]*)"', tag)
    return m.group(1) if m else None


class _SheetPatcher:
    def __init__(self, src, dst, edits, p, col_styles):
        self.src, self.dst, self.edits, self.p = src, dst, edits, p
        self.col_styles = col_styles
        self.buf, self.pos, self.eof = '', 0, False

    def fill(self):
        if self.eof:
            raise ValueError('worksheet XML ended inside sheetData')
        chunk = self.src.read(CHUNK)
        self.buf = self.buf[self.pos:] + chunk
        self.pos = 0
        self.eof = not chunk

    def new_style(self, row_tag, c):
        if row_tag and _attr(row_tag, 'customFormat') in ('1', 'true'):
            return _attr(row_tag, 's')
        for lo, hi, style in self.col_styles:
            if lo <= c <= hi:
                return style
        return None

    def rewrite_row(self, row_xml, r):
        p, edits = self.p, self.edits[r]
        start = re.match(rf'<{p}row\b[^>]*?(/?)>', row_xml)
        tag = start.group(0)
        inner = '' if start.group(1) else row_xml[start.end():-len(f'</{p}row>')]
        cells, col, pos = {}, -1, 0
        for m in re.finditer(rf'<{p}c(?=[\s/>])[^>]*?/>|<{p}c(?=[\s/>])[^>]*>.*?</{p}c>', inner, re.S):
            cell, pos = m.group(0), m.end()
            ref = _attr(cell[:cell.index('>') + 1], 'r')
            col = col_index(re.match('[A-Z]+', ref).group(0)) if ref else col + 1
            if col in edits:
                ref = f'{col_letters(col)}{r + 1}'
                if re.search(rf'<{p}f[\s/>]', cell):
                    raise ValueError(f'cell {ref} holds a formula and is never written')
                cell = _cell_xml(p, ref, _attr(cell[:cell.index('>') + 1], 's'), edits[col])
            elif not ref:
                cell = f'<{p}c r="{col_letters(col)}{r + 1}"' + cell[len(p) + 2:]
            cells[col] = cell
        for c, v in edits.items():
            if c not in cells and v is not None:
                cells[c] = _cell_xml(p, f'{col_letters(c)}{r + 1}', self.new_style(tag, c), v)
        tag = re.sub(r'\sspans="[^"]*"', '', tag).rstrip('/>').rstrip()
        if _attr(tag, 'r') is None:
            tag += f' r="{r + 1}"'
        return f'{tag}>' + ''.join(cells[c] for c in sorted(cells)) + inner[pos:] + f'</{p}row>'

    def new_row(self, r):
        if all(v is None for v in self.edits[r].values()):
            return ''
        return self.rewrite_row(f'<{self.p}row r="{r + 1}"/>', r)

    def run(self):
        p = self.p
        pending = sorted(self.edits)
        token = re.compile(rf'<{p}row(?=[\s/>])|</{p}sheetData>')
        row_re = re.compile(rf'<{p}row\b[^>]*?/>|<{p}row\b[^>]*>.*?</{p}row>', re.S)
        last = -1
        while True:
            m = token.search(self.buf, self.pos)
            if not m:
                self.fill()
                continue
            self.dst.write(self.buf[self.pos:m.start()])
            self.pos = m.start()
            if m.group(0).startswith('</'):
                break
            row = row_re.match(self.buf, self.pos)
            while not row:
                self.fill()
                row = row_re.match(self.buf, self.pos)
            xml = row.group(0)
            ref = _attr(xml[:xml.index('>') + 1], 'r')
            r = int(ref) - 1 if ref else last + 1
            while pending and pending[0] < r:
                self.dst.write(self.new_row(pending.pop(0)))
            if pending and pending[0] == r:
                pending.pop(0)
                xml = self.rewrite_row(xml, r)
            self.dst.write(xml)
            self.pos, last = row.end(), r
        for r in pending:
            self.dst.write(self.new_row(r))
        self.dst.write(self.buf[self.pos:])
        shutil.copyfileobj(self.src, self.dst, CHUNK)


def _patch_head(head, p, edits):
    """Widen <dimension> to cover the written cells; return (head, column styles)."""
    rows = list(edits)
    cols = [c for row in edits.values() for c in row]
    m = re.search(rf'(<{p}dimension\b[^>]*?\sref=")([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?"', head)
    if m:
        c1, r1 = col_index(m.group(2)), int(m.group(3)) - 1
        c2, r2 = (col_index(m.group(4)), int(m.group(5)) - 1) if m.group(4) else (c1, r1)
        c1, r1 = min(c1, *cols), min(r1, *rows)
        c2, r2 = max(c2, *cols), max(r2, *rows)
        ref = f'{col_letters(c1)}{r1 + 1}:{col_letters(c2)}{r2 + 1}'
        head = head[:m.start()] + f'{m.group(1)}{ref}"' + head[m.end():]
    col_styles = []
    for tag in re.findall(rf'<{p}col\b[^>]*>', head):
        if _attr(tag, 'style') is not None:
            col_styles.append((int(_attr(tag, 'min')) - 1, int(_attr(tag, 'max')) - 1, _attr(tag, 'style')))
    return head, col_styles


def _patch_sheet(src, dst, edits):
    buf = ''
    while True:
        m = re.search(r'<(\w+:)?sheetData\b[^>]*?(/?)>', buf)
        if m:
            break
        chunk = src.read(CHUNK)
        if not chunk:
            raise ValueError('worksheet part has no sheetData')
        buf += chunk
    p = m.group(1) or ''
    head, col_styles = _patch_head(buf[:m.start()], p, edits)
    dst.write(head)
    if m.group(2):  # <sheetData/>: expand it
        dst.write(f'<{p}sheetData>')
        patcher = _SheetPatcher(io.StringIO(f'</{p}sheetData>' + buf[m.end():]), dst, edits, p, col_styles)
        patcher.run()
        shutil.copyfileobj(src, dst, CHUNK)
        return
    dst.write(m.group(0))
    patcher = _SheetPatcher(src, dst, edits, p, col_styles)
    patcher.buf = buf[m.end():]
    patcher.run()


def _patch_workbook(data):
    text = data.decode('utf-8')
    m = re.search(r'<(\w+:)?calcPr\b[^>]*?/?>', text)
    if m:
        tag = m.group(0)
        if re.search(r'\sfullCalcOnLoad="[^"]*"', tag):
            tag = re.sub(r'\sfullCalcOnLoad="[^"]*"', ' fullCalcOnLoad="1"', tag)
        else:
            tag = re.sub(r'\s*(/?)>$', r' fullCalcOnLoad="1"\1>', tag, count=1)
        text = text[:m.start()] + tag + text[m.end():]
    else:
        p = (re.search(r'<(\w+:)?workbook\b', text).group(1) or '')
        anchors = [re.search(rf'</{p}{n}>|<{p}{n}\b[^>]*/>', text)
                   for n in ('sheets', 'functionGroups', 'externalReferences', 'definedNames')]
        at = max(a.end() for a in anchors if a)
        text = text[:at] + f'<{p}calcPr fullCalcOnLoad="1"/>' + text[at:]
    return text.encode('utf-8')


def _drop_calc_chain(name, data):
    if name == '[Content_Types].xml':
        return re.sub(rb'<Override\b[^>]*PartName="/xl/calcChain\.xml"[^>]*/>', b'', data)
    return re.sub(rb'<Relationship\b[^>]*/calcChain"[^>]*/>', b'', data)


def _clone(info):
    out = zipfile.ZipInfo(info.filename, info.date_time)
    out.compress_type, out.external_attr, out.create_system = info.compress_type, info.external_attr, info.create_system
    return out


def export_xlsm(template_path, cells, out_path):
    template_path, out_path = Path(template_path), Path(out_path)
    if out_path.resolve() == template_path.resolve():
        raise ValueError('export must write a copy, never the template itself')
    by_sheet = {}
    for cell in cells:
        _check(tuple(cell))
        sheet, r, c, v = cell
        by_sheet.setdefault(sheet, {}).setdefault(r, {})[c] = v
    written = sum(len(row) for rows in by_sheet.values() for row in rows.values())
    fd, tmp = tempfile.mkstemp(suffix='.xlsm.part', dir=out_path.parent)
    os.close(fd)
    try:
        with zipfile.ZipFile(template_path) as zin, zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED) as zout:
            parts = _sheet_parts(zin)
            bad = [s for s in by_sheet if s >= len(parts)]
            if bad:
                raise ValueError(f'sheet index out of range: {bad}')
            touched = {parts[s]: rows for s, rows in by_sheet.items()}
            for info in zin.infolist():
                name = info.filename
                if name == 'xl/calcChain.xml':
                    continue
                if name in touched:
                    big = info.file_size > (1 << 30)
                    with zin.open(info) as raw, zout.open(_clone(info), 'w', force_zip64=big) as out:
                        # shortcut: assumes UTF-8 sheet XML, which Excel and openpyxl always write
                        src = io.TextIOWrapper(raw, encoding='utf-8', newline='')
                        dst = io.TextIOWrapper(out, encoding='utf-8', newline='')
                        _patch_sheet(src, dst, touched[name])
                        dst.flush()
                        dst.detach()
                elif name == 'xl/workbook.xml':
                    zout.writestr(_clone(info), _patch_workbook(zin.read(info)))
                elif name in ('[Content_Types].xml', 'xl/_rels/workbook.xml.rels'):
                    zout.writestr(_clone(info), _drop_calc_chain(name, zin.read(info)))
                else:
                    with zin.open(info) as src, zout.open(_clone(info), 'w', force_zip64=info.file_size > (1 << 30)) as dst:
                        shutil.copyfileobj(src, dst, CHUNK)
        os.replace(tmp, out_path)
    except BaseException:
        Path(tmp).unlink(missing_ok=True)
        raise
    return {'cells_written': written, 'sheets_touched': sorted(by_sheet)}
