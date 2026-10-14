"""Graft a customer's extra sheets and extra columns onto a template package.

Output = the template package + the upload's extra sheets (appended) + the upload's cells at or right
of a given column on named template sheets. Parts the graft does not touch are copied byte-for-byte.
Worksheet rows stream one at a time; the head and tail of each worksheet (outside sheetData) are small
and held in memory. Spooled sheet parts go to temp files, then the package is written in one pass.
"""
import io
import os
import re
import shutil
import tempfile
import zipfile
from dataclasses import dataclass
from pathlib import Path
from xml.sax.saxutils import escape, unescape

import defusedxml.ElementTree as SafeET

from .xlsm import (CHUNK, MAIN_NS, REL_NS, _attr, _clone, _drop_calc_chain, _patch_workbook,
                   _sheet_parts, col_index, col_letters)

WS_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet'
WS_CTYPE = 'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml'
REL_ELEMENTS = 'drawing|legacyDrawing|legacyDrawingHF|tableParts|picture|oleObjects|controls'
ITEMS = (('fonts', 'font'), ('fills', 'fill'), ('borders', 'border'), ('cellXfs', 'xf'), ('dxfs', 'dxf'))
STYLE_ORDER = ('numFmts', 'fonts', 'fills', 'borders', 'cellStyleXfs', 'cellXfs', 'cellStyles', 'dxfs',
               'tableStyles', 'colors', 'extLst')
BOOK_ORDER = ('sheets', 'functionGroups', 'externalReferences', 'definedNames', 'calcPr', 'oleSize',
              'customWorkbookViews', 'pivotCaches', 'extLst')
CELL_RE = re.compile(r'<(?:\w+:)?c(?=[\s/>])[^>]*?/>|<(?P<q>(?:\w+:)?)c(?=[\s/>])[^>]*>.*?</(?P=q)c>', re.S)
F_RE = re.compile(r'<(?:\w+:)?f\b[^>]*?/>|<(?P<q>(?:\w+:)?)f\b[^>]*>.*?</(?P=q)f>', re.S)
ENTITIES = {'&quot;': '"', '&apos;': "'"}


@dataclass(frozen=True)
class GraftPlan:
    sheets: tuple  # upload sheet names appended as new sheets, in this order
    columns: dict  # template sheet name -> first extra column (0-based)


def _int(tag, name):
    return int(_attr(tag, name) or 0)


def _attrs(el, **kv):
    """Set attributes on an element's start tag; the rest of the element is left as it is."""
    end = el.index('>') + 1
    tag = el[:end]
    for name, value in kv.items():
        if re.search(rf'\s{name}="', tag):
            tag = re.sub(rf'\s{name}="[^"]*"', f' {name}="{value}"', tag, count=1)
        elif tag.endswith('/>'):
            tag = f'{tag[:-2]} {name}="{value}"/>'
        else:
            tag = f'{tag[:-1]} {name}="{value}">'
    return tag + el[end:]


def _inner(xml, name):
    m = re.search(rf'<(?P<q>(?:\w+:)?){name}\b[^>]*?(?:/>|>(?P<body>.*?)</(?P=q){name}>)', xml, re.S)
    return (m.group('body') or '') if m else ''


def _elements(inner, name):
    pat = re.compile(rf'<(?:\w+:)?{name}\b[^>]*?/>|<(?P<q>(?:\w+:)?){name}\b[^>]*>.*?</(?P=q){name}>', re.S)
    return [m.group(0) for m in pat.finditer(inner)]


def _read(zf, name):
    return zf.read(name).decode('utf-8') if name in zf.namelist() else None


def _text(zf, name):
    return io.TextIOWrapper(zf.open(name), encoding='utf-8', newline='')


def _sheets(zf):
    """(name, worksheet part) per sheet, in workbook order."""
    wb = SafeET.fromstring(zf.read('xl/workbook.xml'))
    names = [s.get('name') for s in wb.iter(f'{{{MAIN_NS}}}sheet')]
    return list(zip(names, _sheet_parts(zf)))


def _shared_strings(zf):
    text = _read(zf, 'xl/sharedStrings.xml')
    if text is None:
        return []
    return [si[si.index('>') + 1:si.rindex('<')] for si in _elements(_inner(text, 'sst'), 'si')]


class _Sheet:
    """One worksheet part: head text, then <row> elements one at a time, then the tail text."""

    def __init__(self, src):
        self.src = src
        self.buf, self.pos, self.eof, self.last = '', 0, False, -1
        while True:
            m = re.search(r'<((?:\w+:)?)sheetData\b[^>]*?(/?)>', self.buf)
            if m:
                break
            self._need()
        self.p, self.empty = m.group(1), m.group(2) == '/'
        self.head, self.pos = self.buf[:m.start()], m.end()

    def _fill(self):
        chunk = self.src.read(CHUNK)
        self.buf += chunk
        self.eof = not chunk

    def _need(self):
        if self.eof:
            raise ValueError('worksheet XML ended before sheetData closed')
        self._fill()

    def _row_end(self, start):
        gt = self.buf.find('>', start)
        if gt < 0:
            return -1
        if self.buf[gt - 1] == '/':
            return gt + 1
        close = f'</{self.p}row>'
        end = self.buf.find(close, gt)
        return -1 if end < 0 else end + len(close)

    def rows(self):
        """Yield (0-based row number, row XML) until </sheetData>."""
        if self.empty:
            return
        token = re.compile(rf'<{self.p}row(?=[\s/>])|</{self.p}sheetData>')
        while True:
            if self.pos > CHUNK:
                self.buf, self.pos = self.buf[self.pos:], 0
            m = token.search(self.buf, self.pos)
            if m is None:
                self._need()
                continue
            if m.group().startswith('</'):
                self.pos = m.end()
                return
            end = self._row_end(m.start())
            if end < 0:
                self._need()
                continue
            xml = self.buf[m.start():end]
            self.pos = end
            ref = _attr(xml[:xml.index('>') + 1], 'r')
            self.last = int(ref) - 1 if ref else self.last + 1
            yield self.last, xml

    def rest(self):
        """Text after </sheetData>. Held in memory: sheet tails are small next to sheetData."""
        while not self.eof:
            self._fill()
        return self.buf[self.pos:]


def _tags(src, name):
    """Start tags of every <name> element in a text stream, read in chunks."""
    pat = re.compile(rf'<(?:\w+:)?{name}\b[^>]*>')
    carry = ''
    while True:
        chunk = src.read(CHUNK)
        buf = carry + chunk
        yield from pat.findall(buf)
        cut = buf.rfind('<')
        carry = buf[cut:] if cut >= 0 and '>' not in buf[cut:] else ''
        if not chunk:
            return


def _max_shared_si(src):
    best = -1
    for tag in _tags(src, 'f'):
        if _attr(tag, 't') == 'shared' and _attr(tag, 'si') is not None:
            best = max(best, int(_attr(tag, 'si')))
    return best


def _extent(src, first):
    """(c0, r0, c1, r1), 0-based, over upload cells at column >= first; None when there are none."""
    box = None
    for tag in _tags(src, 'c'):
        m = re.match(r'([A-Z]+)(\d+)$', _attr(tag, 'r') or '')
        if not m or col_index(m.group(1)) < first:
            continue
        c, r = col_index(m.group(1)), int(m.group(2)) - 1
        box = (c, r, c, r) if box is None else (min(box[0], c), min(box[1], r), max(box[2], c), max(box[3], r))
    return box


def _extend(text, p, name, item, items, order, counted=False):
    """Append item elements to container <name>, creating the container in schema order when absent."""
    if not items:
        return text
    m = re.search(rf'<{p}{name}\b[^>]*?(/?)>', text)
    body = ''.join(items)
    if m is None:
        cut = text.rindex('</')
        for succ in order[order.index(name) + 1:]:
            hit = re.search(rf'<{p}{succ}\b', text)
            if hit and hit.start() < cut:
                cut = hit.start()
        text = text[:cut] + f'<{p}{name}>{body}</{p}{name}>' + text[cut:]
    elif m.group(1):
        text = text[:m.start()] + m.group(0)[:-2] + '>' + body + f'</{p}{name}>' + text[m.end():]
    else:
        close = text.index(f'</{p}{name}>', m.end())
        text = text[:close] + body + text[close:]
    if counted:
        start = re.search(rf'<{p}{name}\b[^>]*>', text)
        count = len(_elements(_inner(text, name), item))
        text = text[:start.start()] + _attrs(start.group(0), count=count) + text[start.end():]
    return text


class _Styles:
    """Upload style table merged into the template's. Upload indexes map to new template indexes on use."""

    def __init__(self, tmpl, up):
        if tmpl is None:
            raise ValueError('template package has no styles part')
        self.tmpl = tmpl
        self.p = re.search(r'<((?:\w+:)?)styleSheet\b', tmpl).group(1)
        up = up or ''
        self.up = {k: _elements(_inner(up, k), item) for k, item in ITEMS}
        self.up_fmts = {int(_attr(t, 'numFmtId')): t for t in _elements(_inner(up, 'numFmts'), 'numFmt')}
        self.fmt_ids = {_attr(t, 'formatCode'): int(_attr(t, 'numFmtId'))
                        for t in _elements(_inner(tmpl, 'numFmts'), 'numFmt')}
        self.next_fmt = max([163] + list(self.fmt_ids.values())) + 1
        self.base = {k: len(_elements(_inner(tmpl, k), item)) for k, item in ITEMS}
        self.new = {k: [] for k, _ in ITEMS}
        self.new_fmts = []
        self.seen = {}

    def _add(self, kind, el):
        key = (kind, el)
        if key not in self.seen:
            self.seen[key] = self.base[kind] + len(self.new[kind])
            self.new[kind].append(el)
        return self.seen[key]

    def _fmt(self, fid):
        src = self.up_fmts.get(fid)
        if src is None:
            raise ValueError(f'upload cell uses undefined number format {fid}')
        code = _attr(src, 'formatCode')
        if code not in self.fmt_ids:
            self.fmt_ids[code] = self.next_fmt
            self.new_fmts.append(_attrs(src, numFmtId=self.next_fmt))
            self.next_fmt += 1
        return self.fmt_ids[code]

    def xf(self, s):
        """Template cellXfs index for upload index s; None when the upload has no style table."""
        if not self.up['cellXfs']:
            return None
        key = ('xf', s)
        if key not in self.seen:
            el = self.up['cellXfs'][s]
            ids = {'fontId': self._add('fonts', self.up['fonts'][_int(el, 'fontId')]),
                   'fillId': self._add('fills', self.up['fills'][_int(el, 'fillId')]),
                   'borderId': self._add('borders', self.up['borders'][_int(el, 'borderId')]),
                   'xfId': 0}
            if _int(el, 'numFmtId') >= 164:
                ids['numFmtId'] = self._fmt(_int(el, 'numFmtId'))
            self.seen[key] = self._add('cellXfs', _attrs(el, **ids))
        return self.seen[key]

    def dxf(self, d):
        key = ('dxf', d)
        if key not in self.seen:
            self.seen[key] = self._add('dxfs', self.up['dxfs'][d])
        return self.seen[key]

    def text(self):
        out = _extend(self.tmpl, self.p, 'numFmts', 'numFmt', self.new_fmts, STYLE_ORDER, counted=True)
        for kind, item in ITEMS:
            out = _extend(out, self.p, kind, item, self.new[kind], STYLE_ORDER, counted=True)
        return out


class _Copier:
    """Turns upload cells into cells for the output: style remap, inline strings, shared-formula care."""

    def __init__(self, p, styles, strings, shared_offset):
        self.p, self.styles, self.strings, self.offset = p, styles, strings, shared_offset
        self.masters, self.cells, self.dropped = set(), 0, 0

    def cell(self, col, row, raw):
        p = self.p
        open_end = raw.index('>') + 1
        tag = raw[:open_end]
        body = '' if tag.endswith('/>') else raw[open_end:raw.rindex('</')]
        attrs = ''
        style = self.styles.xf(_int(tag, 's'))
        if style is not None:
            attrs += f' s="{style}"'
        kind = _attr(tag, 't')
        if kind == 's':
            idx = int(re.search(r'<(?:\w+:)?v>(\d+)</', body).group(1))
            body = f'<{p}is>{self.strings[idx]}</{p}is>'
            kind = 'inlineStr'
        if kind:
            attrs += f' t="{kind}"'
        body = self._formula(body)
        self.cells += 1
        ref = f'{col_letters(col)}{row + 1}'
        return f'<{p}c r="{ref}"{attrs}>{body}</{p}c>' if body else f'<{p}c r="{ref}"{attrs}/>'

    def _formula(self, body):
        m = F_RE.search(body)
        if m is None or self.offset is None or _attr(m.group(0)[:m.group(0).index('>') + 1], 't') != 'shared':
            return body
        f = m.group(0)
        start = f[:f.index('>') + 1]
        si = int(_attr(start, 'si'))
        if _attr(start, 'ref') is not None:
            self.masters.add(si)
        elif si not in self.masters:
            self.dropped += 1
            return body[:m.start()] + body[m.end():]
        return body[:m.start()] + _attrs(f, si=si + self.offset) + body[m.end():]


def _cells(xml):
    """(0-based column, raw <c> element) for each cell of one <row> element."""
    out, prev = [], -1
    for m in CELL_RE.finditer(xml):
        raw = m.group(0)
        ref = _attr(raw[:raw.index('>') + 1], 'r')
        prev = col_index(re.match(r'[A-Z]+', ref).group(0)) if ref else prev + 1
        out.append((prev, raw))
    return out


def _row_start(tag, styles=None):
    tag = re.sub(r'\sspans="[^"]*"', '', tag)
    if styles is not None and _attr(tag, 's') is not None:
        style = styles.xf(int(_attr(tag, 's')))
        if style is not None:
            tag = _attrs(tag, s=style)
    return tag


def _row_xml(start, cells, p):
    if start.endswith('/>'):
        if not cells:
            return start
        start = start[:-2] + '>'
    return start + ''.join(cells) + f'</{p}row>'


def _merged(tmpl_xml, cells, p):
    merged = dict(_cells(tmpl_xml))
    merged.update(cells)
    return _row_xml(_row_start(tmpl_xml[:tmpl_xml.index('>') + 1]), [merged[c] for c in sorted(merged)], p)


def _copied_rows(rows, copier, first):
    """Upload rows reduced to their cells at column >= first, as (row, [(col, xml)])."""
    for r, xml in rows:
        cells = [(c, copier.cell(c, r, raw)) for c, raw in _cells(xml) if c >= first]
        if cells:
            yield r, cells


def _merge_rows(tmpl, copied, p):
    """Merge-join template rows with the upload's copied cells; both streams ascend by row."""
    t, u = next(tmpl, None), next(copied, None)
    while t is not None or u is not None:
        if u is None or (t is not None and t[0] < u[0]):
            yield t[1]
            t = next(tmpl, None)
        elif t is None or u[0] < t[0]:
            yield _row_xml(f'<{p}row r="{u[0] + 1}">', [x for _, x in u[1]], p)
            u = next(copied, None)
        else:
            yield _merged(t[1], u[1], p)
            t, u = next(tmpl, None), next(copied, None)


def _box(ref):
    pts = [(col_index(c), int(r) - 1) for c, r in re.findall(r'([A-Z]+)(\d+)', ref)]
    return (min(x for x, _ in pts), min(y for _, y in pts), max(x for x, _ in pts), max(y for _, y in pts))


def _widen(head, p, ext):
    m = re.search(rf'<{p}dimension\b[^>]*?\sref="([^"]+)"', head)
    if m is None:
        return head
    c0, r0, c1, r1 = _box(m.group(1))
    ec0, er0, ec1, er1 = ext
    ref = f'{col_letters(min(c0, ec0))}{min(r0, er0) + 1}:{col_letters(max(c1, ec1))}{max(r1, er1) + 1}'
    return head[:m.start(1)] + ref + head[m.end(1):]


def _minus(span, cut):
    a, b = span
    c, d = cut
    if d < a or c > b:
        return [span]
    keep = []
    if a < c:
        keep.append((a, c - 1))
    if d < b:
        keep.append((d + 1, b))
    return keep


def _split_cols(old, extra):
    """Template <col> ranges minus the extra ranges, plus the extra ranges (1-based, inclusive)."""
    out = []
    for lo, hi, tag in old:
        pieces = [(lo, hi)]
        for elo, ehi, _ in extra:
            pieces = [q for span in pieces for q in _minus(span, (elo, ehi))]
        out += [(a, b, _attrs(tag, min=a, max=b)) for a, b in pieces]
    return sorted(out + list(extra), key=lambda item: item[0])


def _style_tag(tag, styles):
    s = _attr(tag, 'style')
    return tag if s is None else _attrs(tag, style=styles.xf(int(s)))


def _merge_cols(head, p, extra):
    if not extra:
        return head
    m = re.search(rf'<{p}cols\b[^>]*?/>|<{p}cols\b[^>]*>.*?</{p}cols>', head, re.S)
    old = []
    if m:
        old = [(int(_attr(t, 'min')), int(_attr(t, 'max')), t)
               for t in re.findall(rf'<{p}col\b[^>]*/>', m.group(0))]
    block = f'<{p}cols>' + ''.join(t for _, _, t in _split_cols(old, extra)) + f'</{p}cols>'
    if m:
        return head[:m.start()] + block + head[m.end():]
    s = re.search(rf'<{p}sheetFormatPr\b[^>]*?/>', head)
    at = s.end() if s else len(head)
    return head[:at] + block + head[at:]


def _extra_cols(usheet, styles, first):
    out = []
    for t in re.findall(rf'<{usheet.p}col\b[^>]*/>', usheet.head):
        lo, hi = max(int(_attr(t, 'min')), first + 1), int(_attr(t, 'max'))
        if lo > hi:
            continue
        t = _style_tag(t, styles)
        out.append((lo, hi, _attrs(t, min=lo, max=hi)))
    return out


def _strip_rels(text, rel):
    """Drop what needs a relationship part: those elements, hyperlinks with r:id, and r:id attributes."""
    text = re.sub(rf'<(?P<q>(?:\w+:)?)(?P<n>{REL_ELEMENTS})\b[^>]*?/>'
                  rf'|<(?P<r>(?:\w+:)?)(?P<m>{REL_ELEMENTS})\b[^>]*>.*?</(?P=r)(?P=m)>', '', text, flags=re.S)
    text = re.sub(rf'<(?:\w+:)?hyperlink\b(?=[^>]*\s{rel}:id=)[^>]*/>', '', text)
    text = re.sub(r'<((?:\w+:)?)hyperlinks\b[^>]*>\s*</\1hyperlinks>', '', text)
    return re.sub(rf'\s{rel}:id="[^"]*"', '', text)


def _remap_dxf(text, styles):
    return re.sub(r'(\sdxfId=")(\d+)"', lambda m: f'{m.group(1)}{styles.dxf(int(m.group(2)))}"', text)


def _graft_columns(tz, uz, tpart, upart, first, styles, strings, spool):
    with _text(uz, upart) as f:
        ext = _extent(f, first)
    with _text(tz, tpart) as f:
        offset = _max_shared_si(f) + 1
    with _text(tz, tpart) as tf, _text(uz, upart) as uf:
        tsheet, usheet = _Sheet(tf), _Sheet(uf)
        copier = _Copier(tsheet.p, styles, strings, offset)
        head = tsheet.head
        if ext is not None:
            head = _merge_cols(_widen(head, tsheet.p, ext), tsheet.p, _extra_cols(usheet, styles, first))
        spool.write(head + f'<{tsheet.p}sheetData>')
        for xml in _merge_rows(tsheet.rows(), _copied_rows(usheet.rows(), copier, first), tsheet.p):
            spool.write(xml)
        spool.write(f'</{tsheet.p}sheetData>' + tsheet.rest())
    return copier.cells, copier.dropped


def _copy_sheet(uz, upart, styles, strings, spool):
    with _text(uz, upart) as uf:
        usheet = _Sheet(uf)
        rel = re.search(rf'xmlns:(\w+)="{REL_NS}"', usheet.head)
        rel = rel.group(1) if rel else 'r'
        head = re.sub(r'\stabSelected="[^"]*"', '', usheet.head)
        head = _strip_rels(_remap_dxf(head, styles), rel)
        head = _style_cols(head, usheet.p, styles)
        copier = _Copier(usheet.p, styles, strings, None)
        spool.write(head + f'<{usheet.p}sheetData>')
        for r, xml in usheet.rows():
            cells = [copier.cell(c, r, raw) for c, raw in _cells(xml)]
            spool.write(_row_xml(_row_start(xml[:xml.index('>') + 1], styles), cells, usheet.p))
        spool.write(f'</{usheet.p}sheetData>' + _strip_rels(_remap_dxf(usheet.rest(), styles), rel))
    return copier.cells


def _style_cols(head, p, styles):
    return re.sub(rf'<{p}col\b[^>]*/>', lambda m: _style_tag(m.group(0), styles), head)


def _new_part(existing):
    n = 1
    while f'xl/worksheets/sheet{n}.xml' in existing:
        n += 1
    existing.add(f'xl/worksheets/sheet{n}.xml')
    return f'xl/worksheets/sheet{n}.xml'


def _grafted_names(uwb, twb, p):
    """Upload global names the template lacks (case-insensitive), skipping #REF! names."""
    have = {n.lower() for n in re.findall(r'<(?:\w+:)?definedName\b[^>]*?\sname="([^"]*)"', twb)}
    out = []
    pat = r'<(?P<q>(?:\w+:)?)definedName\b(?P<attrs>[^>]*?)(?<!/)>(?P<text>.*?)</(?P=q)definedName>'
    for m in re.finditer(pat, uwb, re.S):
        start = m.group(0)[:m.group(0).index('>') + 1]
        name = _attr(start, 'name') or ''
        if 'localSheetId' in start or '#REF!' in m.group('text') or name.lower() in have:
            continue
        have.add(name.lower())
        out.append(f'<{p}definedName{m.group("attrs")}>{m.group("text")}</{p}definedName>')
    return out


def _workbook_xml(twb, uwb, new, rids):
    p = re.search(r'<((?:\w+:)?)workbook\b', twb).group(1)
    rp = re.search(rf'xmlns:(\w+)="{REL_NS}"', twb)
    rp = rp.group(1) if rp else 'r'
    sid = max([0] + [int(x) for x in re.findall(r'sheetId="(\d+)"', twb)])
    tags = []
    for (name, _), rid in zip(new, rids):
        sid += 1
        tags.append(f'<{p}sheet name="{escape(name, {chr(34): "&quot;"})}" sheetId="{sid}" {rp}:id="{rid}"/>')
    out = _extend(twb, p, 'sheets', 'sheet', tags, BOOK_ORDER)
    out = _extend(out, p, 'definedNames', 'definedName', _grafted_names(uwb, twb, p), BOOK_ORDER)
    return _patch_workbook(out.encode('utf-8'))


def _write_spool(zout, zinfo, spool):
    spool.seek(0)
    with zout.open(zinfo, 'w', force_zip64=True) as dst:
        while True:
            chunk = spool.read(CHUNK)
            if not chunk:
                break
            dst.write(chunk.encode('utf-8'))
    spool.close()


def _copy_member(zin, zout, info):
    with zin.open(info) as src, zout.open(_clone(info), 'w', force_zip64=info.file_size > (1 << 30)) as dst:
        shutil.copyfileobj(src, dst, CHUNK)


def _graft(tz, uz, plan, zout, spool_dir):
    tsheets, usheets = _sheets(tz), _sheets(uz)
    tpart_of = dict(tsheets)
    tnames_lower = {n.lower() for n, _ in tsheets}
    uinfo = {n.lower(): (n, part) for n, part in usheets}
    for name, first in plan.columns.items():
        if name not in tpart_of:
            raise ValueError(f'template has no sheet {name!r}')
        if name.lower() not in uinfo:
            raise ValueError(f'upload has no sheet {name!r} for its extra columns')
        if first < 0:
            raise ValueError(f'first extra column for {name!r} must be >= 0')
    seen = set()
    for name in plan.sheets:
        if name.lower() not in uinfo:
            raise ValueError(f'upload has no sheet {name!r}')
        if name.lower() in tnames_lower or name.lower() in seen:
            raise ValueError(f'sheet {name!r} already exists in the template')
        seen.add(name.lower())

    existing = set(tz.namelist())
    new = [(uinfo[name.lower()][0], _new_part(existing)) for name in plan.sheets]
    strings = _shared_strings(uz)
    styles = _Styles(_read(tz, 'xl/styles.xml'), _read(uz, 'xl/styles.xml'))
    spool, cells, dropped = {}, 0, 0
    for name, first in plan.columns.items():
        part = tpart_of[name]
        spool[part] = tempfile.TemporaryFile('w+', encoding='utf-8', newline='', dir=spool_dir)
        c, d = _graft_columns(tz, uz, part, uinfo[name.lower()][1], first, styles, strings, spool[part])
        cells, dropped = cells + c, dropped + d
    for name, part in new:
        spool[part] = tempfile.TemporaryFile('w+', encoding='utf-8', newline='', dir=spool_dir)
        cells += _copy_sheet(uz, uinfo[name.lower()][1], styles, strings, spool[part])

    rels_text = _drop_calc_chain('xl/_rels/workbook.xml.rels', tz.read('xl/_rels/workbook.xml.rels'))
    rels_text = rels_text.decode('utf-8')
    base = max([0] + [int(x) for x in re.findall(r'Id="rId(\d+)"', rels_text)])
    rids = [f'rId{base + i + 1}' for i in range(len(new))]
    adds = ''.join(f'<Relationship Id="{rid}" Type="{WS_REL}" Target="{part[3:]}"/>'
                   for rid, (_, part) in zip(rids, new))
    end = rels_text.rindex('</')
    rels_text = rels_text[:end] + adds + rels_text[end:]
    wb_bytes = _workbook_xml(_read(tz, 'xl/workbook.xml'), _read(uz, 'xl/workbook.xml'), new, rids)
    ct = _drop_calc_chain('[Content_Types].xml', tz.read('[Content_Types].xml')).decode('utf-8')
    overrides = ''.join(f'<Override PartName="/{part}" ContentType="{WS_CTYPE}"/>' for _, part in new)
    end = ct.rindex('</')
    ct = (ct[:end] + overrides + ct[end:]).encode('utf-8')
    replaced = {
        'xl/workbook.xml': wb_bytes,
        'xl/_rels/workbook.xml.rels': rels_text.encode('utf-8'),
        '[Content_Types].xml': ct,
        'xl/styles.xml': styles.text().encode('utf-8'),
    }
    for info in tz.infolist():
        name = info.filename
        if name == 'xl/calcChain.xml':
            continue
        if name in spool:
            _write_spool(zout, _clone(info), spool.pop(name))
        elif name in replaced:
            zout.writestr(_clone(info), replaced[name])
        else:
            _copy_member(tz, zout, info)
    for _, part in new:
        zinfo = zipfile.ZipInfo(part, date_time=(1980, 1, 1, 0, 0, 0))
        zinfo.compress_type = zipfile.ZIP_DEFLATED
        _write_spool(zout, zinfo, spool.pop(part))
    return {'sheets': len(new), 'cells': cells, 'dropped_formulas': dropped}


def graft_extras(template_path, upload_path, plan, out_path):
    template_path, upload_path, out_path = Path(template_path), Path(upload_path), Path(out_path)
    if out_path.resolve() in (template_path.resolve(), upload_path.resolve()):
        raise ValueError('graft must write a new file, never the template or the upload')
    fd, tmp = tempfile.mkstemp(suffix='.xlsx.part', dir=out_path.parent)
    os.close(fd)
    try:
        with zipfile.ZipFile(template_path) as tz, zipfile.ZipFile(upload_path) as uz, \
                zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED) as zout:
            counts = _graft(tz, uz, plan, zout, out_path.parent)
        os.replace(tmp, out_path)
    except BaseException:
        Path(tmp).unlink(missing_ok=True)
        raise
    return counts


def _reordered(wb, order):
    p = re.search(r'<((?:\w+:)?)workbook\b', wb).group(1)
    m = re.search(rf'(<{p}sheets\b[^>]*>)(.*?)(</{p}sheets>)', wb, re.S)
    tags = _elements(m.group(2), 'sheet')
    names = [unescape(_attr(t, 'name'), ENTITIES).lower() for t in tags]
    picked = [names.index(n.lower()) for n in dict.fromkeys(order) if n.lower() in names]
    perm = picked + [i for i in range(len(tags)) if i not in picked]
    where = {old: new for new, old in enumerate(perm)}
    wb = wb[:m.start(2)] + ''.join(tags[i] for i in perm) + wb[m.end(2):]
    wb = re.sub(rf'(<{p}definedName\b[^>]*?\slocalSheetId=")(\d+)"',
                lambda x: f'{x.group(1)}{where[int(x.group(2))]}"', wb)
    for attr in ('activeTab', 'firstSheet'):
        wb = re.sub(rf'(<{p}workbookView\b[^>]*?\s{attr}=")(\d+)"',
                    lambda x: f'{x.group(1)}{where[int(x.group(2))]}"', wb)
    return wb


def reorder_sheets(path, order):
    """Reorder workbook tabs to follow `order` (case-insensitive; unknown names ignored). Unlisted sheets
    keep their relative order at the end."""
    path = Path(path)
    fd, tmp = tempfile.mkstemp(suffix='.xlsx.part', dir=path.parent)
    os.close(fd)
    try:
        with zipfile.ZipFile(path) as zin, zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED) as zout:
            wb = _reordered(zin.read('xl/workbook.xml').decode('utf-8'), tuple(order)).encode('utf-8')
            for info in zin.infolist():
                if info.filename == 'xl/workbook.xml':
                    zout.writestr(_clone(info), wb)
                else:
                    _copy_member(zin, zout, info)
        os.replace(tmp, path)
    except BaseException:
        Path(tmp).unlink(missing_ok=True)
        raise
