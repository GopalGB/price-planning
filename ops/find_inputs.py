# ops/find_inputs.py
"""List input cells: unlocked cells on protected sheets. Streams the XML; never loads the whole sheet."""
import re, sys, zipfile
import xml.etree.ElementTree as ET
from pathlib import Path
NS = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
RNS = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'

def _locked(xf, inherited=True):
    if xf is None or xf.get('applyProtection') in ('0', 'false'):
        return inherited
    protection = xf.find(f'{NS}protection')
    if protection is None:
        return inherited
    return protection.get('locked', '1') not in ('0', 'false')

def _unlocked_style_ids(z):
    root = ET.fromstring(z.read('xl/styles.xml'))
    styles = root.find(f'{NS}cellStyleXfs')
    ids = set()
    for i, xf in enumerate(root.find(f'{NS}cellXfs')):
        parent = int(xf.get('xfId', '0'))
        base = styles[parent] if styles is not None and parent < len(styles) else None
        if not _locked(xf, _locked(base)):
            ids.add(str(i))
    return ids

def _sheets(z):
    wb = ET.fromstring(z.read('xl/workbook.xml'))
    rels = ET.fromstring(z.read('xl/_rels/workbook.xml.rels'))
    target = {r.get('Id'): r.get('Target') for r in rels}
    for s in wb.find(f'{NS}sheets'):
        t = target[s.get(f'{RNS}id')].lstrip('/')
        yield s.get('name'), t if t.startswith('xl/') else 'xl/' + t

def find_inputs(xlsm):
    out = []
    with zipfile.ZipFile(xlsm) as z:
        unlocked = _unlocked_style_ids(z)
        for name, part in _sheets(z):
            protected, cells = False, []
            with z.open(part) as f:
                for _, el in ET.iterparse(f):
                    if el.tag == f'{NS}c':
                        if el.get('s', '0') in unlocked and el.find(f'{NS}f') is None:
                            cells.append(el.get('r'))
                        el.clear()
                    elif el.tag == f'{NS}sheetProtection' and el.get('sheet') in ('1', 'true'):
                        protected = True
                    elif el.tag == f'{NS}row':
                        el.clear()
            if protected:
                quoted = "'" + name.replace("'", "''") + "'"
                out += [f'{quoted}!{c}' for c in cells]
    return sorted(out)

if __name__ == '__main__':
    for ref in find_inputs(Path(sys.argv[1])):
        print(ref)
