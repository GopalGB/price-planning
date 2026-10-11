# ops/fixtures.py
"""Deterministic synthetic workbooks. No company data. Used by CI and by the Univer gate."""
from pathlib import Path
from openpyxl import Workbook
from openpyxl.styles import Protection

def make_semantic(path):
    wb = Workbook()
    inp = wb.active; inp.title = 'Inputs'
    inp.protection.sheet = True
    rows = [('Cost', 10), ('Retail', 20), ('Blank', None), ('EmptyText', ''), ('Rate', '1.234,5')]
    inputs = []
    for i, (label, val) in enumerate(rows, start=2):
        inp.cell(i, 1, label)
        c = inp.cell(i, 2, val)
        c.protection = Protection(locked=False)
        inputs.append(f'Inputs!B{i}')
    lk = wb.create_sheet('Lookup')
    for i in range(1, 11):
        lk.cell(i, 1, f'K{i}'); lk.cell(i, 2, i * 1.5); lk.cell(i, 3, 'A' if i % 2 else 'B')
    calc = wb.create_sheet('Calc')
    formulas = {
        'B2': '=Inputs!B2*2',
        'B3': '=IFERROR(1-Inputs!B2/Inputs!B3,"")',
        'B4': '=IF(Inputs!B4="","blank","set")',
        'B5': '=IF(ISBLANK(Inputs!B4),1,0)',
        'B6': '=IFERROR(1/0,"err")',
        'B7': '=_xlfn.XLOOKUP("K3",Lookup!A1:A10,Lookup!B1:B10,"none")',
        'B8': '=_xlfn.LET(_xlpm.x,Inputs!B2,_xlpm.x+1)',
        'B9': '=_xlfn.SWITCH(Inputs!B2,10,"ten",20,"twenty","other")',
        'B10': '=_xlfn.IFS(Inputs!B2>15,"hi",TRUE,"lo")',
        'B11': '=_xlfn.NUMBERVALUE(Inputs!B6,",",".")',
        'B12': '=SUMIFS(Lookup!B1:B10,Lookup!C1:C10,"A")',
        'B13': '=_xlfn.SUM(_xlfn._xlws.FILTER(Lookup!B1:B10,Lookup!C1:C10="B"))',
        'B14': '=ROWS(_xlfn.UNIQUE(Lookup!C1:C10))',
        'B15': '=SUM(_xlfn.SEQUENCE(3))',
        'B16': '=VLOOKUP("K5",Lookup!A1:C10,2,FALSE)',
        'B17': '=SUBSTITUTE(TRIM("  a  b "),"a","x")',
    }
    for ref, f in formulas.items():
        calc[ref] = f
    wb.save(path)
    return {'inputs': inputs, 'formulas': {f'Calc!{k}': v for k, v in formulas.items()}}

def make_scale(path, rows=4088, cols=312, lookup_rows=100000):
    wb = Workbook(write_only=True)
    lk = wb.create_sheet('Lookup')
    for r in range(1, lookup_rows + 1):
        lk.append([f'K{r}', r * 0.5, f'G{r % 97}'] + [r % 13] * 14)
    grid = wb.create_sheet('Grid')
    for r in range(1, rows + 1):
        row = [f'K{(r * 7) % lookup_rows + 1}', r % 50 + 1]
        for c in range(3, cols + 1):
            row.append(f'=IFERROR(VLOOKUP($A{r},Lookup!$A:$C,2,FALSE)*$B{r}+{c},0)' if c % 3 else f'=SUBSTITUTE(TRIM($A{r}),"K","")')
        grid.append(row)
    wb.save(path)
