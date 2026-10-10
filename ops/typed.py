# ops/typed.py
"""Excel-typed values: blank, empty string, number, text, bool and error are different things."""
import math
ERRORS = {'#NULL!', '#DIV/0!', '#VALUE!', '#REF!', '#NAME?', '#NUM!', '#N/A', '#SPILL!', '#CALC!', '#GETTING_DATA'}

def typed(v):
    if v is None:
        return ('blank', None)
    if isinstance(v, bool):
        return ('bool', v)
    if isinstance(v, (int, float)):
        return ('num', float(v))
    s = str(v)
    if s == '':
        return ('empty', '')
    if s in ERRORS:
        return ('error', s)
    return ('text', s)

def same(a, b, rel=1e-9):
    ta, tb = typed(a), typed(b)
    if ta[0] != tb[0]:
        return False
    if ta[0] == 'num':
        return math.isclose(ta[1], tb[1], rel_tol=rel, abs_tol=1e-12)
    return ta[1] == tb[1]
