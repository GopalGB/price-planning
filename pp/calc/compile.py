"""Turn one formula family into a Python function f(book, r, c), compiled once and run for every member.

Relative references become offsets from the member's own cell, so one function serves the whole family.
"""

from __future__ import annotations

import re

import formualizer

from . import ops
from .functions import _MISSING, FUNCTIONS
from .model import column_number
from .values import Area, XlError

ERRORS = ("#NULL!", "#DIV/0!", "#VALUE!", "#REF!", "#NAME?", "#NUM!", "#N/A", "#SPILL!", "#CALC!")
STRING = re.compile(r'"(?:[^"]|"")*"')
PART = re.compile(r"(\$?)([A-Za-z]{1,3})?(\$?)(\d+)?$")
BINARY = {"+": "add", "-": "sub", "*": "mul", "/": "div", "^": "power", "&": "concat", "=": "eq", "<>": "ne",
          "<": "lt", "<=": "le", ">": "gt", ">=": "ge"}
UNARY = {"-": "neg", "+": "plus", "%": "percent"}
ALL = object()
# Argument positions that take a range as it is; other arguments of a non-array formula are intersected
RANGE_ARGS = {
    "MATCH": {1}, "VLOOKUP": {1}, "HLOOKUP": {1}, "INDEX": {0}, "XLOOKUP": {1, 2}, "SUMIF": {0, 2},
    "COUNTIF": {0}, "UNIQUE": {0}, "SUM": ALL, "MAX": ALL, "MIN": ALL, "COUNTA": ALL, "AND": ALL, "OR": ALL,
    "CONCAT": ALL, "SUMPRODUCT": ALL, "RANK": {1}, "RANK.EQ": {1},
}
ARRAY_ARGS = {"SUMPRODUCT"}  # evaluated as arrays even in a non-array formula
DEFAULTED = {"XLOOKUP", "XMATCH"}  # an omitted argument takes its default, not a blank


class Unsupported(Exception):
    """A formula the calculator cannot evaluate; its cells keep Excel's saved results."""


def _range_args(name: str, count: int):
    if name in ("SUMIFS", "MAXIFS", "MINIFS"):
        return {0} | set(range(1, count, 2))
    if name == "COUNTIFS":
        return set(range(0, count, 2))
    spec = RANGE_ARGS.get(name, set())
    return set(range(count)) if spec is ALL else spec


class Compiler:
    def __init__(self, sheet_names: list[str]):
        self.sheets = {name.casefold(): index for index, name in enumerate(sheet_names)}

    def compile(self, text: str, anchor: tuple[int, int, int], array: bool):
        tree = formualizer.parse(text if text.startswith("=") else "=" + text).to_dict()
        unquoted = STRING.sub("", text)
        state = _State(self, anchor, array, {code for code in ERRORS if code in unquoted})
        body = state.emit(tree, want_area=array, array=array)
        source = f"def formula(book, r, c):\n    return result({body}, r, c, {array})\n"
        namespace = dict(state.constants, result=ops.result, isect=ops.isect, Area=Area)
        namespace.update({name: getattr(ops, name) for name in (*BINARY.values(), *UNARY.values(),
                                                                "fn_if", "fn_iferror", "fn_ifs", "fn_switch")})
        exec(compile(source, f"<family {anchor}>", "exec"), namespace)  # nosec B102 # code generated from parsed template formulas
        return namespace["formula"]


class _State:
    def __init__(self, compiler: Compiler, anchor, array: bool, errors: set):
        self.compiler, self.anchor, self.array, self.errors = compiler, anchor, array, errors
        self.constants, self.names = {}, {}

    def constant(self, value) -> str:
        name = f"k{len(self.constants)}"
        self.constants[name] = value
        return name

    def emit(self, node, want_area: bool, array: bool) -> str:
        kind = node["node_type"]
        if kind == "Literal":
            value = node["value"]
            if isinstance(value, str) and value in self.errors:
                return self.constant(XlError(value))
            if isinstance(value, bool):
                return repr(value)
            if isinstance(value, (int, float)):
                return repr(float(value))
            return repr(value)
        if kind == "Omitted":
            return "None"
        if kind == "Reference":
            code = self.reference(node["reference"])
            return code if want_area or array else f"isect({code}, r, c)"
        if kind == "BinaryOp":
            operator = node["operator"]
            if operator not in BINARY:
                raise Unsupported(f"operator {operator!r}")
            left = self.emit(node["left"], False, array)
            right = self.emit(node["right"], False, array)
            return f"{BINARY[operator]}({left}, {right})"
        if kind == "UnaryOp":
            return f"{UNARY[node['operator']]}({self.emit(node['operand'], False, array)})"
        if kind == "Array":
            rows = [[self._literal(item) for item in row] for row in node["rows"]]
            return self.constant(Area(rows))
        if kind == "Function":
            return self.function(node["name"], node.get("args", []), want_area, array)
        raise Unsupported(f"node {kind}")

    def _literal(self, node):
        if node["node_type"] != "Literal":
            raise Unsupported("array constant with an expression")
        value = node["value"]
        if isinstance(value, str) and value in self.errors:
            return XlError(value)
        return float(value) if isinstance(value, (int, float)) and not isinstance(value, bool) else value

    def function(self, raw: str, args: list, want_area: bool, array: bool) -> str:
        name = raw.upper().removeprefix("_XLFN.").removeprefix("_XLWS.")
        lazy = getattr(self, "fn_" + name.replace(".", "_"), None)
        if lazy is not None:
            return lazy(args, array)
        target = FUNCTIONS.get(name)
        if target is None:
            raise Unsupported(f"function {name}")
        inner = array or name in ARRAY_ARGS
        takes = _range_args(name, len(args))
        parts = [self.constant(_MISSING) if name in DEFAULTED and arg["node_type"] == "Omitted"
                 else self.emit(arg, index in takes, inner) for index, arg in enumerate(args)]
        return f"{self.constant(target)}({', '.join(parts)})"

    def thunk(self, node, array: bool) -> str:
        return f"lambda: {self.emit(node, False, array)}"

    # Functions that decide which arguments to evaluate, or that read references themselves

    def fn_IF(self, args, array):
        test = self.emit(args[0], False, array)
        then = self.thunk(args[1], array) if len(args) > 1 else "lambda: True"
        otherwise = self.thunk(args[2], array) if len(args) > 2 else "lambda: False"
        return f"fn_if({test}, {then}, {otherwise})"

    def fn_IFERROR(self, args, array):
        return f"fn_iferror({self.emit(args[0], False, array)}, {self.thunk(args[1], array)})"

    def fn_IFS(self, args, array):
        pairs = [f"({self.thunk(args[i], array)}, {self.thunk(args[i + 1], array)})" for i in range(0, len(args) - 1, 2)]
        return f"fn_ifs({', '.join(pairs)})"

    def fn_SWITCH(self, args, array):
        subject = self.emit(args[0], False, array)
        rest = args[1:]
        default = self.thunk(rest[-1], array) if len(rest) % 2 else "None"
        cases = [f"({self.thunk(rest[i], array)}, {self.thunk(rest[i + 1], array)})" for i in range(0, len(rest) - 1, 2)]
        return f"fn_switch({subject}, ({', '.join(cases)}{',' if len(cases) == 1 else ''}), {default})"

    def fn_LET(self, args, array):
        saved = dict(self.names)
        bindings = []
        for i in range(0, len(args) - 1, 2):
            label = args[i]["reference"]
            value = self.emit(args[i + 1], True, array)
            variable = f"v{len(self.names)}_{re.sub(r'[^A-Za-z0-9_]', '_', label)}"
            self.names[label.casefold()] = variable
            bindings.append((variable, value))
        body = self.emit(args[-1], array, array)
        self.names = saved
        for variable, value in reversed(bindings):
            body = f"(lambda {variable}: {body})({value})"
        return body

    def fn_COLUMN(self, args, array):
        return self._position(args, array, column=True)

    def fn_ROW(self, args, array):
        return self._position(args, array, column=False)

    def _position(self, args, array, column):
        if not args:
            return "float(c + 1)" if column else "float(r + 1)"
        if args[0]["node_type"] != "Reference":
            raise Unsupported("COLUMN/ROW of an expression")
        sheet, first, last = self._parse(args[0]["reference"])
        index = 1 if column else 0
        start, end = first[index], last[index]
        if start is None:
            return "1.0"
        if array and end != start:
            values = (f"float({self._axis(start, column)} + {k} + 1)" for k in range(self._span(start, end)))
            joined = ", ".join(values)
            return f"Area([[{joined}]])" if column else f"Area([[v] for v in ({joined},)])"
        return f"float({self._axis(start, column)} + 1)"

    def fn_ISFORMULA(self, args, array):
        sheet, first, _ = self._parse(args[0]["reference"])
        return f"book.is_formula({sheet}, {self._axis(first[0], False)}, {self._axis(first[1], True)})"

    # References

    def reference(self, text: str) -> str:
        label = text.casefold()
        if label in self.names:
            return self.names[label]
        sheet, first, last = self._parse(text)
        (r1, c1), (r2, c2) = first, last
        if first == last and r1 is not None and c1 is not None:
            return f"book.cell({sheet}, {self._axis(r1, False)}, {self._axis(c1, True)})"
        top = "0" if r1 is None else self._axis(r1, False)
        # A whole row or column ends where the book being read ends, so a copy's added rows count too
        bottom = f"book.last_row({sheet})" if r2 is None else self._axis(r2, False)
        left = "0" if c1 is None else self._axis(c1, True)
        right = f"book.last_col({sheet})" if c2 is None else self._axis(c2, True)
        return f"book.area({sheet}, {top}, {left}, {bottom}, {right})"

    def _parse(self, text: str):
        sheet = self.anchor[0]
        if "!" in text:
            prefix, _, text = text.rpartition("!")
            key = prefix.strip("'").replace("''", "'").casefold()
            if key not in self.compiler.sheets:
                raise Unsupported(f"reference to an unknown sheet or external book")
            sheet = self.compiler.sheets[key]
        ends = text.split(":")
        if len(ends) > 2:
            raise Unsupported("multi-area reference")
        parsed = [self._point(end) for end in ends]
        first, last = parsed[0], parsed[-1]
        if len(parsed) == 2:
            # Order the corners the way Excel normalises B2:A1
            first, last = self._ordered(first, last)
        return sheet, first, last

    def _point(self, text: str):
        match = PART.fullmatch(text)
        if match is None or (match.group(2) is None and match.group(4) is None):
            raise Unsupported(f"reference {text!r}")
        col_abs, letters, row_abs, digits = match.groups()
        if letters is None:  # a whole-row reference like $2 puts its $ before the row number
            row_abs, col_abs = col_abs, ""
        row = None if digits is None else (("abs", int(digits) - 1) if row_abs else ("rel", int(digits) - 1 - self.anchor[1]))
        col = None if letters is None else (("abs", column_number(letters) - 1) if col_abs else ("rel", column_number(letters) - 1 - self.anchor[2]))
        return row, col

    def _ordered(self, first, last):
        def lower(a, b):
            if a is None or b is None or a[0] != b[0] or a[0] == "rel":
                return a, b
            return (a, b) if a[1] <= b[1] else (b, a)

        r1, r2 = lower(first[0], last[0])
        c1, c2 = lower(first[1], last[1])
        return (r1, c1), (r2, c2)

    def _axis(self, part, column: bool) -> str:
        kind, value = part
        if kind == "abs":
            return str(value)
        variable = "c" if column else "r"
        return variable if value == 0 else f"{variable} + {value}" if value > 0 else f"{variable} - {-value}"

    def _span(self, start, end) -> int:
        return end[1] - start[1] + 1 if start[0] == end[0] else 1
