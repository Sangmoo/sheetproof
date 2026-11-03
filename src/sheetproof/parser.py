"""엑셀 수식 → AST 파서.

``openpyxl.formula.Tokenizer`` 로 토큰을 얻고 Pratt 파서로 AST 를 만든다.
연산자 우선순위는 엑셀과 같다 (높은 것부터):
단항 ``-``/``+`` > ``%`` > ``^`` > ``* /`` > ``+ -`` > ``&`` > 비교.
엑셀의 ``^`` 는 왼쪽 결합이며 ``-2^2`` 는 ``4`` 이다.
"""

from __future__ import annotations

from dataclasses import dataclass, field

from openpyxl.formula.tokenizer import Token, Tokenizer, TokenizerError

FUNC_PREFIXES = ("_xlfn._xlws.", "_xlfn.", "_xlws.", "_xll.")


class FormulaParseError(ValueError):
    pass


# ---------------------------------------------------------------------------
# AST
# ---------------------------------------------------------------------------
@dataclass(frozen=True)
class Number:
    value: float | int


@dataclass(frozen=True)
class String:
    value: str


@dataclass(frozen=True)
class Bool:
    value: bool


@dataclass(frozen=True)
class Error:
    code: str


@dataclass(frozen=True)
class Missing:
    """빈 인수."""


@dataclass(frozen=True)
class Ref:
    """셀/범위 참조. ``sheet`` 가 None 이면 수식이 있는 시트."""

    sheet: str | None
    address: str  # "$A$1", "A1:B3", "A:A", "1:3"
    external: bool = False

    @property
    def is_range(self) -> bool:
        return ":" in self.address


@dataclass(frozen=True)
class Name:
    name: str


@dataclass(frozen=True)
class Unary:
    op: str
    operand: object


@dataclass(frozen=True)
class Postfix:
    op: str
    operand: object


@dataclass(frozen=True)
class Binary:
    op: str
    left: object
    right: object


@dataclass(frozen=True)
class Call:
    name: str
    args: tuple = field(default_factory=tuple)


@dataclass(frozen=True)
class Array:
    rows: tuple  # tuple[tuple[node]]


# 이항 연산자 결합력 (left binding power)
BINARY_BP = {
    "=": 10, "<>": 10, "<": 10, ">": 10, "<=": 10, ">=": 10,
    "&": 20,
    "+": 30, "-": 30,
    "*": 40, "/": 40,
    "^": 50,
}
POSTFIX_BP = 55
PREFIX_BP = 60


def normalize_func_name(raw: str) -> str:
    name = raw.rstrip("(").upper()
    for p in FUNC_PREFIXES:
        if name.startswith(p.upper()):
            name = name[len(p):]
            break
    return name


def _unquote_sheet(s: str) -> str:
    if len(s) >= 2 and s[0] == "'" and s[-1] == "'":
        return s[1:-1].replace("''", "'")
    return s


def split_reference(text: str) -> tuple[str | None, str, bool]:
    """``'My Sheet'!$A$1:B2`` → (sheet, address, external)."""
    if "!" not in text:
        return None, text, False
    # 시트 이름은 따옴표로 감쌀 수 있고 그 안에 ! 가 있을 수 있다
    if text.startswith("'"):
        i = 1
        while i < len(text):
            if text[i] == "'":
                if i + 1 < len(text) and text[i + 1] == "'":
                    i += 2
                    continue
                break
            i += 1
        sheet_part, rest = text[: i + 1], text[i + 1:]
        if not rest.startswith("!"):
            raise FormulaParseError(f"잘못된 참조: {text}")
        address = rest[1:]
    else:
        sheet_part, _, address = text.rpartition("!")
    sheet = _unquote_sheet(sheet_part)
    external = sheet.startswith("[") or "[" in sheet_part
    # Sheet1!A1:Sheet1!B2 형태
    if "!" in address:
        s2, addr2, _ = split_reference(address)
        if s2 is not None and s2 != sheet:
            raise FormulaParseError(f"3D 참조는 지원하지 않습니다: {text}")
        address = addr2
    if ":" in address:
        a, b = address.split(":", 1)
        if "!" in b:
            b = b.rpartition("!")[2]
        address = f"{a}:{b}"
    return sheet, address, external


def _is_ref_text(text: str) -> bool:
    import re

    addr = text.rpartition("!")[2] if "!" in text else text
    return bool(
        re.fullmatch(r"\$?[A-Za-z]{1,3}\$?\d+(:\$?[A-Za-z]{1,3}\$?\d+)?", addr)
        or re.fullmatch(r"\$?[A-Za-z]{1,3}:\$?[A-Za-z]{1,3}", addr)
        or re.fullmatch(r"\$?\d+:\$?\d+", addr)
        or addr.startswith("#REF!")
    )


def _parse_text_literal(v: str) -> str:
    return v[1:-1].replace('""', '"')


# ---------------------------------------------------------------------------
# 파서
# ---------------------------------------------------------------------------
class _Parser:
    def __init__(self, tokens: list[Token], formula: str):
        self.toks = tokens
        self.i = 0
        self.formula = formula

    def peek(self) -> Token | None:
        return self.toks[self.i] if self.i < len(self.toks) else None

    def next(self) -> Token:
        t = self.peek()
        if t is None:
            raise FormulaParseError(f"수식이 예기치 않게 끝났습니다: {self.formula}")
        self.i += 1
        return t

    def fail(self, msg: str):
        raise FormulaParseError(f"{msg}: {self.formula}")

    def parse(self):
        node = self.expr(0)
        if self.peek() is not None:
            self.fail(f"해석할 수 없는 토큰 '{self.peek().value}'")
        return node

    def expr(self, min_bp: int):
        left = self.prefix()
        while True:
            t = self.peek()
            if t is None:
                break
            if t.type == Token.OP_POST:
                if POSTFIX_BP < min_bp:
                    break
                self.next()
                left = Postfix(t.value, left)
                continue
            if t.type == Token.OP_IN:
                op = t.value
                if op == ":" or op == " " or op == ",":
                    self.fail(f"지원하지 않는 참조 연산자 '{op}'")
                bp = BINARY_BP.get(op)
                if bp is None:
                    self.fail(f"알 수 없는 연산자 '{op}'")
                if bp < min_bp:
                    break
                self.next()
                # 모든 이항 연산자는 왼쪽 결합 (엑셀은 ^ 도 왼쪽 결합)
                right = self.expr(bp + 1)
                left = Binary(op, left, right)
                continue
            break
        return left

    def prefix(self):
        t = self.next()
        if t.type == Token.OP_PRE:
            operand = self.expr(PREFIX_BP)
            return Unary(t.value, operand)
        if t.type == Token.OPERAND:
            return self.operand(t)
        if t.type == Token.FUNC and t.subtype == Token.OPEN:
            return self.call(t)
        if t.type == Token.PAREN and t.subtype == Token.OPEN:
            node = self.expr(0)
            close = self.next()
            if not (close.type == Token.PAREN and close.subtype == Token.CLOSE):
                self.fail("괄호가 닫히지 않았습니다")
            return node
        if t.type == Token.ARRAY and t.subtype == Token.OPEN:
            return self.array()
        self.fail(f"예상하지 못한 토큰 '{t.value}'")

    def operand(self, t: Token):
        v = t.value
        st = t.subtype
        if st == Token.NUMBER:
            f = float(v)
            return Number(int(f) if f.is_integer() and "." not in v and "e" not in v.lower() else f)
        if st == Token.TEXT:
            return String(_parse_text_literal(v))
        if st == Token.LOGICAL:
            return Bool(v.upper() == "TRUE")
        if st == Token.ERROR:
            return Error(v.upper())
        if st == Token.RANGE:
            if _is_ref_text(v):
                sheet, address, external = split_reference(v)
                if address.upper().startswith("#REF!"):
                    return Error("#REF!")
                return Ref(sheet, address.upper(), external)
            if "!" in v or v.startswith("["):
                sheet, address, external = split_reference(v)
                if external:
                    return Ref(sheet, address.upper(), True)
                return Name(address)
            return Name(v)
        self.fail(f"알 수 없는 피연산자 '{v}'")

    def call(self, t: Token):
        name = normalize_func_name(t.value)
        args: list = []
        # 인수 없는 호출
        nt = self.peek()
        if nt is not None and nt.type == Token.FUNC and nt.subtype == Token.CLOSE:
            self.next()
            return Call(name, ())
        while True:
            nt = self.peek()
            if nt is None:
                self.fail("함수 괄호가 닫히지 않았습니다")
            if (nt.type == Token.SEP and nt.subtype == Token.ARG) or (
                nt.type == Token.FUNC and nt.subtype == Token.CLOSE
            ):
                args.append(Missing())
            else:
                args.append(self.expr(0))
            sep = self.next()
            if sep.type == Token.SEP and sep.subtype == Token.ARG:
                continue
            if sep.type == Token.FUNC and sep.subtype == Token.CLOSE:
                break
            self.fail(f"함수 인수 구분자가 필요합니다 ('{sep.value}')")
        return Call(name, tuple(args))

    def array(self):
        rows: list[list] = [[]]
        while True:
            t = self.peek()
            if t is None:
                self.fail("배열 상수가 닫히지 않았습니다")
            if t.type == Token.ARRAY and t.subtype == Token.CLOSE:
                self.next()
                break
            if t.type == Token.SEP:
                self.next()
                if t.subtype == Token.ROW:
                    rows.append([])
                continue
            rows[-1].append(self.expr(0))
        return Array(tuple(tuple(r) for r in rows))


def parse_formula(formula: str):
    """``=A1+1`` 형태의 수식을 AST 로 만든다."""
    if not formula.startswith("="):
        formula = "=" + formula
    try:
        tok = Tokenizer(formula)
    except TokenizerError as e:  # pragma: no cover - 토크나이저 메시지 그대로 전달
        raise FormulaParseError(f"토큰화 실패: {e}") from e
    tokens = [t for t in tok.items if t.type != Token.WSPACE]
    if not tokens:
        raise FormulaParseError(f"빈 수식: {formula}")
    return _Parser(tokens, formula).parse()


# ---------------------------------------------------------------------------
# AST → 수식 문자열 (왕복 테스트, 보고서 표시용)
# ---------------------------------------------------------------------------
def _quote_sheet(s: str) -> str:
    import re

    if re.fullmatch(r"[A-Za-z_][A-Za-z0-9_.]*", s) and not _is_ref_text(s):
        return s
    return "'" + s.replace("'", "''") + "'"


def _num_text(v) -> str:
    if isinstance(v, float):
        r = repr(v)
        return r[:-2] if r.endswith(".0") else r
    return str(v)


def to_formula(node, top: bool = True) -> str:
    s = _fmt(node, 0)
    return "=" + s if top else s


def _fmt(node, parent_bp: int) -> str:
    if isinstance(node, Number):
        return _num_text(node.value)
    if isinstance(node, String):
        return '"' + node.value.replace('"', '""') + '"'
    if isinstance(node, Bool):
        return "TRUE" if node.value else "FALSE"
    if isinstance(node, Error):
        return node.code
    if isinstance(node, Missing):
        return ""
    if isinstance(node, Ref):
        return (f"{_quote_sheet(node.sheet)}!" if node.sheet is not None else "") + node.address
    if isinstance(node, Name):
        return node.name
    if isinstance(node, Unary):
        return node.op + _fmt(node.operand, PREFIX_BP)
    if isinstance(node, Postfix):
        return _fmt(node.operand, POSTFIX_BP) + node.op
    if isinstance(node, Binary):
        bp = BINARY_BP[node.op]
        s = f"{_fmt(node.left, bp)}{node.op}{_fmt(node.right, bp + 1)}"
        return f"({s})" if bp < parent_bp else s
    if isinstance(node, Call):
        return f"{node.name}(" + ",".join(_fmt(a, 0) for a in node.args) + ")"
    if isinstance(node, Array):
        return "{" + ";".join(",".join(_fmt(x, 0) for x in r) for r in node.rows) + "}"
    raise TypeError(node)


def walk(node):
    """AST 의 모든 노드를 순회한다."""
    yield node
    if isinstance(node, (Unary, Postfix)):
        yield from walk(node.operand)
    elif isinstance(node, Binary):
        yield from walk(node.left)
        yield from walk(node.right)
    elif isinstance(node, Call):
        for a in node.args:
            yield from walk(a)
    elif isinstance(node, Array):
        for r in node.rows:
            for x in r:
                yield from walk(x)
