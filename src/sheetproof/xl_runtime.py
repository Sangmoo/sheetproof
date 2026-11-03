"""Excel 의미론 런타임 (sheetproof 가 생성한 코드가 사용하는 헬퍼).

이 파일은 생성된 모델 코드와 함께 그대로 내보내지므로 표준 라이브러리만 사용한다.

값 표현
- 빈 셀: ``None``
- 숫자: ``int`` / ``float`` (날짜는 1900 체계 일련번호)
- 문자열: ``str``, 논리값: ``bool``
- 오류값: :class:`XLError` (``#N/A``, ``#DIV/0!`` ...)
- 범위/배열: 2차원 리스트 ``list[list[value]]``

모든 공개 함수는 예외를 던지지 않고 오류값을 반환하는 것을 원칙으로 한다.
"""

from __future__ import annotations

import calendar
import datetime as _dt
import functools
import math
import re
from decimal import ROUND_DOWN, ROUND_HALF_UP, ROUND_UP, Decimal, InvalidOperation

__all__ = ["XLError", "rng", "values_equal", "normalize", "FUNCTIONS"]


# ---------------------------------------------------------------------------
# 오류값
# ---------------------------------------------------------------------------
class XLError:
    """엑셀 오류값. 같은 코드면 같은 값으로 취급한다."""

    __slots__ = ("code",)
    CODES = ("#NULL!", "#DIV/0!", "#VALUE!", "#REF!", "#NAME?", "#NUM!", "#N/A")

    def __init__(self, code: str):
        self.code = code

    def __eq__(self, other):
        return isinstance(other, XLError) and other.code == self.code

    def __hash__(self):
        return hash(("XLError", self.code))

    def __repr__(self):
        return f"xl.err({self.code!r})"

    def __str__(self):
        return self.code


def err(code: str) -> XLError:
    return XLError(code)


NULL = XLError("#NULL!")
DIV0 = XLError("#DIV/0!")
VALUE = XLError("#VALUE!")
REF = XLError("#REF!")
NAME = XLError("#NAME?")
NUM = XLError("#NUM!")
NA = XLError("#N/A")


class _Missing:
    """함수 인수가 비어 있는 경우 (예: ``VLOOKUP(a, b, 2, )``)."""

    _inst = None

    def __new__(cls):
        if cls._inst is None:
            cls._inst = super().__new__(cls)
        return cls._inst

    def __repr__(self):
        return "xl.MISSING"


MISSING = _Missing()


def _is_err(x) -> bool:
    return isinstance(x, XLError)


def _is_num(x) -> bool:
    return isinstance(x, (int, float)) and not isinstance(x, bool)


def _safe(fn):
    """예상치 못한 파이썬 예외를 #VALUE! 로 바꾼다."""

    @functools.wraps(fn)
    def wrapper(*args):
        try:
            return fn(*args)
        except (ZeroDivisionError,):
            return DIV0
        except (OverflowError,):
            return NUM
        except Exception:  # noqa: BLE001 - 엑셀 함수는 예외 대신 오류값을 돌려준다
            return VALUE

    return wrapper


def guard(thunk):
    """LLM 등 외부에서 만든 식을 감싸 예외를 #VALUE! 로 바꾼다."""
    try:
        return thunk()
    except ZeroDivisionError:
        return DIV0
    except Exception:  # noqa: BLE001
        return VALUE


# ---------------------------------------------------------------------------
# 범위
# ---------------------------------------------------------------------------
_COL_RE = re.compile(r"^\$?([A-Za-z]{1,3})\$?(\d+)$")


def _col_to_num(col: str) -> int:
    n = 0
    for ch in col.upper():
        n = n * 26 + (ord(ch) - 64)
    return n


def _num_to_col(n: int) -> str:
    s = ""
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


@functools.lru_cache(maxsize=4096)
def _range_keys(address: str):
    sheet, _, ref = address.rpartition("!")
    parts = ref.replace("$", "").split(":")
    m1 = _COL_RE.match(parts[0])
    m2 = _COL_RE.match(parts[-1])
    c1, r1 = _col_to_num(m1.group(1)), int(m1.group(2))
    c2, r2 = _col_to_num(m2.group(1)), int(m2.group(2))
    c1, c2 = min(c1, c2), max(c1, c2)
    r1, r2 = min(r1, r2), max(r1, r2)
    prefix = f"{sheet}!" if sheet else ""
    return tuple(
        tuple(f"{prefix}{_num_to_col(c)}{r}" for c in range(c1, c2 + 1))
        for r in range(r1, r2 + 1)
    )


def rng(values: dict, address: str):
    """``"Sheet1!A1:B3"`` 범위를 2차원 리스트로 만든다."""
    get = values.get
    return [[get(k) for k in row] for row in _range_keys(address)]


def _flatten(x):
    if isinstance(x, list):
        for row in x:
            yield from row
    else:
        yield x


def _shape(x):
    if isinstance(x, list):
        return len(x), (len(x[0]) if x else 0)
    return 1, 1


def _scalar(x):
    """범위를 스칼라 문맥에서 쓸 때. 1x1 만 허용한다."""
    if isinstance(x, list):
        if len(x) == 1 and len(x[0]) == 1:
            return x[0][0]
        return VALUE
    if x is MISSING:
        return None
    return x


def _vector(x):
    """1차원(행 또는 열) 범위를 리스트로."""
    if not isinstance(x, list):
        return [x]
    if len(x) == 1:
        return list(x[0])
    if all(len(r) == 1 for r in x):
        return [r[0] for r in x]
    return None


def _broadcast(fn, *args):
    """배열 인수가 있으면 원소 단위로 적용한다 (엑셀 배열 연산)."""
    if not any(isinstance(a, list) for a in args):
        return fn(*args)
    shapes = [_shape(a) for a in args]
    rows = max(s[0] for s in shapes)
    cols = max(s[1] for s in shapes)

    def pick(a, s, i, j):
        if not isinstance(a, list):
            return a
        ii = 0 if s[0] == 1 else i
        jj = 0 if s[1] == 1 else j
        if ii >= s[0] or jj >= s[1]:
            return NA
        return a[ii][jj]

    return [
        [fn(*(pick(a, s, i, j) for a, s in zip(args, shapes))) for j in range(cols)]
        for i in range(rows)
    ]


# ---------------------------------------------------------------------------
# 형 변환
# ---------------------------------------------------------------------------
def _parse_number_text(s: str):
    t = s.strip()
    if not t:
        return None
    pct = t.endswith("%")
    if pct:
        t = t[:-1]
    t = t.replace(",", "")
    try:
        v = float(t)
    except ValueError:
        return None
    if pct:
        v /= 100
    return int(v) if v.is_integer() and "." not in t and "e" not in t.lower() and not pct else v


def to_number(x):
    x = _scalar(x)
    if x is None:
        return 0
    if isinstance(x, bool):
        return int(x)
    if _is_num(x):
        return x
    if _is_err(x):
        return x
    if isinstance(x, str):
        v = _parse_number_text(x)
        return VALUE if v is None else v
    return VALUE


def _fmt_general(x) -> str:
    if isinstance(x, int):
        return str(x)
    if x == int(x) and abs(x) < 1e15:
        return str(int(x))
    s = format(x, ".15g")
    if "e" in s:
        mant, exp = s.split("e")
        sign = "-" if exp.startswith("-") else "+"
        return f"{mant.upper()}E{sign}{abs(int(exp)):02d}"
    return s


def to_text(x):
    x = _scalar(x)
    if x is None:
        return ""
    if isinstance(x, bool):
        return "TRUE" if x else "FALSE"
    if _is_num(x):
        return _fmt_general(x)
    if _is_err(x):
        return x
    return str(x)


def to_bool(x):
    x = _scalar(x)
    if x is None:
        return False
    if isinstance(x, bool):
        return x
    if _is_num(x):
        return x != 0
    if _is_err(x):
        return x
    if isinstance(x, str):
        u = x.strip().upper()
        if u == "TRUE":
            return True
        if u == "FALSE":
            return False
    return VALUE


def _to_int(x):
    v = to_number(x)
    if _is_err(v):
        return v
    return int(math.floor(v))


def _num_result(v):
    if isinstance(v, float):
        if math.isnan(v) or math.isinf(v):
            return NUM
        if v.is_integer() and abs(v) < 2**53:
            return v
    return v


def normalize(x):
    """셀에 최종 저장될 값으로 정규화한다 (빈 참조 → 0 등)."""
    if isinstance(x, list):
        if len(x) == 1 and len(x[0]) == 1:
            x = x[0][0]
        else:
            return VALUE
    if x is None or x is MISSING:
        return 0
    if isinstance(x, float) and (math.isnan(x) or math.isinf(x)):
        return NUM
    return x


# ---------------------------------------------------------------------------
# 연산자
# ---------------------------------------------------------------------------
def _arith(op):
    def scalar(a, b):
        if _is_err(a):
            return a
        if _is_err(b):
            return b
        x = to_number(a)
        if _is_err(x):
            return x
        y = to_number(b)
        if _is_err(y):
            return y
        return _num_result(op(x, y))

    @_safe
    def wrapper(a, b):
        return _broadcast(scalar, a, b)

    return wrapper


def _div(x, y):
    if y == 0:
        return DIV0
    return x / y


def _pow(x, y):
    if x == 0 and y == 0:
        return NUM
    if x == 0 and y < 0:
        return DIV0
    if x < 0 and not float(y).is_integer():
        return NUM
    r = x**y
    if isinstance(r, complex):
        return NUM
    return r


add = _arith(lambda x, y: x + y)
sub = _arith(lambda x, y: x - y)
mul = _arith(lambda x, y: x * y)
div = _arith(_div)
pow = _arith(_pow)  # noqa: A001 - 엑셀 ^ 연산자


def _neg_s(a):
    x = to_number(a)
    return x if _is_err(x) else -x


def _pct_s(a):
    x = to_number(a)
    return x if _is_err(x) else x / 100


def _pos_s(a):
    return _scalar(a)


@_safe
def neg(a):
    return _broadcast(_neg_s, a)


@_safe
def percent(a):
    return _broadcast(_pct_s, a)


@_safe
def pos(a):
    return _broadcast(_pos_s, a)


def _concat_s(a, b):
    x = to_text(a)
    if _is_err(x):
        return x
    y = to_text(b)
    if _is_err(y):
        return y
    return x + y


@_safe
def concat(a, b):
    return _broadcast(_concat_s, a, b)


def _type_rank(x):
    if isinstance(x, bool):
        return 2
    if isinstance(x, str):
        return 1
    return 0


def _compare(a, b):
    """-1 / 0 / 1 또는 오류값."""
    a = _scalar(a)
    b = _scalar(b)
    if _is_err(a):
        return a
    if _is_err(b):
        return b
    if a is None and b is None:
        return 0
    if a is None:
        a = "" if isinstance(b, str) else (False if isinstance(b, bool) else 0)
    if b is None:
        b = "" if isinstance(a, str) else (False if isinstance(a, bool) else 0)
    ra, rb = _type_rank(a), _type_rank(b)
    if ra != rb:
        return -1 if ra < rb else 1
    if isinstance(a, str):
        a, b = a.lower(), b.lower()
    if a == b:
        return 0
    return -1 if a < b else 1


def _cmp_op(test):
    def scalar(a, b):
        c = _compare(a, b)
        if _is_err(c):
            return c
        return test(c)

    @_safe
    def wrapper(a, b):
        return _broadcast(scalar, a, b)

    return wrapper


eq = _cmp_op(lambda c: c == 0)
ne = _cmp_op(lambda c: c != 0)
lt = _cmp_op(lambda c: c < 0)
le = _cmp_op(lambda c: c <= 0)
gt = _cmp_op(lambda c: c > 0)
ge = _cmp_op(lambda c: c >= 0)


# ---------------------------------------------------------------------------
# 반올림 (엑셀은 은행가 반올림이 아니라 0에서 먼 쪽으로 반올림)
# ---------------------------------------------------------------------------
def _decimal(x) -> Decimal:
    return Decimal(format(x, ".15g"))


def _round(x, digits, mode):
    v = to_number(x)
    if _is_err(v):
        return v
    n = to_number(digits)
    if _is_err(n):
        return n
    n = int(n) if n >= 0 else -int(-n)
    try:
        q = Decimal(1).scaleb(-n)
        r = _decimal(v).quantize(q, rounding=mode)
    except InvalidOperation:
        return NUM
    f = float(r)
    return int(f) if n <= 0 and abs(f) < 2**53 else f


@_safe
def ROUND(x, digits=0):
    return _broadcast(lambda a, b: _round(a, b, ROUND_HALF_UP), x, digits)


@_safe
def ROUNDUP(x, digits=0):
    return _broadcast(lambda a, b: _round(a, b, ROUND_UP), x, digits)


@_safe
def ROUNDDOWN(x, digits=0):
    return _broadcast(lambda a, b: _round(a, b, ROUND_DOWN), x, digits)


def _int_s(x):
    v = to_number(x)
    if _is_err(v):
        return v
    return int(math.floor(v))


@_safe
def INT(x):
    return _broadcast(_int_s, x)


def _abs_s(x):
    v = to_number(x)
    return v if _is_err(v) else abs(v)


@_safe
def ABS(x):
    return _broadcast(_abs_s, x)


def _mod_s(a, b):
    x = to_number(a)
    if _is_err(x):
        return x
    y = to_number(b)
    if _is_err(y):
        return y
    if y == 0:
        return DIV0
    r = x - y * math.floor(x / y)
    return _num_result(r)


@_safe
def MOD(a, b):
    return _broadcast(_mod_s, a, b)


# ---------------------------------------------------------------------------
# 집계
# ---------------------------------------------------------------------------
def _numbers(args, *, count_text_direct=True):
    """집계 함수용 숫자 목록. 범위 안의 텍스트/논리값/빈 셀은 무시하고,
    직접 인수는 변환한다. 오류는 즉시 반환."""
    out = []
    for a in args:
        if a is MISSING:
            continue
        if isinstance(a, list):
            for x in _flatten(a):
                if _is_err(x):
                    return x
                if _is_num(x):
                    out.append(x)
        else:
            if _is_err(a):
                return a
            if a is None:
                continue
            v = to_number(a)
            if _is_err(v):
                if count_text_direct:
                    return v
                continue
            out.append(v)
    return out


@_safe
def SUM(*args):
    nums = _numbers(args)
    if _is_err(nums):
        return nums
    return _num_result(math.fsum(nums)) if any(isinstance(n, float) for n in nums) else sum(nums)


@_safe
def AVERAGE(*args):
    nums = _numbers(args)
    if _is_err(nums):
        return nums
    if not nums:
        return DIV0
    return _num_result(math.fsum(nums) / len(nums))


@_safe
def MIN(*args):
    nums = _numbers(args)
    if _is_err(nums):
        return nums
    return min(nums) if nums else 0


@_safe
def MAX(*args):
    nums = _numbers(args)
    if _is_err(nums):
        return nums
    return max(nums) if nums else 0


@_safe
def COUNT(*args):
    n = 0
    for a in args:
        if isinstance(a, list):
            n += sum(1 for x in _flatten(a) if _is_num(x))
        elif a is not MISSING and a is not None and not _is_err(a):
            if _is_num(a) or isinstance(a, bool) or (isinstance(a, str) and _parse_number_text(a) is not None):
                n += 1
    return n


@_safe
def COUNTA(*args):
    n = 0
    for a in args:
        if isinstance(a, list):
            n += sum(1 for x in _flatten(a) if x is not None)
        elif a is not MISSING:
            n += 1
    return n


@_safe
def SUMPRODUCT(*arrays):
    if not arrays:
        return VALUE
    shapes = {_shape(a) for a in arrays}
    if len(shapes) != 1:
        return VALUE
    flat = [list(_flatten(a)) for a in arrays]
    total = 0.0
    any_float = False
    for items in zip(*flat):
        prod = 1
        for x in items:
            if _is_err(x):
                return x
            if not _is_num(x):
                prod = 0
                continue
            if isinstance(x, float):
                any_float = True
            prod *= x
        total += prod
    if not any_float and float(total).is_integer():
        return int(total)
    return _num_result(total)


# ---------------------------------------------------------------------------
# 조건 (SUMIF / COUNTIF 계열)
# ---------------------------------------------------------------------------
def _wildcard_regex(pattern: str):
    out = []
    i = 0
    while i < len(pattern):
        ch = pattern[i]
        if ch == "~" and i + 1 < len(pattern) and pattern[i + 1] in "*?~":
            out.append(re.escape(pattern[i + 1]))
            i += 2
            continue
        if ch == "*":
            out.append(".*")
        elif ch == "?":
            out.append(".")
        else:
            out.append(re.escape(ch))
        i += 1
    return re.compile("".join(out), re.IGNORECASE | re.DOTALL)


def _has_wildcard(s: str) -> bool:
    return any(c in s for c in "*?")


def _criteria(crit):
    """조건 값을 판정 함수로 바꾼다."""
    crit = _scalar(crit)
    if _is_err(crit):
        return crit
    if crit is None or crit is MISSING:
        crit = 0 if crit is MISSING else ""
    if isinstance(crit, bool):
        return lambda x: isinstance(x, bool) and x == crit
    if _is_num(crit):
        def num_eq(x, c=crit):
            if _is_num(x):
                return x == c
            if isinstance(x, str):
                v = _parse_number_text(x)
                return v is not None and v == c
            return False
        return num_eq

    s = str(crit)
    op = "="
    for cand in ("<=", ">=", "<>", "<", ">", "="):
        if s.startswith(cand):
            op, s = cand, s[len(cand):]
            break
    num = _parse_number_text(s)
    if s == "":
        if op == "=":
            return lambda x: x is None or x == ""
        if op == "<>":
            return lambda x: not (x is None or x == "")
    if num is not None:
        tests = {
            "=": lambda v: v == num, "<>": lambda v: v != num,
            "<": lambda v: v < num, ">": lambda v: v > num,
            "<=": lambda v: v <= num, ">=": lambda v: v >= num,
        }
        t = tests[op]
        if op == "=":
            return lambda x: (_is_num(x) and t(x)) or (
                isinstance(x, str) and _parse_number_text(x) == num)
        if op == "<>":
            return lambda x: not ((_is_num(x) and x == num) or (
                isinstance(x, str) and _parse_number_text(x) == num))
        return lambda x: _is_num(x) and t(x)
    up = s.upper()
    if up in ("TRUE", "FALSE") and op in ("=", "<>"):
        b = up == "TRUE"
        if op == "=":
            return lambda x: isinstance(x, bool) and x == b
        return lambda x: not (isinstance(x, bool) and x == b)
    if op in ("=", "<>"):
        rx = _wildcard_regex(s)
        if op == "=":
            return lambda x: isinstance(x, str) and rx.fullmatch(x) is not None
        return lambda x: not (isinstance(x, str) and rx.fullmatch(x) is not None)
    low = s.lower()
    tests = {
        "<": lambda v: v < low, ">": lambda v: v > low,
        "<=": lambda v: v <= low, ">=": lambda v: v >= low,
    }
    t = tests[op]
    return lambda x: isinstance(x, str) and t(x.lower())


def _as_grid(x):
    if isinstance(x, list):
        return x
    return [[x]]


def _cell_at(grid, i, j):
    if i < len(grid) and j < len(grid[i]):
        return grid[i][j]
    return None


def _ifs_mask(pairs):
    """[(range, criteria), ...] → (rows, cols, mask) 또는 오류."""
    rows = cols = None
    preds = []
    for r, c in pairs:
        g = _as_grid(r)
        shp = _shape(g)
        if rows is None:
            rows, cols = shp
        elif shp != (rows, cols):
            return VALUE
        p = _criteria(c)
        if _is_err(p):
            return p
        preds.append((g, p))
    mask = [[all(p(g[i][j]) for g, p in preds) for j in range(cols)] for i in range(rows)]
    return rows, cols, mask


def _agg_masked(target, mask_res, kind):
    if _is_err(mask_res):
        return mask_res
    rows, cols, mask = mask_res
    tgrid = _as_grid(target) if target is not None else None
    vals = []
    count = 0
    for i in range(rows):
        for j in range(cols):
            if not mask[i][j]:
                continue
            count += 1
            if tgrid is None:
                continue
            x = _cell_at(tgrid, i, j)
            if _is_err(x):
                return x
            if _is_num(x):
                vals.append(x)
    if kind == "count":
        return count
    if kind == "sum":
        s = math.fsum(vals) if any(isinstance(v, float) for v in vals) else sum(vals)
        return _num_result(s)
    if not vals:
        return DIV0
    return _num_result(math.fsum(vals) / len(vals))


@_safe
def SUMIF(range_, criteria, sum_range=MISSING):
    target = range_ if sum_range is MISSING or sum_range is None else sum_range
    return _agg_masked(target, _ifs_mask([(range_, criteria)]), "sum")


@_safe
def SUMIFS(sum_range, *pairs):
    if len(pairs) % 2:
        return VALUE
    m = _ifs_mask(list(zip(pairs[::2], pairs[1::2])))
    if not _is_err(m) and _shape(_as_grid(sum_range)) != (m[0], m[1]):
        return VALUE
    return _agg_masked(sum_range, m, "sum")


@_safe
def COUNTIF(range_, criteria):
    return _agg_masked(None, _ifs_mask([(range_, criteria)]), "count")


@_safe
def COUNTIFS(*pairs):
    if not pairs or len(pairs) % 2:
        return VALUE
    return _agg_masked(None, _ifs_mask(list(zip(pairs[::2], pairs[1::2]))), "count")


@_safe
def AVERAGEIF(range_, criteria, average_range=MISSING):
    target = range_ if average_range is MISSING or average_range is None else average_range
    return _agg_masked(target, _ifs_mask([(range_, criteria)]), "avg")


@_safe
def AVERAGEIFS(average_range, *pairs):
    if len(pairs) % 2:
        return VALUE
    m = _ifs_mask(list(zip(pairs[::2], pairs[1::2])))
    if not _is_err(m) and _shape(_as_grid(average_range)) != (m[0], m[1]):
        return VALUE
    return _agg_masked(average_range, m, "avg")


# ---------------------------------------------------------------------------
# 논리
# ---------------------------------------------------------------------------
def _if_s(cond, a, b):
    c = to_bool(cond)
    if _is_err(c):
        return c
    r = a if c else b
    return 0 if r is MISSING else r


@_safe
def IF(cond, value_if_true=True, value_if_false=False):
    if isinstance(cond, list):
        return _broadcast(_if_s, cond, value_if_true, value_if_false)
    return _if_s(cond, value_if_true, value_if_false)


def _iferror_s(v, alt):
    if _is_err(v):
        return 0 if alt is MISSING else alt
    return v


@_safe
def IFERROR(value, value_if_error):
    if isinstance(value, list):
        return _broadcast(_iferror_s, value, value_if_error)
    return _iferror_s(value, value_if_error)


@_safe
def IFS(*args):
    if not args or len(args) % 2:
        return VALUE
    for cond, val in zip(args[::2], args[1::2]):
        c = to_bool(cond)
        if _is_err(c):
            return c
        if c:
            return 0 if val is MISSING else val
    return NA


def _logic_values(args):
    vals = []
    for a in args:
        if a is MISSING:
            continue
        if isinstance(a, list):
            for x in _flatten(a):
                if _is_err(x):
                    return x
                if isinstance(x, bool):
                    vals.append(x)
                elif _is_num(x):
                    vals.append(x != 0)
        else:
            if a is None:
                continue
            b = to_bool(a)
            if _is_err(b):
                return b
            vals.append(b)
    if not vals:
        return VALUE
    return vals


@_safe
def AND(*args):
    v = _logic_values(args)
    return v if _is_err(v) else all(v)


@_safe
def OR(*args):
    v = _logic_values(args)
    return v if _is_err(v) else any(v)


def _not_s(x):
    b = to_bool(x)
    return b if _is_err(b) else (not b)


@_safe
def NOT(x):
    return _broadcast(_not_s, x)


@_safe
def CHOOSE(index, *values):
    i = _to_int(index)
    if _is_err(i):
        return i
    if i < 1 or i > len(values):
        return VALUE
    v = values[i - 1]
    return 0 if v is MISSING else v


# ---------------------------------------------------------------------------
# 정보
# ---------------------------------------------------------------------------
@_safe
def ISBLANK(x):
    return _broadcast(lambda v: v is None, x)


@_safe
def ISNUMBER(x):
    return _broadcast(_is_num, x)


@_safe
def ISTEXT(x):
    return _broadcast(lambda v: isinstance(v, str), x)


@_safe
def ISERROR(x):
    return _broadcast(_is_err, x)


@_safe
def ISNA(x):
    return _broadcast(lambda v: v == NA, x)


@_safe
def VALUE_(x):
    v = _scalar(x)
    if _is_err(v):
        return v
    if v is None:
        return 0
    if _is_num(v):
        return v
    if isinstance(v, str):
        n = _parse_number_text(v)
        return VALUE if n is None else n
    return VALUE


# ---------------------------------------------------------------------------
# 조회
# ---------------------------------------------------------------------------
def _lookup_key(x):
    if x is None:
        return 0
    return x


def _lookup_eq(item, val):
    if _is_err(item):
        return False
    item = _lookup_key(item)
    if isinstance(val, str):
        if not isinstance(item, str):
            return False
        if _has_wildcard(val):
            return _wildcard_regex(val).fullmatch(item) is not None
        return item.lower() == val.lower()
    if isinstance(val, bool) or isinstance(item, bool):
        return isinstance(val, bool) and isinstance(item, bool) and val == item
    if _is_num(val) and _is_num(item):
        return item == val
    return False


def _same_type(a, b):
    return _type_rank(_lookup_key(a)) == _type_rank(_lookup_key(b))


def _approx_pos(items, val, descending=False):
    """정렬된 목록에서 근사 일치 위치 (0-based) 또는 None."""
    best = None
    for i, item in enumerate(items):
        if item is None or _is_err(item) or not _same_type(item, val):
            continue
        c = _compare(item, val)
        if _is_err(c):
            continue
        if not descending:
            if c <= 0:
                best = i
            else:
                break
        else:
            if c >= 0:
                best = i
            else:
                break
    return best


def _exact_pos(items, val, reverse=False):
    order = range(len(items) - 1, -1, -1) if reverse else range(len(items))
    for i in order:
        if _lookup_eq(items[i], val):
            return i
    return None


@_safe
def VLOOKUP(lookup_value, table, col_index, range_lookup=True):
    val = _scalar(lookup_value)
    if _is_err(val):
        return val
    val = _lookup_key(val)
    grid = _as_grid(table)
    col = _to_int(col_index)
    if _is_err(col):
        return col
    if col < 1:
        return VALUE
    if col > (len(grid[0]) if grid else 0):
        return REF
    approx = True if range_lookup is MISSING else to_bool(range_lookup)
    if range_lookup is None:
        approx = False
    if _is_err(approx):
        return approx
    first = [row[0] for row in grid]
    pos = _approx_pos(first, val) if approx else _exact_pos(first, val)
    if pos is None:
        return NA
    return grid[pos][col - 1]


@_safe
def HLOOKUP(lookup_value, table, row_index, range_lookup=True):
    grid = _as_grid(table)
    transposed = [list(r) for r in zip(*grid)]
    return VLOOKUP(lookup_value, transposed, row_index, range_lookup)


@_safe
def MATCH(lookup_value, lookup_array, match_type=1):
    val = _scalar(lookup_value)
    if _is_err(val):
        return val
    val = _lookup_key(val)
    items = _vector(lookup_array)
    if items is None:
        return NA
    mt = 1 if match_type is MISSING else to_number(match_type)
    if _is_err(mt):
        return mt
    if mt == 0:
        pos = _exact_pos(items, val)
    elif mt > 0:
        pos = _approx_pos(items, val)
    else:
        pos = _approx_pos(items, val, descending=True)
    return NA if pos is None else pos + 1


@_safe
def INDEX(array, row_num, col_num=MISSING):
    grid = _as_grid(array)
    rows, cols = _shape(grid)
    r = 0 if row_num is MISSING else _to_int(row_num)
    if _is_err(r):
        return r
    if col_num is MISSING:
        if rows == 1 and cols > 1:
            r, c = 1, r
        else:
            c = 1 if cols == 1 else 0
    else:
        c = _to_int(col_num)
        if _is_err(c):
            return c
    if r < 0 or c < 0 or r > rows or c > cols:
        return REF
    if r == 0 and c == 0:
        return grid
    if r == 0:
        return [[grid[i][c - 1]] for i in range(rows)]
    if c == 0:
        return [list(grid[r - 1])]
    return grid[r - 1][c - 1]


@_safe
def XLOOKUP(lookup_value, lookup_array, return_array, if_not_found=MISSING,
            match_mode=0, search_mode=1):
    val = _scalar(lookup_value)
    if _is_err(val):
        return val
    val = _lookup_key(val)
    items = _vector(lookup_array)
    if items is None:
        return VALUE
    mm = 0 if match_mode is MISSING else _to_int(match_mode)
    sm = 1 if search_mode is MISSING else _to_int(search_mode)
    if _is_err(mm):
        return mm
    if _is_err(sm):
        return sm
    reverse = sm < 0
    pos = None
    if mm in (0, 2):
        if mm == 0 and isinstance(val, str):
            order = range(len(items) - 1, -1, -1) if reverse else range(len(items))
            for i in order:
                it = items[i]
                if isinstance(it, str) and it.lower() == val.lower():
                    pos = i
                    break
        else:
            pos = _exact_pos(items, val, reverse)
    elif mm in (-1, 1):
        pos = _exact_pos(items, val, reverse)
        if pos is None:
            best = None
            for i, it in enumerate(items):
                if it is None or _is_err(it) or not _same_type(it, val):
                    continue
                c = _compare(it, val)
                if (mm == -1 and c < 0) or (mm == 1 and c > 0):
                    if best is None:
                        best = i
                    else:
                        cb = _compare(it, items[best])
                        if (mm == -1 and cb > 0) or (mm == 1 and cb < 0):
                            best = i
            pos = best
    else:
        return VALUE
    if pos is None:
        return NA if if_not_found is MISSING else if_not_found
    grid = _as_grid(return_array)
    rows, cols = _shape(grid)
    if isinstance(lookup_array, list) and len(lookup_array) == 1 and _shape(lookup_array)[1] > 1:
        # 가로 방향 조회 → 열 반환
        if pos >= cols:
            return VALUE
        col = [[grid[i][pos]] for i in range(rows)]
        return col if rows > 1 else col[0][0]
    if pos >= rows:
        return VALUE
    row = list(grid[pos])
    return [row] if cols > 1 else row[0]


# ---------------------------------------------------------------------------
# 텍스트
# ---------------------------------------------------------------------------
def _text_args(*args):
    out = []
    for a in args:
        t = to_text(a)
        if _is_err(t):
            return t
        out.append(t)
    return out


def _left_s(text, n=1):
    t = to_text(text)
    if _is_err(t):
        return t
    k = 1 if n is MISSING else _to_int(n)
    if _is_err(k):
        return k
    if k < 0:
        return VALUE
    return t[:k]


def _right_s(text, n=1):
    t = to_text(text)
    if _is_err(t):
        return t
    k = 1 if n is MISSING else _to_int(n)
    if _is_err(k):
        return k
    if k < 0:
        return VALUE
    return t[len(t) - k:] if k else ""


def _mid_s(text, start, n):
    t = to_text(text)
    if _is_err(t):
        return t
    s = _to_int(start)
    if _is_err(s):
        return s
    k = _to_int(n)
    if _is_err(k):
        return k
    if s < 1 or k < 0:
        return VALUE
    return t[s - 1:s - 1 + k]


@_safe
def LEFT(text, n=1):
    return _broadcast(_left_s, text, n)


@_safe
def RIGHT(text, n=1):
    return _broadcast(_right_s, text, n)


@_safe
def MID(text, start, n):
    return _broadcast(_mid_s, text, start, n)


def _len_s(text):
    t = to_text(text)
    return t if _is_err(t) else len(t)


@_safe
def LEN(text):
    return _broadcast(_len_s, text)


def _trim_s(text):
    t = to_text(text)
    return t if _is_err(t) else re.sub(r" +", " ", t.strip(" "))


@_safe
def TRIM(text):
    return _broadcast(_trim_s, text)


def _upper_s(text):
    t = to_text(text)
    return t if _is_err(t) else t.upper()


def _lower_s(text):
    t = to_text(text)
    return t if _is_err(t) else t.lower()


@_safe
def UPPER(text):
    return _broadcast(_upper_s, text)


@_safe
def LOWER(text):
    return _broadcast(_lower_s, text)


@_safe
def CONCATENATE(*args):
    parts = _text_args(*[_scalar(a) for a in args if a is not MISSING])
    return parts if _is_err(parts) else "".join(parts)


@_safe
def CONCAT(*args):
    out = []
    for a in args:
        if a is MISSING:
            continue
        for x in _flatten(a):
            t = to_text(x)
            if _is_err(t):
                return t
            out.append(t)
    return "".join(out)


# ---------------------------------------------------------------------------
# 날짜 (1900 체계, 1900-02-29 버그 포함)
# ---------------------------------------------------------------------------
_EPOCH = _dt.date(1899, 12, 31)


def date_to_serial(y: int, m: int, d: int) -> int:
    """정상 날짜 → 엑셀 일련번호."""
    n = (_dt.date(y, m, d) - _EPOCH).days
    if (y, m, d) >= (1900, 3, 1):
        n += 1
    return n


def serial_to_ymd(serial):
    """엑셀 일련번호 → (년, 월, 일). 0 → 1900-01-00, 60 → 1900-02-29."""
    n = int(math.floor(serial))
    if n < 0:
        raise ValueError("negative serial")
    if n == 0:
        return 1900, 1, 0
    if n == 60:
        return 1900, 2, 29
    if n > 60:
        n -= 1
    d = _EPOCH + _dt.timedelta(days=n)
    return d.year, d.month, d.day


def _serial_arg(x):
    v = to_number(x)
    if _is_err(v):
        return v
    if v < 0 or v >= 2958466:
        return NUM
    return v


@_safe
def DATE(year, month, day):
    def s(y, m, d):
        y, m, d = _to_int(y), _to_int(m), _to_int(d)
        for v in (y, m, d):
            if _is_err(v):
                return v
        if 0 <= y < 1900:
            y += 1900
        if y < 0 or y >= 10000:
            return NUM
        y += (m - 1) // 12
        m = (m - 1) % 12 + 1
        if y < 1900:
            return NUM
        serial = date_to_serial(y, m, 1) + d - 1
        if serial < 0:
            return NUM
        return serial

    return _broadcast(s, year, month, day)


def _ymd_part(idx):
    def scalar(x):
        v = _serial_arg(x)
        if _is_err(v):
            return v
        return serial_to_ymd(v)[idx]

    @_safe
    def fn(x):
        return _broadcast(scalar, x)

    return fn


YEAR = _ymd_part(0)
MONTH = _ymd_part(1)
DAY = _ymd_part(2)


def _add_months(serial, months, end_of_month):
    v = _serial_arg(serial)
    if _is_err(v):
        return v
    k = _to_int(months) if not _is_num(months) else int(months)
    if _is_err(k):
        return k
    y, m, d = serial_to_ymd(v)
    if d == 0:
        y, m, d = 1899, 12, 31
    total = y * 12 + (m - 1) + k
    ny, nm = divmod(total, 12)
    nm += 1
    if ny < 1900:
        return NUM
    last = calendar.monthrange(ny, nm)[1]
    nd = last if end_of_month else min(d, last)
    return date_to_serial(ny, nm, nd)


@_safe
def EDATE(start_date, months):
    return _broadcast(lambda s, k: _add_months(s, k, False), start_date, months)


@_safe
def EOMONTH(start_date, months):
    return _broadcast(lambda s, k: _add_months(s, k, True), start_date, months)


# ---------------------------------------------------------------------------
# TEXT (기본 서식만)
# ---------------------------------------------------------------------------
_MONTHS = ["January", "February", "March", "April", "May", "June", "July",
           "August", "September", "October", "November", "December"]
_DAYS = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]
_DATE_TOKEN = re.compile(r"yyyy|yy|mmmm|mmm|mm|m|dddd|ddd|dd|d|hh|h|ss|s|AM/PM|\"[^\"]*\"|.", re.I)


def _format_date(serial, fmt):
    y, mo, d = serial_to_ymd(serial)
    frac = serial - math.floor(serial)
    secs = int(round(frac * 86400))
    hh, rem = divmod(secs, 3600)
    mi, ss = divmod(rem, 60)
    toks = _DATE_TOKEN.findall(fmt)
    ampm = any(t.upper() == "AM/PM" for t in toks)
    out = []
    for i, t in enumerate(toks):
        low = t.lower()
        prev = next((x.lower() for x in reversed(toks[:i]) if x.strip() and x[0].isalpha()), "")
        nxt = next((x.lower() for x in toks[i + 1:] if x.strip() and x[0].isalpha()), "")
        if low in ("m", "mm") and (prev.startswith("h") or nxt.startswith("s")):
            out.append(f"{mi:02d}" if low == "mm" else str(mi))
        elif low == "yyyy":
            out.append(f"{y:04d}")
        elif low == "yy":
            out.append(f"{y % 100:02d}")
        elif low == "mmmm":
            out.append(_MONTHS[mo - 1])
        elif low == "mmm":
            out.append(_MONTHS[mo - 1][:3])
        elif low == "mm":
            out.append(f"{mo:02d}")
        elif low == "m":
            out.append(str(mo))
        elif low in ("dddd", "ddd"):
            if d == 0:
                wd = 5
            else:
                wd = _dt.date(y, mo, d).weekday() if (y, mo, d) != (1900, 2, 29) else 2
            out.append(_DAYS[wd] if low == "dddd" else _DAYS[wd][:3])
        elif low == "dd":
            out.append(f"{d:02d}")
        elif low == "d":
            out.append(str(d))
        elif low in ("hh", "h"):
            h = hh % 12 or 12 if ampm else hh
            out.append(f"{h:02d}" if low == "hh" else str(h))
        elif low in ("ss", "s"):
            out.append(f"{ss:02d}" if low == "ss" else str(ss))
        elif low == "am/pm":
            out.append("AM" if hh < 12 else "PM")
        elif t.startswith('"'):
            out.append(t[1:-1])
        else:
            out.append(t)
    return "".join(out)


def _format_number(v, fmt):
    lit_prefix, lit_suffix, core = "", "", ""
    # 따옴표 리터럴과 숫자 자리표시자를 분리한다
    tokens = re.findall(r'"[^"]*"|\\.|[0#?.,%]+|.', fmt)
    started = ended = False
    for t in tokens:
        lit = t[1:-1] if t.startswith('"') else (t[1:] if t.startswith("\\") else None)
        if lit is None and re.fullmatch(r"[0#?.,%]+", t) and not ended:
            core += t
            started = True
        else:
            text = lit if lit is not None else t
            if started:
                ended = True
                lit_suffix += text
            else:
                lit_prefix += text
    if not core:
        return lit_prefix + lit_suffix
    pct = core.count("%")
    core = core.replace("%", "")
    if pct:
        v = v * (100 ** pct)
    scale = 0
    while core.endswith(","):
        core = core[:-1]
        scale += 1
    v = v / (1000 ** scale)
    int_part, _, dec_part = core.partition(".")
    decimals = len(dec_part)
    thousands = "," in int_part
    min_int = int_part.count("0")
    q = Decimal(1).scaleb(-decimals)
    r = _decimal(abs(v)).quantize(q, rounding=ROUND_HALF_UP)
    s = format(r, "f")
    ip, _, dp = s.partition(".")
    if ip == "0" and min_int == 0:
        ip = ""
    ip = ip.rjust(min_int, "0")
    if thousands and ip:
        ip = f"{int(ip):,}".rjust(min_int, "0")
    num = ip + ("." + dp if decimals else "")
    sign = "-" if v < 0 and float(r) != 0 else ""
    return sign + lit_prefix + num + "%" * pct + lit_suffix


def _text_s(value, fmt):
    v = _scalar(value)
    f = to_text(fmt)
    if _is_err(v):
        return v
    if _is_err(f):
        return f
    sections = re.split(r';(?=(?:[^"]*"[^"]*")*[^"]*$)', f)
    num = None
    if _is_num(v) or isinstance(v, bool) or v is None:
        num = to_number(v) if not isinstance(v, bool) else None
    elif isinstance(v, str):
        num = _parse_number_text(v)
    if num is None:
        if len(sections) >= 4:
            return sections[3].replace("@", to_text(v)).replace('"', "")
        if "@" in f:
            return f.replace("@", to_text(v)).replace('"', "")
        return to_text(v)
    sec = sections[0]
    if len(sections) >= 2 and num < 0:
        sec, num = sections[1], -num
    elif len(sections) >= 3 and num == 0:
        sec = sections[2]
    low = sec.lower()
    if low in ("general", "g/표준", ""):
        return _fmt_general(num)
    if low == "@":
        return to_text(v)
    stripped = re.sub(r'"[^"]*"', "", low)
    if re.search(r"[ymdhs]", stripped) and not re.search(r"[0#]", stripped):
        if num < 0:
            return VALUE
        return _format_date(num, sec)
    return _format_number(num, sec)


@_safe
def TEXT(value, format_text):
    return _broadcast(_text_s, value, format_text)


# ---------------------------------------------------------------------------
# 비교 (검증 엔진 / 생성된 pytest 공용)
# ---------------------------------------------------------------------------
def values_equal(expected, actual, rel_tol=1e-9, abs_tol=1e-9) -> bool:
    """엑셀 저장값(expected)과 계산값(actual)이 같은지 판정한다."""
    if isinstance(expected, list) or isinstance(actual, list):
        return False
    if _is_err(expected) or _is_err(actual):
        return expected == actual
    if expected in (None, "") and actual in (None, ""):
        return True
    if isinstance(expected, bool) or isinstance(actual, bool):
        return isinstance(expected, bool) and isinstance(actual, bool) and expected == actual
    if _is_num(expected) and _is_num(actual):
        return math.isclose(expected, actual, rel_tol=rel_tol, abs_tol=abs_tol)
    if isinstance(expected, str) and isinstance(actual, str):
        return expected == actual
    return False


# ---------------------------------------------------------------------------
# 함수 레지스트리 (엑셀 이름 → 파이썬 함수)
# ---------------------------------------------------------------------------
FUNCTIONS = {
    "SUM": SUM, "SUMIF": SUMIF, "SUMIFS": SUMIFS, "SUMPRODUCT": SUMPRODUCT,
    "COUNT": COUNT, "COUNTA": COUNTA, "COUNTIF": COUNTIF, "COUNTIFS": COUNTIFS,
    "AVERAGE": AVERAGE, "AVERAGEIF": AVERAGEIF, "AVERAGEIFS": AVERAGEIFS,
    "MIN": MIN, "MAX": MAX,
    "ROUND": ROUND, "ROUNDUP": ROUNDUP, "ROUNDDOWN": ROUNDDOWN,
    "INT": INT, "ABS": ABS, "MOD": MOD,
    "IF": IF, "IFS": IFS, "IFERROR": IFERROR, "AND": AND, "OR": OR, "NOT": NOT,
    "CHOOSE": CHOOSE,
    "VLOOKUP": VLOOKUP, "HLOOKUP": HLOOKUP, "INDEX": INDEX, "MATCH": MATCH,
    "XLOOKUP": XLOOKUP,
    "LEFT": LEFT, "RIGHT": RIGHT, "MID": MID, "LEN": LEN, "TRIM": TRIM,
    "UPPER": UPPER, "LOWER": LOWER, "CONCATENATE": CONCATENATE, "CONCAT": CONCAT,
    "TEXT": TEXT,
    "DATE": DATE, "YEAR": YEAR, "MONTH": MONTH, "DAY": DAY,
    "EDATE": EDATE, "EOMONTH": EOMONTH,
    "ISBLANK": ISBLANK, "ISNUMBER": ISNUMBER, "ISTEXT": ISTEXT,
    "ISERROR": ISERROR, "ISNA": ISNA, "VALUE": VALUE_,
}

#: 생성 코드에서 사용할 파이썬 이름 (엑셀 이름과 다른 경우만)
PY_NAMES = {"VALUE": "VALUE_"}

#: 셀 참조 인수를 "범위"로 넘겨야 의미가 맞는 함수
#: (예: SUM(A1) 에서 A1 이 텍스트면 무시되지만 SUM("3") 은 3)
RANGE_ARG_FUNCTIONS = frozenset({
    "SUM", "COUNT", "COUNTA", "AVERAGE", "MIN", "MAX", "SUMPRODUCT", "AND", "OR",
    "CONCAT", "SUMIF", "SUMIFS", "COUNTIF", "COUNTIFS", "AVERAGEIF", "AVERAGEIFS",
    "VLOOKUP", "HLOOKUP", "INDEX", "MATCH", "XLOOKUP",
})
