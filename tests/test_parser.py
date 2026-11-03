import pytest

from sheetproof import parser as P

ROUND_TRIP = [
    "=A1+B2*C3",
    "=(A1+B2)*C3",
    "=-2^2",
    "=2^3^2",
    "=50%*A1",
    "=SUM(A1:B10,Sheet2!C3,'My Sheet'!$D$4:$E$9)",
    '=IF(A1>=10,"big","small")',
    '=A1&" - "&B1',
    "=VLOOKUP(A2,Table!$A$2:$D$6,2,)",
    "={1,2;3,4}",
    "=IFERROR(1/0,#N/A)",
    "=TRUE<>FALSE",
    "=A:A",
    "=SUM(1:3)",
    "=TaxRate*A1",
    "=-(A1+B1)%",
    '="He said ""hi"""',
    "=TODAY()",
]


@pytest.mark.parametrize("formula", ROUND_TRIP)
def test_round_trip(formula):
    ast1 = P.parse_formula(formula)
    text = P.to_formula(ast1)
    assert P.parse_formula(text) == ast1


def test_precedence_unary_binds_tighter_than_power():
    ast = P.parse_formula("=-2^2")
    assert isinstance(ast, P.Binary) and ast.op == "^"
    assert isinstance(ast.left, P.Unary)


def test_power_is_left_associative():
    ast = P.parse_formula("=2^3^2")
    assert ast.op == "^" and isinstance(ast.left, P.Binary)


def test_mul_before_add_and_concat_last():
    ast = P.parse_formula('=1+2*3&"x"')
    assert ast.op == "&"
    assert ast.left.op == "+" and ast.left.right.op == "*"


def test_comparison_lowest():
    ast = P.parse_formula("=A1&B1=C1")
    assert ast.op == "="


def test_sheet_reference_parsing():
    ast = P.parse_formula("='It''s'!A1:B2")
    assert ast == P.Ref("It's", "A1:B2")
    ast = P.parse_formula("=단가표!$A$2")
    assert ast == P.Ref("단가표", "$A$2")


def test_function_prefix_stripped():
    ast = P.parse_formula("=_xlfn.XLOOKUP(1,A1:A3,B1:B3)")
    assert ast.name == "XLOOKUP"


def test_missing_argument():
    ast = P.parse_formula("=IF(A1,,2)")
    assert isinstance(ast.args[1], P.Missing)


def test_names_and_external():
    assert P.parse_formula("=Rate") == P.Name("Rate")
    ext = P.parse_formula("=[1]Sheet1!A1")
    assert isinstance(ext, P.Ref) and ext.external


def test_array_constant():
    ast = P.parse_formula("={1,2;3,4}")
    assert ast == P.Array(((P.Number(1), P.Number(2)), (P.Number(3), P.Number(4))))


@pytest.mark.parametrize("bad", ["=SUM(1,2", "=1+", "=(1+2"])
def test_parse_errors(bad):
    with pytest.raises(P.FormulaParseError):
        P.parse_formula(bad)
