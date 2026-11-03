"""엑셀 의미론 단위 테스트. 기대값은 Excel 동작 기준으로 직접 작성했다."""

import pytest

from sheetproof import xl_runtime as xl

R = [[1], [2], [None], ["3"], [True]]  # 범위: 숫자, 빈 셀, 텍스트 숫자, 논리값


class TestBlankAndCoercion:
    def test_blank_is_zero_in_arithmetic(self):
        assert xl.add(None, 5) == 5
        assert xl.mul(None, 3) == 0

    def test_blank_is_empty_string_in_concat(self):
        assert xl.concat(None, "a") == "a"

    def test_numeric_text_coerces(self):
        assert xl.add("3", 1) == 4
        assert xl.add(" 2.5 ", 1) == 3.5

    def test_non_numeric_text_is_value_error(self):
        assert xl.add("abc", 1) == xl.VALUE

    def test_bool_arithmetic(self):
        assert xl.add(True, 1) == 2

    def test_normalize_blank_reference_is_zero(self):
        assert xl.normalize(None) == 0
        assert xl.normalize([[None]]) == 0

    def test_general_number_text(self):
        assert xl.concat("v=", 1 / 3) == "v=0.333333333333333"
        assert xl.concat(1.0, "") == "1"
        assert xl.concat(True, "") == "TRUE"


class TestErrors:
    def test_div_zero(self):
        assert xl.div(1, 0) == xl.DIV0
        assert xl.div(1, None) == xl.DIV0

    def test_left_error_wins(self):
        assert xl.add(xl.NA, xl.DIV0) == xl.NA

    def test_error_propagates_through_functions(self):
        assert xl.SUM([[1], [xl.NA]]) == xl.NA
        assert xl.ROUND(xl.DIV0, 2) == xl.DIV0
        assert xl.LEN(xl.VALUE) == xl.VALUE

    def test_iferror(self):
        assert xl.IFERROR(xl.DIV0, "x") == "x"
        assert xl.IFERROR(5, "x") == 5

    def test_if_does_not_propagate_unused_branch(self):
        assert xl.IF(True, 1, xl.DIV0) == 1

    def test_if_error_condition(self):
        assert xl.IF(xl.NA, 1, 2) == xl.NA

    def test_power_errors(self):
        assert xl.pow(0, 0) == xl.NUM
        assert xl.pow(-8, 0.5) == xl.NUM


class TestRound:
    @pytest.mark.parametrize("x,n,expected", [
        (2.5, 0, 3), (-2.5, 0, -3), (0.5, 0, 1), (1.5, 0, 2),  # 은행가 반올림 아님
        (2.675, 2, 2.68), (1.005, 2, 1.01), (1234.5678, -2, 1200), (-1.25, 1, -1.3),
    ])
    def test_round_half_away_from_zero(self, x, n, expected):
        assert xl.ROUND(x, n) == expected

    def test_roundup_rounddown(self):
        assert xl.ROUNDUP(3.14159, 3) == 3.142
        assert xl.ROUNDUP(-3.14159, 1) == -3.2
        assert xl.ROUNDDOWN(-3.789, 1) == -3.7
        assert xl.ROUNDDOWN(3.999, 0) == 3

    def test_int_mod(self):
        assert xl.INT(-7.5) == -8
        assert xl.MOD(-7.5, 4) == 0.5
        assert xl.MOD(7, -3) == -2
        assert xl.MOD(1, 0) == xl.DIV0


class TestAggregates:
    def test_sum_ignores_text_and_bool_in_ranges(self):
        assert xl.SUM(R) == 3

    def test_sum_direct_args_coerce(self):
        assert xl.SUM("3", True, 1) == 5

    def test_count_counta(self):
        assert xl.COUNT(R) == 2
        assert xl.COUNTA(R) == 4
        assert xl.COUNTA([[""]]) == 1  # 수식 결과 "" 도 센다

    def test_average_min_max(self):
        assert xl.AVERAGE(R) == 1.5
        assert xl.AVERAGE([[None]]) == xl.DIV0
        assert xl.MIN([[None]]) == 0
        assert xl.MAX(R, 10) == 10

    def test_sumproduct_treats_non_numbers_as_zero(self):
        assert xl.SUMPRODUCT([[1], [2]], [[3], ["x"]]) == 3
        assert xl.SUMPRODUCT([[True], [True]], [[1], [2]]) == 0
        assert xl.SUMPRODUCT([[1, 2]], [[1], [2]]) == xl.VALUE

    def test_array_ops_broadcast(self):
        mask = xl.eq([["a"], ["b"], ["A"]], "a")
        assert mask == [[True], [False], [True]]
        assert xl.SUMPRODUCT(xl.mul(mask, [[1], [2], [3]])) == 4


class TestCriteria:
    data = [["East"], ["west"], ["East"], [None], [""], [5], ["5"], [12]]

    def test_case_insensitive_equality(self):
        assert xl.COUNTIF(self.data, "east") == 2

    def test_numeric_criteria_matches_numeric_text(self):
        assert xl.COUNTIF(self.data, 5) == 2
        assert xl.COUNTIF(self.data, "=5") == 2

    def test_operators(self):
        assert xl.COUNTIF(self.data, ">5") == 1
        assert xl.COUNTIF(self.data, "<>East") == 6
        assert xl.COUNTIF(self.data, "") == 2
        assert xl.COUNTIF(self.data, "<>") == 6

    def test_wildcards(self):
        assert xl.COUNTIF([["Kim"], ["Jim"], ["Kimchi"], ["k*"]], "?im") == 2
        assert xl.COUNTIF([["Kim"], ["Kimchi"]], "Kim*") == 2
        assert xl.COUNTIF([["a*b"], ["axb"]], "a~*b") == 1

    def test_sumifs_and_averageifs(self):
        rng = [["a"], ["b"], ["a"]]
        amt = [[1], [2], [3]]
        assert xl.SUMIF(rng, "a", amt) == 4
        assert xl.SUMIFS(amt, rng, "a", amt, ">1") == 3
        assert xl.AVERAGEIF(rng, "a", amt) == 2
        assert xl.AVERAGEIFS(amt, rng, "z") == xl.DIV0
        assert xl.SUMIFS(amt, [["a"]], "a") == xl.VALUE


class TestLookup:
    table = [["A100", "Alpha", 1000], ["A200", "Beta", 2500], ["B100", "Gamma", 780]]
    brackets = [[0, 0], [10, 0.03], [50, 0.05], [100, 0.08]]

    def test_vlookup_exact_case_insensitive(self):
        assert xl.VLOOKUP("a200", self.table, 3, False) == 2500
        assert xl.VLOOKUP("Z", self.table, 2, False) == xl.NA

    def test_vlookup_approximate(self):
        assert xl.VLOOKUP(49, self.brackets, 2, True) == 0.03
        assert xl.VLOOKUP(500, self.brackets, 2) == 0.08
        assert xl.VLOOKUP(-1, self.brackets, 2) == xl.NA

    def test_vlookup_col_errors(self):
        assert xl.VLOOKUP("A100", self.table, 4, False) == xl.REF
        assert xl.VLOOKUP("A100", self.table, 0, False) == xl.VALUE

    def test_vlookup_wildcard(self):
        assert xl.VLOOKUP("Ga*", [[r[1], r[2]] for r in self.table], 2, False) == 780

    def test_hlookup(self):
        assert xl.HLOOKUP("B", [["A", "B"], [3, 2]], 2, False) == 2

    def test_match_index(self):
        col = [["A100"], ["A200"], ["B100"]]
        assert xl.MATCH("B100", col, 0) == 3
        assert xl.MATCH(75, [[0], [10], [50], [100]], 1) == 3
        assert xl.MATCH(30, [[100], [50], [10]], -1) == 2
        assert xl.INDEX(self.table, 2, 3) == 2500
        assert xl.INDEX(col, 2) == "A200"
        assert xl.INDEX([[1, 2, 3]], 3) == 3
        assert xl.INDEX(self.table, 9, 1) == xl.REF

    def test_xlookup(self):
        keys = [["A100"], ["A200"], ["B100"]]
        prices = [[1000], [2500], [780]]
        assert xl.XLOOKUP("A200", keys, prices) == 2500
        assert xl.XLOOKUP("Z", keys, prices) == xl.NA
        assert xl.XLOOKUP("Z", keys, prices, -1) == -1
        assert xl.XLOOKUP(2000, prices, keys, "없음", 1) == "A200"   # 다음으로 큰 값
        assert xl.XLOOKUP(2000, prices, keys, "없음", -1) == "A100"  # 다음으로 작은 값
        assert xl.XLOOKUP(1, [[1], [2], [1]], [["x"], ["y"], ["z"]], xl.MISSING, 0, -1) == "z"
        assert xl.XLOOKUP("b*", [["a1"], ["b1"]], [[1], [2]], xl.MISSING, 2) == 2


class TestText:
    def test_basic(self):
        assert xl.LEFT("sheetproof", 5) == "sheet"
        assert xl.RIGHT("sheetproof", 5) == "proof"
        assert xl.MID("sheetproof", 3, 4) == "eetp"
        assert xl.LEFT("abc") == "a"
        assert xl.LEN("  a ") == 4
        assert xl.TRIM("  Hello   World  ") == "Hello World"
        assert xl.UPPER("abc") == "ABC"
        assert xl.CONCATENATE(10, "/", True) == "10/TRUE"
        assert xl.CONCAT([["a"], [None], ["b"]], 1) == "ab1"

    @pytest.mark.parametrize("value,fmt,expected", [
        (1234567.891, "#,##0.00", "1,234,567.89"),
        (0.256, "0.0%", "25.6%"),
        (-1234.5, "#,##0", "-1,235"),
        (3.14159, "0.000", "3.142"),
        (5, "000", "005"),
        (45477, "yyyy-mm-dd", "2024-07-04"),
        (45477, "yyyy/mm/dd", "2024/07/04"),
        (45477, "mm/dd/yyyy", "07/04/2024"),
        (12000, '#,##0"원"', "12,000원"),
        (0.5, "0%", "50%"),
        (1.5, "General", "1.5"),
    ])
    def test_text_format(self, value, fmt, expected):
        assert xl.TEXT(value, fmt) == expected


class TestDates:
    def test_date_serials(self):
        assert xl.DATE(2024, 1, 1) == 45292
        assert xl.DATE(1900, 1, 1) == 1
        assert xl.DATE(1900, 2, 29) == 60  # 엑셀의 1900 윤년 버그
        assert xl.DATE(1900, 3, 1) == 61
        assert xl.DATE(99, 1, 1) == xl.DATE(1999, 1, 1)

    def test_date_normalizes_overflow(self):
        assert xl.DATE(2024, 14, 1) == xl.DATE(2025, 2, 1)
        assert xl.DATE(2024, 3, 0) == xl.DATE(2024, 2, 29)

    def test_parts(self):
        d = xl.DATE(2024, 7, 4)
        assert (xl.YEAR(d), xl.MONTH(d), xl.DAY(d)) == (2024, 7, 4)
        assert xl.DAY(60) == 29 and xl.MONTH(60) == 2
        assert xl.MONTH(0) == 1 and xl.DAY(0) == 0
        assert xl.YEAR(-1) == xl.NUM

    def test_edate_eomonth(self):
        assert xl.EDATE(xl.DATE(2024, 1, 31), 1) == xl.DATE(2024, 2, 29)
        assert xl.EDATE(xl.DATE(2024, 2, 29), 12) == xl.DATE(2025, 2, 28)
        assert xl.EOMONTH(xl.DATE(2024, 1, 15), 1) == xl.DATE(2024, 2, 29)
        assert xl.EOMONTH(xl.DATE(2020, 8, 31), -6) == xl.DATE(2020, 2, 29)


class TestLogicAndCompare:
    def test_compare_types(self):
        assert xl.lt(1, "a") is True        # 숫자 < 텍스트
        assert xl.lt("z", True) is True     # 텍스트 < 논리값
        assert xl.eq("abc", "ABC") is True  # 대소문자 무시
        assert xl.eq(None, "") is True
        assert xl.eq(None, 0) is True

    def test_and_or_not(self):
        assert xl.AND(True, 1, [["x"]]) is True  # 범위 안 텍스트 무시
        assert xl.OR(False, 0) is False
        assert xl.AND([["x"]]) == xl.VALUE
        assert xl.NOT(0) is True

    def test_ifs_choose(self):
        assert xl.IFS(False, 1, True, 2) == 2
        assert xl.IFS(False, 1) == xl.NA
        assert xl.CHOOSE(2, "a", "b") == "b"
        assert xl.CHOOSE(3, "a", "b") == xl.VALUE

    def test_info(self):
        assert xl.ISBLANK(None) is True
        assert xl.ISBLANK("") is False
        assert xl.ISNUMBER(True) is False
        assert xl.ISNUMBER(1.5) is True


class TestValuesEqual:
    def test_tolerance(self):
        assert xl.values_equal(0.3, 0.1 + 0.2)
        assert not xl.values_equal(1.0, 1.001)
        assert xl.values_equal(1.0, 1.001, rel_tol=0.01)

    def test_types(self):
        assert xl.values_equal(xl.NA, xl.NA)
        assert not xl.values_equal(xl.NA, xl.DIV0)
        assert not xl.values_equal(1, True)
        assert not xl.values_equal("1", 1)
        assert xl.values_equal(None, "")
