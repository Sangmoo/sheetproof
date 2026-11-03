"""테스트용 샘플 통합문서를 만든다 (수식만, cached value 없음).

openpyxl 로 만든 파일에는 계산 결과가 없으므로, 만든 뒤 반드시 실제 Excel 로 재계산·저장한다:

    python scripts/make_samples.py
    powershell -ExecutionPolicy Bypass -File scripts/excel_recalc.ps1

(Excel 이 없으면 LibreOffice headless 로도 가능: README 참고)
"""

from __future__ import annotations

import datetime as dt
import sys
from pathlib import Path

from openpyxl import Workbook
from openpyxl.workbook.defined_name import DefinedName

OUT = Path(__file__).resolve().parent.parent / "samples" / "_raw"


def settlement():
    """단순 정산표: 수량×단가, 할인, 부가세(이름 정의), ROUND, 합계."""
    wb = Workbook()
    ws = wb.active
    ws.title = "정산"
    ws.append(["품목", "수량", "단가", "금액", "할인", "공급가", "부가세", "합계"])
    items = [("볼트", 120, 35.5), ("너트", 300, 12.25), ("와셔", 0, 8), ("앵커", 45, 1250),
             ("브라켓", 17, 4399.9), ("케이블", 250, 2.675)]
    for i, (name, q, p) in enumerate(items, start=2):
        ws.append([name, q, p])
        ws[f"D{i}"] = f"=B{i}*C{i}"
        ws[f"E{i}"] = f"=IF(D{i}>=10000,D{i}*할인율,0)"
        ws[f"F{i}"] = f"=D{i}-E{i}"
        ws[f"G{i}"] = f"=ROUND(F{i}*세율,0)"
        ws[f"H{i}"] = f"=F{i}+G{i}"
    last = 1 + len(items)
    ws[f"A{last + 1}"] = "합계"
    for col in "BDEFGH":
        ws[f"{col}{last + 1}"] = f"=SUM({col}2:{col}{last})"
    ws["J1"], ws["K1"] = "세율", 0.1
    ws["J2"], ws["K2"] = "할인율", 0.05
    ws["J4"] = "평균 단가"
    ws["K4"] = f"=ROUND(AVERAGE(C2:C{last}),2)"
    ws["J5"] = "최대 금액"
    ws["K5"] = f"=MAX(D2:D{last})"
    ws["J6"] = "판매 품목 수"
    ws["K6"] = f'=COUNTIF(B2:B{last},">0")'
    ws["J7"] = "부가세 비율"
    ws["K7"] = f"=G{last + 1}/F{last + 1}"
    ws["J8"] = "반올림 검사"
    ws["K8"] = "=ROUND(2.675,2)+ROUND(-2.5,0)+ROUNDUP(3.14159,3)+ROUNDDOWN(-3.789,1)"
    ws["J9"] = "백분율"
    ws["K9"] = "=50%*K5+2^3^2-(-2)^2+-2^2"
    wb.defined_names["세율"] = DefinedName("세율", attr_text="정산!$K$1")
    wb.defined_names["할인율"] = DefinedName("할인율", attr_text="정산!$K$2")
    wb.save(OUT / "settlement.xlsx")


def price_lookup():
    """단가표 조회: VLOOKUP(정확/근사), HLOOKUP, INDEX/MATCH, XLOOKUP, IFERROR, 시트 간 참조."""
    wb = Workbook()
    t = wb.active
    t.title = "단가표"
    t.append(["코드", "품명", "단가", "등급"])
    rows = [("A100", "Alpha", 1000, "B"), ("A200", "Beta", 2500, "A"), ("B100", "Gamma", 780, "C"),
            ("B200", "Delta", 12000, "A"), ("C300", "Epsilon", 455.5, "B")]
    for r in rows:
        t.append(list(r))
    t["F1"], t["G1"] = "수량 하한", "할인율"
    for i, (lo, rate) in enumerate([(0, 0), (10, 0.03), (50, 0.05), (100, 0.08), (500, 0.12)], start=2):
        t[f"F{i}"], t[f"G{i}"] = lo, rate
    t["I1"], t["J1"], t["K1"], t["L1"] = "등급", "A", "B", "C"
    t["I2"], t["J2"], t["K2"], t["L2"] = "포인트", 3, 2, 1

    o = wb.create_sheet("주문")
    o.append(["코드", "수량", "품명", "단가", "할인율", "금액", "등급", "포인트", "XLOOKUP 단가", "위치"])
    orders = [("B200", 5), ("a100", 120), ("C300", 49), ("Z999", 10), ("A200", 500), ("B100", 0)]
    for i, (code, q) in enumerate(orders, start=2):
        o.append([code, q])
        o[f"C{i}"] = f'=IFERROR(VLOOKUP(A{i},단가표!$A$2:$D$6,2,FALSE),"미등록")'
        o[f"D{i}"] = f"=IFERROR(VLOOKUP(A{i},단가표!$A$2:$D$6,3,0),0)"
        o[f"E{i}"] = f"=VLOOKUP(B{i},단가표!$F$2:$G$6,2,TRUE)"
        o[f"F{i}"] = f"=ROUND(B{i}*D{i}*(1-E{i}),0)"
        o[f"G{i}"] = f"=INDEX(단가표!$D$2:$D$6,MATCH(A{i},단가표!$A$2:$A$6,0))"
        o[f"H{i}"] = f"=IFERROR(HLOOKUP(G{i},단가표!$J$1:$L$2,2,FALSE),0)"
        o[f"I{i}"] = f'=_xlfn.XLOOKUP(A{i},단가표!$A$2:$A$6,단가표!$C$2:$C$6,-1)'
        o[f"J{i}"] = f"=IFERROR(MATCH(A{i},단가표!$A$2:$A$6,0),-1)"
    n = 1 + len(orders)
    o[f"A{n + 2}"] = "총액"
    o[f"F{n + 2}"] = f"=SUM(F2:F{n})"
    o[f"A{n + 3}"] = "최근접 단가"
    o[f"F{n + 3}"] = '=_xlfn.XLOOKUP(2000,단가표!$C$2:$C$6,단가표!$B$2:$B$6,"없음",1)'
    o[f"A{n + 4}"] = "정렬 MATCH"
    o[f"F{n + 4}"] = "=MATCH(75,단가표!$F$2:$F$6,1)"
    o[f"A{n + 5}"] = "INDEX 2D"
    o[f"F{n + 5}"] = "=INDEX(단가표!$A$2:$D$6,4,3)"
    o[f"A{n + 6}"] = "와일드카드"
    o[f"F{n + 6}"] = '=VLOOKUP("Ep*",단가표!$B$2:$C$6,2,FALSE)'
    wb.save(OUT / "price_lookup.xlsx")


def sumifs_summary():
    """SUMIF(S), COUNTIF(S), AVERAGEIF(S), SUMPRODUCT 집계."""
    wb = Workbook()
    d = wb.active
    d.title = "Data"
    d.append(["Region", "Rep", "Product", "Qty", "Amount", "Date"])
    data = [
        ("East", "Kim", "Widget", 10, 1500.5, dt.date(2024, 1, 15)),
        ("West", "Lee", "Gadget", 5, 820, dt.date(2024, 1, 20)),
        ("East", "Park", "Widget", 7, 1050, dt.date(2024, 2, 3)),
        ("North", "Kim", "Gizmo", 12, 2400, dt.date(2024, 2, 14)),
        ("West", "Choi", "Widget", 3, 450.25, dt.date(2024, 3, 1)),
        ("East", "Lee", "Gadget", None, 0, dt.date(2024, 3, 9)),
        ("South", "Park", "Gizmo", 8, 1600, dt.date(2024, 3, 30)),
        ("east", "Kim", "Widget Pro", 4, 990, dt.date(2024, 4, 2)),
        ("West", "Kim", "Gadget", 9, "n/a", dt.date(2024, 4, 18)),
        ("North", "Lee", "Widget", 6, 900, dt.date(2024, 5, 5)),
    ]
    for r in data:
        d.append(list(r))
    for i in range(2, 2 + len(data)):
        d[f"F{i}"].number_format = "yyyy-mm-dd"
    s = wb.create_sheet("Summary")
    s.append(["Region", "Amount", "Count", "Avg", "Qty>5", "Widget*"])
    for i, reg in enumerate(["East", "West", "North", "South", "Central"], start=2):
        s[f"A{i}"] = reg
        s[f"B{i}"] = f"=SUMIF(Data!$A$2:$A$11,A{i},Data!$E$2:$E$11)"
        s[f"C{i}"] = f"=COUNTIF(Data!$A$2:$A$11,A{i})"
        s[f"D{i}"] = f'=IFERROR(AVERAGEIF(Data!$A$2:$A$11,A{i},Data!$E$2:$E$11),"-")'
        s[f"E{i}"] = f'=COUNTIFS(Data!$A$2:$A$11,A{i},Data!$D$2:$D$11,">5")'
        s[f"F{i}"] = f'=SUMIFS(Data!$E$2:$E$11,Data!$A$2:$A$11,A{i},Data!$C$2:$C$11,"Widget*")'
    s["H1"], s["I1"] = "지표", "값"
    metrics = [
        ("Kim 매출", '=SUMIFS(Data!E2:E11,Data!B2:B11,"Kim")'),
        ("1분기 매출", '=SUMIFS(Data!E2:E11,Data!F2:F11,">="&DATE(2024,1,1),Data!F2:F11,"<"&DATE(2024,4,1))'),
        ("Qty 합(SUMPRODUCT)", '=SUMPRODUCT((Data!A2:A11="East")*(Data!D2:D11))'),
        ("금액×수량", "=SUMPRODUCT(Data!D2:D11,Data!E2:E11)"),
        ("Qty 빈칸 수", '=COUNTIF(Data!D2:D11,"")'),
        ("Qty 숫자 수", "=COUNT(Data!D2:D11)"),
        ("Amount 비어있지 않음", "=COUNTA(Data!E2:E11)"),
        ("Gadget 아닌 행", '=COUNTIF(Data!C2:C11,"<>Gadget")'),
        ("평균(1000초과)", '=AVERAGEIFS(Data!E2:E11,Data!E2:E11,">1000")'),
        ("최소 금액", "=MIN(Data!E2:E11)"),
        ("? 와일드카드", '=COUNTIF(Data!B2:B11,"?im")'),
        ("Qty>=8 금액", '=SUMIF(Data!D2:D11,">=8",Data!E2:E11)'),
        ("문자 비교", '=COUNTIF(Data!A2:A11,">N")'),
    ]
    for i, (label, f) in enumerate(metrics, start=2):
        s[f"H{i}"], s[f"I{i}"] = label, f
    wb.save(OUT / "sumifs_summary.xlsx")


def dates():
    """날짜 계산: DATE, YEAR/MONTH/DAY, EDATE, EOMONTH, TEXT, 1900 체계."""
    wb = Workbook()
    ws = wb.active
    ws.title = "계약"
    ws.append(["계약", "시작일", "개월", "만료일", "월말", "년", "월", "일", "표시", "경과일", "분기"])
    contracts = [("C-1", dt.date(2024, 1, 31), 1), ("C-2", dt.date(2023, 11, 15), 3),
                 ("C-3", dt.date(2024, 2, 29), 12), ("C-4", dt.date(2020, 8, 31), -6),
                 ("C-5", dt.date(1999, 12, 31), 25)]
    for i, (c, start, m) in enumerate(contracts, start=2):
        ws.append([c, start, m])
        ws[f"B{i}"].number_format = "yyyy-mm-dd"
        ws[f"D{i}"] = f"=EDATE(B{i},C{i})"
        ws[f"D{i}"].number_format = "yyyy-mm-dd"
        ws[f"E{i}"] = f"=EOMONTH(B{i},C{i})"
        ws[f"E{i}"].number_format = "yyyy-mm-dd"
        ws[f"F{i}"] = f"=YEAR(D{i})"
        ws[f"G{i}"] = f"=MONTH(D{i})"
        ws[f"H{i}"] = f"=DAY(D{i})"
        ws[f"I{i}"] = f'=TEXT(D{i},"yyyy-mm-dd")&" ("&TEXT(C{i},"0")&"개월)"'
        ws[f"J{i}"] = f"=E{i}-B{i}"
        ws[f"K{i}"] = f'="Q"&INT((MONTH(B{i})-1)/3)+1'
    ws["M1"], ws["N1"] = "검사", "값"
    checks = [
        ("DATE 월 넘김", "=DATE(2024,14,1)"),
        ("DATE 음수 일", "=DATE(2024,3,0)"),
        ("DATE 2자리 연도", "=DATE(99,1,1)"),
        ("1900-02-29", "=DATE(1900,2,29)"),
        ("1900-03-01", "=DATE(1900,3,1)"),
        ("일련번호 60 일", "=DAY(60)"),
        ("일련번호 0 월", "=MONTH(0)"),
        ("윤년 월말", "=EOMONTH(DATE(2024,1,15),1)"),
        ("TEXT 숫자", '=TEXT(1234567.891,"#,##0.00")'),
        ("TEXT 퍼센트", '=TEXT(0.256,"0.0%")'),
        ("TEXT 날짜", '=TEXT(DATE(2024,7,4),"yyyy/mm/dd")'),
        ("TEXT 정수", '=TEXT(-1234.5,"#,##0")'),
        ("날짜 차이", "=DATE(2024,12,25)-DATE(2024,1,1)"),
    ]
    for i, (label, f) in enumerate(checks, start=2):
        ws[f"M{i}"], ws[f"N{i}"] = label, f
    wb.save(OUT / "dates.xlsx")


def errors_text():
    """오류값 전파, 텍스트 함수, 논리 함수, 휘발성/미지원 함수."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Checks"
    ws.append(["입력", "값"])
    inputs = [("a", 10), ("b", 0), ("c", "  Hello   World  "), ("d", "123"), ("e", None),
              ("f", True), ("g", -7.5), ("h", "sheetproof")]
    for i, (k, v) in enumerate(inputs, start=2):
        ws[f"A{i}"] = k
        if v is not None:
            ws[f"B{i}"] = v
    ws["D1"], ws["E1"] = "검사", "결과"
    checks = [
        ("0 나누기", "=B2/B3"),
        ("오류 전파", "=E2+1"),
        ("IFERROR", '=IFERROR(B2/B3,"div0")'),
        ("N/A 전파", '=VLOOKUP("zzz",A2:B9,2,FALSE)'),
        ("ISNA", "=ISNA(E5)"),
        ("텍스트+숫자", "=B5+1"),
        ("VALUE 오류", '=B4+"abc"'),
        ("빈 셀 산술", "=B6*5+1"),
        ("빈 셀 연결", '="["&B6&"]"'),
        ("ISBLANK", "=ISBLANK(B6)"),
        ("ISNUMBER", "=ISNUMBER(B5)"),
        ("TRIM", "=TRIM(B4)"),
        ("LEFT/RIGHT/MID", '=LEFT(B9,5)&"-"&RIGHT(B9,5)&"-"&MID(B9,3,4)'),
        ("LEN", "=LEN(B4)"),
        ("UPPER/LOWER", '=UPPER(B9)&LOWER("ABC")'),
        ("CONCATENATE", '=CONCATENATE(B2,"/",B7,"/",B8)'),
        ("CONCAT 범위", "=_xlfn.CONCAT(A2:A5)"),
        ("논리값 산술", "=B7+1"),
        ("AND/OR/NOT", "=AND(B2>5,OR(B3=1,NOT(B7=FALSE)))"),
        ("IFS", '=_xlfn.IFS(B8>0,"양수",B8<0,"음수",TRUE,"영")'),
        ("CHOOSE", '=CHOOSE(3,"x","y","z")'),
        ("INT 음수", "=INT(B8)"),
        ("MOD 음수", "=MOD(B8,4)+MOD(7,-3)"),
        ("ABS", "=ABS(B8)"),
        ("문자 비교", '="abc"="ABC"'),
        ("타입 비교", '=1<"a"'),
        ("빈 셀 비교", '=B6=""'),
        ("숫자 텍스트", '=TEXT(B2/3,"0.000")'),
        ("오류 비교", "=IF(ISERROR(E2),-1,1)"),
        ("일반 서식", '="v="&(1/3)'),
        ("미지원 함수", '=PROPER(B9)&REPT("!",2)'),
        ("휘발성", "=TODAY()>0"),
        ("미지원 의존", "=LEN(E32)"),
    ]
    for i, (label, f) in enumerate(checks, start=2):
        ws[f"D{i}"], ws[f"E{i}"] = label, f
    wb.save(OUT / "errors_text.xlsx")


def no_cache():
    """cached value 가 없는 파일 (openpyxl 로 만들고 재계산하지 않음)."""
    wb = Workbook()
    ws = wb.active
    ws["A1"], ws["A2"], ws["A3"] = 1, 2, "=A1+A2"
    wb.save(OUT.parent / "no_cache.xlsx")


def main():
    OUT.mkdir(parents=True, exist_ok=True)
    for fn in (settlement, price_lookup, sumifs_summary, dates, errors_text, no_cache):
        fn()
        print("made", fn.__name__)


if __name__ == "__main__":
    sys.exit(main())
