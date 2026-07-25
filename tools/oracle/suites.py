"""Probe suites — the corpus the engine is scored against.

Each suite is a shared grid of literal values plus a list of formulas. The same
spec drives both sides of the comparison: Python writes it into a real `.xlsx`
which LibreOffice recalculates, and Node builds the identical workbook in the
engine. Any disagreement is then a genuine semantic difference, not a difference
in how the two sides were set up.

Probes are chosen for the places an engine goes *quietly* wrong — coercion,
blank-versus-empty-string, rounding at the decimal boundary, sign conventions —
rather than for coverage-by-function-count.
"""

# A shared grid every suite gets. Column letters here are referenced directly by
# the probe formulas, so the layout is part of the contract.
BASE_GRID = {
    # A: plain numbers
    "A1": 1, "A2": 2, "A3": 3, "A4": 4, "A5": 5,
    # B: mixed types — the direct-versus-range rule lives here
    "B1": 10, "B2": "5", "B3": True, "B4": None, "B5": "text",
    # C: numbers with awkward decimals
    "C1": 2.675, "C2": 1.005, "C3": -2.5, "C4": 0.1, "C5": 0.2,
    # D: text
    "D1": "North", "D2": "South", "D3": "Northeast", "D4": "  padded  ", "D5": "MiXeD",
    # E: money-ish
    "E1": 1827.6, "E2": 1078.6, "E3": -168.9, "E4": 0, "E5": 749.0,
    # F: a sorted lookup key column, G its payload
    "F1": 1, "F2": 5, "F3": 9, "F4": 20, "F5": 50,
    "G1": "one", "G2": "five", "G3": "nine", "G4": "twenty", "G5": "fifty",
    # H: dates as serials
    "H1": 45000, "H2": 45291, "H3": 46228, "H4": 60, "H5": 1,
    # I: cash flows
    "I1": -1000, "I2": 400, "I3": 400, "I4": 400, "I5": 400,
    # J: an empty column, on purpose
    # L: ascending dates, for XNPV/XIRR
    "L1": 45000, "L2": 45365, "L3": 45730, "L4": 46095, "L5": 46460,
    # K: negatives and zero
    "K1": -1, "K2": 0, "K3": -0.5, "K4": 100, "K5": -100,
}

SUITES: dict[str, list[str]] = {}


def suite(name: str, formulas: list[str]) -> None:
    SUITES[name] = formulas


# ---------------------------------------------------------------------------

suite("operators", [
    "=1+1", "=10-3", "=6*7", "=10/4", "=2^10", "=10/3",
    # The two precedence traps.
    "=-2^2", "=2^3^2", "=-2^2+1", "=2^-1", "=1--1", "=-(2^2)",
    "=50%", "=200*5%", "=-2%", "=50%%",
    '="a"&"b"', '=1&2', '=1&""', '=TRUE&""',
    "=1=1", "=1<>1", "=1<2", "=2>1", "=1<=1", "=1>=2",
    "=(1+2)*3", "=1+2*3", "=2*3^2",
    "=1/0", "=0/0", "=A1/K2",
])

suite("coercion", [
    '="5"+1', '="5"-1', '="5"*2', '="5"/5', '="5"=5', '="5"<5', '="05"=5',
    "=TRUE+1", "=TRUE*2", "=TRUE=1", "=FALSE=0", "=TRUE>1", '=TRUE>"z"',
    '="a"+1', '=""+1', '=" "+1', '="1,000"+0', '="5%"+0', '="$5"+0',
    '="a">1', '="a"="A"', '=EXACT("a","A")',
    "=B4+1", "=B4=0", '=B4=""', "=B4&1", "=ISBLANK(B4)",
    "=J1+1", "=J1=0", '=J1=""', "=ISBLANK(J1)",
    '=SUM(B1:B5)', '=SUM("5")', "=SUM(TRUE)", '=SUM("a")',
    "=COUNT(B1:B5)", "=COUNTA(B1:B5)", "=COUNTBLANK(B1:B5)",
    "=AND(B1:B5)", "=OR(A1:A5)",
])

suite("blank_semantics", [
    "=IF(J1=0,\"\",A1/J1-1)",
    "=IF(B4=0,\"empty\",\"filled\")",
    "=J1", "=J1*2", "=LEN(J1)", "=N(J1)", "=T(J1)",
    "=COUNTIF(B1:B5,\"\")", "=COUNTIF(B1:B5,\"<>\")",
    "=SUMIF(A1:A5,\">0\",J1:J5)",
    "=MIN(J1:J5)", "=MAX(J1:J5)", "=AVERAGE(J1:J5)",
])

suite("rounding", [
    "=ROUND(2.675,2)", "=ROUND(C1,2)", "=ROUND(1.005,2)", "=ROUND(C2,2)",
    "=ROUND(2.5,0)", "=ROUND(-2.5,0)", "=ROUND(C3,0)",
    "=ROUND(1234.5678,-2)", "=ROUNDUP(1234.5678,-2)", "=ROUNDDOWN(1234.5678,-2)",
    "=ROUNDUP(2.001,2)", "=ROUNDDOWN(-2.999,2)", "=ROUND(0,2)",
    "=INT(-2.5)", "=INT(2.5)", "=TRUNC(-2.5)", "=TRUNC(2.999,2)",
    "=MOD(-3,2)", "=MOD(3,-2)", "=MOD(3,2)", "=MOD(3,0)",
    "=CEILING(4.2,1)", "=FLOOR(4.8,1)", "=CEILING(-4.2,-1)", "=MROUND(10,3)",
    "=EVEN(1.5)", "=ODD(1.5)", "=EVEN(-1.5)", "=ODD(-1.5)", "=EVEN(2)", "=ODD(3)",
    "=C4+C5-0.3", "=ROUND(C4+C5,10)",
    "=ABS(K1)", "=SIGN(K1)", "=SIGN(K2)", "=SQRT(4)", "=SQRT(K1)",
    "=POWER(2,10)", "=EXP(1)", "=LN(1)", "=LN(0)", "=LOG(100)", "=LOG(8,2)", "=LOG10(1000)",
])

suite("aggregates", [
    "=SUM(A1:A5)", "=SUM(A1:A5,10)", "=PRODUCT(A1:A5)", "=SUMSQ(A1:A3)",
    "=AVERAGE(A1:A5)", "=MEDIAN(A1:A5)", "=MIN(A1:A5)", "=MAX(A1:A5)",
    "=COUNT(A1:A5)", "=COUNTA(A1:A5)",
    "=STDEV.S(A1:A5)", "=STDEV.P(A1:A5)", "=VAR.S(A1:A5)", "=VAR.P(A1:A5)",
    "=LARGE(A1:A5,2)", "=SMALL(A1:A5,2)", "=RANK(A3,A1:A5)", "=RANK(A3,A1:A5,1)",
    "=SUMPRODUCT(A1:A3,A3:A5)", "=SUMPRODUCT(A1:A3)",
    "=SUBTOTAL(9,A1:A5)", "=SUBTOTAL(1,A1:A5)", "=SUBTOTAL(4,A1:A5)",
    "=SUM(A:A)", "=SUM(A1:A5 A3:A5)",
    "=PERCENTILE.INC(A1:A5,0.5)", "=QUARTILE.INC(A1:A5,1)",
    "=CORREL(A1:A5,I1:I5)", "=SLOPE(I1:I5,A1:A5)", "=INTERCEPT(I1:I5,A1:A5)",
])

suite("conditional", [
    '=SUMIF(A1:A5,">2")', '=SUMIF(A1:A5,">=3")', '=SUMIF(A1:A5,"<>3")',
    '=SUMIF(D1:D5,"North*",A1:A5)', '=SUMIF(D1:D5,"North",A1:A5)',
    '=SUMIF(D1:D5,"?orth",A1:A5)', '=SUMIF(A1:A5,3)',
    '=COUNTIF(A1:A5,">2")', '=COUNTIF(D1:D5,"North*")', '=COUNTIF(D1:D5,"*e*")',
    '=COUNTIF(B1:B5,TRUE)', '=COUNTIF(B1:B5,"text")',
    '=AVERAGEIF(A1:A5,">2")',
    '=SUMIFS(A1:A5,A1:A5,">1",A1:A5,"<5")',
    '=COUNTIFS(A1:A5,">1",D1:D5,"*o*")',
    '=MAXIFS(A1:A5,A1:A5,"<4")', '=MINIFS(A1:A5,A1:A5,">2")',
])

suite("logical", [
    "=IF(TRUE,1,2)", "=IF(FALSE,1,2)", "=IF(A1=1,\"yes\",\"no\")",
    "=IF(TRUE,1,1/0)", "=IF(A1>0,A1/A1,1/0)",
    "=IF(1,\"t\")", "=IF(0,\"t\")",
    "=AND(TRUE,TRUE)", "=AND(TRUE,FALSE)", "=OR(FALSE,TRUE)", "=NOT(TRUE)",
    "=XOR(TRUE,TRUE)", "=XOR(TRUE,FALSE)",
    "=IFERROR(1/0,\"trapped\")", "=IFERROR(1,\"trapped\")",
    "=IFNA(NA(),\"trapped\")", "=IFNA(1/0,\"trapped\")",
    "=IFS(A1=2,\"two\",A1=1,\"one\")", "=IFS(FALSE,1)",
    "=SWITCH(A1,1,\"one\",2,\"two\",\"other\")", "=SWITCH(A5,1,\"one\",\"other\")",
    "=CHOOSE(2,\"a\",\"b\",\"c\")", "=CHOOSE(1,A1,1/0)",
    "=ISNUMBER(A1)", "=ISTEXT(D1)", "=ISERROR(1/0)", "=ISERR(NA())", "=ISNA(NA())",
    "=ISLOGICAL(B3)", "=ISEVEN(2)", "=ISODD(3)", "=TYPE(A1)", "=TYPE(D1)",
    "=ERROR.TYPE(1/0)",
])

suite("text", [
    "=LEN(D1)", "=LEFT(D1,3)", "=RIGHT(D1,3)", "=MID(D1,2,3)", "=LEFT(D1)",
    "=UPPER(D5)", "=LOWER(D5)", "=PROPER(D5)", "=TRIM(D4)",
    '=FIND("o",D1)', '=FIND("O",D1)', '=SEARCH("O",D1)', '=SEARCH("n*h",D3)',
    '=SUBSTITUTE(D1,"o","0")', '=SUBSTITUTE("aaa","a","b",2)',
    '=REPLACE(D1,1,1,"X")', '=REPT("ab",3)',
    '=CONCATENATE(D1,"-",D2)', '=CONCAT(D1:D2)', '=TEXTJOIN(",",TRUE,D1:D3)',
    '=TEXTJOIN(",",TRUE,B1:B5)', '=TEXTJOIN(",",FALSE,B1:B5)',
    "=CHAR(65)", "=CODE(\"A\")", "=EXACT(D1,D1)",
    '=VALUE("5")', '=VALUE("5%")', '=VALUE("abc")',
    "=FIXED(1234.5678,2)", "=FIXED(1234.5678,2,TRUE)",
    "=A1&\"\"", "=C1&\"\"", "=E1&\"\"", "=(1/3)&\"\"", "=(2/3)&\"\"",
    "=1E20&\"\"", "=0.00001&\"\"", "=123456789012345678&\"\"",
    # Pin the exact points where General switches to scientific notation.
    "=0.0001&\"\"", "=0.000001&\"\"", "=0.0000001&\"\"", "=1E-10&\"\"", "=1E-11&\"\"",
    "=1E15&\"\"", "=1E16&\"\"", "=1E17&\"\"", "=1E21&\"\"", "=1E22&\"\"",
    "=1234567890123456&\"\"", "=(1/7)&\"\"", "=(-1/3)&\"\"", "=1.5&\"\"", "=(-0)&\"\"",
    "=1E-12&\"\"", "=1E-13&\"\"", "=1E-14&\"\"", "=1E-15&\"\"", "=1E-20&\"\"",
    "=1E14&\"\"", "=1E15&\"\"",
])

suite("dates", [
    "=DATE(2026,7,25)", "=DATE(1900,1,1)", "=DATE(1900,3,1)", "=DATE(1900,2,28)",
    "=DATE(2026,13,1)", "=DATE(2026,1,32)", "=DATE(26,1,1)",
    "=YEAR(H3)", "=MONTH(H3)", "=DAY(H3)", "=YEAR(H5)", "=MONTH(H5)", "=DAY(H5)",
    "=WEEKDAY(H3)", "=WEEKDAY(H3,2)", "=WEEKDAY(H3,3)",
    "=EOMONTH(H3,0)", "=EOMONTH(H3,1)", "=EOMONTH(H3,-1)",
    "=EDATE(H3,1)", "=EDATE(DATE(2026,1,31),1)", "=EOMONTH(DATE(2026,1,31),1)",
    "=DAYS(H3,H1)", "=H3-H1",
    "=DAYS360(H1,H3)", "=DAYS360(H1,H3,TRUE)",
    '=DATEDIF(H1,H3,"D")', '=DATEDIF(H1,H3,"M")', '=DATEDIF(H1,H3,"Y")',
    "=YEARFRAC(H1,H3)", "=YEARFRAC(H1,H3,1)", "=YEARFRAC(H1,H3,2)",
    "=YEARFRAC(H1,H3,3)", "=YEARFRAC(H1,H3,4)",
    "=NETWORKDAYS(H1,H2)", "=WORKDAY(H1,10)",
    "=TIME(13,30,0)", "=HOUR(0.5)", "=MINUTE(0.5)", "=SECOND(0.5)",
])

suite("lookup", [
    "=VLOOKUP(5,F1:G5,2,FALSE)", "=VLOOKUP(5,F1:G5,2,TRUE)",
    "=VLOOKUP(7,F1:G5,2,TRUE)", "=VLOOKUP(0,F1:G5,2,TRUE)",
    "=VLOOKUP(7,F1:G5,2,FALSE)", "=VLOOKUP(50,F1:G5,2)",
    "=HLOOKUP(1,A1:E1,1,FALSE)",
    "=MATCH(9,F1:F5,0)", "=MATCH(7,F1:F5,1)", "=MATCH(7,F1:F5)",
    '=MATCH("nine",G1:G5,0)', '=MATCH("n*",G1:G5,0)',
    "=INDEX(G1:G5,3)", "=INDEX(F1:G5,3,2)", "=INDEX(A1:E1,1,2)",
    "=INDEX(G1:G5,MATCH(9,F1:F5,0))",
    "=LOOKUP(7,F1:F5,G1:G5)",
    "=ROWS(F1:G5)", "=COLUMNS(F1:G5)", "=ROW(F3)", "=COLUMN(G1)",
    "=XLOOKUP(9,F1:F5,G1:G5)", '=XLOOKUP(7,F1:F5,G1:G5,"none")',
    "=XLOOKUP(7,F1:F5,G1:G5,,-1)", "=XLOOKUP(7,F1:F5,G1:G5,,1)",
    "=XMATCH(9,F1:F5)",
])

suite("financial", [
    "=PMT(0.05/12,360,-300000)", "=PMT(0.05,10,-1000)", "=PMT(0,10,-1000)",
    "=PMT(0.05,10,-1000,0,1)",
    "=PV(0.05,10,-100)", "=FV(0.05,10,-100)", "=FV(0.05,10,-100,0,1)",
    "=NPER(0.05,-100,1000)", "=RATE(10,-100,1000)",
    "=NPV(0.1,I2:I5)", "=NPV(0.1,I1:I5)", "=I1+NPV(0.1,I2:I5)",
    "=IRR(I1:I5)", "=IRR(I1:I5,0.5)", "=MIRR(I1:I5,0.1,0.12)",
    "=IPMT(0.05,1,10,-1000)", "=IPMT(0.05,5,10,-1000)",
    "=PPMT(0.05,1,10,-1000)", "=PPMT(0.05,5,10,-1000)",
    "=SLN(10000,1000,5)", "=SYD(10000,1000,5,1)", "=SYD(10000,1000,5,5)",
    "=DDB(10000,1000,5,1)", "=DDB(10000,1000,5,2)",
    "=EFFECT(0.05,12)", "=NOMINAL(0.0512,12)", "=RRI(10,1000,2000)",
    "=PDURATION(0.05,1000,2000)",
    "=XNPV(0.1,I1:I5,L1:L5)", "=XNPV(0.1,I1:I5,H1:H5)", "=XIRR(I1:I5,L1:L5)",
    "=CUMIPMT(0.05,10,1000,1,3,0)", "=CUMPRINC(0.05,10,1000,1,3,0)",
])

suite("errors", [
    "=1/0", "=SQRT(-1)", "=NA()", '=VLOOKUP("zzz",F1:G5,2,FALSE)',
    "=1/0+1", "=SUM(1/0,1)", "=IFERROR(1/0,0)",
    "=#DIV/0!", "=#N/A", "=#VALUE!",
    "=A1:A2 D1:D2",
    "=INDEX(A1:A5,99)", "=VLOOKUP(1,F1:G5,9,FALSE)",
])

suite("model_shapes", [
    # The shapes that actually appear in a generated financial model.
    "=E1-E2", "=(E1-E2)/E1", "=E1*0.35", "=SUM(A1:A5)/COUNT(A1:A5)",
    "=IF(E1=0,0,E3/E1)", "=IF(E4=0,\"n/a\",E1/E4)",
    "=ROUND(E1*1.05,1)", "=E1*(1+0.05)^3",
    "=-E3", "=ABS(E3)", "=MAX(0,E3)", "=MIN(E1,E2)",
    "=E1/1000", "=E1&\" total\"",
    "=SUM(A1:A5)-SUM(A1:A3)",
    "=IFERROR(E3/E4,0)",
])
