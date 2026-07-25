# Capability floor

Generated from the function registry by `capability.test.ts`. Do not edit by
hand — run `UPDATE_CAPABILITY=1 npm test`.

Anything outside this list renders ⚠ with a stated reason. It is never guessed
at, and never silently treated as zero.

## Implemented (197)

```
ABS ACOS AND AREAS ASIN ATAN ATAN2 AVERAGE
AVERAGEA AVERAGEIF AVERAGEIFS CEILING CEILING.MATH CHAR CHOOSE CLEAN
CODE COLUMN COLUMNS COMBIN CONCAT CONCATENATE CORREL COS
COSH COUNT COUNTA COUNTBLANK COUNTIF COUNTIFS CUMIPMT CUMPRINC
DATE DATEDIF DATEVALUE DAY DAYS DAYS360 DDB DEGREES
DOLLAR EDATE EFFECT EOMONTH ERROR.TYPE EVEN EXACT EXP
FACT FALSE FIND FIXED FLOOR FLOOR.MATH FORECAST FV
GCD HLOOKUP HOUR IF IFERROR IFNA IFS INDEX
INT INTERCEPT IPMT IRR ISBLANK ISERR ISERROR ISEVEN
ISLOGICAL ISNA ISNONTEXT ISNUMBER ISODD ISREF ISTEXT LARGE
LCM LEFT LEN LN LOG LOG10 LOOKUP LOWER
MATCH MAX MAXA MAXIFS MEDIAN MID MIN MINA
MINIFS MINUTE MIRR MOD MONTH MROUND N NA
NETWORKDAYS NOMINAL NOT NOW NPER NPV NUMBERVALUE ODD
OR PDURATION PERCENTILE PERCENTILE.EXC PERCENTILE.INC PI PMT POWER
PPMT PRODUCT PROPER PV QUARTILE QUARTILE.EXC QUARTILE.INC QUOTIENT
RADIANS RAND RANDBETWEEN RANK RANK.EQ RATE REPLACE REPT
RIGHT ROUND ROUNDDOWN ROUNDUP ROW ROWS RRI RSQ
SEARCH SECOND SIGN SIN SINH SLN SLOPE SMALL
SQRT STDEV STDEV.P STDEV.S STDEVP SUBSTITUTE SUBTOTAL SUM
SUMIF SUMIFS SUMPRODUCT SUMSQ SWITCH SYD T TAN
TANH TEXT TEXTJOIN TIME TIMEVALUE TODAY TRANSPOSE TRIM
TRUE TRUNC TYPE UNICHAR UNICODE UPPER VALUE VAR
VAR.P VAR.S VARP VLOOKUP WEEKDAY WORKDAY XIRR XLOOKUP
XMATCH XNPV XOR YEAR YEARFRAC
```

## Refused, deliberately (26)

These are recognised Excel functions the engine declines to evaluate. The
distinction from an unknown name matters for the coverage backlog: an unknown
name might be a typo in the generated model, while a refusal is our own gap with
a stated reason.

```
AGGREGATE CELL FILTER FREQUENCY GETPIVOTDATA GROWTH HYPERLINK INDIRECT
INFO ISFORMULA LINEST LOGEST MDETERM MINVERSE MMULT OFFSET
RANDARRAY RTD SEQUENCE SORT SORTBY TOCOL TOROW TREND
UNIQUE WEBSERVICE
```

Broadly they fall into three groups:

- **Dynamic references** — `INDIRECT`, `OFFSET`. They build references at
  evaluation time, which a static dependency graph cannot express.
- **Spilling results** — `FILTER`, `SORT`, `UNIQUE`, `SEQUENCE`, `MMULT`,
  `TREND` and friends. They change the shape of the grid rather than one cell's
  value. A saved file already contains whatever the writer spilled, so refusing
  costs nothing and prevents a half-right render.
- **State the engine cannot see** — `AGGREGATE` (hidden rows), `CELL`/`INFO`
  (workbook and environment), `GETPIVOTDATA` (pivot cache), `WEBSERVICE`/`RTD`
  (network).

Two further refusals are conditional rather than whole-function, so they do not
appear in the list above:

- `SUBTOTAL(101–111)` — the hidden-row variants only; 1–11 are implemented.
- **Approximate-match lookups over unsorted data** (`VLOOKUP`/`HLOOKUP` with the
  4th argument omitted or TRUE, `MATCH` type 1 or -1). The sorted contract is
  implemented exactly; unsorted input is refused, because Excel answers it from
  its binary search's probe order and that answer is arbitrary.
