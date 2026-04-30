# Excel Runtime Calculation Surface

`excel-runtime` now includes a numeric-first calculation slice inside the main runtime crate.
This is not full Excel parity, but it is no longer accurate to describe the runtime as "load/save only".

## Entry Points

- `Application.Calculate`: recalculates supported in-cell formulas.
- `Application.Evaluate`: evaluates a formula expression against the active sheet context.
- `Worksheet.Evaluate`: evaluates a formula expression against the target worksheet context.
- `Range.Formula` / `Range.FormulaR1C1`: stored formulas participate in recalculation through the same evaluator.

## Current Formula Scope

### Core expression support

- arithmetic operators: `+`, `-`, `*`, `/`
- comparison operators: `=`, `<>`, `<`, `<=`, `>`, `>=`
- parentheses and unary sign
- boolean literals: `TRUE`, `FALSE`
- A1 single-cell references
- sheet-qualified A1 references
- in-cell `FormulaR1C1` formulas via conversion to A1 before evaluation

### Math and scalar helpers

- `ABS`
- `INT`
- `ROUND`
- `ROUNDUP`
- `ROUNDDOWN`
- `TRUNC`
- `CEILING`
- `CEILING.MATH`
- `CEILING.PRECISE`
- `ISO.CEILING`
- `FLOOR`
- `FLOOR.MATH`
- `FLOOR.PRECISE`
- `MROUND`
- `QUOTIENT`
- `MOD`
- `SIGN`
- `EVEN`
- `ODD`
- `FACT`
- `FACTDOUBLE`
- `COMBIN`
- `COMBINA`
- `PERMUT`
- `PERMUTATIONA`
- `MULTINOMIAL`
- `POWER`
- `SQRT`
- `EXP`
- `LN`
- `LOG`
- `LOG10`
- `PI`
- `SIN`
- `COS`
- `TAN`
- `SEC`
- `CSC`
- `COT`
- `ASIN`
- `ACOS`
- `ATAN`
- `ATAN2`
- `ACOT`
- `SINH`
- `COSH`
- `TANH`
- `SECH`
- `CSCH`
- `COTH`
- `ASINH`
- `ACOSH`
- `ATANH`
- `ACOTH`
- `DEGREES`
- `RADIANS`
- `ISEVEN`
- `ISODD`
- `GCD`
- `LCM`
- `SERIESSUM`
- `BITAND`
- `BITOR`
- `BITXOR`
- `BITLSHIFT`
- `BITRSHIFT`
- `DELTA`
- `GESTEP`
- `DECIMAL`
- `ARABIC`
- `ROMAN`
- `BIN2DEC`
- `BIN2HEX`
- `BIN2OCT`
- `DEC2BIN`
- `DEC2HEX`
- `DEC2OCT`
- `HEX2BIN`
- `HEX2DEC`
- `HEX2OCT`
- `OCT2BIN`
- `OCT2DEC`
- `OCT2HEX`

Roman numeral notes:

- `ROMAN` supports classic and simplified forms `0` through `4`; omitted form and `TRUE` use classic form, while `FALSE` uses simplified form.
- `ARABIC` accepts classic and simplified Roman numerals, ignores case and surrounding spaces, returns `0` for empty text, and supports a leading negative sign.
- `SERIESSUM` supports scalar or rectangular coefficient inputs and ignores non-numeric cells in coefficient ranges.

### Financial helpers

- `DOLLARDE`
- `DOLLARFR`
- `FVSCHEDULE`
- `NPV`
- `XNPV`
- `IRR`
- `XIRR`
- `MIRR`
- `DISC`
- `INTRATE`
- `RECEIVED`
- `PRICEDISC`
- `YIELDDISC`
- `PRICEMAT`
- `YIELDMAT`
- `ACCRINT`
- `ACCRINTM`
- `TBILLEQ`
- `TBILLPRICE`
- `TBILLYIELD`
- `COUPDAYBS`
- `COUPDAYS`
- `COUPDAYSNC`
- `COUPNCD`
- `COUPNUM`
- `COUPPCD`
- `PRICE`
- `YIELD`
- `DURATION`
- `MDURATION`
- `FV`
- `PV`
- `PMT`
- `IPMT`
- `PPMT`
- `CUMIPMT`
- `CUMPRINC`
- `SLN`
- `SYD`
- `DB`
- `DDB`
- `AMORLINC`
- `EFFECT`
- `NOMINAL`
- `RRI`
- `PDURATION`
- `NPER`
- `RATE`
- `ISPMT`

Financial notes:

- `DOLLARDE` and `DOLLARFR` truncate the denominator argument before conversion, return `#NUM!` for negative denominators, and return `#DIV/0!` when the truncated denominator is zero.
- `FVSCHEDULE` compounds the principal by each scheduled rate. Range schedules treat blank cells as zero rates and return `#VALUE!` for text or logical cells.
- `NPV` discounts ordered end-of-period cash flows and ignores empty, logical, text, and error values inside value arguments.
- `XNPV` discounts irregular cash flows over a 365-day year, truncates date serials to integers, and returns `#NUM!` for length mismatches or dates before the first date.
- `IRR` preserves value order, ignores text, logical, and empty cells in range inputs, and uses a 20-iteration solve with the documented `10%` default guess.
- `XIRR` solves the rate for irregular dated cash flows with the same date validation as `XNPV` and uses a 100-iteration solve with the documented `10%` default guess.
- `MIRR` preserves value order, includes zero cash flows as periods, and ignores text, logical, and empty cells in range inputs.
- `DISC`, `INTRATE`, `RECEIVED`, `PRICEDISC`, and `YIELDDISC` cover discounted securities using the same basis values as `YEARFRAC`.
- `PRICEMAT` and `YIELDMAT` cover securities that pay interest at maturity, validating issue, settlement, and maturity dates against the same basis values.
- `ACCRINT` calculates periodic accrued interest across normal and odd first coupon periods, honoring the optional `calc_method` argument.
- `ACCRINTM` calculates interest accrued to maturity using `YEARFRAC` basis values and defaults omitted par to `1000`.
- `TBILLEQ`, `TBILLPRICE`, and `TBILLYIELD` cover Treasury bill formulas, using actual days and rejecting maturities more than 365 days after settlement.
- `COUPDAYBS`, `COUPDAYS`, `COUPDAYSNC`, `COUPNCD`, `COUPNUM`, and `COUPPCD` cover regular coupon schedules with annual, semiannual, and quarterly frequencies and basis values `0` through `4`.
- `PRICE` and `YIELD` calculate regular coupon security price and yield over the same schedule model, including one-coupon and multi-coupon pricing paths.
- `DURATION` and `MDURATION` calculate Macauley and modified duration over the same regular coupon schedules.
- `FV`, `PV`, `PMT`, `NPER`, and `RATE` support omitted future/present value and payment timing arguments with the standard `0`/`1` timing modes. `RATE` uses a 20-iteration solve with the documented `10%` default guess. `ISPMT` uses Excel's zero-based period convention.
- `IPMT` and `PPMT` split the standard `PMT` result into interest and principal portions for one-based periods in the range `1..nper`.
- `CUMIPMT` and `CUMPRINC` sum the corresponding interest or principal portions across a one-based inclusive period range.
- `SLN` and `SYD` cover straight-line and sum-of-years' digits depreciation with `#NUM!` for non-positive life or out-of-range periods.
- `DB` covers fixed-declining balance depreciation with the optional first-year month argument and Excel's three-decimal fixed rate rounding.
- `DDB` covers double-declining balance depreciation with an optional factor and caps each period so accumulated depreciation does not drop below salvage value.
- `AMORLINC` covers French-accounting linear depreciation, including prorated first periods and salvage-value caps.
- `EFFECT` and `NOMINAL` truncate compounding periods to an integer. `RRI` and `PDURATION` cover direct logarithmic/compound-growth rate helpers.

### Logical and control helpers

- `TRUE`
- `FALSE`
- `AND`
- `OR`
- `NOT`
- `XOR`
- `IF`
- `IFS`
- `SWITCH`

Logical notes:

- `AND`, `OR`, and `XOR` support scalar arguments and rectangular range arguments.
- Text and blank cells inside logical range arguments are ignored; a logical range with no numeric or boolean values returns `#VALUE!`.
- `IF`, `IFS`, and `SWITCH` can return scalar text when the selected branch or result is text.

### Aggregate and count helpers

- `SUM`
- `PRODUCT`
- `SUMSQ`
- `SUMPRODUCT`
- `SUBTOTAL`
- `SUMXMY2`
- `SUMX2MY2`
- `SUMX2PY2`
- `MIN`
- `MAX`
- `MINA`
- `MAXA`
- `MEDIAN`
- `AVERAGE`
- `AVERAGEA`
- `GEOMEAN`
- `HARMEAN`
- `MODE`
- `MODE.SNGL`
- `TRIMMEAN`
- `AVEDEV`
- `DEVSQ`
- `VAR`
- `VAR.P`
- `VAR.S`
- `VARP`
- `VARA`
- `VARPA`
- `STDEV`
- `STDEV.P`
- `STDEV.S`
- `STDEVP`
- `STDEVA`
- `STDEVPA`
- `CORREL`
- `PEARSON`
- `COVAR`
- `COVARIANCE.P`
- `COVARIANCE.S`
- `SLOPE`
- `INTERCEPT`
- `RSQ`
- `LARGE`
- `SMALL`
- `PERCENTILE`
- `PERCENTILE.INC`
- `PERCENTILE.EXC`
- `PERCENTRANK`
- `PERCENTRANK.INC`
- `PERCENTRANK.EXC`
- `QUARTILE`
- `QUARTILE.INC`
- `QUARTILE.EXC`
- `RANK`
- `RANK.EQ`
- `RANK.AVG`
- `COUNT`
- `COUNTA`
- `COUNTBLANK`

Aggregate notes:

- `SUMPRODUCT` supports scalar and rectangular range arguments with matching shapes; non-numeric range entries are treated as zero and error cells are propagated.
- `SUBTOTAL` supports function numbers `1` through `11` and `101` through `111`, skips nested top-level `SUBTOTAL` formulas in range arguments, and currently treats hidden or filtered rows the same as visible rows because row visibility/filter state is not modeled yet.

### Criteria aggregate helpers

- `COUNTIF`
- `SUMIF`
- `AVERAGEIF`
- `COUNTIFS`
- `SUMIFS`
- `AVERAGEIFS`
- `MINIFS`
- `MAXIFS`

### Lookup and reference helpers

- `INDEX`
- `MATCH`
- `XMATCH`
- `VLOOKUP`
- `HLOOKUP`
- `XLOOKUP`
- `ROW`
- `COLUMN`
- `ROWS`
- `COLUMNS`
- `AREAS`
- `ADDRESS`
- `SHEET`
- `SHEETS`
- `CHOOSE`

Lookup notes:

- `MATCH` supports exact match, ascending approximate match, and descending approximate match over one-dimensional ranges.
- `XMATCH` supports exact match, wildcard match, exact-or-next-smaller, exact-or-next-larger, and forward or reverse linear search over one-dimensional ranges.
- `VLOOKUP` and `HLOOKUP` support exact and ascending approximate table lookup.
- `XLOOKUP` supports scalar one-dimensional lookup and return arrays, `if_not_found`, exact match, wildcard match, exact-or-next-smaller, exact-or-next-larger, and forward or reverse linear search.
- `ROW()` and `COLUMN()` without arguments resolve against the formula cell position during recalculation.
- `AREAS` returns `1` for supported single-area references; `ADDRESS` supports A1/R1C1 text output with absolute/relative flags and optional sheet text.
- `SHEET()` returns the current worksheet's 1-based workbook position; `SHEETS()` returns the workbook worksheet count, and reference arguments are supported for single-sheet references.
- Lookup comparisons support numbers, booleans, and case-insensitive text.
- `INDEX`, `VLOOKUP`, `HLOOKUP`, and `XLOOKUP` return scalar text results as `CellValue::Text`; numeric results still flow through the numeric evaluator.
- `CHOOSE` can return scalar text when the selected argument is text.

### Date and time serial helpers

- `DATE`
- `YEAR`
- `MONTH`
- `DAY`
- `DAYS`
- `DATEDIF`
- `DAYS360`
- `YEARFRAC`
- `EDATE`
- `EOMONTH`
- `DATEVALUE`
- `WEEKDAY`
- `WEEKNUM`
- `ISOWEEKNUM`
- `WORKDAY`
- `WORKDAY.INTL`
- `NETWORKDAYS`
- `NETWORKDAYS.INTL`
- `TIME`
- `TIMEVALUE`
- `HOUR`
- `MINUTE`
- `SECOND`

Date/time notes:

- Date helpers use Excel's 1900 date system, including the compatibility serial `60` for `1900-02-29`.
- `DATE` supports month and day rollover, such as month `13` and day `0`.
- `DAYS360` supports the U.S. (NASD) and European methods.
- `YEARFRAC` supports bases `0` through `4`, with basis `0`/`4` using the matching 30/360 day-count variants.
- `WEEKDAY` supports return types `1`, `2`, `3`, and `11` through `17`; `WEEKNUM` supports return types `1`, `2`, `11` through `17`, and ISO return type `21`.
- `WORKDAY.INTL` and `NETWORKDAYS.INTL` support Excel weekend codes `1` through `7` and `11` through `17`, plus seven-character weekend masks.
- `TIME` produces fractional-day serial values, with hour/minute/second rollover for non-negative arguments.
- `DATEVALUE` supports ISO-like `yyyy-mm-dd`, `yyyy/m/d`, and numeric `m/d/yyyy` text.
- `TIMEVALUE` supports `h:mm`, `h:mm:ss`, and `AM`/`PM` suffixes.
- Locale-sensitive date/time text parsing beyond that focused subset and volatile clock functions are not implemented.

Criteria notes:

- `*IFS` family currently requires criteria ranges and value ranges to have the same shape.
- criteria strings support numeric comparison prefixes such as `">3"`, text equality/inequality prefixes such as `"=north*"` and `"<>north*"`, the blank-string criterion `""`, and the nonblank criterion `"<>"`.
- Text criteria support Excel-style `*` and `?` wildcards with `~` escaping.

### Error and information helpers

- `NA`
- `N`
- `TYPE`
- `ERROR.TYPE`
- `IFERROR`
- `IFNA`
- `ISERROR`
- `ISERR`
- `ISNA`
- `ISBLANK`
- `ISLOGICAL`
- `ISNONTEXT`
- `ISNUMBER`
- `ISREF`
- `ISFORMULA`
- `ISTEXT`
- `FORMULATEXT`

### Text helpers

- `CONCAT`
- `CONCATENATE`
- `BASE`
- `CHAR`
- `CLEAN`
- `CODE`
- `DOLLAR`
- `FIXED`
- `LEFT`
- `RIGHT`
- `MID`
- `T`
- `UNICHAR`
- `UNICODE`
- `UPPER`
- `LOWER`
- `PROPER`
- `TRIM`
- `TEXTJOIN`
- `TEXTBEFORE`
- `TEXTAFTER`
- `REPT`
- `REPLACE`
- `SUBSTITUTE`
- `LEN`
- `FIND`
- `SEARCH`
- `EXACT`
- `VALUE`
- `NUMBERVALUE`

Text notes:

- text-returning helpers produce `CellValue::Text` and support scalar literals, booleans, numbers, single-cell references, and nested text helpers.
- `FIND` is case-sensitive, `SEARCH` is case-insensitive, and both return 1-based character positions.
- `SEARCH`, `XLOOKUP`, and `XMATCH` support Excel-style `*` and `?` wildcards with `~` escaping in their focused scalar paths.
- `TEXTJOIN` supports scalar values and rectangular ranges, and `TEXTBEFORE` / `TEXTAFTER` support scalar delimiters, instance numbers, case-sensitivity mode, match-end mode, and scalar `if_not_found` fallbacks.
- `VALUE` supports plain decimal numeric text and a trailing percent sign; `NUMBERVALUE` additionally supports configurable decimal and group separators. Currency symbols, date/time text, and locale-specific formatting beyond those focused paths are not implemented.
- `FIXED` uses period decimal text and optional comma grouping; `DOLLAR` uses the invariant `$` currency symbol and Excel-style parentheses for negative values.
- `REPT` returns `#VALUE!` for outputs beyond Excel's 32,767-character cell text limit.
- Range flattening, locale-aware formatting, and full Excel text coercion are not implemented yet.

## Current Boundaries

- The evaluator is still numeric-first. It includes focused scalar text and date/time text-parse subsets, but broader string semantics, locale-sensitive date/time parsing, and richer coercion rules are not implemented.
- Dynamic array behavior and broader `Formula2` parity are not implemented.
- Lookup/reference support is still a focused scalar subset. It does not model named ranges, array-returning `INDEX(..., 0, ...)` or `XLOOKUP`, binary search modes, external references, or broader lookup/reference families yet.
- Named ranges, multi-area references, and 3D references are not implemented.
- Unsupported formulas should fail predictably, but the runtime does not yet aim for 1:1 compatibility with Excel's full calculation engine.

## Test Strategy

- Calculation regressions currently live in `crates/excel-runtime/src/lib.rs` as synthetic workbook tests.
- New formula support should continue to land with targeted regression coverage and a full `cargo test --workspace --quiet` pass.
