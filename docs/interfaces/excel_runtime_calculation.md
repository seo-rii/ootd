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
- `CONFIDENCE` / `CONFIDENCE.NORM`
- `PERMUT`
- `PERMUTATIONA`
- `MULTINOMIAL`
- `POWER`
- `SQRT`
- `SQRTPI`
- `ERF`
- `ERF.PRECISE`
- `ERFC`
- `ERFC.PRECISE`
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
- `RAND`
- `RANDBETWEEN`
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
- `BESSELI`
- `BESSELJ`
- `BESSELK`
- `BESSELY`
- `BETA.DIST` / `BETADIST`
- `BETA.INV` / `BETAINV`
- `BIN2DEC`
- `BIN2HEX`
- `BIN2OCT`
- `CONVERT`
- `EUROCONVERT`
- `DEC2BIN`
- `DEC2HEX`
- `DEC2OCT`
- `HEX2BIN`
- `HEX2DEC`
- `HEX2OCT`
- `OCT2BIN`
- `OCT2DEC`
- `OCT2HEX`
- `COMPLEX`
- `IMABS`
- `IMAGINARY`
- `IMARGUMENT`
- `IMCONJUGATE`
- `IMCOS`
- `IMCOSH`
- `IMCOT`
- `IMCSC`
- `IMCSCH`
- `IMDIV`
- `IMEXP`
- `IMLN`
- `IMLOG10`
- `IMLOG2`
- `IMPOWER`
- `IMPRODUCT`
- `IMREAL`
- `IMSEC`
- `IMSECH`
- `IMSIN`
- `IMSINH`
- `IMSQRT`
- `IMSUB`
- `IMSUM`
- `IMTAN`
- `BINOM.DIST` / `BINOMDIST`
- `BINOM.DIST.RANGE`
- `BINOM.INV` / `CRITBINOM`
- `CHIDIST`
- `CHIINV`
- `CHISQ.DIST`
- `CHISQ.DIST.RT`
- `CHISQ.INV`
- `CHISQ.INV.RT`
- `CHISQ.TEST` / `CHITEST`
- `CONFIDENCE.T`
- `F.DIST`
- `F.DIST.RT` / `FDIST`
- `F.INV`
- `F.INV.RT` / `FINV`
- `F.TEST` / `FTEST`
- `FISHER`
- `FISHERINV`
- `FORECAST`
- `FORECAST.LINEAR`
- `GAMMA`
- `GAMMA.DIST` / `GAMMADIST`
- `GAMMA.INV` / `GAMMAINV`
- `GAMMALN`
- `GAMMALN.PRECISE`
- `EXPON.DIST` / `EXPONDIST`
- `GAUSS`
- `HYPGEOM.DIST` / `HYPGEOMDIST`
- `KURT`
- `MDETERM`
- `GROWTH`
- `LINEST`
- `LOGEST`
- `LOGNORM.DIST` / `LOGNORMDIST`
- `LOGNORM.INV` / `LOGINV`
- `NEGBINOM.DIST` / `NEGBINOMDIST`
- `NORM.DIST` / `NORMDIST`
- `NORM.INV` / `NORMINV`
- `NORM.S.DIST` / `NORMSDIST`
- `NORM.S.INV` / `NORMSINV`
- `PHI`
- `POISSON.DIST` / `POISSON`
- `PROB`
- `PERCENTOF`
- `SKEW`
- `SKEW.P`
- `STANDARDIZE`
- `STEYX`
- `T.DIST`
- `T.DIST.2T`
- `T.DIST.RT`
- `T.INV`
- `T.INV.2T`
- `T.TEST` / `TTEST`
- `TDIST`
- `TINV`
- `TREND`
- `WEIBULL.DIST` / `WEIBULL`
- `Z.TEST` / `ZTEST`

Roman numeral notes:

- `ROMAN` supports classic and simplified forms `0` through `4`; omitted form and `TRUE` use classic form, while `FALSE` uses simplified form.
- `ARABIC` accepts classic and simplified Roman numerals, ignores case and surrounding spaces, returns `0` for empty text, and supports a leading negative sign.
- `SERIESSUM` supports scalar or rectangular coefficient inputs and ignores non-numeric cells in coefficient ranges.
- `FISHER` and `FISHERINV` cover the Fisher transformation and its inverse for scalar numeric values.
- `ERF`, `ERF.PRECISE`, `ERFC`, and `ERFC.PRECISE` cover scalar error-function calculations. `ERF` also supports the legacy lower/upper integration form.
- `GAMMALN` and `GAMMALN.PRECISE` return the natural logarithm of the gamma function for positive scalar inputs.
- `BETA.DIST`, `CHISQ.DIST`, `EXPON.DIST`, `F.DIST`, `GAMMA.DIST`, `LOGNORM.DIST`, `NORM.DIST`, `NORM.S.DIST`, `POISSON.DIST`, `T.DIST`, and `WEIBULL.DIST` cover scalar probability and cumulative distribution calculations, with legacy function aliases mapped to the same implementations where their argument shapes match.
- `CHISQ.TEST` / `CHITEST` compare same-shaped observed and expected ranges and return the right-tailed chi-square probability.
- `F.TEST` / `FTEST`, `T.TEST` / `TTEST`, and `Z.TEST` / `ZTEST` cover variance, mean, and z-test probability calculations over numeric arguments and ranges.
- `BINOM.DIST`, `BINOM.DIST.RANGE`, `BINOM.INV`, `HYPGEOM.DIST`, and `NEGBINOM.DIST` cover scalar discrete statistical distribution calculations, including compatibility aliases where Excel exposes them.
- `KURT`, `SKEW`, `SKEW.P`, `PROB`, `PERCENTOF`, `MODE.MULT`, and `MDETERM` cover focused scalar statistical and matrix calculations over numeric arguments and ranges.
- `GAMMA`, `CONFIDENCE.NORM`, `CONFIDENCE.T`, `FORECAST` / `FORECAST.LINEAR`, `LINEST`, `LOGEST`, `TREND`, `GROWTH`, and `STEYX` cover scalar gamma, confidence interval, and linear/exponential regression calculations.
- `BETA.INV`, `CHISQ.INV`, `F.INV`, `GAMMA.INV`, `LOGNORM.INV`, `NORM.INV`, `NORM.S.INV`, and `T.INV` cover scalar inverse cumulative distribution calculations and include the legacy aliases currently listed above.
- Continuous distribution helpers use iterative approximations for regularized beta/gamma and inverse CDF calculations, so they target practical worksheet compatibility rather than bit-for-bit Excel parity.
- `GAUSS`, `PHI`, and `STANDARDIZE` cover scalar standard-normal helpers.
- Complex engineering helpers parse Excel-style text values with lowercase `i` or `j` suffixes; text-returning `IM*` helpers preserve the suffix when the result includes an imaginary component.
- `BESSELI`, `BESSELJ`, `BESSELK`, and `BESSELY` support non-negative integer orders, truncating fractional order arguments to match Excel's scalar argument handling.
- `CONVERT` supports Excel's documented measurement groups, case-sensitive unit aliases, metric prefixes, binary prefixes for information units, and affine temperature conversions.
- `EUROCONVERT` supports the documented legacy euro member currency codes, fixed EU rates, currency-specific rounding, full-precision mode, and triangulation precision for member-currency-to-member-currency conversion.

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
- `AMORDEGRC`
- `AMORLINC`
- `VDB`
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
- `AMORDEGRC` covers French-accounting degressive depreciation, including coefficient selection, integer depreciation rounding, and final-period rate changes.
- `AMORLINC` covers French-accounting linear depreciation, including prorated first periods and salvage-value caps.
- `VDB` covers variable declining-balance depreciation over full or partial periods, including the optional straight-line switch.
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
- `AGGREGATE`
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
- `MODE.MULT`
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
- `AGGREGATE` supports function numbers `1` through `19`; options that ignore nested `SUBTOTAL` / `AGGREGATE` formulas and error values are honored, while hidden-row options are treated as visible because row visibility/filter state is not modeled yet.
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
- `LOOKUP`
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
- `CELL`
- `SHEET`
- `SHEETS`
- `CHOOSE`

Lookup notes:

- `MATCH` supports exact match, ascending approximate match, and descending approximate match over one-dimensional ranges.
- `LOOKUP` supports the vector form with ascending approximate match, using the lookup vector as the result vector when the third argument is omitted.
- `XMATCH` supports exact match, wildcard match, exact-or-next-smaller, exact-or-next-larger, and forward or reverse linear search over one-dimensional ranges.
- `VLOOKUP` and `HLOOKUP` support exact and ascending approximate table lookup.
- `XLOOKUP` supports scalar one-dimensional lookup and return arrays, `if_not_found`, exact match, wildcard match, exact-or-next-smaller, exact-or-next-larger, and forward or reverse linear search.
- `ROW()` and `COLUMN()` without arguments resolve against the formula cell position during recalculation.
- `AREAS` returns `1` for supported single-area references; `ADDRESS` supports A1/R1C1 text output with absolute/relative flags and optional sheet text.
- `CELL` supports focused metadata types such as `address`, `row`, `col`, `contents`, `type`, `format`, `filename`, `width`, `prefix`, `color`, `parentheses`, and `protect` for the upper-left cell of a reference.
- `SHEET()` returns the current worksheet's 1-based workbook position; `SHEETS()` returns the workbook worksheet count, and reference arguments are supported for single-sheet references.
- Lookup comparisons support numbers, booleans, and case-insensitive text.
- `INDEX`, `LOOKUP`, `VLOOKUP`, `HLOOKUP`, and `XLOOKUP` return scalar text results as `CellValue::Text`; numeric results still flow through the numeric evaluator.
- `CHOOSE` can return scalar text when the selected argument is text.

### Date and time serial helpers

- `DATE`
- `TODAY`
- `NOW`
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
- `TODAY` and `NOW` return volatile serial values from the host clock; `NOW` includes the fractional day.
- Locale-sensitive date/time text parsing beyond that focused subset is not implemented.

Criteria notes:

- `*IFS` family currently requires criteria ranges and value ranges to have the same shape.
- criteria strings support numeric comparison prefixes such as `">3"`, text equality/inequality prefixes such as `"=north*"` and `"<>north*"`, the blank-string criterion `""`, and the nonblank criterion `"<>"`.
- Text criteria support Excel-style `*` and `?` wildcards with `~` escaping.

### Database helpers

- `DAVERAGE`
- `DCOUNT`
- `DCOUNTA`
- `DGET`
- `DMAX`
- `DMIN`
- `DPRODUCT`
- `DSTDEV`
- `DSTDEVP`
- `DSUM`
- `DVAR`
- `DVARP`

Database notes:

- `D*` helpers support a rectangular database with header row, numeric field indexes or field-name text, and criteria ranges where rows are OR clauses and nonblank columns are AND criteria.
- Database criteria reuse the same numeric/text/wildcard matching rules as the criteria aggregate helpers.

### Error and information helpers

- `NA`
- `INFO`
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

Information notes:

- `INFO` supports focused workbook/runtime metadata types including `directory`, `numfile`, `origin`, `osversion`, `recalc`, `release`, and `system`.

### Text helpers

- `ARRAYTOTEXT`
- `ASC`
- `BAHTTEXT`
- `CONCAT`
- `CONCATENATE`
- `BASE`
- `CHAR`
- `CLEAN`
- `CODE`
- `DBCS`
- `DOLLAR`
- `FIXED`
- `HYPERLINK`
- `LEFT`
- `LEFTB`
- `RIGHT`
- `RIGHTB`
- `MID`
- `MIDB`
- `JIS`
- `PHONETIC`
- `T`
- `TEXT`
- `UNICHAR`
- `UNICODE`
- `UPPER`
- `LOWER`
- `PROPER`
- `TRIM`
- `TEXTJOIN`
- `TEXTBEFORE`
- `TEXTAFTER`
- `REGEXEXTRACT`
- `REGEXREPLACE`
- `REGEXTEST`
- `REPT`
- `REPLACE`
- `REPLACEB`
- `SUBSTITUTE`
- `LEN`
- `LENB`
- `FIND`
- `FINDB`
- `SEARCH`
- `SEARCHB`
- `EXACT`
- `VALUE`
- `VALUETOTEXT`
- `NUMBERVALUE`

Text notes:

- text-returning helpers produce `CellValue::Text` and support scalar literals, booleans, numbers, single-cell references, and nested text helpers.
- `FIND` is case-sensitive, `SEARCH` is case-insensitive, and both return 1-based character positions. The `*B` variants use focused byte-width positions, treating ASCII as width 1 and non-ASCII scalar values as width 2.
- `SEARCH`, `XLOOKUP`, and `XMATCH` support Excel-style `*` and `?` wildcards with `~` escaping in their focused scalar paths.
- `REGEXTEST`, `REGEXEXTRACT`, and `REGEXREPLACE` support scalar regex matching, extraction, and replacement using Rust regex syntax; array-spill return modes currently resolve to a scalar focused result.
- `TEXTJOIN` supports scalar values and rectangular ranges, and `TEXTBEFORE` / `TEXTAFTER` support scalar delimiters, instance numbers, case-sensitivity mode, match-end mode, and scalar `if_not_found` fallbacks.
- `VALUE` supports plain decimal numeric text and a trailing percent sign; `NUMBERVALUE` additionally supports configurable decimal and group separators. Currency symbols, date/time text, and locale-specific formatting beyond those focused paths are not implemented.
- `FIXED` uses period decimal text and optional comma grouping; `DOLLAR` uses the invariant `$` currency symbol and Excel-style parentheses for negative values.
- `TEXT` supports a focused invariant numeric-format subset with `0` / `#` placeholders, decimal places, comma grouping, `%`, `$`, quoted literals, escapes, and positive/negative/zero sections; date/time and locale-specific format codes still return `#VALUE!`.
- `BAHTTEXT` rounds to satang precision and returns Thai Baht/Satang text with `บาทถ้วน` or `สตางค์` suffixes.
- `HYPERLINK` returns the friendly name when supplied, otherwise the link text; it does not model worksheet hyperlink navigation metadata.
- `PHONETIC` preserves scalar text and concatenates text cells from a reference because the current object model does not store phonetic guide metadata.
- `VALUETOTEXT` and `ARRAYTOTEXT` support concise and strict formatting for scalar values and rectangular references; strict text values use Excel-style doubled quotes.
- `ASC`, `DBCS`, and `JIS` preserve text in the current non-DBCS focused path.
- `REPT` returns `#VALUE!` for outputs beyond Excel's 32,767-character cell text limit.
- General-purpose range flattening beyond the listed helpers, locale-aware formatting, and full Excel text coercion are not implemented yet.

### Web helpers

- `ENCODEURL`

Web notes:

- `ENCODEURL` percent-encodes UTF-8 bytes and leaves ASCII letters, digits, `-`, `_`, `.`, and `~` unescaped.

## Current Boundaries

- The evaluator is still numeric-first. It includes focused scalar text and date/time text-parse subsets, but broader string semantics, locale-sensitive date/time parsing, and richer coercion rules are not implemented.
- Dynamic array behavior and broader `Formula2` parity are not implemented.
- Lookup/reference support is still a focused scalar subset. It does not model named ranges, array-returning `INDEX(..., 0, ...)` or `XLOOKUP`, binary search modes, external references, or broader lookup/reference families yet.
- Named ranges, multi-area references, and 3D references are not implemented.
- Unsupported formulas should fail predictably, but the runtime does not yet aim for 1:1 compatibility with Excel's full calculation engine.

## Test Strategy

- Calculation regressions currently live in `crates/excel-runtime/src/lib.rs` as synthetic workbook tests.
- New formula support should continue to land with targeted regression coverage and a full `cargo test --workspace --quiet` pass.
