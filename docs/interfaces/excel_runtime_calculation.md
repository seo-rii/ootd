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
- `MOD`
- `SIGN`
- `POWER`
- `SQRT`
- `ISEVEN`
- `ISODD`

### Logical and control helpers

- `AND`
- `OR`
- `NOT`
- `IF`

### Aggregate and count helpers

- `SUM`
- `PRODUCT`
- `MIN`
- `MAX`
- `AVERAGE`
- `COUNT`
- `COUNTA`
- `COUNTBLANK`

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
- `VLOOKUP`
- `HLOOKUP`
- `ROW`
- `COLUMN`
- `ROWS`
- `COLUMNS`
- `CHOOSE`

Lookup notes:

- `MATCH` supports exact match, ascending approximate match, and descending approximate match over one-dimensional ranges.
- `VLOOKUP` and `HLOOKUP` support exact and ascending approximate table lookup.
- `ROW()` and `COLUMN()` without arguments resolve against the formula cell position during recalculation.
- Lookup comparisons support numbers, booleans, and case-insensitive text.
- `INDEX`, `VLOOKUP`, and `HLOOKUP` return scalar text results as `CellValue::Text`; numeric results still flow through the numeric evaluator.
- `IF` and `CHOOSE` can return scalar text when the selected branch or argument is text.

### Date and time serial helpers

- `DATE`
- `YEAR`
- `MONTH`
- `DAY`
- `DAYS`
- `EDATE`
- `EOMONTH`
- `TIME`
- `HOUR`
- `MINUTE`
- `SECOND`

Date/time notes:

- Date helpers use Excel's 1900 date system, including the compatibility serial `60` for `1900-02-29`.
- `DATE` supports month and day rollover, such as month `13` and day `0`.
- `TIME` produces fractional-day serial values, with hour/minute/second rollover for non-negative arguments.
- Locale-sensitive date/time text parsing and volatile clock functions are not implemented.

Criteria notes:

- `*IFS` family currently requires criteria ranges and value ranges to have the same shape.
- criteria strings support numeric comparison prefixes such as `">3"` and the blank-string criterion `""`.

### Error and information helpers

- `NA`
- `IFERROR`
- `IFNA`
- `ISERROR`
- `ISERR`
- `ISNA`
- `ISBLANK`
- `ISNUMBER`
- `ISTEXT`

### Text helpers

- `CONCAT`
- `CONCATENATE`
- `LEFT`
- `RIGHT`
- `MID`
- `UPPER`
- `LOWER`
- `TRIM`
- `REPT`
- `REPLACE`
- `SUBSTITUTE`
- `LEN`
- `FIND`
- `SEARCH`
- `EXACT`
- `VALUE`

Text notes:

- text-returning helpers produce `CellValue::Text` and support scalar literals, booleans, numbers, single-cell references, and nested text helpers.
- `FIND` is case-sensitive, `SEARCH` is case-insensitive, and both return 1-based character positions.
- `VALUE` supports plain decimal numeric text and a trailing percent sign, but not currency symbols, date/time text, or locale-specific separators.
- `REPT` returns `#VALUE!` for outputs beyond Excel's 32,767-character cell text limit.
- Range flattening, wildcard matching in `SEARCH`, locale-aware formatting, and full Excel text coercion are not implemented yet.

## Current Boundaries

- The evaluator is still numeric-first. It includes a focused scalar text helper subset, but broader string semantics, date/time text parsing, and richer coercion rules are not implemented.
- Dynamic array behavior and broader `Formula2` parity are not implemented.
- Lookup/reference support is still a focused scalar subset. It does not model named ranges, array-returning `INDEX(..., 0, ...)`, external references, or broader lookup families such as `XLOOKUP` yet.
- Named ranges, multi-area references, and 3D references are not implemented.
- Unsupported formulas should fail predictably, but the runtime does not yet aim for 1:1 compatibility with Excel's full calculation engine.

## Test Strategy

- Calculation regressions currently live in `crates/excel-runtime/src/lib.rs` as synthetic workbook tests.
- New formula support should continue to land with targeted regression coverage and a full `cargo test --workspace --quiet` pass.
