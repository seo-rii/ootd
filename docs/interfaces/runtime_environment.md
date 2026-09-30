# Runtime Environment And Calculation Context

Status: synthetic contract (`OOTD-038`/`OOTD-039`/`OOTD-047`). Desktop Excel evidence has not been
captured for any row below.

Formula evaluation does not read ambient process state. The host supplies a
`RuntimeEnvironment`, and every calculation cycle evaluates against one `CalcContext` derived from
it.

## RuntimeEnvironment

`ExcelRuntime::environment()` returns the active environment and
`ExcelRuntime::set_environment(environment)` replaces it.

| Field | Meaning | Default |
| --- | --- | --- |
| `clock` | `RuntimeClock::System` reads the host clock; `RuntimeClock::Fixed(instant)` reports one instant | `System` |
| `utc_offset_minutes` | Offset added to the clock reading to obtain workbook-local time | `0` (UTC) |
| `random_seed` | Seed of the session random stream; `None` seeds once from the system clock | `None` |

`set_environment` restarts the session random stream from the new environment's seed, so the same
seed and the same sequence of calculations reproduce the same random values.

## CalcContext

One context is opened per calculation cycle:

- `Application.Calculate`, automatic recalculation, and `calculate_workbook_with_report` open one
  context per workbook, shared by all of its worksheets;
- `Worksheet.Calculate` and `Range.Calculate` open one context for that sheet or range;
- `Application.Evaluate`, `Worksheet.Evaluate`, and `Application.WorksheetFunction` open one
  context per expression.

The context captures:

- **Clock:** one reading of the environment clock, shifted by `utc_offset_minutes` and converted
  to a serial in the workbook date system (1900, or 1904 when `Workbook.Date1904` is true, which
  subtracts 1,462 days). `NOW()` returns that serial and `TODAY()` its floor, so every cell in one
  cycle observes the same instant. A clock reading before the Unix epoch yields `#NUM!`.
- **Random stream:** `RAND`, `RANDBETWEEN`, and `RANDARRAY` draw from the session stream held by
  the runtime. The context works on a copy and the runtime stores the advanced state when the
  cycle ends, so consecutive cycles continue one deterministic sequence. Draws made while `Find`
  inspects formula results are discarded.

The regressions are in `crates/excel-runtime/src/tests/runtime_environment.rs`.

## Remaining Boundaries

- The offset is fixed; daylight-saving transitions and named time zones are the host's
  responsibility.
- Locale-sensitive parsing and formatting are not part of the environment yet (`OOTD-076`).
- Date functions other than `NOW`/`TODAY` still interpret serials in the 1900 date system.
