use super::*;
use crate::{RuntimeClock, RuntimeEnvironment};
use std::time::Duration;

/// 2026-01-02T03:04:05Z, which is 2026-01-02 12:04:05 at UTC+09:00.
const FIXED_UNIX_SECONDS: u64 = 1_767_323_045;
const FIXED_LOCAL_SERIAL: f64 = 46_024.0 + (12.0 * 3_600.0 + 4.0 * 60.0 + 5.0) / 86_400.0;

fn open_runtime(
    environment: Option<RuntimeEnvironment>,
) -> (ExcelRuntime, WorkbookHandle, ObjectHandle) {
    let mut runtime = ExcelRuntime::new();
    if let Some(environment) = environment {
        runtime.set_environment(environment);
    }
    let workbook = runtime
        .open_workbook(OpenWorkbookSpec {
            bytes: synthetic_workbook_bytes(),
            format_hint: Some(FileFormat::Xlsx),
            profile: ExcelProfile::Excel365,
            read_only: false,
        })
        .expect("open environment workbook");
    let active_sheet = expect_object_handle(
        runtime
            .dispatch_get(runtime.root_application(), "ActiveSheet", &[])
            .expect("ActiveSheet"),
    );
    (runtime, workbook, active_sheet)
}

fn fixed_environment(utc_offset_minutes: i32, random_seed: Option<u64>) -> RuntimeEnvironment {
    RuntimeEnvironment {
        clock: RuntimeClock::Fixed(UNIX_EPOCH + Duration::from_secs(FIXED_UNIX_SECONDS)),
        utc_offset_minutes,
        random_seed,
    }
}

fn range(runtime: &mut ExcelRuntime, sheet: ObjectHandle, address: &str) -> ObjectHandle {
    expect_object_handle(
        runtime
            .dispatch_invoke(sheet, "Range", &[OmValue::Text(address.to_string())])
            .unwrap_or_else(|error| panic!("Range({address}): {error:?}")),
    )
}

fn set_formula(runtime: &mut ExcelRuntime, sheet: ObjectHandle, address: &str, formula: &str) {
    let target = range(runtime, sheet, address);
    runtime
        .dispatch_set(target, "Formula", OmValue::Text(formula.to_string()), &[])
        .unwrap_or_else(|error| panic!("set {address} formula: {error:?}"));
}

fn calculate(runtime: &mut ExcelRuntime) {
    runtime
        .dispatch_invoke(runtime.root_application(), "Calculate", &[])
        .expect("Calculate");
}

fn values(runtime: &mut ExcelRuntime, sheet: ObjectHandle, addresses: &[&str]) -> Vec<f64> {
    addresses
        .iter()
        .map(|address| {
            let cell = range(runtime, sheet, address);
            expect_number(
                runtime
                    .dispatch_get(cell, "Value2", &[])
                    .unwrap_or_else(|error| panic!("{address} Value2: {error:?}")),
            )
        })
        .collect()
}

#[test]
fn runtime_environment_clock_and_offset_drive_now_and_today() {
    let (mut runtime, workbook, sheet) = open_runtime(Some(fixed_environment(540, None)));
    set_formula(&mut runtime, sheet, "M1", "=NOW()");
    set_formula(&mut runtime, sheet, "M2", "=TODAY()");
    calculate(&mut runtime);
    let [now, today] = values(&mut runtime, sheet, &["M1", "M2"])[..] else {
        unreachable!()
    };
    assert!((now - FIXED_LOCAL_SERIAL).abs() < 1e-9, "NOW() = {now}");
    assert_eq!(today, 46_024.0);

    runtime
        .dispatch_set(workbook.0, "Date1904", OmValue::Bool(true), &[])
        .expect("switch to the 1904 date system");
    calculate(&mut runtime);
    let [now, today] = values(&mut runtime, sheet, &["M1", "M2"])[..] else {
        unreachable!()
    };
    assert!(
        (now - (FIXED_LOCAL_SERIAL - 1_462.0)).abs() < 1e-9,
        "1904 NOW() = {now}"
    );
    assert_eq!(today, 46_024.0 - 1_462.0);
}

#[test]
fn runtime_environment_reads_the_clock_once_per_calculation() {
    let (mut runtime, _, sheet) = open_runtime(None);
    let addresses = (1..=40).map(|row| format!("M{row}")).collect::<Vec<_>>();
    for address in &addresses {
        set_formula(&mut runtime, sheet, address, "=NOW()+0");
    }
    calculate(&mut runtime);
    let addresses = addresses.iter().map(String::as_str).collect::<Vec<_>>();
    let readings = values(&mut runtime, sheet, &addresses);
    assert!(
        readings.iter().all(|reading| *reading == readings[0]),
        "one calculation must observe one clock reading: {readings:?}"
    );
}

#[test]
fn runtime_environment_seed_makes_random_functions_reproducible() {
    let formulas = [
        ("M1", "=RAND()"),
        ("M2", "=RAND()"),
        ("M3", "=RANDBETWEEN(1,1000000000)"),
        ("M4", "=SUM(RANDARRAY(2,2))"),
    ];
    let addresses = formulas.map(|(address, _)| address);
    let sample = |seed: u64| {
        let (mut runtime, _, sheet) = open_runtime(Some(fixed_environment(0, Some(seed))));
        for (address, formula) in formulas {
            set_formula(&mut runtime, sheet, address, formula);
        }
        calculate(&mut runtime);
        let first = values(&mut runtime, sheet, &addresses);
        calculate(&mut runtime);
        let second = values(&mut runtime, sheet, &addresses);
        (first, second)
    };

    let (first, second) = sample(7);
    assert_eq!(sample(7), (first.clone(), second.clone()));
    assert_ne!(first, second, "each calculation advances the random stream");
    assert_ne!(
        sample(8).0,
        first,
        "a different seed yields a different stream"
    );
}

#[test]
fn runtime_environment_round_trips_through_the_public_api() {
    let mut runtime = ExcelRuntime::new();
    assert_eq!(runtime.environment(), &RuntimeEnvironment::default());
    assert_eq!(RuntimeEnvironment::default().clock, RuntimeClock::System);
    let environment = fixed_environment(-300, Some(11));
    runtime.set_environment(environment.clone());
    assert_eq!(runtime.environment(), &environment);
}
