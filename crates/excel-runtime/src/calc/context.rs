//! The per-cycle calculation context: clock reading, date system, random stream, and memoized
//! formula results.

use super::*;

/// The 1900-system serial of the system clock in UTC, used by tests that bracket `NOW()` under the
/// default environment.
#[cfg(test)]
pub(crate) fn formula_current_excel_serial() -> Result<f64, FormulaEvalError> {
    let elapsed = std::time::SystemTime::now()
        .duration_since(std::time::UNIX_EPOCH)
        .map_err(|_| FormulaEvalError::Num)?;
    Ok(25_569.0 + elapsed.as_secs_f64() / 86_400.0)
}

/// The deterministic inputs of one calculation cycle: a single clock reading projected into the
/// workbook date system, and the session random stream, which advances only through this context.
pub(crate) struct CalcContext {
    now_serial: Option<f64>,
    date_system: DateSystem,
    random_state: std::cell::Cell<u64>,
    cell_results: std::cell::RefCell<CalcCellResults>,
}

pub(super) type CalcCellResults = std::collections::BTreeMap<(SheetId, u32, u32), CalcCellResult>;
pub(super) type CalcCellResult = Result<CellValue, FormulaEvalError>;

impl CalcContext {
    pub(crate) fn new(environment: &RuntimeEnvironment, date1904: bool, random_state: u64) -> Self {
        let date_system_offset = if date1904 { 1_462.0 } else { 0.0 };
        Self {
            now_serial: environment
                .local_unix_seconds()
                .map(|seconds| 25_569.0 + seconds / 86_400.0 - date_system_offset),
            date_system: if date1904 {
                DateSystem::Excel1904
            } else {
                DateSystem::Excel1900
            },
            random_state: std::cell::Cell::new(random_state),
            cell_results: std::cell::RefCell::new(CalcCellResults::new()),
        }
    }

    /// Discards memoized formula results after the cycle changes cell values they may read, such
    /// as newly committed spill ranges.
    pub(crate) fn forget_cell_results(&self) {
        self.cell_results.borrow_mut().clear();
    }

    pub(super) fn cell_result(&self, key: (SheetId, u32, u32)) -> Option<CalcCellResult> {
        self.cell_results.borrow().get(&key).cloned()
    }

    pub(super) fn remember_cell_result(&self, key: (SheetId, u32, u32), result: &CalcCellResult) {
        self.cell_results.borrow_mut().insert(key, result.clone());
    }

    pub(super) fn forget_cell_result(&self, key: (SheetId, u32, u32)) {
        self.cell_results.borrow_mut().remove(&key);
    }

    pub(super) fn date_system(&self) -> DateSystem {
        self.date_system
    }

    /// The random stream state after every draw made through this context.
    pub(crate) fn random_state(&self) -> u64 {
        self.random_state.get()
    }

    pub(super) fn now_serial(&self) -> Result<f64, FormulaEvalError> {
        self.now_serial.ok_or(FormulaEvalError::Num)
    }

    pub(super) fn next_random_u64(&self) -> u64 {
        let next = self
            .random_state
            .get()
            .wrapping_mul(6_364_136_223_846_793_005)
            .wrapping_add(1_442_695_040_888_963_407);
        self.random_state.set(next);
        next
    }

    pub(super) fn rand(&self) -> f64 {
        const SCALE: f64 = 1.0 / ((1_u64 << 53) as f64);
        ((self.next_random_u64() >> 11) as f64) * SCALE
    }

    pub(super) fn rand_between(&self, bottom: f64, top: f64) -> Result<f64, FormulaEvalError> {
        formula_rand_between(bottom, top, || self.next_random_u64())
    }
}

pub(super) fn formula_rand_between(
    bottom: f64,
    top: f64,
    next_random_u64: impl FnOnce() -> u64,
) -> Result<f64, FormulaEvalError> {
    if !bottom.is_finite() || !top.is_finite() {
        return Err(FormulaEvalError::Value);
    }
    let bottom = bottom.trunc();
    let top = top.trunc();
    if bottom < i64::MIN as f64
        || bottom > i64::MAX as f64
        || top < i64::MIN as f64
        || top > i64::MAX as f64
    {
        return Err(FormulaEvalError::Num);
    }
    let bottom = bottom as i64;
    let top = top as i64;
    if bottom > top {
        return Err(FormulaEvalError::Num);
    }
    let span = i128::from(top) - i128::from(bottom) + 1;
    let span = u64::try_from(span).map_err(|_| FormulaEvalError::Num)?;
    let offset = (next_random_u64() % span) as i128;
    Ok((i128::from(bottom) + offset) as f64)
}
