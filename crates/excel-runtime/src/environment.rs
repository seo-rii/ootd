use std::time::{SystemTime, UNIX_EPOCH};

/// The time source formula evaluation reads for `NOW` and `TODAY`.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum RuntimeClock {
    /// Read the host system clock once at the start of every calculation.
    #[default]
    System,
    /// Report one fixed instant for every calculation.
    Fixed(SystemTime),
}

/// Host-supplied inputs that formula evaluation must not read from ambient process state.
///
/// Each calculation cycle captures one clock reading, applies `utc_offset_minutes` to project it
/// into workbook-local time, and draws random numbers from a session stream. A `random_seed`
/// makes that stream reproducible; without one the stream is seeded once from the system clock
/// when the runtime is created.
#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct RuntimeEnvironment {
    pub clock: RuntimeClock,
    pub utc_offset_minutes: i32,
    pub random_seed: Option<u64>,
}

impl RuntimeEnvironment {
    /// Seconds since the Unix epoch in workbook-local time, or `None` before the epoch.
    pub(crate) fn local_unix_seconds(&self) -> Option<f64> {
        let instant = match self.clock {
            RuntimeClock::System => SystemTime::now(),
            RuntimeClock::Fixed(instant) => instant,
        };
        let elapsed = instant.duration_since(UNIX_EPOCH).ok()?;
        Some(elapsed.as_secs_f64() + f64::from(self.utc_offset_minutes) * 60.0)
    }

    /// The initial random stream state for this environment.
    pub(crate) fn initial_random_state(&self) -> u64 {
        let seed = self.random_seed.unwrap_or_else(|| {
            SystemTime::now()
                .duration_since(UNIX_EPOCH)
                .map(|duration| duration.as_nanos() as u64)
                .unwrap_or(0x9e37_79b9_7f4a_7c15)
        });
        let state = seed ^ 0xa076_1d64_78bd_642f;
        if state == 0 {
            0xe703_7ed1_a0b4_28db
        } else {
            state
        }
    }
}
