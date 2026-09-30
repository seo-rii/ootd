//! Serial/calendar conversion in the 1900 and 1904 date systems, weekday and workday rules,
//! and date/time text parsing.

use super::*;

pub(super) fn formula_datevalue_text(
    date_system: DateSystem,
    locale: RuntimeLocale,
    text: &str,
) -> Result<f64, FormulaEvalError> {
    let trimmed = text.trim();
    if trimmed.chars().any(|ch| ch.is_ascii_alphabetic()) {
        let normalized = trimmed.replace([',', '-', '/'], " ");
        let parts = normalized.split_whitespace().collect::<Vec<_>>();
        if parts.len() != 3 {
            return Err(FormulaEvalError::Value);
        }
        let month_names = [
            ("JAN", "JANUARY"),
            ("FEB", "FEBRUARY"),
            ("MAR", "MARCH"),
            ("APR", "APRIL"),
            ("MAY", "MAY"),
            ("JUN", "JUNE"),
            ("JUL", "JULY"),
            ("AUG", "AUGUST"),
            ("SEP", "SEPTEMBER"),
            ("OCT", "OCTOBER"),
            ("NOV", "NOVEMBER"),
            ("DEC", "DECEMBER"),
        ];
        let mut month_index = None;
        for (index, part) in parts.iter().enumerate() {
            let token = part.trim_end_matches('.').to_ascii_uppercase();
            if let Some((month_zero_based, _)) =
                month_names.iter().enumerate().find(|(_, (short, long))| {
                    token == *short || token == *long || (token == "SEPT" && *long == "SEPTEMBER")
                })
            {
                if month_index.is_some() {
                    return Err(FormulaEvalError::Value);
                }
                month_index = Some((index, month_zero_based as i64 + 1));
            }
        }
        let Some((month_position, month)) = month_index else {
            return Err(FormulaEvalError::Value);
        };
        let parse_numeric_part = |part: &str| -> Result<i64, FormulaEvalError> {
            let value = part.trim_end_matches('.');
            let lower = value.to_ascii_lowercase();
            let value = if lower.ends_with("st")
                || lower.ends_with("nd")
                || lower.ends_with("rd")
                || lower.ends_with("th")
            {
                &value[..value.len() - 2]
            } else {
                value
            };
            if value.is_empty() {
                return Err(FormulaEvalError::Value);
            }
            value.parse::<i64>().map_err(|_| FormulaEvalError::Value)
        };
        let (year, day) = match month_position {
            0 => (parse_numeric_part(parts[2])?, parse_numeric_part(parts[1])?),
            1 if parts[0].trim().len() == 4 => {
                (parse_numeric_part(parts[0])?, parse_numeric_part(parts[2])?)
            }
            1 => (parse_numeric_part(parts[2])?, parse_numeric_part(parts[0])?),
            _ => return Err(FormulaEvalError::Value),
        };
        if !(1900..=9999).contains(&year) || day < 1 {
            return Err(FormulaEvalError::Value);
        }
        if year == 1900 && month == 2 && day == 29 {
            return Ok(60.0);
        }
        if day > i64::from(days_in_excel_month(year, month as u32)) {
            return Err(FormulaEvalError::Value);
        }
        return date_system.serial_from_args(year as f64, month as f64, day as f64);
    }
    // `.` separates dates only where it cannot be a decimal point.
    let separator = if trimmed.contains('-') {
        '-'
    } else if trimmed.contains('/') {
        '/'
    } else if trimmed.contains('.') && locale.decimal_separator != '.' {
        '.'
    } else {
        return Err(FormulaEvalError::Value);
    };
    let parts = trimmed.split(separator).collect::<Vec<_>>();
    if parts.len() != 3 || parts.iter().any(|part| part.trim().is_empty()) {
        return Err(FormulaEvalError::Value);
    }
    let parse_part = |part: &str| -> Result<i64, FormulaEvalError> {
        part.trim()
            .parse::<i64>()
            .map_err(|_| FormulaEvalError::Value)
    };
    let first = parse_part(parts[0])?;
    let second = parse_part(parts[1])?;
    let third = parse_part(parts[2])?;
    let (year, month, day) = if parts[0].trim().len() == 4 {
        (first, second, third)
    } else {
        match locale.date_order {
            RuntimeDateOrder::MonthDayYear => (third, first, second),
            RuntimeDateOrder::DayMonthYear => (third, second, first),
            RuntimeDateOrder::YearMonthDay => (first, second, third),
        }
    };
    if !(1900..=9999).contains(&year) || !(1..=12).contains(&month) || day < 1 {
        return Err(FormulaEvalError::Value);
    }
    if year == 1900 && month == 2 && day == 29 {
        return Ok(60.0);
    }
    if day > i64::from(days_in_excel_month(year, month as u32)) {
        return Err(FormulaEvalError::Value);
    }
    date_system.serial_from_args(year as f64, month as f64, day as f64)
}

pub(super) fn formula_timevalue_text(text: &str) -> Result<f64, FormulaEvalError> {
    let trimmed = text.trim();
    if trimmed.is_empty() {
        return Err(FormulaEvalError::Value);
    }
    let upper = trimmed.to_ascii_uppercase();
    let (body, pm_marker) = if upper.ends_with("AM") {
        (trimmed[..trimmed.len() - 2].trim_end(), Some(false))
    } else if upper.ends_with("PM") {
        (trimmed[..trimmed.len() - 2].trim_end(), Some(true))
    } else {
        (trimmed, None)
    };
    let parts = body.split(':').collect::<Vec<_>>();
    if !(2..=3).contains(&parts.len()) || parts.iter().any(|part| part.trim().is_empty()) {
        return Err(FormulaEvalError::Value);
    }
    let parse_part = |part: &str| -> Result<i64, FormulaEvalError> {
        part.trim()
            .parse::<i64>()
            .map_err(|_| FormulaEvalError::Value)
    };
    let mut hour = parse_part(parts[0])?;
    let minute = parse_part(parts[1])?;
    let second = if parts.len() == 3 {
        parse_part(parts[2])?
    } else {
        0
    };
    if !(0..=59).contains(&minute) || !(0..=59).contains(&second) {
        return Err(FormulaEvalError::Value);
    }
    if let Some(is_pm) = pm_marker {
        if !(1..=12).contains(&hour) {
            return Err(FormulaEvalError::Value);
        }
        if hour == 12 {
            hour = 0;
        }
        if is_pm {
            hour += 12;
        }
    } else if !(0..=23).contains(&hour) {
        return Err(FormulaEvalError::Value);
    }
    Ok((hour * 3600 + minute * 60 + second) as f64 / 86_400.0)
}

pub(crate) fn formula_date_serial_from_args(
    year: f64,
    month: f64,
    day: f64,
) -> Result<f64, FormulaEvalError> {
    let mut year = formula_integer_argument(year)?;
    let month = formula_integer_argument(month)?;
    let day = formula_integer_argument(day)?;
    if (0..=1899).contains(&year) {
        year += 1900;
    } else if !(1900..=9999).contains(&year) {
        return Err(FormulaEvalError::Num);
    }
    let total_months = year
        .checked_mul(12)
        .and_then(|value| value.checked_add(month - 1))
        .ok_or(FormulaEvalError::Num)?;
    let normalized_year = div_floor(total_months, 12);
    let normalized_month = total_months - normalized_year * 12 + 1;
    if normalized_year == 1900 && normalized_month == 2 && day == 29 {
        return Ok(60.0);
    }
    let days = days_from_civil(normalized_year, normalized_month as u32, 1)
        .checked_add(day - 1)
        .ok_or(FormulaEvalError::Num)?;
    formula_serial_from_civil_days(days).map(|serial| serial as f64)
}

pub(super) fn formula_ymd_from_serial(serial: f64) -> Result<(i64, u32, u32), FormulaEvalError> {
    let serial = formula_serial_integer(serial)?;
    if serial == 60 {
        return Ok((1900, 2, 29));
    }
    let base_days = days_from_civil(1899, 12, 31);
    let adjusted_serial = if serial > 60 { serial - 1 } else { serial };
    let days = base_days
        .checked_add(adjusted_serial)
        .ok_or(FormulaEvalError::Num)?;
    let (year, month, day) = civil_from_days(days);
    if !(1900..=9999).contains(&year) {
        return Err(FormulaEvalError::Num);
    }
    Ok((year, month, day))
}

pub(super) fn formula_serial_integer(serial: f64) -> Result<i64, FormulaEvalError> {
    if !serial.is_finite() {
        return Err(FormulaEvalError::Value);
    }
    let serial = serial.floor();
    if serial < 1.0 || serial > i64::MAX as f64 {
        return Err(FormulaEvalError::Num);
    }
    Ok(serial as i64)
}

/// The workbook date system a calculation interprets serials in. Calendar conversions run in the
/// 1900 system; a 1904 serial is the 1900 serial minus 1,462 days, and 1904 serials are never
/// negative.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(super) enum DateSystem {
    Excel1900,
    Excel1904,
}

impl DateSystem {
    const EXCEL_1904_OFFSET_DAYS: i64 = 1_462;

    pub(super) fn offset_days(self) -> i64 {
        match self {
            Self::Excel1900 => 0,
            Self::Excel1904 => Self::EXCEL_1904_OFFSET_DAYS,
        }
    }

    /// The 1900-system serial of a whole-day serial in this system.
    pub(super) fn serial_1900(self, serial: i64) -> i64 {
        serial + self.offset_days()
    }

    pub(super) fn ymd(self, serial: f64) -> Result<(i64, u32, u32), FormulaEvalError> {
        if self == Self::Excel1904 && serial.is_finite() && serial < 0.0 {
            return Err(FormulaEvalError::Num);
        }
        formula_ymd_from_serial(serial + self.offset_days() as f64)
    }

    pub(super) fn serial_from_args(
        self,
        year: f64,
        month: f64,
        day: f64,
    ) -> Result<f64, FormulaEvalError> {
        let serial = formula_date_serial_from_args(year, month, day)? - self.offset_days() as f64;
        if serial < 0.0 {
            return Err(FormulaEvalError::Num);
        }
        Ok(serial)
    }

    pub(super) fn edate(self, serial: f64, months: f64) -> Result<f64, FormulaEvalError> {
        let (year, month, day) = self.ymd(serial)?;
        let months = formula_integer_argument(months)?;
        let (target_year, target_month) = normalize_year_month(year, i64::from(month) + months)?;
        let target_day = i64::from(day.min(days_in_excel_month(target_year, target_month)));
        self.serial_from_args(target_year as f64, target_month as f64, target_day as f64)
    }

    pub(super) fn eomonth(self, serial: f64, months: f64) -> Result<f64, FormulaEvalError> {
        let (year, month, _) = self.ymd(serial)?;
        let months = formula_integer_argument(months)?;
        let (target_year, target_month) = normalize_year_month(year, i64::from(month) + months)?;
        let target_day = i64::from(days_in_excel_month(target_year, target_month));
        self.serial_from_args(target_year as f64, target_month as f64, target_day as f64)
    }
}

pub(super) fn formula_weekday_monday0_from_serial(serial: i64) -> i64 {
    let adjusted_serial = if serial > 60 { serial - 1 } else { serial };
    (adjusted_serial - 1).rem_euclid(7)
}

pub(super) fn formula_standard_weekend_mask() -> [bool; 7] {
    [false, false, false, false, false, true, true]
}

pub(super) fn formula_weekend_mask_from_code(code: i64) -> Result<[bool; 7], FormulaEvalError> {
    match code {
        1 => Ok(formula_standard_weekend_mask()),
        2 => Ok([true, false, false, false, false, false, true]),
        3 => Ok([true, true, false, false, false, false, false]),
        4 => Ok([false, true, true, false, false, false, false]),
        5 => Ok([false, false, true, true, false, false, false]),
        6 => Ok([false, false, false, true, true, false, false]),
        7 => Ok([false, false, false, false, true, true, false]),
        11 => Ok([false, false, false, false, false, false, true]),
        12 => Ok([true, false, false, false, false, false, false]),
        13 => Ok([false, true, false, false, false, false, false]),
        14 => Ok([false, false, true, false, false, false, false]),
        15 => Ok([false, false, false, true, false, false, false]),
        16 => Ok([false, false, false, false, true, false, false]),
        17 => Ok([false, false, false, false, false, true, false]),
        _ => Err(FormulaEvalError::Value),
    }
}

pub(super) fn formula_weekend_mask_from_string(value: &str) -> Result<[bool; 7], FormulaEvalError> {
    if value.chars().count() != 7 {
        return Err(FormulaEvalError::Value);
    }
    let mut mask = [false; 7];
    let mut workday_count = 0_u8;
    for (index, ch) in value.chars().enumerate() {
        mask[index] = match ch {
            '0' => {
                workday_count += 1;
                false
            }
            '1' => true,
            _ => return Err(FormulaEvalError::Value),
        };
    }
    if workday_count == 0 {
        return Err(FormulaEvalError::Value);
    }
    Ok(mask)
}

pub(super) fn formula_is_workday_serial(
    date_system: DateSystem,
    serial: i64,
    holidays: &[i64],
    weekend: &[bool; 7],
) -> bool {
    !weekend[formula_weekday_monday0_from_serial(date_system.serial_1900(serial)) as usize]
        && !holidays.contains(&serial)
}

pub(super) fn formula_networkdays(
    date_system: DateSystem,
    start_serial: i64,
    end_serial: i64,
    holidays: &[i64],
) -> Result<f64, FormulaEvalError> {
    formula_networkdays_with_weekend(
        date_system,
        start_serial,
        end_serial,
        holidays,
        &formula_standard_weekend_mask(),
    )
}

pub(super) fn formula_networkdays_with_weekend(
    date_system: DateSystem,
    start_serial: i64,
    end_serial: i64,
    holidays: &[i64],
    weekend: &[bool; 7],
) -> Result<f64, FormulaEvalError> {
    let (first, last, sign) = if start_serial <= end_serial {
        (start_serial, end_serial, 1.0)
    } else {
        (end_serial, start_serial, -1.0)
    };
    date_system.ymd(first as f64)?;
    date_system.ymd(last as f64)?;
    let mut count = 0_u64;
    for serial in first..=last {
        if formula_is_workday_serial(date_system, serial, holidays, weekend) {
            count += 1;
        }
    }
    Ok(count as f64 * sign)
}

pub(super) fn formula_workday(
    date_system: DateSystem,
    start_serial: i64,
    days: i64,
    holidays: &[i64],
) -> Result<f64, FormulaEvalError> {
    formula_workday_with_weekend(
        date_system,
        start_serial,
        days,
        holidays,
        &formula_standard_weekend_mask(),
    )
}

pub(super) fn formula_workday_with_weekend(
    date_system: DateSystem,
    start_serial: i64,
    days: i64,
    holidays: &[i64],
    weekend: &[bool; 7],
) -> Result<f64, FormulaEvalError> {
    date_system.ymd(start_serial as f64)?;
    let direction = if days < 0 { -1 } else { 1 };
    let mut serial = start_serial;
    let mut remaining = days.unsigned_abs();
    while remaining > 0 {
        serial = serial.checked_add(direction).ok_or(FormulaEvalError::Num)?;
        date_system.ymd(serial as f64)?;
        if formula_is_workday_serial(date_system, serial, holidays, weekend) {
            remaining -= 1;
        }
    }
    Ok(serial as f64)
}

pub(super) fn formula_time_serial_from_args(
    hour: f64,
    minute: f64,
    second: f64,
) -> Result<f64, FormulaEvalError> {
    let hour = formula_time_argument(hour)?;
    let minute = formula_time_argument(minute)?;
    let second = formula_time_argument(second)?;
    let total_seconds = hour
        .checked_mul(3600)
        .and_then(|value| {
            minute
                .checked_mul(60)
                .and_then(|minute| value.checked_add(minute))
        })
        .and_then(|value| value.checked_add(second))
        .ok_or(FormulaEvalError::Num)?;
    Ok((total_seconds.rem_euclid(86_400)) as f64 / 86_400.0)
}

pub(super) fn formula_time_parts_from_serial(
    serial: f64,
) -> Result<(u32, u32, u32), FormulaEvalError> {
    if !serial.is_finite() {
        return Err(FormulaEvalError::Value);
    }
    if serial < 0.0 {
        return Err(FormulaEvalError::Num);
    }
    let fraction = serial - serial.floor();
    let mut total_seconds = (fraction * 86_400.0).round() as i64;
    if total_seconds >= 86_400 {
        total_seconds = 0;
    }
    let hour = total_seconds / 3600;
    let minute = (total_seconds % 3600) / 60;
    let second = total_seconds % 60;
    Ok((hour as u32, minute as u32, second as u32))
}

pub(super) fn formula_time_argument(value: f64) -> Result<i64, FormulaEvalError> {
    let value = formula_integer_argument(value)?;
    if !(0..=32_767).contains(&value) {
        return Err(FormulaEvalError::Num);
    }
    Ok(value)
}

pub(super) fn normalize_year_month(year: i64, month: i64) -> Result<(i64, u32), FormulaEvalError> {
    let total_months = year
        .checked_mul(12)
        .and_then(|value| value.checked_add(month - 1))
        .ok_or(FormulaEvalError::Num)?;
    let normalized_year = div_floor(total_months, 12);
    let normalized_month = (total_months - normalized_year * 12 + 1) as u32;
    if !(1900..=9999).contains(&normalized_year) {
        return Err(FormulaEvalError::Num);
    }
    Ok((normalized_year, normalized_month))
}

pub(super) fn formula_serial_from_civil_days(days: i64) -> Result<i64, FormulaEvalError> {
    let min_days = days_from_civil(1900, 1, 1);
    let max_days = days_from_civil(9999, 12, 31);
    if days < min_days || days > max_days {
        return Err(FormulaEvalError::Num);
    }
    let base_days = days_from_civil(1899, 12, 31);
    let mut serial = days - base_days;
    if days >= days_from_civil(1900, 3, 1) {
        serial += 1;
    }
    Ok(serial)
}

pub(super) fn days_in_excel_month(year: i64, month: u32) -> u32 {
    match month {
        1 | 3 | 5 | 7 | 8 | 10 | 12 => 31,
        4 | 6 | 9 | 11 => 30,
        2 if year == 1900 => 29,
        2 if is_gregorian_leap_year(year) => 29,
        2 => 28,
        _ => 0,
    }
}

pub(super) fn is_gregorian_leap_year(year: i64) -> bool {
    (year % 4 == 0 && year % 100 != 0) || year % 400 == 0
}

pub(super) fn days_from_civil(year: i64, month: u32, day: u32) -> i64 {
    let year = year - i64::from(month <= 2);
    let era = div_floor(year, 400);
    let year_of_era = year - era * 400;
    let month = i64::from(month);
    let day_of_year = (153 * (month + if month > 2 { -3 } else { 9 }) + 2) / 5 + i64::from(day) - 1;
    let day_of_era = year_of_era * 365 + year_of_era / 4 - year_of_era / 100 + day_of_year;
    era * 146_097 + day_of_era - 719_468
}

pub(super) fn civil_from_days(days: i64) -> (i64, u32, u32) {
    let days = days + 719_468;
    let era = div_floor(days, 146_097);
    let day_of_era = days - era * 146_097;
    let year_of_era =
        (day_of_era - day_of_era / 1_460 + day_of_era / 36_524 - day_of_era / 146_096) / 365;
    let mut year = year_of_era + era * 400;
    let day_of_year = day_of_era - (365 * year_of_era + year_of_era / 4 - year_of_era / 100);
    let month_prime = (5 * day_of_year + 2) / 153;
    let day = day_of_year - (153 * month_prime + 2) / 5 + 1;
    let month = month_prime + if month_prime < 10 { 3 } else { -9 };
    year += i64::from(month <= 2);
    (year, month as u32, day as u32)
}
