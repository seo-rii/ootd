//! Annuity helpers shared by FV/PV/PMT/IPMT and the fractional-dollar functions.

use super::*;

pub(super) fn formula_dollar_fraction_parts(fraction: f64) -> Result<(f64, f64), FormulaEvalError> {
    if !fraction.is_finite() {
        return Err(FormulaEvalError::Value);
    }
    if fraction < 0.0 {
        return Err(FormulaEvalError::Num);
    }
    let denominator = fraction.trunc();
    if denominator < 1.0 {
        return Err(FormulaEvalError::Div0);
    }
    let scale = 10_f64.powf(denominator.log10().ceil());
    if scale.is_finite() {
        Ok((denominator, scale))
    } else {
        Err(FormulaEvalError::Num)
    }
}

pub(super) fn formula_dollar_fraction_near_integer(value: f64) -> f64 {
    let rounded = value.round();
    if (value - rounded).abs() <= 1e-9 {
        rounded
    } else {
        value
    }
}

pub(super) fn formula_financial_type_argument(value: f64) -> Result<f64, FormulaEvalError> {
    if !value.is_finite() {
        return Err(FormulaEvalError::Value);
    }
    match value.trunc() as i64 {
        0 => Ok(0.0),
        1 => Ok(1.0),
        _ => Err(FormulaEvalError::Num),
    }
}

pub(super) fn formula_annuity_growth(rate: f64, nper: f64) -> Result<f64, FormulaEvalError> {
    let growth = (1.0 + rate).powf(nper);
    if growth.is_finite() {
        Ok(growth)
    } else {
        Err(FormulaEvalError::Num)
    }
}

pub(super) fn formula_fv_value(
    rate: f64,
    nper: f64,
    pmt: f64,
    pv: f64,
    payment_type: f64,
) -> Result<f64, FormulaEvalError> {
    if ![rate, nper, pmt, pv].iter().all(|value| value.is_finite()) {
        return Err(FormulaEvalError::Value);
    }
    let value = if rate == 0.0 {
        -(pv + pmt * nper)
    } else {
        let growth = formula_annuity_growth(rate, nper)?;
        -(pv * growth + pmt * (1.0 + rate * payment_type) * (growth - 1.0) / rate)
    };
    if value.is_finite() {
        Ok(value)
    } else {
        Err(FormulaEvalError::Num)
    }
}

pub(super) fn formula_pmt_value(
    rate: f64,
    nper: f64,
    pv: f64,
    fv: f64,
    payment_type: f64,
) -> Result<f64, FormulaEvalError> {
    if ![rate, nper, pv, fv].iter().all(|value| value.is_finite()) {
        return Err(FormulaEvalError::Value);
    }
    if nper == 0.0 {
        return Err(FormulaEvalError::Div0);
    }
    let value = if rate == 0.0 {
        -(pv + fv) / nper
    } else {
        let growth = formula_annuity_growth(rate, nper)?;
        let denominator = (1.0 + rate * payment_type) * (growth - 1.0);
        if denominator == 0.0 {
            return Err(FormulaEvalError::Div0);
        }
        -(pv * growth + fv) * rate / denominator
    };
    if value.is_finite() {
        Ok(value)
    } else {
        Err(FormulaEvalError::Num)
    }
}

pub(super) fn formula_ipmt_value(
    rate: f64,
    period: f64,
    nper: f64,
    pv: f64,
    fv: f64,
    payment_type: f64,
) -> Result<f64, FormulaEvalError> {
    if ![rate, period, nper, pv, fv]
        .iter()
        .all(|value| value.is_finite())
    {
        return Err(FormulaEvalError::Value);
    }
    if period < 1.0 || period > nper || nper <= 0.0 {
        return Err(FormulaEvalError::Num);
    }
    if rate == 0.0 {
        return Ok(0.0);
    }
    if payment_type == 1.0 && period == 1.0 {
        return Ok(0.0);
    }
    let payment = formula_pmt_value(rate, nper, pv, fv, payment_type)?;
    let mut value = formula_fv_value(rate, period - 1.0, payment, pv, payment_type)? * rate;
    if payment_type == 1.0 {
        value /= 1.0 + rate;
    }
    if value.is_finite() {
        Ok(value)
    } else {
        Err(FormulaEvalError::Num)
    }
}
