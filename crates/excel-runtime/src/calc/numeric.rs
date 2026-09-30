//! Rounding and special-function helpers: error function, gamma, beta, and t/F distributions.

use super::*;

pub(super) fn round_half_away_from_zero(value: f64) -> f64 {
    if value.is_sign_negative() {
        -((-value) + 0.5).floor()
    } else {
        (value + 0.5).floor()
    }
}

pub(super) fn round_away_from_zero(value: f64) -> f64 {
    if value.is_sign_negative() {
        value.floor()
    } else {
        value.ceil()
    }
}

pub(super) fn round_toward_zero(value: f64) -> f64 {
    if value.is_sign_negative() {
        value.ceil()
    } else {
        value.floor()
    }
}

pub(super) fn formula_round_factor(digits: f64) -> Result<f64, FormulaEvalError> {
    if !digits.is_finite() || digits.fract() != 0.0 {
        return Err(FormulaEvalError::Value);
    }
    if digits < i32::MIN as f64 || digits > i32::MAX as f64 {
        return Err(FormulaEvalError::Num);
    }
    let factor = 10_f64.powi(digits as i32);
    if !factor.is_finite() || factor == 0.0 {
        return Err(FormulaEvalError::Num);
    }
    Ok(factor)
}

pub(super) fn formula_checked_numeric_result(value: f64) -> Result<f64, FormulaEvalError> {
    if value.is_finite() {
        Ok(value)
    } else {
        Err(FormulaEvalError::Num)
    }
}

pub(super) fn formula_erf_approx(value: f64) -> f64 {
    let sign = if value.is_sign_negative() { -1.0 } else { 1.0 };
    let x = value.abs();
    let t = 1.0 / (1.0 + 0.5 * x);
    let tau = t
        * (-x * x - 1.26551223
            + t * (1.00002368
                + t * (0.37409196
                    + t * (0.09678418
                        + t * (-0.18628806
                            + t * (0.27886807
                                + t * (-1.13520398
                                    + t * (1.48851587 + t * (-0.82215223 + t * 0.17087277)))))))))
            .exp();
    sign * (1.0 - tau)
}

pub(super) fn formula_standard_normal_cdf(z: f64) -> Result<f64, FormulaEvalError> {
    if !z.is_finite() {
        return Err(FormulaEvalError::Value);
    }
    formula_checked_numeric_result(0.5 * (1.0 + formula_erf_approx(z / std::f64::consts::SQRT_2)))
}

pub(super) fn formula_gamma_ln_value(value: f64) -> f64 {
    const COEFFICIENTS: [f64; 9] = [
        0.9999999999998099,
        676.5203681218851,
        -1259.1392167224028,
        771.3234287776531,
        -176.6150291621406,
        12.507343278686905,
        -0.13857109526572012,
        0.000009984369578019572,
        0.00000015056327351493116,
    ];
    let lanczos = |input: f64| {
        let z = input - 1.0;
        let mut x = COEFFICIENTS[0];
        for (index, coefficient) in COEFFICIENTS.iter().enumerate().skip(1) {
            x += coefficient / (z + index as f64);
        }
        let t = z + 7.5;
        0.5 * (2.0 * std::f64::consts::PI).ln() + (z + 0.5) * t.ln() - t + x.ln()
    };
    if value < 0.5 {
        std::f64::consts::PI.ln() - (std::f64::consts::PI * value).sin().ln() - lanczos(1.0 - value)
    } else {
        lanczos(value)
    }
}

pub(super) fn formula_beta_fraction(
    alpha: f64,
    beta: f64,
    x: f64,
) -> Result<f64, FormulaEvalError> {
    const EPSILON: f64 = 1e-14;
    const FLOOR: f64 = 1e-300;
    const MAX_ITERATIONS: usize = 200;
    let qab = alpha + beta;
    let qap = alpha + 1.0;
    let qam = alpha - 1.0;
    let mut c = 1.0;
    let mut d = 1.0 - qab * x / qap;
    if d.abs() < FLOOR {
        d = FLOOR;
    }
    d = 1.0 / d;
    let mut h = d;
    for m in 1..=MAX_ITERATIONS {
        let m_f = m as f64;
        let m2 = 2.0 * m_f;
        let mut aa = m_f * (beta - m_f) * x / ((qam + m2) * (alpha + m2));
        d = 1.0 + aa * d;
        if d.abs() < FLOOR {
            d = FLOOR;
        }
        c = 1.0 + aa / c;
        if c.abs() < FLOOR {
            c = FLOOR;
        }
        d = 1.0 / d;
        h *= d * c;
        aa = -(alpha + m_f) * (qab + m_f) * x / ((alpha + m2) * (qap + m2));
        d = 1.0 + aa * d;
        if d.abs() < FLOOR {
            d = FLOOR;
        }
        c = 1.0 + aa / c;
        if c.abs() < FLOOR {
            c = FLOOR;
        }
        d = 1.0 / d;
        let delta = d * c;
        h *= delta;
        if (delta - 1.0).abs() <= EPSILON {
            return formula_checked_numeric_result(h);
        }
    }
    formula_checked_numeric_result(h)
}

pub(super) fn formula_regularized_beta(
    x: f64,
    alpha: f64,
    beta: f64,
) -> Result<f64, FormulaEvalError> {
    if ![x, alpha, beta].iter().all(|value| value.is_finite()) {
        return Err(FormulaEvalError::Value);
    }
    if alpha <= 0.0 || beta <= 0.0 || !(0.0..=1.0).contains(&x) {
        return Err(FormulaEvalError::Num);
    }
    if x == 0.0 || x == 1.0 {
        return Ok(x);
    }
    let log_beta = formula_gamma_ln_value(alpha) + formula_gamma_ln_value(beta)
        - formula_gamma_ln_value(alpha + beta);
    let front = (alpha * x.ln() + beta * (-x).ln_1p() - log_beta).exp();
    if x < (alpha + 1.0) / (alpha + beta + 2.0) {
        formula_checked_numeric_result(
            (front * formula_beta_fraction(alpha, beta, x)? / alpha).clamp(0.0, 1.0),
        )
    } else {
        formula_checked_numeric_result(
            (1.0 - front * formula_beta_fraction(beta, alpha, 1.0 - x)? / beta).clamp(0.0, 1.0),
        )
    }
}

pub(super) fn formula_student_t_cdf(x: f64, degrees: f64) -> Result<f64, FormulaEvalError> {
    if !x.is_finite() || !degrees.is_finite() {
        return Err(FormulaEvalError::Value);
    }
    if degrees < 1.0 {
        return Err(FormulaEvalError::Num);
    }
    let degrees = degrees.trunc();
    let beta_x = degrees / (degrees + x * x);
    let tail = 0.5 * formula_regularized_beta(beta_x, degrees / 2.0, 0.5)?;
    Ok(if x >= 0.0 { 1.0 - tail } else { tail })
}

pub(super) fn formula_student_t_right_tail_from_abs(
    t: f64,
    degrees: f64,
) -> Result<f64, FormulaEvalError> {
    formula_student_t_cdf(t.abs(), degrees).map(|value| (1.0 - value).clamp(0.0, 1.0))
}

pub(super) fn formula_f_right_tail(
    x: f64,
    degrees1: f64,
    degrees2: f64,
) -> Result<f64, FormulaEvalError> {
    if ![x, degrees1, degrees2]
        .iter()
        .all(|value| value.is_finite())
    {
        return Err(FormulaEvalError::Value);
    }
    if x < 0.0 || degrees1 < 1.0 || degrees2 < 1.0 {
        return Err(FormulaEvalError::Num);
    }
    let degrees1 = degrees1.trunc();
    let degrees2 = degrees2.trunc();
    let transformed = degrees1 * x / (degrees1 * x + degrees2);
    formula_regularized_beta(transformed, degrees1 / 2.0, degrees2 / 2.0)
        .map(|value| (1.0 - value).clamp(0.0, 1.0))
}

pub(super) fn formula_sample_mean_and_variance(
    values: &[f64],
) -> Result<(f64, f64), FormulaEvalError> {
    if values.len() < 2 {
        return Err(FormulaEvalError::Div0);
    }
    if values.iter().any(|value| !value.is_finite()) {
        return Err(FormulaEvalError::Value);
    }
    let mean = values.iter().sum::<f64>() / values.len() as f64;
    let deviation_sum = values
        .iter()
        .map(|value| {
            let deviation = value - mean;
            deviation * deviation
        })
        .sum::<f64>();
    formula_checked_numeric_result(deviation_sum / (values.len() - 1) as f64)
        .map(|variance| (mean, variance))
}
