//! Complex numbers, Bessel functions, unit and euro conversion, and radix/bitwise/Roman helpers.

use super::*;

#[derive(Debug, Clone, Copy, PartialEq)]
pub(crate) struct FormulaComplexNumber {
    pub(crate) real: f64,
    pub(crate) imaginary: f64,
    pub(super) suffix: Option<char>,
}

pub(super) fn formula_complex_clean_component(value: f64) -> f64 {
    if value.abs() < 1e-12 { 0.0 } else { value }
}

pub(super) fn formula_complex_number(
    real: f64,
    imaginary: f64,
    suffix: Option<char>,
) -> Result<FormulaComplexNumber, FormulaEvalError> {
    if !real.is_finite() || !imaginary.is_finite() {
        return Err(FormulaEvalError::Num);
    }
    if suffix.is_some_and(|suffix| !matches!(suffix, 'i' | 'j')) {
        return Err(FormulaEvalError::Value);
    }
    Ok(FormulaComplexNumber {
        real: formula_complex_clean_component(real),
        imaginary: formula_complex_clean_component(imaginary),
        suffix,
    })
}

pub(super) fn formula_complex_suffix(text: &str) -> Result<char, FormulaEvalError> {
    match text {
        "i" => Ok('i'),
        "j" => Ok('j'),
        _ => Err(FormulaEvalError::Value),
    }
}

pub(super) fn formula_complex_component(text: &str) -> Result<f64, FormulaEvalError> {
    let value = text.parse::<f64>().map_err(|_| FormulaEvalError::Value)?;
    if value.is_finite() {
        Ok(value)
    } else {
        Err(FormulaEvalError::Num)
    }
}

pub(super) fn formula_complex_imaginary_coefficient(text: &str) -> Result<f64, FormulaEvalError> {
    match text {
        "" | "+" => Ok(1.0),
        "-" => Ok(-1.0),
        _ => formula_complex_component(text),
    }
}

pub(super) fn formula_complex_split_imaginary(text: &str) -> Option<usize> {
    let mut split = None;
    for (index, ch) in text.char_indices().skip(1) {
        if matches!(ch, '+' | '-')
            && !text[..index]
                .chars()
                .next_back()
                .is_some_and(|previous| matches!(previous, 'e' | 'E'))
        {
            split = Some(index);
        }
    }
    split
}

pub(crate) fn formula_complex_from_text(
    text: &str,
) -> Result<FormulaComplexNumber, FormulaEvalError> {
    let text = text.trim();
    if text.is_empty() || text.chars().any(char::is_whitespace) {
        return Err(FormulaEvalError::Value);
    }
    if text.ends_with('I') || text.ends_with('J') {
        return Err(FormulaEvalError::Value);
    }
    let Some(suffix) = text.chars().last().filter(|ch| matches!(ch, 'i' | 'j')) else {
        return formula_complex_number(formula_complex_component(text)?, 0.0, None);
    };
    let value = &text[..text.len() - suffix.len_utf8()];
    let (real, imaginary) = if let Some(split) = formula_complex_split_imaginary(value) {
        (
            formula_complex_component(&value[..split])?,
            formula_complex_imaginary_coefficient(&value[split..])?,
        )
    } else {
        (0.0, formula_complex_imaginary_coefficient(value)?)
    };
    formula_complex_number(real, imaginary, Some(suffix))
}

pub(super) fn formula_complex_format(
    value: FormulaComplexNumber,
) -> Result<String, FormulaEvalError> {
    let real = formula_complex_clean_component(value.real);
    let imaginary = formula_complex_clean_component(value.imaginary);
    if imaginary == 0.0 {
        return formula_text_from_number(real);
    }
    let suffix = value.suffix.unwrap_or('i');
    let imaginary_text = |magnitude: f64| -> Result<String, FormulaEvalError> {
        if magnitude == 1.0 {
            Ok(suffix.to_string())
        } else {
            Ok(format!(
                "{}{}",
                formula_text_from_number(magnitude)?,
                suffix
            ))
        }
    };
    if real == 0.0 {
        if imaginary < 0.0 {
            return Ok(format!("-{}", imaginary_text(-imaginary)?));
        }
        return imaginary_text(imaginary);
    }
    let sign = if imaginary < 0.0 { "-" } else { "+" };
    Ok(format!(
        "{}{}{}",
        formula_text_from_number(real)?,
        sign,
        imaginary_text(imaginary.abs())?
    ))
}

pub(super) fn formula_complex_join_suffix(
    left: FormulaComplexNumber,
    right: FormulaComplexNumber,
) -> Result<Option<char>, FormulaEvalError> {
    match (left.suffix, right.suffix) {
        (Some(left), Some(right)) if left != right => Err(FormulaEvalError::Value),
        (Some(suffix), _) | (_, Some(suffix)) => Ok(Some(suffix)),
        (None, None) => Ok(None),
    }
}

pub(super) fn formula_complex_add(
    left: FormulaComplexNumber,
    right: FormulaComplexNumber,
) -> Result<FormulaComplexNumber, FormulaEvalError> {
    formula_complex_number(
        left.real + right.real,
        left.imaginary + right.imaginary,
        formula_complex_join_suffix(left, right)?,
    )
}

pub(super) fn formula_complex_subtract(
    left: FormulaComplexNumber,
    right: FormulaComplexNumber,
) -> Result<FormulaComplexNumber, FormulaEvalError> {
    formula_complex_number(
        left.real - right.real,
        left.imaginary - right.imaginary,
        formula_complex_join_suffix(left, right)?,
    )
}

pub(super) fn formula_complex_multiply(
    left: FormulaComplexNumber,
    right: FormulaComplexNumber,
) -> Result<FormulaComplexNumber, FormulaEvalError> {
    formula_complex_number(
        left.real * right.real - left.imaginary * right.imaginary,
        left.real * right.imaginary + left.imaginary * right.real,
        formula_complex_join_suffix(left, right)?,
    )
}

pub(super) fn formula_complex_divide(
    left: FormulaComplexNumber,
    right: FormulaComplexNumber,
) -> Result<FormulaComplexNumber, FormulaEvalError> {
    let denominator = right.real * right.real + right.imaginary * right.imaginary;
    if denominator == 0.0 {
        return Err(FormulaEvalError::Num);
    }
    formula_complex_number(
        (left.real * right.real + left.imaginary * right.imaginary) / denominator,
        (left.imaginary * right.real - left.real * right.imaginary) / denominator,
        formula_complex_join_suffix(left, right)?,
    )
}

pub(super) fn formula_complex_reciprocal(
    value: FormulaComplexNumber,
) -> Result<FormulaComplexNumber, FormulaEvalError> {
    formula_complex_divide(formula_complex_number(1.0, 0.0, value.suffix)?, value)
}

pub(super) fn formula_complex_exp(
    value: FormulaComplexNumber,
) -> Result<FormulaComplexNumber, FormulaEvalError> {
    let magnitude = value.real.exp();
    formula_complex_number(
        magnitude * value.imaginary.cos(),
        magnitude * value.imaginary.sin(),
        value.suffix,
    )
}

pub(super) fn formula_complex_ln(
    value: FormulaComplexNumber,
) -> Result<FormulaComplexNumber, FormulaEvalError> {
    let magnitude = value.real.hypot(value.imaginary);
    if magnitude == 0.0 {
        return Err(FormulaEvalError::Num);
    }
    formula_complex_number(
        magnitude.ln(),
        value.imaginary.atan2(value.real),
        value.suffix,
    )
}

pub(super) fn formula_complex_power(
    value: FormulaComplexNumber,
    power: f64,
) -> Result<FormulaComplexNumber, FormulaEvalError> {
    if !power.is_finite() {
        return Err(FormulaEvalError::Value);
    }
    let magnitude = value.real.hypot(value.imaginary);
    if magnitude == 0.0 {
        if power <= 0.0 {
            return Err(FormulaEvalError::Num);
        }
        return formula_complex_number(0.0, 0.0, value.suffix);
    }
    let powered_magnitude = magnitude.powf(power);
    let argument = value.imaginary.atan2(value.real) * power;
    formula_complex_number(
        powered_magnitude * argument.cos(),
        powered_magnitude * argument.sin(),
        value.suffix,
    )
}

pub(super) fn formula_complex_sin(
    value: FormulaComplexNumber,
) -> Result<FormulaComplexNumber, FormulaEvalError> {
    formula_complex_number(
        value.real.sin() * value.imaginary.cosh(),
        value.real.cos() * value.imaginary.sinh(),
        value.suffix,
    )
}

pub(super) fn formula_complex_cos(
    value: FormulaComplexNumber,
) -> Result<FormulaComplexNumber, FormulaEvalError> {
    formula_complex_number(
        value.real.cos() * value.imaginary.cosh(),
        -value.real.sin() * value.imaginary.sinh(),
        value.suffix,
    )
}

pub(super) fn formula_complex_sinh(
    value: FormulaComplexNumber,
) -> Result<FormulaComplexNumber, FormulaEvalError> {
    formula_complex_number(
        value.real.sinh() * value.imaginary.cos(),
        value.real.cosh() * value.imaginary.sin(),
        value.suffix,
    )
}

pub(super) fn formula_complex_cosh(
    value: FormulaComplexNumber,
) -> Result<FormulaComplexNumber, FormulaEvalError> {
    formula_complex_number(
        value.real.cosh() * value.imaginary.cos(),
        value.real.sinh() * value.imaginary.sin(),
        value.suffix,
    )
}

pub(super) fn formula_complex_sqrt(
    value: FormulaComplexNumber,
) -> Result<FormulaComplexNumber, FormulaEvalError> {
    let magnitude = value.real.hypot(value.imaginary);
    let real = ((magnitude + value.real) / 2.0).sqrt();
    let imaginary_sign = if value.imaginary < 0.0 { -1.0 } else { 1.0 };
    let imaginary = imaginary_sign * ((magnitude - value.real) / 2.0).sqrt();
    formula_complex_number(real, imaginary, value.suffix)
}

pub(super) fn formula_bessel_order(value: f64) -> Result<usize, FormulaEvalError> {
    if !value.is_finite() {
        return Err(FormulaEvalError::Value);
    }
    let value = value.trunc();
    if value < 0.0 {
        return Err(FormulaEvalError::Num);
    }
    if value > 10_000.0 {
        return Err(FormulaEvalError::Num);
    }
    Ok(value as usize)
}

pub(super) fn formula_bessel_j0(value: f64) -> f64 {
    let absolute = value.abs();
    if absolute < 8.0 {
        let y = value * value;
        let numerator = 57_568_490_574.0
            + y * (-13_362_590_354.0
                + y * (651_619_640.7
                    + y * (-11_214_424.18 + y * (77_392.33017 + y * -184.9052456))));
        let denominator = 57_568_490_411.0
            + y * (1_029_532_985.0
                + y * (9_494_680.718 + y * (59_272.64853 + y * (267.8532712 + y))));
        numerator / denominator
    } else {
        let z = 8.0 / absolute;
        let y = z * z;
        let angle = absolute - std::f64::consts::FRAC_PI_4;
        let first = 1.0
            + y * (-0.001098628627
                + y * (0.00002734510407 + y * (-0.000002073370639 + y * 0.0000002093887211)));
        let second = -0.01562499995
            + y * (0.0001430488765
                + y * (-0.000006911147651 + y * (0.0000007621095161 - y * 0.0000000934945152)));
        (0.636619772 / absolute).sqrt() * (angle.cos() * first - z * angle.sin() * second)
    }
}

pub(super) fn formula_bessel_j1(value: f64) -> f64 {
    let absolute = value.abs();
    let result = if absolute < 8.0 {
        let y = value * value;
        let numerator = absolute
            * (72_362_614_232.0
                + y * (-7_895_059_235.0
                    + y * (242_396_853.1
                        + y * (-2_972_611.439 + y * (15_704.48260 + y * -30.16036606)))));
        let denominator = 144_725_228_442.0
            + y * (2_300_535_178.0
                + y * (18_583_304.74 + y * (99_447.43394 + y * (376.9991397 + y))));
        numerator / denominator
    } else {
        let z = 8.0 / absolute;
        let y = z * z;
        let angle = absolute - 3.0 * std::f64::consts::FRAC_PI_4;
        let first = 1.0
            + y * (0.00183105
                + y * (-0.00003516396496 + y * (0.000002457520174 - y * 0.000000240337019)));
        let second = 0.04687499995
            + y * (-0.0002002690873
                + y * (0.000008449199096 + y * (-0.00000088228987 + y * 0.000000105787412)));
        (0.636619772 / absolute).sqrt() * (angle.cos() * first - z * angle.sin() * second)
    };
    if value < 0.0 { -result } else { result }
}

pub(super) fn formula_bessel_y0(value: f64) -> Result<f64, FormulaEvalError> {
    if value <= 0.0 || !value.is_finite() {
        return Err(if value.is_finite() {
            FormulaEvalError::Num
        } else {
            FormulaEvalError::Value
        });
    }
    if value < 8.0 {
        let y = value * value;
        let numerator = -2_957_821_389.0
            + y * (7_062_834_065.0
                + y * (-512_359_803.6
                    + y * (10_879_881.29 + y * (-86_327.92757 + y * 228.4622733))));
        let denominator = 40_076_544_269.0
            + y * (745_249_964.8
                + y * (7_189_466.438 + y * (47_447.26470 + y * (226.1030244 + y))));
        Ok(numerator / denominator + 0.636619772 * formula_bessel_j0(value) * value.ln())
    } else {
        let z = 8.0 / value;
        let y = z * z;
        let angle = value - std::f64::consts::FRAC_PI_4;
        let first = 1.0
            + y * (-0.001098628627
                + y * (0.00002734510407 + y * (-0.000002073370639 + y * 0.0000002093887211)));
        let second = -0.01562499995
            + y * (0.0001430488765
                + y * (-0.000006911147651 + y * (0.0000007621095161 - y * 0.0000000934945152)));
        Ok((0.636619772 / value).sqrt() * (angle.sin() * first + z * angle.cos() * second))
    }
}

pub(super) fn formula_bessel_y1(value: f64) -> Result<f64, FormulaEvalError> {
    if value <= 0.0 || !value.is_finite() {
        return Err(if value.is_finite() {
            FormulaEvalError::Num
        } else {
            FormulaEvalError::Value
        });
    }
    if value < 8.0 {
        let y = value * value;
        let numerator = value
            * (-4_900_604_943_000.0
                + y * (1_275_274_390_000.0
                    + y * (-51_534_381_390.0
                        + y * (734_926_455.1 + y * (-4_237_922.726 + y * 8_511.937935)))));
        let denominator = 24_995_805_700_000.0
            + y * (424_441_966_400.0
                + y * (3_733_650_367.0
                    + y * (22_459_040.02 + y * (102_042.6050 + y * (354.9632885 + y)))));
        Ok(numerator / denominator
            + 0.636619772 * (formula_bessel_j1(value) * value.ln() - 1.0 / value))
    } else {
        let z = 8.0 / value;
        let y = z * z;
        let angle = value - 3.0 * std::f64::consts::FRAC_PI_4;
        let first = 1.0
            + y * (0.00183105
                + y * (-0.00003516396496 + y * (0.000002457520174 - y * 0.000000240337019)));
        let second = 0.04687499995
            + y * (-0.0002002690873
                + y * (0.000008449199096 + y * (-0.00000088228987 + y * 0.000000105787412)));
        Ok((0.636619772 / value).sqrt() * (angle.sin() * first + z * angle.cos() * second))
    }
}

pub(super) fn formula_bessel_i0(value: f64) -> f64 {
    let absolute = value.abs();
    if absolute < 3.75 {
        let y = (value / 3.75).powi(2);
        1.0 + y
            * (3.5156229
                + y * (3.0899424
                    + y * (1.2067492 + y * (0.2659732 + y * (0.0360768 + y * 0.0045813)))))
    } else {
        let y = 3.75 / absolute;
        absolute.exp() / absolute.sqrt()
            * (0.39894228
                + y * (0.01328592
                    + y * (0.00225319
                        + y * (-0.00157565
                            + y * (0.00916281
                                + y * (-0.02057706
                                    + y * (0.02635537 + y * (-0.01647633 + y * 0.00392377))))))))
    }
}

pub(super) fn formula_bessel_i1(value: f64) -> f64 {
    let absolute = value.abs();
    let result = if absolute < 3.75 {
        let y = (value / 3.75).powi(2);
        absolute
            * (0.5
                + y * (0.87890594
                    + y * (0.51498869
                        + y * (0.15084934 + y * (0.02658733 + y * (0.00301532 + y * 0.00032411))))))
    } else {
        let y = 3.75 / absolute;
        absolute.exp() / absolute.sqrt()
            * (0.39894228
                + y * (-0.03988024
                    + y * (-0.00362018
                        + y * (0.00163801
                            + y * (-0.01031555
                                + y * (0.02282967
                                    + y * (-0.02895312 + y * (0.01787654 - y * 0.00420059))))))))
    };
    if value < 0.0 { -result } else { result }
}

pub(super) fn formula_bessel_k0(value: f64) -> Result<f64, FormulaEvalError> {
    if value <= 0.0 || !value.is_finite() {
        return Err(if value.is_finite() {
            FormulaEvalError::Num
        } else {
            FormulaEvalError::Value
        });
    }
    if value <= 2.0 {
        let y = value * value / 4.0;
        Ok(-(value / 2.0).ln() * formula_bessel_i0(value)
            + (-0.57721566
                + y * (0.42278420
                    + y * (0.23069756
                        + y * (0.03488590
                            + y * (0.00262698 + y * (0.00010750 + y * 0.00000740)))))))
    } else {
        let y = 2.0 / value;
        Ok((-value).exp() / value.sqrt()
            * (1.25331414
                + y * (-0.07832358
                    + y * (0.02189568
                        + y * (-0.01062446
                            + y * (0.00587872 + y * (-0.00251540 + y * 0.00053208)))))))
    }
}

pub(super) fn formula_bessel_k1(value: f64) -> Result<f64, FormulaEvalError> {
    if value <= 0.0 || !value.is_finite() {
        return Err(if value.is_finite() {
            FormulaEvalError::Num
        } else {
            FormulaEvalError::Value
        });
    }
    if value <= 2.0 {
        let y = value * value / 4.0;
        Ok((value / 2.0).ln() * formula_bessel_i1(value)
            + (1.0 / value)
                * (1.0
                    + y * (0.15443144
                        + y * (-0.67278579
                            + y * (-0.18156897
                                + y * (-0.01919402 + y * (-0.00110404 - y * 0.00004686)))))))
    } else {
        let y = 2.0 / value;
        Ok((-value).exp() / value.sqrt()
            * (1.25331414
                + y * (0.23498619
                    + y * (-0.03655620
                        + y * (0.01504268
                            + y * (-0.00780353 + y * (0.00325614 - y * 0.00068245)))))))
    }
}

pub(super) fn formula_bessel_i(value: f64, order: usize) -> Result<f64, FormulaEvalError> {
    if !value.is_finite() {
        return Err(FormulaEvalError::Value);
    }
    if order == 0 {
        return Ok(formula_bessel_i0(value));
    }
    if order == 1 {
        return Ok(formula_bessel_i1(value));
    }
    let absolute = value.abs();
    if absolute == 0.0 {
        return Ok(0.0);
    }
    const BIGNO: f64 = 1e100;
    const BIGNI: f64 = 1e-100;
    let tox = 2.0 / absolute;
    let mut bip = 0.0;
    let mut bi = 1.0;
    let mut answer = 0.0;
    let m = 2 * (order + (40.0 * order as f64).sqrt() as usize);
    for j in (1..=m).rev() {
        let bim = bip + j as f64 * tox * bi;
        bip = bi;
        bi = bim;
        if bi.abs() > BIGNO {
            answer *= BIGNI;
            bi *= BIGNI;
            bip *= BIGNI;
        }
        if j == order {
            answer = bip;
        }
    }
    answer *= formula_bessel_i0(absolute) / bi;
    if value < 0.0 && order % 2 == 1 {
        answer = -answer;
    }
    if answer.is_finite() {
        Ok(answer)
    } else {
        Err(FormulaEvalError::Num)
    }
}

pub(super) fn formula_bessel_j(value: f64, order: usize) -> Result<f64, FormulaEvalError> {
    if !value.is_finite() {
        return Err(FormulaEvalError::Value);
    }
    if order == 0 {
        return Ok(formula_bessel_j0(value));
    }
    if order == 1 {
        return Ok(formula_bessel_j1(value));
    }
    let absolute = value.abs();
    if absolute == 0.0 {
        return Ok(0.0);
    }
    let tox = 2.0 / absolute;
    let mut answer;
    if absolute > order as f64 {
        let mut previous = formula_bessel_j0(absolute);
        let mut current = formula_bessel_j1(absolute);
        for j in 1..order {
            let next = j as f64 * tox * current - previous;
            previous = current;
            current = next;
        }
        answer = current;
    } else {
        const BIGNO: f64 = 1e100;
        const BIGNI: f64 = 1e-100;
        let mut next = 0.0;
        let mut current = 1.0;
        let mut sum = 0.0;
        let mut include_in_sum = false;
        answer = 0.0;
        let m = 2 * ((order + (40.0 * order as f64).sqrt() as usize) / 2);
        for j in (1..=m).rev() {
            let previous = j as f64 * tox * current - next;
            next = current;
            current = previous;
            if current.abs() > BIGNO {
                answer *= BIGNI;
                current *= BIGNI;
                next *= BIGNI;
                sum *= BIGNI;
            }
            if include_in_sum {
                sum += current;
            }
            include_in_sum = !include_in_sum;
            if j == order {
                answer = next;
            }
        }
        sum = 2.0 * sum - current;
        answer /= sum;
    }
    if value < 0.0 && order % 2 == 1 {
        answer = -answer;
    }
    if answer.is_finite() {
        Ok(answer)
    } else {
        Err(FormulaEvalError::Num)
    }
}

pub(super) fn formula_bessel_k(value: f64, order: usize) -> Result<f64, FormulaEvalError> {
    if order == 0 {
        return formula_bessel_k0(value);
    }
    if order == 1 {
        return formula_bessel_k1(value);
    }
    let mut previous = formula_bessel_k0(value)?;
    let mut current = formula_bessel_k1(value)?;
    let tox = 2.0 / value;
    for j in 1..order {
        let next = previous + j as f64 * tox * current;
        previous = current;
        current = next;
        if !current.is_finite() {
            return Err(FormulaEvalError::Num);
        }
    }
    Ok(current)
}

pub(super) fn formula_bessel_y(value: f64, order: usize) -> Result<f64, FormulaEvalError> {
    if order == 0 {
        return formula_bessel_y0(value);
    }
    if order == 1 {
        return formula_bessel_y1(value);
    }
    let mut previous = formula_bessel_y0(value)?;
    let mut current = formula_bessel_y1(value)?;
    let tox = 2.0 / value;
    for j in 1..order {
        let next = j as f64 * tox * current - previous;
        previous = current;
        current = next;
        if !current.is_finite() {
            return Err(FormulaEvalError::Num);
        }
    }
    Ok(current)
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(super) enum FormulaConvertDimension {
    Area,
    Distance,
    Energy,
    Force,
    Information,
    Magnetism,
    Mass,
    Power,
    Pressure,
    Speed,
    Temperature,
    Time,
    Volume,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(super) enum FormulaConvertTemperatureUnit {
    Celsius,
    Fahrenheit,
    Kelvin,
    Rankine,
    Reaumur,
}

#[derive(Debug, Clone, Copy)]
pub(super) enum FormulaConvertScale {
    Ratio(f64),
    Temperature(FormulaConvertTemperatureUnit),
}

#[derive(Debug, Clone, Copy)]
pub(super) struct FormulaConvertUnit {
    pub(super) dimension: FormulaConvertDimension,
    pub(super) scale: FormulaConvertScale,
    pub(super) metric_power: i32,
    pub(super) binary_prefixable: bool,
}

pub(super) fn formula_convert_ratio_unit(
    dimension: FormulaConvertDimension,
    factor: f64,
    metric_power: i32,
    binary_prefixable: bool,
) -> FormulaConvertUnit {
    FormulaConvertUnit {
        dimension,
        scale: FormulaConvertScale::Ratio(factor),
        metric_power,
        binary_prefixable,
    }
}

pub(super) fn formula_convert_temperature_unit(
    unit: FormulaConvertTemperatureUnit,
) -> FormulaConvertUnit {
    FormulaConvertUnit {
        dimension: FormulaConvertDimension::Temperature,
        scale: FormulaConvertScale::Temperature(unit),
        metric_power: 0,
        binary_prefixable: false,
    }
}

pub(super) fn formula_convert_exact_unit(unit: &str) -> Option<FormulaConvertUnit> {
    use FormulaConvertDimension::*;
    const INCH: f64 = 0.0254;
    const FOOT: f64 = 0.3048;
    const YARD: f64 = 0.9144;
    const MILE: f64 = 1609.344;
    const SURVEY_MILE: f64 = 1609.3472186944373;
    const NAUTICAL_MILE: f64 = 1852.0;
    const LIGHT_YEAR: f64 = 9_460_730_472_580_800.0;
    const PARSEC: f64 = 30_856_775_814_913_670.0;
    const PICA_POINT: f64 = INCH / 72.0;
    const PICA: f64 = INCH / 6.0;
    const US_FLUID_OUNCE: f64 = 0.0000295735295625;
    const UK_PINT: f64 = 0.00056826125;
    const UK_GALLON: f64 = 0.00454609;
    let unit = match unit {
        "g" => formula_convert_ratio_unit(Mass, 1.0, 1, false),
        "sg" => formula_convert_ratio_unit(Mass, 14_593.90294, 0, false),
        "lbm" => formula_convert_ratio_unit(Mass, 453.59237, 0, false),
        "u" => formula_convert_ratio_unit(Mass, 1.660_539_066_6e-24, 0, false),
        "ozm" => formula_convert_ratio_unit(Mass, 28.349523125, 0, false),
        "grain" => formula_convert_ratio_unit(Mass, 0.06479891, 0, false),
        "cwt" | "shweight" => formula_convert_ratio_unit(Mass, 45_359.237, 0, false),
        "uk_cwt" | "lcwt" | "hweight" => formula_convert_ratio_unit(Mass, 50_802.34544, 0, false),
        "stone" => formula_convert_ratio_unit(Mass, 6_350.29318, 0, false),
        "ton" => formula_convert_ratio_unit(Mass, 907_184.74, 0, false),
        "uk_ton" | "LTON" | "brton" => formula_convert_ratio_unit(Mass, 1_016_046.9088, 0, false),
        "m" => formula_convert_ratio_unit(Distance, 1.0, 1, false),
        "mi" => formula_convert_ratio_unit(Distance, MILE, 0, false),
        "Nmi" => formula_convert_ratio_unit(Distance, NAUTICAL_MILE, 0, false),
        "in" => formula_convert_ratio_unit(Distance, INCH, 0, false),
        "ft" => formula_convert_ratio_unit(Distance, FOOT, 0, false),
        "yd" => formula_convert_ratio_unit(Distance, YARD, 0, false),
        "ang" => formula_convert_ratio_unit(Distance, 1e-10, 0, false),
        "ell" => formula_convert_ratio_unit(Distance, 45.0 * INCH, 0, false),
        "ly" => formula_convert_ratio_unit(Distance, LIGHT_YEAR, 0, false),
        "parsec" | "pc" => formula_convert_ratio_unit(Distance, PARSEC, 0, false),
        "Picapt" | "Pica" => formula_convert_ratio_unit(Distance, PICA_POINT, 0, false),
        "pica" => formula_convert_ratio_unit(Distance, PICA, 0, false),
        "survey_mi" => formula_convert_ratio_unit(Distance, SURVEY_MILE, 0, false),
        "yr" => formula_convert_ratio_unit(Time, 31_557_600.0, 0, false),
        "day" | "d" => formula_convert_ratio_unit(Time, 86_400.0, 0, false),
        "hr" => formula_convert_ratio_unit(Time, 3_600.0, 0, false),
        "mn" | "min" => formula_convert_ratio_unit(Time, 60.0, 0, false),
        "sec" | "s" => formula_convert_ratio_unit(Time, 1.0, 1, false),
        "Pa" | "p" => formula_convert_ratio_unit(Pressure, 1.0, 1, false),
        "atm" | "at" => formula_convert_ratio_unit(Pressure, 101_325.0, 0, false),
        "mmHg" => formula_convert_ratio_unit(Pressure, 133.322, 0, false),
        "psi" => formula_convert_ratio_unit(Pressure, 6_894.757293168361, 0, false),
        "Torr" => formula_convert_ratio_unit(Pressure, 101_325.0 / 760.0, 0, false),
        "N" => formula_convert_ratio_unit(Force, 1.0, 1, false),
        "dyn" | "dy" => formula_convert_ratio_unit(Force, 1e-5, 1, false),
        "lbf" => formula_convert_ratio_unit(Force, 4.4482216152605, 0, false),
        "J" => formula_convert_ratio_unit(Energy, 1.0, 1, false),
        "e" => formula_convert_ratio_unit(Energy, 1e-7, 1, false),
        "c" => formula_convert_ratio_unit(Energy, 4.184, 1, false),
        "cal" => formula_convert_ratio_unit(Energy, 4.1868, 1, false),
        "eV" | "ev" => formula_convert_ratio_unit(Energy, 1.602_176_634e-19, 1, false),
        "HPh" | "hh" => formula_convert_ratio_unit(Energy, 2_684_519.538, 0, false),
        "Wh" | "wh" => formula_convert_ratio_unit(Energy, 3_600.0, 1, false),
        "flb" => formula_convert_ratio_unit(Energy, 1.3558179483314004, 0, false),
        "BTU" | "btu" => formula_convert_ratio_unit(Energy, 1_055.05585262, 0, false),
        "HP" | "h" => formula_convert_ratio_unit(Power, 745.6998715822702, 0, false),
        "PS" => formula_convert_ratio_unit(Power, 735.49875, 0, false),
        "W" | "w" => formula_convert_ratio_unit(Power, 1.0, 1, false),
        "T" => formula_convert_ratio_unit(Magnetism, 1.0, 1, false),
        "ga" => formula_convert_ratio_unit(Magnetism, 1e-4, 0, false),
        "C" | "cel" => formula_convert_temperature_unit(FormulaConvertTemperatureUnit::Celsius),
        "F" | "fah" => formula_convert_temperature_unit(FormulaConvertTemperatureUnit::Fahrenheit),
        "K" | "kel" => formula_convert_temperature_unit(FormulaConvertTemperatureUnit::Kelvin),
        "Rank" => formula_convert_temperature_unit(FormulaConvertTemperatureUnit::Rankine),
        "Reau" => formula_convert_temperature_unit(FormulaConvertTemperatureUnit::Reaumur),
        "tsp" => formula_convert_ratio_unit(Volume, US_FLUID_OUNCE / 6.0, 0, false),
        "tspm" => formula_convert_ratio_unit(Volume, 0.000005, 0, false),
        "tbs" => formula_convert_ratio_unit(Volume, US_FLUID_OUNCE / 2.0, 0, false),
        "oz" => formula_convert_ratio_unit(Volume, US_FLUID_OUNCE, 0, false),
        "cup" => formula_convert_ratio_unit(Volume, US_FLUID_OUNCE * 8.0, 0, false),
        "pt" | "us_pt" => formula_convert_ratio_unit(Volume, US_FLUID_OUNCE * 16.0, 0, false),
        "uk_pt" => formula_convert_ratio_unit(Volume, UK_PINT, 0, false),
        "qt" => formula_convert_ratio_unit(Volume, US_FLUID_OUNCE * 32.0, 0, false),
        "uk_qt" => formula_convert_ratio_unit(Volume, UK_PINT * 2.0, 0, false),
        "gal" => formula_convert_ratio_unit(Volume, US_FLUID_OUNCE * 128.0, 0, false),
        "uk_gal" => formula_convert_ratio_unit(Volume, UK_GALLON, 0, false),
        "l" | "L" | "lt" => formula_convert_ratio_unit(Volume, 0.001, 1, false),
        "ang3" | "ang^3" => formula_convert_ratio_unit(Volume, 1e-30, 0, false),
        "barrel" => formula_convert_ratio_unit(Volume, US_FLUID_OUNCE * 128.0 * 42.0, 0, false),
        "bushel" => formula_convert_ratio_unit(Volume, 2_150.42 * INCH.powi(3), 0, false),
        "ft3" | "ft^3" => formula_convert_ratio_unit(Volume, FOOT.powi(3), 0, false),
        "in3" | "in^3" => formula_convert_ratio_unit(Volume, INCH.powi(3), 0, false),
        "ly3" | "ly^3" => formula_convert_ratio_unit(Volume, LIGHT_YEAR.powi(3), 0, false),
        "m3" | "m^3" => formula_convert_ratio_unit(Volume, 1.0, 3, false),
        "mi3" | "mi^3" => formula_convert_ratio_unit(Volume, MILE.powi(3), 0, false),
        "yd3" | "yd^3" => formula_convert_ratio_unit(Volume, YARD.powi(3), 0, false),
        "Nmi3" | "Nmi^3" => formula_convert_ratio_unit(Volume, NAUTICAL_MILE.powi(3), 0, false),
        "Picapt3" | "Picapt^3" | "Pica3" | "Pica^3" => {
            formula_convert_ratio_unit(Volume, PICA_POINT.powi(3), 0, false)
        }
        "GRT" | "regton" => formula_convert_ratio_unit(Volume, 100.0 * FOOT.powi(3), 0, false),
        "MTON" => formula_convert_ratio_unit(Volume, 40.0 * FOOT.powi(3), 0, false),
        "uk_acre" => formula_convert_ratio_unit(Area, 4_046.8564224, 0, false),
        "us_acre" => formula_convert_ratio_unit(Area, 4_046.872609874252, 0, false),
        "ang2" | "ang^2" => formula_convert_ratio_unit(Area, 1e-20, 0, false),
        "ar" => formula_convert_ratio_unit(Area, 100.0, 1, false),
        "ft2" | "ft^2" => formula_convert_ratio_unit(Area, FOOT.powi(2), 0, false),
        "ha" => formula_convert_ratio_unit(Area, 10_000.0, 0, false),
        "in2" | "in^2" => formula_convert_ratio_unit(Area, INCH.powi(2), 0, false),
        "ly2" | "ly^2" => formula_convert_ratio_unit(Area, LIGHT_YEAR.powi(2), 0, false),
        "m2" | "m^2" => formula_convert_ratio_unit(Area, 1.0, 2, false),
        "Morgen" => formula_convert_ratio_unit(Area, 2_500.0, 0, false),
        "mi2" | "mi^2" => formula_convert_ratio_unit(Area, MILE.powi(2), 0, false),
        "Nmi2" | "Nmi^2" => formula_convert_ratio_unit(Area, NAUTICAL_MILE.powi(2), 0, false),
        "Picapt2" | "Pica2" | "Pica^2" | "Picapt^2" => {
            formula_convert_ratio_unit(Area, PICA_POINT.powi(2), 0, false)
        }
        "yd2" | "yd^2" => formula_convert_ratio_unit(Area, YARD.powi(2), 0, false),
        "bit" => formula_convert_ratio_unit(Information, 1.0, 1, true),
        "byte" => formula_convert_ratio_unit(Information, 8.0, 1, true),
        "admkn" => formula_convert_ratio_unit(Speed, 6080.0 * FOOT / 3600.0, 0, false),
        "kn" => formula_convert_ratio_unit(Speed, NAUTICAL_MILE / 3600.0, 0, false),
        "m/h" | "m/hr" => formula_convert_ratio_unit(Speed, 1.0 / 3600.0, 1, false),
        "m/s" | "m/sec" => formula_convert_ratio_unit(Speed, 1.0, 1, false),
        "mph" => formula_convert_ratio_unit(Speed, MILE / 3600.0, 0, false),
        _ => return None,
    };
    Some(unit)
}

pub(super) fn formula_convert_metric_prefix(unit: &str) -> Option<(f64, &str)> {
    const PREFIXES: [(&str, f64); 21] = [
        ("da", 1e1),
        ("Y", 1e24),
        ("Z", 1e21),
        ("E", 1e18),
        ("P", 1e15),
        ("T", 1e12),
        ("G", 1e9),
        ("M", 1e6),
        ("k", 1e3),
        ("h", 1e2),
        ("e", 1e1),
        ("d", 1e-1),
        ("c", 1e-2),
        ("m", 1e-3),
        ("u", 1e-6),
        ("n", 1e-9),
        ("p", 1e-12),
        ("f", 1e-15),
        ("a", 1e-18),
        ("z", 1e-21),
        ("y", 1e-24),
    ];
    PREFIXES
        .iter()
        .find_map(|(prefix, factor)| unit.strip_prefix(prefix).map(|rest| (*factor, rest)))
        .filter(|(_, rest)| !rest.is_empty())
}

pub(super) fn formula_convert_binary_prefix(unit: &str) -> Option<(f64, &str)> {
    const PREFIXES: [(&str, f64); 8] = [
        ("Yi", 1_208_925_819_614_629_174_706_176.0),
        ("Zi", 1_180_591_620_717_411_303_424.0),
        ("Ei", 1_152_921_504_606_846_976.0),
        ("Pi", 1_125_899_906_842_624.0),
        ("Ti", 1_099_511_627_776.0),
        ("Gi", 1_073_741_824.0),
        ("Mi", 1_048_576.0),
        ("ki", 1_024.0),
    ];
    PREFIXES
        .iter()
        .find_map(|(prefix, factor)| unit.strip_prefix(prefix).map(|rest| (*factor, rest)))
        .filter(|(_, rest)| !rest.is_empty())
}

pub(super) fn formula_convert_unit(unit: &str) -> Result<FormulaConvertUnit, FormulaEvalError> {
    if let Some(unit) = formula_convert_exact_unit(unit) {
        return Ok(unit);
    }
    if let Some((factor, suffix)) = formula_convert_binary_prefix(unit) {
        if let Some(mut unit) = formula_convert_exact_unit(suffix) {
            if unit.binary_prefixable {
                if let FormulaConvertScale::Ratio(base_factor) = unit.scale {
                    unit.scale = FormulaConvertScale::Ratio(base_factor * factor);
                    return Ok(unit);
                }
            }
        }
        return Err(FormulaEvalError::NA);
    }
    if let Some((factor, suffix)) = formula_convert_metric_prefix(unit) {
        if let Some(mut unit) = formula_convert_exact_unit(suffix) {
            if unit.metric_power > 0 {
                if let FormulaConvertScale::Ratio(base_factor) = unit.scale {
                    unit.scale =
                        FormulaConvertScale::Ratio(base_factor * factor.powi(unit.metric_power));
                    return Ok(unit);
                }
            }
        }
        return Err(FormulaEvalError::NA);
    }
    Err(FormulaEvalError::NA)
}

pub(super) fn formula_convert_temperature_to_kelvin(
    value: f64,
    unit: FormulaConvertTemperatureUnit,
) -> f64 {
    match unit {
        FormulaConvertTemperatureUnit::Celsius => value + 273.15,
        FormulaConvertTemperatureUnit::Fahrenheit => (value + 459.67) * 5.0 / 9.0,
        FormulaConvertTemperatureUnit::Kelvin => value,
        FormulaConvertTemperatureUnit::Rankine => value * 5.0 / 9.0,
        FormulaConvertTemperatureUnit::Reaumur => value * 1.25 + 273.15,
    }
}

pub(super) fn formula_convert_temperature_from_kelvin(
    value: f64,
    unit: FormulaConvertTemperatureUnit,
) -> f64 {
    match unit {
        FormulaConvertTemperatureUnit::Celsius => value - 273.15,
        FormulaConvertTemperatureUnit::Fahrenheit => value * 9.0 / 5.0 - 459.67,
        FormulaConvertTemperatureUnit::Kelvin => value,
        FormulaConvertTemperatureUnit::Rankine => value * 9.0 / 5.0,
        FormulaConvertTemperatureUnit::Reaumur => (value - 273.15) * 0.8,
    }
}

pub(super) fn formula_convert_value(
    value: f64,
    from_unit: &str,
    to_unit: &str,
) -> Result<f64, FormulaEvalError> {
    if !value.is_finite() {
        return Err(FormulaEvalError::Value);
    }
    let from_unit = formula_convert_unit(from_unit)?;
    let to_unit = formula_convert_unit(to_unit)?;
    if from_unit.dimension != to_unit.dimension {
        return Err(FormulaEvalError::NA);
    }
    let result = match (from_unit.scale, to_unit.scale) {
        (FormulaConvertScale::Ratio(from_factor), FormulaConvertScale::Ratio(to_factor)) => {
            value * from_factor / to_factor
        }
        (
            FormulaConvertScale::Temperature(from_unit),
            FormulaConvertScale::Temperature(to_unit),
        ) => formula_convert_temperature_from_kelvin(
            formula_convert_temperature_to_kelvin(value, from_unit),
            to_unit,
        ),
        _ => return Err(FormulaEvalError::NA),
    };
    if result.is_finite() {
        Ok(result)
    } else {
        Err(FormulaEvalError::Num)
    }
}

#[derive(Debug, Clone, Copy)]
pub(super) struct FormulaEuroCurrency {
    pub(super) rate: f64,
    pub(super) calculation_precision: i32,
}

pub(super) fn formula_euro_currency(code: &str) -> Option<FormulaEuroCurrency> {
    let currency = match code {
        "ATS" => FormulaEuroCurrency {
            rate: 13.7603,
            calculation_precision: 2,
        },
        "BEF" | "LUF" => FormulaEuroCurrency {
            rate: 40.3399,
            calculation_precision: 0,
        },
        "DEM" => FormulaEuroCurrency {
            rate: 1.95583,
            calculation_precision: 2,
        },
        "ESP" => FormulaEuroCurrency {
            rate: 166.386,
            calculation_precision: 0,
        },
        "EUR" => FormulaEuroCurrency {
            rate: 1.0,
            calculation_precision: 2,
        },
        "FIM" => FormulaEuroCurrency {
            rate: 5.94573,
            calculation_precision: 2,
        },
        "FRF" => FormulaEuroCurrency {
            rate: 6.55957,
            calculation_precision: 2,
        },
        "GRD" => FormulaEuroCurrency {
            rate: 340.75,
            calculation_precision: 0,
        },
        "IEP" => FormulaEuroCurrency {
            rate: 0.787564,
            calculation_precision: 2,
        },
        "ITL" => FormulaEuroCurrency {
            rate: 1936.27,
            calculation_precision: 0,
        },
        "NLG" => FormulaEuroCurrency {
            rate: 2.20371,
            calculation_precision: 2,
        },
        "PTE" => FormulaEuroCurrency {
            rate: 200.482,
            calculation_precision: 0,
        },
        "SIT" => FormulaEuroCurrency {
            rate: 239.64,
            calculation_precision: 2,
        },
        _ => return None,
    };
    Some(currency)
}

pub(super) fn formula_round_to_decimal_places(
    value: f64,
    places: i32,
) -> Result<f64, FormulaEvalError> {
    let factor = 10_f64.powi(places);
    if !factor.is_finite() || factor == 0.0 {
        return Err(FormulaEvalError::Num);
    }
    formula_checked_numeric_result(round_half_away_from_zero(value * factor) / factor)
}

pub(super) fn formula_round_to_significant_digits(
    value: f64,
    digits: i64,
) -> Result<f64, FormulaEvalError> {
    if digits < 1 || digits > i64::from(i32::MAX) {
        return Err(FormulaEvalError::Value);
    }
    if value == 0.0 {
        return Ok(0.0);
    }
    let magnitude = value.abs().log10().floor();
    if !magnitude.is_finite() || magnitude < i32::MIN as f64 || magnitude > i32::MAX as f64 {
        return Err(FormulaEvalError::Num);
    }
    let places = i32::try_from(digits - 1).map_err(|_| FormulaEvalError::Value)? - magnitude as i32;
    formula_round_to_decimal_places(value, places)
}

pub(super) fn formula_euroconvert_value(
    value: f64,
    source_code: &str,
    target_code: &str,
    full_precision: bool,
    triangulation_precision: Option<i64>,
) -> Result<f64, FormulaEvalError> {
    if !value.is_finite() {
        return Err(FormulaEvalError::Value);
    }
    let source_code = source_code.to_ascii_uppercase();
    let target_code = target_code.to_ascii_uppercase();
    if source_code == target_code {
        return Ok(value);
    }
    let source = formula_euro_currency(source_code.as_str()).ok_or(FormulaEvalError::Value)?;
    let target = formula_euro_currency(target_code.as_str()).ok_or(FormulaEvalError::Value)?;
    let mut euros = value / source.rate;
    if let Some(precision) = triangulation_precision {
        if precision < 3 {
            return Err(FormulaEvalError::Value);
        }
        if source_code != "EUR" {
            euros = formula_round_to_significant_digits(euros, precision)?;
        }
    }
    let result = euros * target.rate;
    if full_precision {
        formula_checked_numeric_result(result)
    } else {
        formula_round_to_decimal_places(result, target.calculation_precision)
    }
}

pub(super) fn formula_bitwise_argument(value: f64) -> Result<u64, FormulaEvalError> {
    let value = formula_integer_argument(value)?;
    let value = u64::try_from(value).map_err(|_| FormulaEvalError::Num)?;
    if value > ((1_u64 << 48) - 1) {
        return Err(FormulaEvalError::Num);
    }
    Ok(value)
}

pub(super) fn formula_bit_shift_argument(value: f64) -> Result<i64, FormulaEvalError> {
    let value = formula_integer_argument(value)?;
    if !(-53..=53).contains(&value) {
        return Err(FormulaEvalError::Num);
    }
    Ok(value)
}

pub(super) fn formula_engineering_input(
    text: &str,
    radix: u32,
    bits: u32,
    max_digits: usize,
) -> Result<i64, FormulaEvalError> {
    let text = text.trim();
    if text.is_empty() || text.len() > max_digits {
        return Err(FormulaEvalError::Num);
    }
    let mut value = 0_u64;
    for ch in text.chars() {
        let Some(digit) = ch.to_digit(radix) else {
            return Err(FormulaEvalError::Num);
        };
        value = value
            .checked_mul(u64::from(radix))
            .and_then(|current| current.checked_add(u64::from(digit)))
            .ok_or(FormulaEvalError::Num)?;
    }
    let sign_threshold = 1_u64 << (bits - 1);
    let modulus = 1_u64 << bits;
    if value >= modulus {
        return Err(FormulaEvalError::Num);
    }
    if value >= sign_threshold {
        Ok(value as i64 - modulus as i64)
    } else {
        Ok(value as i64)
    }
}

pub(super) fn formula_engineering_format(
    value: i64,
    radix: u32,
    bits: u32,
    max_digits: usize,
    places: Option<usize>,
) -> Result<String, FormulaEvalError> {
    let minimum = -(1_i64 << (bits - 1));
    let maximum = (1_i64 << (bits - 1)) - 1;
    if value < minimum || value > maximum {
        return Err(FormulaEvalError::Num);
    }
    let unsigned = if value < 0 {
        ((1_i128 << bits) + i128::from(value)) as u128
    } else {
        value as u128
    };
    let mut output = formula_unsigned_radix_text(unsigned, radix);
    if value < 0 {
        if output.len() < max_digits {
            output = "0".repeat(max_digits - output.len()) + output.as_str();
        }
        return Ok(output);
    }
    if let Some(places) = places {
        if places == 0 || output.len() > places || places > max_digits {
            return Err(FormulaEvalError::Num);
        }
        if output.len() < places {
            output = "0".repeat(places - output.len()) + output.as_str();
        }
    }
    Ok(output)
}

pub(super) fn formula_unsigned_radix_text(mut value: u128, radix: u32) -> String {
    if value == 0 {
        return "0".to_string();
    }
    let mut output = String::new();
    while value > 0 {
        let digit = (value % u128::from(radix)) as u32;
        output.push(
            char::from_digit(digit, radix)
                .expect("digit")
                .to_ascii_uppercase(),
        );
        value /= u128::from(radix);
    }
    output.chars().rev().collect()
}

pub(super) fn formula_roman_text(value: i64, form: usize) -> Result<String, FormulaEvalError> {
    if !(0..=3999).contains(&value) || form > 4 {
        return Err(FormulaEvalError::Value);
    }
    if value == 0 {
        return Ok(String::new());
    }
    let mut remaining = value;
    let mut output = String::new();
    for (candidate, text) in [
        (1000, "M"),
        (900, "CM"),
        (500, "D"),
        (400, "CD"),
        (100, "C"),
        (90, "XC"),
        (50, "L"),
        (40, "XL"),
        (10, "X"),
        (9, "IX"),
        (5, "V"),
        (4, "IV"),
        (1, "I"),
    ] {
        while remaining >= candidate {
            output.push_str(text);
            remaining -= candidate;
        }
    }
    if form == 0 {
        return Ok(output);
    }
    const CONCISE_REPLACEMENTS: [&[(&str, &str)]; 5] = [
        &[
            ("XLV", "VL"),
            ("XCV", "VC"),
            ("CDL", "LD"),
            ("CML", "LM"),
            ("CMVC", "LMVL"),
        ],
        &[
            ("CDXC", "LDXL"),
            ("CDVC", "LDVL"),
            ("CMXC", "LMXL"),
            ("XCIX", "VCIV"),
            ("XLIX", "VLIV"),
        ],
        &[
            ("XLIX", "IL"),
            ("XCIX", "IC"),
            ("CDXC", "XD"),
            ("CDVC", "XDV"),
            ("CDIC", "XDIX"),
            ("LMVL", "XMV"),
            ("CMIC", "XMIX"),
            ("CMXC", "XM"),
        ],
        &[
            ("XDV", "VD"),
            ("XDIX", "VDIV"),
            ("XMV", "VM"),
            ("XMIX", "VMIV"),
        ],
        &[("VDIV", "ID"), ("VMIV", "IM")],
    ];
    for (index, replacements) in CONCISE_REPLACEMENTS.iter().enumerate().take(form + 1) {
        if index == 1 && form > 1 {
            continue;
        }
        for (from, to) in *replacements {
            output = output.replace(from, to);
        }
    }
    Ok(output)
}
