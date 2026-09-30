//! The aggregate worksheet function table (SUM, AVERAGE, STDEV, and relatives).

use super::*;

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(super) enum FormulaAggregateFunction {
    Sum,
    Product,
    SumSq,
    Min,
    Max,
    Median,
    Average,
    Gcd,
    GeoMean,
    HarMean,
    Kurt,
    Lcm,
    ModeMult,
    ModeSngl,
    Skew,
    SkewP,
    AveDev,
    DevSq,
    VarP,
    VarS,
    StDevP,
    StDevS,
    Count,
}

impl FormulaAggregateFunction {
    pub(super) fn from_name(name: &str) -> Option<Self> {
        if name.eq_ignore_ascii_case("SUM") {
            Some(Self::Sum)
        } else if name.eq_ignore_ascii_case("PRODUCT") {
            Some(Self::Product)
        } else if name.eq_ignore_ascii_case("SUMSQ") {
            Some(Self::SumSq)
        } else if name.eq_ignore_ascii_case("MIN") {
            Some(Self::Min)
        } else if name.eq_ignore_ascii_case("MAX") {
            Some(Self::Max)
        } else if name.eq_ignore_ascii_case("MEDIAN") {
            Some(Self::Median)
        } else if name.eq_ignore_ascii_case("AVERAGE") {
            Some(Self::Average)
        } else if name.eq_ignore_ascii_case("GCD") {
            Some(Self::Gcd)
        } else if name.eq_ignore_ascii_case("GEOMEAN") {
            Some(Self::GeoMean)
        } else if name.eq_ignore_ascii_case("HARMEAN") {
            Some(Self::HarMean)
        } else if name.eq_ignore_ascii_case("KURT") {
            Some(Self::Kurt)
        } else if name.eq_ignore_ascii_case("LCM") {
            Some(Self::Lcm)
        } else if name.eq_ignore_ascii_case("MODE.MULT") {
            Some(Self::ModeMult)
        } else if name.eq_ignore_ascii_case("MODE") || name.eq_ignore_ascii_case("MODE.SNGL") {
            Some(Self::ModeSngl)
        } else if name.eq_ignore_ascii_case("SKEW") {
            Some(Self::Skew)
        } else if name.eq_ignore_ascii_case("SKEW.P") {
            Some(Self::SkewP)
        } else if name.eq_ignore_ascii_case("AVEDEV") {
            Some(Self::AveDev)
        } else if name.eq_ignore_ascii_case("DEVSQ") {
            Some(Self::DevSq)
        } else if name.eq_ignore_ascii_case("VAR.P") || name.eq_ignore_ascii_case("VARP") {
            Some(Self::VarP)
        } else if name.eq_ignore_ascii_case("VAR.S") || name.eq_ignore_ascii_case("VAR") {
            Some(Self::VarS)
        } else if name.eq_ignore_ascii_case("STDEV.P") || name.eq_ignore_ascii_case("STDEVP") {
            Some(Self::StDevP)
        } else if name.eq_ignore_ascii_case("STDEV.S") || name.eq_ignore_ascii_case("STDEV") {
            Some(Self::StDevS)
        } else if name.eq_ignore_ascii_case("COUNT") {
            Some(Self::Count)
        } else {
            None
        }
    }

    pub(super) fn evaluate(self, values: &[f64]) -> Result<f64, FormulaEvalError> {
        let mean = |values: &[f64]| -> Result<f64, FormulaEvalError> {
            if values.is_empty() {
                Err(FormulaEvalError::Div0)
            } else {
                Ok(values.iter().sum::<f64>() / values.len() as f64)
            }
        };
        let deviation_sum = |values: &[f64]| -> Result<f64, FormulaEvalError> {
            let mean = mean(values)?;
            Ok(values
                .iter()
                .map(|value| {
                    let deviation = value - mean;
                    deviation * deviation
                })
                .sum())
        };
        let checked_numeric_result = |value: f64| -> Result<f64, FormulaEvalError> {
            if value.is_finite() {
                Ok(value)
            } else {
                Err(FormulaEvalError::Num)
            }
        };
        const EXCEL_INTEGER_LIMIT: u64 = 1_u64 << 53;
        let trunc_excel_nonnegative_integer = |value: f64| -> Result<u64, FormulaEvalError> {
            if !value.is_finite() {
                return Err(FormulaEvalError::Value);
            }
            let value = value.trunc();
            if value < 0.0 || value >= EXCEL_INTEGER_LIMIT as f64 {
                return Err(FormulaEvalError::Num);
            }
            Ok(value as u64)
        };
        let gcd_u64 = |mut left: u64, mut right: u64| -> u64 {
            while right != 0 {
                let next = left % right;
                left = right;
                right = next;
            }
            left
        };

        match self {
            FormulaAggregateFunction::Sum => Ok(values.iter().sum()),
            FormulaAggregateFunction::Product => Ok(values.iter().product()),
            FormulaAggregateFunction::SumSq => Ok(values.iter().map(|value| value * value).sum()),
            FormulaAggregateFunction::Min => {
                Ok(values.iter().copied().reduce(f64::min).unwrap_or(0.0))
            }
            FormulaAggregateFunction::Max => {
                Ok(values.iter().copied().reduce(f64::max).unwrap_or(0.0))
            }
            FormulaAggregateFunction::Median => {
                if values.is_empty() {
                    return Err(FormulaEvalError::Num);
                }
                let mut values = values.to_vec();
                values.sort_by(|left, right| left.total_cmp(right));
                let midpoint = values.len() / 2;
                if values.len() % 2 == 0 {
                    Ok((values[midpoint - 1] + values[midpoint]) / 2.0)
                } else {
                    Ok(values[midpoint])
                }
            }
            FormulaAggregateFunction::Average => {
                if values.is_empty() {
                    Err(FormulaEvalError::Div0)
                } else {
                    Ok(values.iter().sum::<f64>() / values.len() as f64)
                }
            }
            FormulaAggregateFunction::Gcd => {
                if values.is_empty() {
                    return Err(FormulaEvalError::Value);
                }
                let mut result = 0_u64;
                for value in values {
                    result = gcd_u64(result, trunc_excel_nonnegative_integer(*value)?);
                }
                Ok(result as f64)
            }
            FormulaAggregateFunction::GeoMean => {
                if values.is_empty() {
                    return Err(FormulaEvalError::Div0);
                }
                let mut log_sum = 0.0;
                for value in values {
                    if *value <= 0.0 {
                        return Err(FormulaEvalError::Num);
                    }
                    log_sum += value.ln();
                }
                Ok((log_sum / values.len() as f64).exp())
            }
            FormulaAggregateFunction::HarMean => {
                if values.is_empty() {
                    return Err(FormulaEvalError::Div0);
                }
                let mut reciprocal_sum = 0.0;
                for value in values {
                    if *value <= 0.0 {
                        return Err(FormulaEvalError::Num);
                    }
                    reciprocal_sum += 1.0 / value;
                }
                Ok(values.len() as f64 / reciprocal_sum)
            }
            FormulaAggregateFunction::Kurt => {
                let count = values.len();
                if count < 4 {
                    return Err(FormulaEvalError::Div0);
                }
                let mean = mean(values)?;
                let mut deviation_square_sum = 0.0_f64;
                let mut deviation_fourth_sum = 0.0_f64;
                for value in values {
                    let deviation = value - mean;
                    let deviation_square = deviation * deviation;
                    deviation_square_sum += deviation_square;
                    deviation_fourth_sum += deviation_square * deviation_square;
                }
                if deviation_square_sum == 0.0 {
                    return Err(FormulaEvalError::Div0);
                }
                let count = count as f64;
                let sample_variance = deviation_square_sum / (count - 1.0);
                let kurtosis = count * (count + 1.0) * deviation_fourth_sum
                    / ((count - 1.0)
                        * (count - 2.0)
                        * (count - 3.0)
                        * sample_variance
                        * sample_variance)
                    - 3.0 * (count - 1.0) * (count - 1.0) / ((count - 2.0) * (count - 3.0));
                checked_numeric_result(kurtosis)
            }
            FormulaAggregateFunction::Lcm => {
                if values.is_empty() {
                    return Err(FormulaEvalError::Value);
                }
                let mut result = 1_u64;
                for value in values {
                    let value = trunc_excel_nonnegative_integer(*value)?;
                    if value == 0 {
                        return Ok(0.0);
                    }
                    let next = (result / gcd_u64(result, value))
                        .checked_mul(value)
                        .ok_or(FormulaEvalError::Num)?;
                    if next >= EXCEL_INTEGER_LIMIT {
                        return Err(FormulaEvalError::Num);
                    }
                    result = next;
                }
                Ok(result as f64)
            }
            FormulaAggregateFunction::ModeMult | FormulaAggregateFunction::ModeSngl => {
                let mut mode = None;
                let mut mode_count = 1_usize;
                for value in values {
                    let count = values
                        .iter()
                        .filter(|candidate| **candidate == *value)
                        .count();
                    if count > mode_count {
                        mode = Some(*value);
                        mode_count = count;
                    }
                }
                mode.ok_or(FormulaEvalError::NA)
            }
            FormulaAggregateFunction::Skew => {
                let count = values.len();
                if count < 3 {
                    return Err(FormulaEvalError::Div0);
                }
                let mean = mean(values)?;
                let mut deviation_square_sum = 0.0_f64;
                let mut deviation_cube_sum = 0.0_f64;
                for value in values {
                    let deviation = value - mean;
                    deviation_square_sum += deviation * deviation;
                    deviation_cube_sum += deviation * deviation * deviation;
                }
                if deviation_square_sum == 0.0 {
                    return Err(FormulaEvalError::Div0);
                }
                let count = count as f64;
                let sample_standard_deviation = (deviation_square_sum / (count - 1.0)).sqrt();
                checked_numeric_result(
                    count * deviation_cube_sum
                        / ((count - 1.0)
                            * (count - 2.0)
                            * sample_standard_deviation
                            * sample_standard_deviation
                            * sample_standard_deviation),
                )
            }
            FormulaAggregateFunction::SkewP => {
                let count = values.len();
                if count < 3 {
                    return Err(FormulaEvalError::Div0);
                }
                let mean = mean(values)?;
                let mut deviation_square_sum = 0.0_f64;
                let mut deviation_cube_sum = 0.0_f64;
                for value in values {
                    let deviation = value - mean;
                    deviation_square_sum += deviation * deviation;
                    deviation_cube_sum += deviation * deviation * deviation;
                }
                if deviation_square_sum == 0.0 {
                    return Err(FormulaEvalError::Div0);
                }
                let count = count as f64;
                let population_standard_deviation = (deviation_square_sum / count).sqrt();
                checked_numeric_result(
                    deviation_cube_sum
                        / count
                        / (population_standard_deviation
                            * population_standard_deviation
                            * population_standard_deviation),
                )
            }
            FormulaAggregateFunction::AveDev => {
                let mean = mean(values)?;
                Ok(values.iter().map(|value| (value - mean).abs()).sum::<f64>()
                    / values.len() as f64)
            }
            FormulaAggregateFunction::DevSq => deviation_sum(values),
            FormulaAggregateFunction::VarP => Ok(deviation_sum(values)? / values.len() as f64),
            FormulaAggregateFunction::VarS => {
                if values.len() < 2 {
                    Err(FormulaEvalError::Div0)
                } else {
                    Ok(deviation_sum(values)? / (values.len() - 1) as f64)
                }
            }
            FormulaAggregateFunction::StDevP => {
                Ok((deviation_sum(values)? / values.len() as f64).sqrt())
            }
            FormulaAggregateFunction::StDevS => {
                if values.len() < 2 {
                    Err(FormulaEvalError::Div0)
                } else {
                    Ok((deviation_sum(values)? / (values.len() - 1) as f64).sqrt())
                }
            }
            FormulaAggregateFunction::Count => Ok(values.len() as f64),
        }
    }
}
