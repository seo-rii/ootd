//! Formula parser methods for conditional aggregates, SUBTOTAL/AGGREGATE, statistical tests, regression and forecasting, GROUPBY/PIVOTBY, and cube stubs.

use super::*;

impl<'a, 'b, 'state> FormulaParser<'a, 'b, 'state> {
    pub(super) fn parse_forecast_ets_function(
        &mut self,
        name: &str,
    ) -> Result<f64, FormulaEvalError> {
        macro_rules! parse_optional_number {
            ($default:expr) => {{
                self.skip_whitespace();
                if self.peek_char().is_some_and(|ch| matches!(ch, ',' | ')')) {
                    $default
                } else {
                    self.parse_comparison()?
                }
            }};
        }

        let is_forecast = name.eq_ignore_ascii_case("FORECAST.ETS");
        let is_confint = name.eq_ignore_ascii_case("FORECAST.ETS.CONFINT");
        let is_seasonality = name.eq_ignore_ascii_case("FORECAST.ETS.SEASONALITY");
        let is_stat = name.eq_ignore_ascii_case("FORECAST.ETS.STAT");

        let mut target_date = None;
        if is_forecast || is_confint {
            let target = self.parse_comparison()?;
            if !target.is_finite() {
                return Err(FormulaEvalError::Value);
            }
            target_date = Some(target);
            self.skip_whitespace();
            if !self.consume_char(',') {
                return Err(FormulaEvalError::Unsupported);
            }
        }

        let values = self.parse_aggregate_argument()?;
        self.skip_whitespace();
        if !self.consume_char(',') {
            return Err(FormulaEvalError::Unsupported);
        }
        let timeline = self.parse_aggregate_argument()?;

        let mut confidence_level = 0.95_f64;
        let mut statistic_type = None;
        let mut seasonality = 1.0_f64;
        let mut data_completion = 1.0_f64;
        let mut aggregation = 0.0_f64;

        if is_confint {
            self.skip_whitespace();
            if !self.consume_char(')') {
                if !self.consume_char(',') {
                    return Err(FormulaEvalError::Unsupported);
                }
                confidence_level = parse_optional_number!(0.95_f64);
                self.skip_whitespace();
                if !self.consume_char(')') {
                    if !self.consume_char(',') {
                        return Err(FormulaEvalError::Unsupported);
                    }
                    seasonality = parse_optional_number!(1.0_f64);
                    self.skip_whitespace();
                    if !self.consume_char(')') {
                        if !self.consume_char(',') {
                            return Err(FormulaEvalError::Unsupported);
                        }
                        data_completion = parse_optional_number!(1.0_f64);
                        self.skip_whitespace();
                        if !self.consume_char(')') {
                            if !self.consume_char(',') {
                                return Err(FormulaEvalError::Unsupported);
                            }
                            aggregation = parse_optional_number!(0.0_f64);
                            self.skip_whitespace();
                            if !self.consume_char(')') {
                                return Err(FormulaEvalError::Unsupported);
                            }
                        }
                    }
                }
            }
        } else if is_stat {
            self.skip_whitespace();
            if !self.consume_char(',') {
                return Err(FormulaEvalError::Unsupported);
            }
            let statistic = self.parse_comparison()?;
            if !statistic.is_finite() {
                return Err(FormulaEvalError::Value);
            }
            statistic_type = Some(formula_integer_argument(statistic)?);
            self.skip_whitespace();
            if !self.consume_char(')') {
                if !self.consume_char(',') {
                    return Err(FormulaEvalError::Unsupported);
                }
                seasonality = parse_optional_number!(1.0_f64);
                self.skip_whitespace();
                if !self.consume_char(')') {
                    if !self.consume_char(',') {
                        return Err(FormulaEvalError::Unsupported);
                    }
                    data_completion = parse_optional_number!(1.0_f64);
                    self.skip_whitespace();
                    if !self.consume_char(')') {
                        if !self.consume_char(',') {
                            return Err(FormulaEvalError::Unsupported);
                        }
                        aggregation = parse_optional_number!(0.0_f64);
                        self.skip_whitespace();
                        if !self.consume_char(')') {
                            return Err(FormulaEvalError::Unsupported);
                        }
                    }
                }
            }
        } else if is_seasonality {
            self.skip_whitespace();
            if !self.consume_char(')') {
                if !self.consume_char(',') {
                    return Err(FormulaEvalError::Unsupported);
                }
                data_completion = parse_optional_number!(1.0_f64);
                self.skip_whitespace();
                if !self.consume_char(')') {
                    if !self.consume_char(',') {
                        return Err(FormulaEvalError::Unsupported);
                    }
                    aggregation = parse_optional_number!(0.0_f64);
                    self.skip_whitespace();
                    if !self.consume_char(')') {
                        return Err(FormulaEvalError::Unsupported);
                    }
                }
            }
        } else {
            self.skip_whitespace();
            if !self.consume_char(')') {
                if !self.consume_char(',') {
                    return Err(FormulaEvalError::Unsupported);
                }
                seasonality = parse_optional_number!(1.0_f64);
                self.skip_whitespace();
                if !self.consume_char(')') {
                    if !self.consume_char(',') {
                        return Err(FormulaEvalError::Unsupported);
                    }
                    data_completion = parse_optional_number!(1.0_f64);
                    self.skip_whitespace();
                    if !self.consume_char(')') {
                        if !self.consume_char(',') {
                            return Err(FormulaEvalError::Unsupported);
                        }
                        aggregation = parse_optional_number!(0.0_f64);
                        self.skip_whitespace();
                        if !self.consume_char(')') {
                            return Err(FormulaEvalError::Unsupported);
                        }
                    }
                }
            }
        }

        if values.len() != timeline.len() || values.is_empty() {
            return Err(FormulaEvalError::NA);
        }
        if values
            .iter()
            .chain(timeline.iter())
            .any(|value| !value.is_finite())
        {
            return Err(FormulaEvalError::Value);
        }
        if is_confint
            && (!confidence_level.is_finite() || confidence_level <= 0.0 || confidence_level >= 1.0)
        {
            return Err(FormulaEvalError::Num);
        }
        let statistic_type = if let Some(statistic_type) = statistic_type {
            if !(1..=8).contains(&statistic_type) {
                return Err(FormulaEvalError::Num);
            }
            Some(statistic_type)
        } else {
            None
        };

        let seasonality_setting = formula_integer_argument(seasonality)?;
        if !(0..=8760).contains(&seasonality_setting) {
            return Err(FormulaEvalError::Num);
        }
        let data_completion = formula_integer_argument(data_completion)?;
        if data_completion != 0 && data_completion != 1 {
            return Err(FormulaEvalError::Num);
        }
        let aggregation = formula_integer_argument(aggregation)?;
        if !(0..=7).contains(&aggregation) {
            return Err(FormulaEvalError::Num);
        }

        let mut pairs = timeline
            .iter()
            .copied()
            .zip(values.iter().copied())
            .collect::<Vec<_>>();
        pairs.sort_by(|left, right| left.0.partial_cmp(&right.0).unwrap_or(Ordering::Equal));

        let mut aggregated_times = Vec::new();
        let mut aggregated_values = Vec::new();
        let mut index = 0usize;
        while index < pairs.len() {
            let time = pairs[index].0;
            let mut group = Vec::new();
            while index < pairs.len() && pairs[index].0 == time {
                group.push(pairs[index].1);
                index += 1;
            }
            let value = match aggregation {
                0 | 1 => group.iter().sum::<f64>() / group.len() as f64,
                2 | 3 => group.len() as f64,
                4 => group.iter().copied().fold(f64::NEG_INFINITY, f64::max),
                5 => {
                    group.sort_by(|left, right| left.partial_cmp(right).unwrap_or(Ordering::Equal));
                    let mid = group.len() / 2;
                    if group.len() % 2 == 0 {
                        (group[mid - 1] + group[mid]) / 2.0
                    } else {
                        group[mid]
                    }
                }
                6 => group.iter().copied().fold(f64::INFINITY, f64::min),
                7 => group.iter().sum::<f64>(),
                _ => unreachable!("aggregation was validated"),
            };
            aggregated_times.push(time);
            aggregated_values.push(value);
        }

        if aggregated_times.len() < 2 {
            return Err(FormulaEvalError::Num);
        }
        let mut step = f64::INFINITY;
        for window in aggregated_times.windows(2) {
            let diff = window[1] - window[0];
            if diff <= 0.0 {
                return Err(FormulaEvalError::Value);
            }
            step = step.min(diff);
        }
        if !step.is_finite() || step == 0.0 {
            return Err(FormulaEvalError::Num);
        }

        let first_time = aggregated_times[0];
        let last_time = *aggregated_times
            .last()
            .expect("timeline has at least two points");
        let total_slots_float = (last_time - first_time) / step;
        let total_slots_rounded = total_slots_float.round();
        if (total_slots_float - total_slots_rounded).abs() > 1e-7 {
            return Err(FormulaEvalError::Num);
        }
        let total_slots = total_slots_rounded as usize + 1;
        if total_slots < aggregated_values.len() {
            return Err(FormulaEvalError::Num);
        }
        let missing_slots = total_slots - aggregated_values.len();
        if missing_slots > 0 && (missing_slots as f64 / total_slots as f64) > 0.30 {
            return Err(FormulaEvalError::Num);
        }

        let mut completed = vec![None; total_slots];
        for (time, value) in aggregated_times.iter().zip(aggregated_values.iter()) {
            let slot_float = (*time - first_time) / step;
            let slot_rounded = slot_float.round();
            if (slot_float - slot_rounded).abs() > 1e-7 || slot_rounded < 0.0 {
                return Err(FormulaEvalError::Num);
            }
            let slot = slot_rounded as usize;
            if slot >= completed.len() {
                return Err(FormulaEvalError::Num);
            }
            completed[slot] = Some(*value);
        }
        for slot in 0..completed.len() {
            if completed[slot].is_some() {
                continue;
            }
            completed[slot] = Some(if data_completion == 0 {
                0.0
            } else {
                let previous = (0..slot).rev().find_map(|index| completed[index]);
                let next = (slot + 1..completed.len()).find_map(|index| completed[index]);
                match (previous, next) {
                    (Some(left), Some(right)) => (left + right) / 2.0,
                    (Some(left), None) => left,
                    (None, Some(right)) => right,
                    (None, None) => 0.0,
                }
            });
        }
        let completed_values = completed
            .into_iter()
            .map(|value| value.expect("missing slots were completed"))
            .collect::<Vec<_>>();

        let fit_model = |period: usize| -> (Vec<(f64, f64)>, Vec<f64>, f64, f64) {
            let period = period.max(1);
            let mut coefficients = vec![(0.0_f64, 0.0_f64); period];
            if period == 1 {
                let x_values = formula_regression_default_x(completed_values.len());
                let (slope, intercept) = formula_regression_slope_intercept(
                    completed_values.as_slice(),
                    x_values.as_slice(),
                    true,
                    false,
                )
                .unwrap_or((
                    0.0,
                    completed_values.iter().sum::<f64>() / completed_values.len() as f64,
                ));
                coefficients[0] = (slope, intercept);
            } else {
                for slot in 0..period {
                    let mut x_values = Vec::new();
                    let mut y_values = Vec::new();
                    for (index, value) in completed_values.iter().enumerate() {
                        if index % period == slot {
                            x_values.push(index as f64 + 1.0);
                            y_values.push(*value);
                        }
                    }
                    coefficients[slot] = if y_values.len() >= 2 {
                        formula_regression_slope_intercept(
                            y_values.as_slice(),
                            x_values.as_slice(),
                            true,
                            false,
                        )
                        .unwrap_or((0.0, y_values.iter().sum::<f64>() / y_values.len() as f64))
                    } else {
                        (
                            0.0,
                            y_values.first().copied().unwrap_or_else(|| {
                                completed_values.iter().sum::<f64>() / completed_values.len() as f64
                            }),
                        )
                    };
                }
            }

            let fitted = completed_values
                .iter()
                .enumerate()
                .map(|(index, _)| {
                    let (slope, intercept) = coefficients[index % period];
                    intercept + slope * (index as f64 + 1.0)
                })
                .collect::<Vec<_>>();
            let mae = completed_values
                .iter()
                .zip(fitted.iter())
                .map(|(actual, fitted)| (actual - fitted).abs())
                .sum::<f64>()
                / completed_values.len() as f64;
            let rmse = (completed_values
                .iter()
                .zip(fitted.iter())
                .map(|(actual, fitted)| {
                    let error = actual - fitted;
                    error * error
                })
                .sum::<f64>()
                / completed_values.len() as f64)
                .sqrt();
            (coefficients, fitted, mae, rmse)
        };

        let detect_period = || -> usize {
            if completed_values.len() < 4 {
                return 1;
            }
            let max_period = (completed_values.len() / 2).min(24).min(8760);
            let mut best_period = 1usize;
            let (_, _, _, mut best_rmse) = fit_model(1);
            for period in 2..=max_period {
                let (_, _, _, rmse) = fit_model(period);
                if rmse + 1e-9 < best_rmse {
                    best_rmse = rmse;
                    best_period = period;
                }
            }
            best_period
        };

        let period = if seasonality_setting == 0 {
            1usize
        } else if seasonality_setting == 1 {
            detect_period()
        } else {
            let period = usize::try_from(seasonality_setting).map_err(|_| FormulaEvalError::Num)?;
            if period > completed_values.len() || period > completed_values.len() / 2 {
                return Err(FormulaEvalError::Num);
            }
            period
        };

        let (coefficients, fitted, mae, rmse) = fit_model(period);
        let target_index = if let Some(target_date) = target_date {
            if target_date < last_time {
                return Err(FormulaEvalError::Num);
            }
            let offset = (target_date - first_time) / step;
            let rounded = offset.round();
            if (offset - rounded).abs() > 1e-7 || rounded < 0.0 {
                return Err(FormulaEvalError::Num);
            }
            rounded as usize
        } else {
            completed_values.len()
        };
        let (target_slope, target_intercept) = coefficients[target_index % period];
        let forecast = target_intercept + target_slope * (target_index as f64 + 1.0);

        if is_forecast {
            return formula_checked_numeric_result(forecast);
        }
        if is_seasonality {
            return Ok(period as f64);
        }
        if is_confint {
            let inverse_standard_normal = |probability: f64| -> Result<f64, FormulaEvalError> {
                if !probability.is_finite() {
                    return Err(FormulaEvalError::Value);
                }
                if probability <= 0.0 || probability >= 1.0 {
                    return Err(FormulaEvalError::Num);
                }
                const A: [f64; 6] = [
                    -3.969683028665376e1,
                    2.209460984245205e2,
                    -2.759285104469687e2,
                    1.383577518672690e2,
                    -3.066479806614716e1,
                    2.506628277459239,
                ];
                const B: [f64; 5] = [
                    -5.447609879822406e1,
                    1.615858368580409e2,
                    -1.556989798598866e2,
                    6.680131188771972e1,
                    -1.328068155288572e1,
                ];
                const C: [f64; 6] = [
                    -7.784894002430293e-3,
                    -3.223964580411365e-1,
                    -2.400758277161838,
                    -2.549732539343734,
                    4.374664141464968,
                    2.938163982698783,
                ];
                const D: [f64; 4] = [
                    7.784695709041462e-3,
                    3.224671290700398e-1,
                    2.445134137142996,
                    3.754408661907416,
                ];
                const P_LOW: f64 = 0.02425;
                const P_HIGH: f64 = 1.0 - P_LOW;
                if probability < P_LOW {
                    let q = (-2.0 * probability.ln()).sqrt();
                    let numerator =
                        (((((C[0] * q + C[1]) * q + C[2]) * q + C[3]) * q + C[4]) * q) + C[5];
                    let denominator = ((((D[0] * q + D[1]) * q + D[2]) * q + D[3]) * q) + 1.0;
                    Ok(numerator / denominator)
                } else if probability <= P_HIGH {
                    let q = probability - 0.5;
                    let r = q * q;
                    let numerator =
                        (((((A[0] * r + A[1]) * r + A[2]) * r + A[3]) * r + A[4]) * r) + A[5];
                    let denominator =
                        (((((B[0] * r + B[1]) * r + B[2]) * r + B[3]) * r + B[4]) * r) + 1.0;
                    Ok(numerator * q / denominator)
                } else {
                    let q = (-2.0 * (1.0 - probability).ln()).sqrt();
                    let numerator =
                        (((((C[0] * q + C[1]) * q + C[2]) * q + C[3]) * q + C[4]) * q) + C[5];
                    let denominator = ((((D[0] * q + D[1]) * q + D[2]) * q + D[3]) * q) + 1.0;
                    Ok(-numerator / denominator)
                }
            };
            let horizon = target_index.saturating_sub(completed_values.len() - 1) as f64;
            return formula_checked_numeric_result(
                inverse_standard_normal(0.5 + confidence_level / 2.0)?
                    * rmse
                    * (1.0 + horizon / completed_values.len() as f64).sqrt(),
            );
        }

        let statistic_type = statistic_type.expect("statistic type was parsed");
        match statistic_type {
            1 => Ok(0.5),
            2 => {
                let trend_weight = coefficients
                    .iter()
                    .any(|(slope, _)| slope.abs() > f64::EPSILON);
                Ok(if trend_weight { 0.1 } else { 0.0 })
            }
            3 => Ok(if period > 1 { 0.1 } else { 0.0 }),
            4 => {
                let lag = period.max(1);
                if completed_values.len() <= lag {
                    return Err(FormulaEvalError::Div0);
                }
                let scale = (lag..completed_values.len())
                    .map(|index| (completed_values[index] - completed_values[index - lag]).abs())
                    .sum::<f64>()
                    / (completed_values.len() - lag) as f64;
                if scale == 0.0 {
                    return if mae == 0.0 {
                        Ok(0.0)
                    } else {
                        Err(FormulaEvalError::Div0)
                    };
                }
                formula_checked_numeric_result(mae / scale)
            }
            5 => {
                let mut total = 0.0_f64;
                let mut count = 0usize;
                for (actual, forecast) in completed_values.iter().zip(fitted.iter()) {
                    let denominator = actual.abs() + forecast.abs();
                    if denominator != 0.0 {
                        total += 200.0 * (actual - forecast).abs() / denominator;
                        count += 1;
                    }
                }
                Ok(if count == 0 {
                    0.0
                } else {
                    total / count as f64
                })
            }
            6 => Ok(mae),
            7 => Ok(rmse),
            8 => Ok(step),
            _ => unreachable!("statistic type was validated"),
        }
    }

    pub(super) fn parse_regression_coefficient_function(
        &mut self,
        exponential: bool,
    ) -> Result<f64, FormulaEvalError> {
        let known_y = self.parse_aggregate_argument()?;
        let mut known_x = formula_regression_default_x(known_y.len());
        let mut constant = true;
        self.skip_whitespace();
        if !self.consume_char(')') {
            if !self.consume_char(',') {
                return Err(FormulaEvalError::Unsupported);
            }
            self.skip_whitespace();
            if !self.peek_char().is_some_and(|ch| matches!(ch, ',' | ')')) {
                known_x = self.parse_aggregate_argument()?;
            }
            self.skip_whitespace();
            if !self.consume_char(')') {
                if !self.consume_char(',') {
                    return Err(FormulaEvalError::Unsupported);
                }
                self.skip_whitespace();
                if !self.peek_char().is_some_and(|ch| matches!(ch, ',' | ')')) {
                    constant = self.parse_comparison()? != 0.0;
                }
                self.skip_whitespace();
                if !self.consume_char(')') {
                    if !self.consume_char(',') {
                        return Err(FormulaEvalError::Unsupported);
                    }
                    self.skip_whitespace();
                    if !self.consume_char(')') {
                        self.parse_comparison()?;
                        self.skip_whitespace();
                        if !self.consume_char(')') {
                            return Err(FormulaEvalError::Unsupported);
                        }
                    }
                }
            }
        }

        let (slope, _) = formula_regression_slope_intercept(
            known_y.as_slice(),
            known_x.as_slice(),
            constant,
            exponential,
        )?;
        if exponential {
            formula_checked_numeric_result(slope.exp())
        } else {
            formula_checked_numeric_result(slope)
        }
    }

    pub(super) fn parse_regression_prediction_function(
        &mut self,
        exponential: bool,
    ) -> Result<f64, FormulaEvalError> {
        let known_y = self.parse_aggregate_argument()?;
        let mut known_x = formula_regression_default_x(known_y.len());
        let mut new_x = None;
        let mut constant = true;
        self.skip_whitespace();
        if !self.consume_char(')') {
            if !self.consume_char(',') {
                return Err(FormulaEvalError::Unsupported);
            }
            self.skip_whitespace();
            if !self.peek_char().is_some_and(|ch| matches!(ch, ',' | ')')) {
                known_x = self.parse_aggregate_argument()?;
            }
            self.skip_whitespace();
            if !self.consume_char(')') {
                if !self.consume_char(',') {
                    return Err(FormulaEvalError::Unsupported);
                }
                self.skip_whitespace();
                if !self.peek_char().is_some_and(|ch| matches!(ch, ',' | ')')) {
                    new_x = self.parse_aggregate_argument()?.first().copied();
                }
                self.skip_whitespace();
                if !self.consume_char(')') {
                    if !self.consume_char(',') {
                        return Err(FormulaEvalError::Unsupported);
                    }
                    self.skip_whitespace();
                    if !self.consume_char(')') {
                        constant = self.parse_comparison()? != 0.0;
                        self.skip_whitespace();
                        if !self.consume_char(')') {
                            return Err(FormulaEvalError::Unsupported);
                        }
                    }
                }
            }
        }

        let new_x = new_x
            .or_else(|| known_x.first().copied())
            .ok_or(FormulaEvalError::NA)?;
        if !new_x.is_finite() {
            return Err(FormulaEvalError::Value);
        }
        let (slope, intercept) = formula_regression_slope_intercept(
            known_y.as_slice(),
            known_x.as_slice(),
            constant,
            exponential,
        )?;
        let prediction = intercept + slope * new_x;
        if exponential {
            formula_checked_numeric_result(prediction.exp())
        } else {
            formula_checked_numeric_result(prediction)
        }
    }

    pub(super) fn parse_countif_function(&mut self) -> Result<f64, FormulaEvalError> {
        let (sheet_id, rect) = self.parse_reference_argument()?;
        self.skip_whitespace();
        if !self.consume_char(',') {
            return Err(FormulaEvalError::Unsupported);
        }
        let criteria = self.parse_criteria_argument()?;
        self.skip_whitespace();
        if !self.consume_char(')') {
            return Err(FormulaEvalError::Unsupported);
        }
        Ok(self
            .evaluator
            .countif_values_in_rect(sheet_id, rect, &criteria)? as f64)
    }

    pub(super) fn parse_getpivotdata_value_function(
        &mut self,
    ) -> Result<FormulaValueProbe, FormulaEvalError> {
        let data_field = self.parse_text_value_argument()?;
        if data_field.trim().is_empty() {
            return Err(FormulaEvalError::Value);
        }
        self.skip_whitespace();
        if !self.consume_char(',') {
            return Err(FormulaEvalError::Unsupported);
        }
        let (pivot_sheet_id, pivot_rect) = self.parse_reference_argument()?;
        loop {
            self.skip_whitespace();
            if self.consume_char(')') {
                let value = self.evaluator.cell_value_or_blank(
                    pivot_sheet_id,
                    pivot_rect.row_first,
                    pivot_rect.col_first,
                )?;
                return Ok(formula_value_probe_from_cell_value(value));
            }
            if !self.consume_char(',') {
                return Err(FormulaEvalError::Unsupported);
            }
            self.parse_value_probe_argument()?;
            self.skip_whitespace();
            if !self.consume_char(',') {
                return Err(FormulaEvalError::Value);
            }
            self.parse_value_probe_argument()?;
        }
    }

    pub(super) fn parse_cube_caption_text_function(
        &mut self,
        name: &str,
    ) -> Result<String, FormulaEvalError> {
        self.parse_text_value_argument()?;
        self.skip_whitespace();
        if !self.consume_char(',') {
            return Err(FormulaEvalError::Unsupported);
        }
        let fallback = self.parse_text_value_argument()?;
        if name.eq_ignore_ascii_case("CUBEKPIMEMBER") {
            self.skip_whitespace();
            if !self.consume_char(',') {
                return Err(FormulaEvalError::Unsupported);
            }
            self.parse_value_probe_argument()?;
        } else if name.eq_ignore_ascii_case("CUBERANKEDMEMBER") {
            self.skip_whitespace();
            if !self.consume_char(',') {
                return Err(FormulaEvalError::Unsupported);
            }
            self.parse_comparison()?;
        }

        self.skip_whitespace();
        if self.consume_char(')') {
            return Ok(fallback);
        }
        if !self.consume_char(',') {
            return Err(FormulaEvalError::Unsupported);
        }
        self.skip_whitespace();
        let caption = if self.peek_char().is_some_and(|ch| matches!(ch, ',' | ')')) {
            fallback
        } else {
            self.parse_text_value_argument()?
        };
        self.consume_remaining_optional_value_arguments()?;
        Ok(caption)
    }

    pub(super) fn parse_cubememberproperty_text_function(
        &mut self,
    ) -> Result<String, FormulaEvalError> {
        self.parse_text_value_argument()?;
        self.skip_whitespace();
        if !self.consume_char(',') {
            return Err(FormulaEvalError::Unsupported);
        }
        self.parse_text_value_argument()?;
        self.skip_whitespace();
        if !self.consume_char(',') {
            return Err(FormulaEvalError::Unsupported);
        }
        self.parse_text_value_argument()?;
        self.skip_whitespace();
        if !self.consume_char(')') {
            return Err(FormulaEvalError::Unsupported);
        }
        Err(FormulaEvalError::NA)
    }

    pub(super) fn parse_cubesetcount_function(&mut self) -> Result<f64, FormulaEvalError> {
        let set = self.parse_text_value_argument()?;
        self.skip_whitespace();
        if !self.consume_char(')') {
            return Err(FormulaEvalError::Unsupported);
        }
        Ok(if set.trim().is_empty() { 0.0 } else { 1.0 })
    }

    pub(super) fn parse_groupby_aggregation_argument(
        &mut self,
    ) -> Result<FormulaGroupByAggregation, FormulaEvalError> {
        self.skip_whitespace();
        let checkpoint = self.index;
        if let Some(name) = self.parse_identifier() {
            self.skip_whitespace();
            if self.peek_char().is_none_or(|ch| matches!(ch, ',' | ')')) {
                return FormulaGroupByAggregation::from_name(name.as_str())
                    .ok_or(FormulaEvalError::Value);
            }
        }
        self.index = checkpoint;
        match self.parse_value_probe_argument()? {
            FormulaValueProbe::Text(name) => {
                FormulaGroupByAggregation::from_name(name.as_str()).ok_or(FormulaEvalError::Value)
            }
            FormulaValueProbe::Error(error) => Err(error),
            _ => Err(FormulaEvalError::Value),
        }
    }

    pub(super) fn rect_row_key_values(
        &mut self,
        sheet_id: SheetId,
        rect: Rect,
        row_offset: u32,
    ) -> Result<Vec<FormulaValueProbe>, FormulaEvalError> {
        let mut key = Vec::with_capacity(rect.width() as usize);
        let row = rect.row_first + row_offset;
        for col in rect.col_first..=rect.col_last {
            let value = self.evaluator.cell_value_or_blank(sheet_id, row, col)?;
            key.push(formula_value_probe_from_cell_value(value));
        }
        Ok(key)
    }

    pub(super) fn rect_row_key_matches(
        &mut self,
        sheet_id: SheetId,
        rect: Rect,
        row_offset: u32,
        target: &[FormulaValueProbe],
    ) -> Result<bool, FormulaEvalError> {
        let candidate = self.rect_row_key_values(sheet_id, rect, row_offset)?;
        if candidate.len() != target.len() {
            return Ok(false);
        }
        for (left, right) in candidate.iter().zip(target.iter()) {
            if !formula_value_probe_exact_match(left, right)? {
                return Ok(false);
            }
        }
        Ok(true)
    }

    pub(super) fn collect_first_group_values(
        &mut self,
        row_sheet_id: SheetId,
        row_rect: Rect,
        value_sheet_id: SheetId,
        value_rect: Rect,
        column_filter: Option<(SheetId, Rect, Vec<FormulaValueProbe>)>,
    ) -> Result<Vec<FormulaValueProbe>, FormulaEvalError> {
        if row_rect.height() != value_rect.height() {
            return Err(FormulaEvalError::Value);
        }
        if let Some((_, column_rect, _)) = &column_filter {
            if column_rect.height() != value_rect.height() {
                return Err(FormulaEvalError::Value);
            }
        }
        let first_row_key = self.rect_row_key_values(row_sheet_id, row_rect, 0)?;
        let mut values = Vec::new();
        for row_offset in 0..value_rect.height() {
            if !self.rect_row_key_matches(row_sheet_id, row_rect, row_offset, &first_row_key)? {
                continue;
            }
            if let Some((column_sheet_id, column_rect, column_key)) = &column_filter
                && !self.rect_row_key_matches(
                    *column_sheet_id,
                    *column_rect,
                    row_offset,
                    column_key,
                )?
            {
                continue;
            }
            let value = self.evaluator.cell_value_or_blank(
                value_sheet_id,
                value_rect.row_first + row_offset,
                value_rect.col_first,
            )?;
            values.push(formula_value_probe_from_cell_value(value));
        }
        Ok(values)
    }

    pub(super) fn parse_groupby_value_function(
        &mut self,
        row_sheet_id: SheetId,
        row_rect: Rect,
    ) -> Result<FormulaValueProbe, FormulaEvalError> {
        self.skip_whitespace();
        if !self.consume_char(',') {
            return Err(FormulaEvalError::Unsupported);
        }
        let (value_sheet_id, value_rect) = self.parse_reference_argument()?;
        self.skip_whitespace();
        if !self.consume_char(',') {
            return Err(FormulaEvalError::Unsupported);
        }
        let aggregation = self.parse_groupby_aggregation_argument()?;
        self.consume_remaining_optional_value_arguments()?;
        let values = self.collect_first_group_values(
            row_sheet_id,
            row_rect,
            value_sheet_id,
            value_rect,
            None,
        )?;
        aggregation.evaluate(values.as_slice())
    }

    pub(super) fn parse_pivotby_value_function(
        &mut self,
        row_sheet_id: SheetId,
        row_rect: Rect,
    ) -> Result<FormulaValueProbe, FormulaEvalError> {
        self.skip_whitespace();
        if !self.consume_char(',') {
            return Err(FormulaEvalError::Unsupported);
        }
        let (column_sheet_id, column_rect) = self.parse_reference_argument()?;
        self.skip_whitespace();
        if !self.consume_char(',') {
            return Err(FormulaEvalError::Unsupported);
        }
        let (value_sheet_id, value_rect) = self.parse_reference_argument()?;
        self.skip_whitespace();
        if !self.consume_char(',') {
            return Err(FormulaEvalError::Unsupported);
        }
        let aggregation = self.parse_groupby_aggregation_argument()?;
        self.consume_remaining_optional_value_arguments()?;
        let first_column_key = self.rect_row_key_values(column_sheet_id, column_rect, 0)?;
        let values = self.collect_first_group_values(
            row_sheet_id,
            row_rect,
            value_sheet_id,
            value_rect,
            Some((column_sheet_id, column_rect, first_column_key)),
        )?;
        aggregation.evaluate(values.as_slice())
    }

    pub(super) fn parse_f_test_function(&mut self) -> Result<f64, FormulaEvalError> {
        let first_values = self.parse_aggregate_argument()?;
        self.skip_whitespace();
        if !self.consume_char(',') {
            return Err(FormulaEvalError::Unsupported);
        }
        let second_values = self.parse_aggregate_argument()?;
        self.skip_whitespace();
        if !self.consume_char(')') {
            return Err(FormulaEvalError::Unsupported);
        }

        let (_, first_variance) = formula_sample_mean_and_variance(first_values.as_slice())?;
        let (_, second_variance) = formula_sample_mean_and_variance(second_values.as_slice())?;
        if first_variance <= 0.0 || second_variance <= 0.0 {
            return Err(FormulaEvalError::Div0);
        }
        let (ratio, degrees1, degrees2) = if first_variance >= second_variance {
            (
                first_variance / second_variance,
                first_values.len() as f64 - 1.0,
                second_values.len() as f64 - 1.0,
            )
        } else {
            (
                second_variance / first_variance,
                second_values.len() as f64 - 1.0,
                first_values.len() as f64 - 1.0,
            )
        };
        formula_f_right_tail(ratio, degrees1, degrees2)
            .and_then(|tail| formula_checked_numeric_result((2.0 * tail).min(1.0)))
    }

    pub(super) fn parse_t_test_function(&mut self) -> Result<f64, FormulaEvalError> {
        let first_values = self.parse_aggregate_argument()?;
        self.skip_whitespace();
        if !self.consume_char(',') {
            return Err(FormulaEvalError::Unsupported);
        }
        let second_values = self.parse_aggregate_argument()?;
        self.skip_whitespace();
        if !self.consume_char(',') {
            return Err(FormulaEvalError::Unsupported);
        }
        let tails = formula_integer_argument(self.parse_comparison()?)?;
        self.skip_whitespace();
        if !self.consume_char(',') {
            return Err(FormulaEvalError::Unsupported);
        }
        let test_type = formula_integer_argument(self.parse_comparison()?)?;
        self.skip_whitespace();
        if !self.consume_char(')') {
            return Err(FormulaEvalError::Unsupported);
        }
        if !matches!(tails, 1 | 2) || !matches!(test_type, 1 | 2 | 3) {
            return Err(FormulaEvalError::Num);
        }

        let (t_statistic, degrees) = match test_type {
            1 => {
                if first_values.len() != second_values.len() {
                    return Err(FormulaEvalError::NA);
                }
                let differences = first_values
                    .iter()
                    .zip(second_values.iter())
                    .map(|(first, second)| first - second)
                    .collect::<Vec<_>>();
                let (mean, variance) = formula_sample_mean_and_variance(differences.as_slice())?;
                if variance <= 0.0 {
                    return Err(FormulaEvalError::Div0);
                }
                (
                    mean.abs() / (variance / differences.len() as f64).sqrt(),
                    differences.len() as f64 - 1.0,
                )
            }
            2 => {
                let (first_mean, first_variance) =
                    formula_sample_mean_and_variance(first_values.as_slice())?;
                let (second_mean, second_variance) =
                    formula_sample_mean_and_variance(second_values.as_slice())?;
                let degrees = first_values.len() as f64 + second_values.len() as f64 - 2.0;
                let pooled_variance = ((first_values.len() - 1) as f64 * first_variance
                    + (second_values.len() - 1) as f64 * second_variance)
                    / degrees;
                if pooled_variance <= 0.0 {
                    return Err(FormulaEvalError::Div0);
                }
                (
                    (first_mean - second_mean).abs()
                        / (pooled_variance
                            * (1.0 / first_values.len() as f64 + 1.0 / second_values.len() as f64))
                            .sqrt(),
                    degrees,
                )
            }
            3 => {
                let (first_mean, first_variance) =
                    formula_sample_mean_and_variance(first_values.as_slice())?;
                let (second_mean, second_variance) =
                    formula_sample_mean_and_variance(second_values.as_slice())?;
                let first_component = first_variance / first_values.len() as f64;
                let second_component = second_variance / second_values.len() as f64;
                let denominator = (first_component + second_component).sqrt();
                let degrees_denominator = first_component * first_component
                    / (first_values.len() as f64 - 1.0)
                    + second_component * second_component / (second_values.len() as f64 - 1.0);
                if denominator == 0.0 || degrees_denominator == 0.0 {
                    return Err(FormulaEvalError::Div0);
                }
                (
                    (first_mean - second_mean).abs() / denominator,
                    (first_component + second_component) * (first_component + second_component)
                        / degrees_denominator,
                )
            }
            _ => return Err(FormulaEvalError::Num),
        };
        if !t_statistic.is_finite() || !degrees.is_finite() {
            return Err(FormulaEvalError::Num);
        }
        let tail = formula_student_t_right_tail_from_abs(t_statistic, degrees)?;
        formula_checked_numeric_result(if tails == 1 {
            tail
        } else {
            (2.0 * tail).min(1.0)
        })
    }

    pub(super) fn parse_z_test_function(&mut self) -> Result<f64, FormulaEvalError> {
        let values = self.parse_aggregate_argument()?;
        self.skip_whitespace();
        if !self.consume_char(',') {
            return Err(FormulaEvalError::Unsupported);
        }
        let x = self.parse_comparison()?;
        if !x.is_finite() {
            return Err(FormulaEvalError::Value);
        }
        self.skip_whitespace();
        let sigma = if self.consume_char(')') {
            None
        } else {
            if !self.consume_char(',') {
                return Err(FormulaEvalError::Unsupported);
            }
            let sigma = self.parse_comparison()?;
            if !sigma.is_finite() {
                return Err(FormulaEvalError::Value);
            }
            if sigma <= 0.0 {
                return Err(FormulaEvalError::Num);
            }
            self.skip_whitespace();
            if !self.consume_char(')') {
                return Err(FormulaEvalError::Unsupported);
            }
            Some(sigma)
        };
        if values.is_empty() {
            return Err(FormulaEvalError::NA);
        }
        if values.iter().any(|value| !value.is_finite()) {
            return Err(FormulaEvalError::Value);
        }
        let mean = values.iter().sum::<f64>() / values.len() as f64;
        let sigma = if let Some(sigma) = sigma {
            sigma
        } else {
            let (_, variance) = formula_sample_mean_and_variance(values.as_slice())?;
            if variance <= 0.0 {
                return Err(FormulaEvalError::Div0);
            }
            variance.sqrt()
        };
        let denominator = sigma / (values.len() as f64).sqrt();
        if denominator == 0.0 {
            return Err(FormulaEvalError::Div0);
        }
        let z = (mean - x) / denominator;
        formula_standard_normal_cdf(z)
            .and_then(|cdf| formula_checked_numeric_result((1.0 - cdf).clamp(0.0, 1.0)))
    }

    pub(super) fn parse_chisq_test_function(&mut self) -> Result<f64, FormulaEvalError> {
        let (observed_sheet_id, observed_rect) = self.parse_reference_argument()?;
        self.skip_whitespace();
        if !self.consume_char(',') {
            return Err(FormulaEvalError::Unsupported);
        }
        let (expected_sheet_id, expected_rect) = self.parse_reference_argument()?;
        self.skip_whitespace();
        if !self.consume_char(')') {
            return Err(FormulaEvalError::Unsupported);
        }
        if observed_rect.width() != expected_rect.width()
            || observed_rect.height() != expected_rect.height()
        {
            return Err(FormulaEvalError::NA);
        }
        let degrees = if observed_rect.height() == 1 {
            u64::from(observed_rect.width().saturating_sub(1))
        } else if observed_rect.width() == 1 {
            u64::from(observed_rect.height().saturating_sub(1))
        } else {
            u64::from(observed_rect.height() - 1)
                .checked_mul(u64::from(observed_rect.width() - 1))
                .ok_or(FormulaEvalError::Num)?
        };
        if degrees == 0 {
            return Err(FormulaEvalError::Div0);
        }

        let mut statistic = 0.0_f64;
        for row_offset in 0..observed_rect.height() {
            for col_offset in 0..observed_rect.width() {
                let observed = self.evaluator.numeric_cell_value(
                    observed_sheet_id,
                    observed_rect.row_first + row_offset,
                    observed_rect.col_first + col_offset,
                )?;
                let expected = self.evaluator.numeric_cell_value(
                    expected_sheet_id,
                    expected_rect.row_first + row_offset,
                    expected_rect.col_first + col_offset,
                )?;
                if !observed.is_finite() || !expected.is_finite() {
                    return Err(FormulaEvalError::Value);
                }
                if observed < 0.0 || expected <= 0.0 {
                    return Err(FormulaEvalError::Num);
                }
                statistic += (observed - expected).powi(2) / expected;
            }
        }
        if !statistic.is_finite() {
            return Err(FormulaEvalError::Num);
        }

        let checked_numeric_result = |value: f64| -> Result<f64, FormulaEvalError> {
            if value.is_finite() {
                Ok(value)
            } else {
                Err(FormulaEvalError::Num)
            }
        };
        let gamma_ln_value = |value: f64| {
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
                std::f64::consts::PI.ln()
                    - (std::f64::consts::PI * value).sin().ln()
                    - lanczos(1.0 - value)
            } else {
                lanczos(value)
            }
        };
        let regularized_gamma_p = |shape: f64, x: f64| -> Result<f64, FormulaEvalError> {
            if !shape.is_finite() || !x.is_finite() {
                return Err(FormulaEvalError::Value);
            }
            if shape <= 0.0 || x < 0.0 {
                return Err(FormulaEvalError::Num);
            }
            if x == 0.0 {
                return Ok(0.0);
            }
            const EPSILON: f64 = 1e-14;
            const FLOOR: f64 = 1e-300;
            const MAX_ITERATIONS: usize = 200;
            let gamma_ln = gamma_ln_value(shape);
            if x < shape + 1.0 {
                let mut term = 1.0 / shape;
                let mut sum = term;
                let mut ap = shape;
                for _ in 0..MAX_ITERATIONS {
                    ap += 1.0;
                    term *= x / ap;
                    sum += term;
                    if term.abs() <= sum.abs() * EPSILON {
                        return checked_numeric_result(
                            (sum * (-x + shape * x.ln() - gamma_ln).exp()).clamp(0.0, 1.0),
                        );
                    }
                }
                return checked_numeric_result(
                    (sum * (-x + shape * x.ln() - gamma_ln).exp()).clamp(0.0, 1.0),
                );
            }

            let mut b = x + 1.0 - shape;
            let mut c = 1.0 / FLOOR;
            let mut d = 1.0 / b.max(FLOOR);
            let mut h = d;
            for i in 1..=MAX_ITERATIONS {
                let i = i as f64;
                let an = -i * (i - shape);
                b += 2.0;
                d = an * d + b;
                if d.abs() < FLOOR {
                    d = FLOOR;
                }
                c = b + an / c;
                if c.abs() < FLOOR {
                    c = FLOOR;
                }
                d = 1.0 / d;
                let delta = d * c;
                h *= delta;
                if (delta - 1.0).abs() <= EPSILON {
                    let q = (-x + shape * x.ln() - gamma_ln).exp() * h;
                    return checked_numeric_result((1.0 - q).clamp(0.0, 1.0));
                }
            }
            let q = (-x + shape * x.ln() - gamma_ln).exp() * h;
            checked_numeric_result((1.0 - q).clamp(0.0, 1.0))
        };
        regularized_gamma_p(degrees as f64 / 2.0, statistic / 2.0)
            .map(|value| (1.0 - value).clamp(0.0, 1.0))
    }

    pub(super) fn parse_database_value_function(
        &mut self,
        name: &str,
    ) -> Result<FormulaValueProbe, FormulaEvalError> {
        let (database_sheet_id, database_rect) = self.parse_reference_argument()?;
        if database_rect.height() < 2 || database_rect.width() < 1 {
            return Err(FormulaEvalError::Value);
        }
        self.skip_whitespace();
        if !self.consume_char(',') {
            return Err(FormulaEvalError::Unsupported);
        }
        let field = self.parse_value_probe_argument()?;
        self.skip_whitespace();
        if !self.consume_char(',') {
            return Err(FormulaEvalError::Unsupported);
        }
        let (criteria_sheet_id, criteria_rect) = self.parse_reference_argument()?;
        if criteria_rect.height() < 1 || criteria_rect.width() < 1 {
            return Err(FormulaEvalError::Value);
        }
        self.skip_whitespace();
        if !self.consume_char(')') {
            return Err(FormulaEvalError::Unsupported);
        }

        macro_rules! database_header_text {
            ($sheet_id:expr, $row:expr, $col:expr) => {{
                let value = self.evaluator.cell_value_or_blank($sheet_id, $row, $col)?;
                formula_text_from_value_probe(formula_value_probe_from_cell_value(value))?
            }};
        }
        macro_rules! database_field_offset {
            ($field:expr) => {{
                match $field {
                    FormulaValueProbe::Number(value) => {
                        let index = formula_integer_argument(value)?;
                        if index < 1 || index > i64::from(database_rect.width()) {
                            return Err(FormulaEvalError::Value);
                        }
                        index as u32 - 1
                    }
                    FormulaValueProbe::Text(label) => {
                        let mut found = None;
                        for col_offset in 0..database_rect.width() {
                            let header = database_header_text!(
                                database_sheet_id,
                                database_rect.row_first,
                                database_rect.col_first + col_offset
                            );
                            if header.eq_ignore_ascii_case(label.as_str()) {
                                found = Some(col_offset);
                                break;
                            }
                        }
                        found.ok_or(FormulaEvalError::Value)?
                    }
                    FormulaValueProbe::Error(error) => return Err(error),
                    FormulaValueProbe::Blank
                    | FormulaValueProbe::Bool(_)
                    | FormulaValueProbe::Omitted
                    | FormulaValueProbe::Lambda { .. } => {
                        return Err(FormulaEvalError::Value);
                    }
                }
            }};
        }
        let field_offset = database_field_offset!(field);

        let mut criteria_rows = Vec::<Vec<(u32, FormulaCriteria)>>::new();
        if criteria_rect.height() == 1 {
            criteria_rows.push(Vec::new());
        } else {
            for criteria_row in criteria_rect.row_first + 1..=criteria_rect.row_last {
                let mut terms = Vec::new();
                for criteria_col in criteria_rect.col_first..=criteria_rect.col_last {
                    let criteria_value = self.evaluator.cell_value_or_blank(
                        criteria_sheet_id,
                        criteria_row,
                        criteria_col,
                    )?;
                    let criteria = match criteria_value {
                        CellValue::Blank => continue,
                        CellValue::Number(value) => FormulaCriteria::from_numeric_value(value),
                        CellValue::Bool(value) => {
                            FormulaCriteria::from_numeric_value(if value { 1.0 } else { 0.0 })
                        }
                        CellValue::Text(value) => FormulaCriteria::from_string_literal(value),
                        CellValue::IsoDateTime(value) => {
                            FormulaCriteria::from_string_literal(value.into_string())
                        }
                        CellValue::RichText(value) => {
                            FormulaCriteria::from_string_literal(value.into_string())
                        }
                        CellValue::Error(error) => {
                            return Err(formula_eval_error_from_cell_error(error));
                        }
                    };
                    let label = database_header_text!(
                        criteria_sheet_id,
                        criteria_rect.row_first,
                        criteria_col
                    );
                    let col_offset = database_field_offset!(FormulaValueProbe::Text(label));
                    terms.push((col_offset, criteria));
                }
                criteria_rows.push(terms);
            }
        }

        let mut numeric_values = Vec::new();
        let mut counta = 0_u64;
        let mut dget_value = None::<FormulaValueProbe>;
        let mut dget_count = 0_u64;
        for row in database_rect.row_first + 1..=database_rect.row_last {
            let mut matches_any_criteria_row = false;
            for criteria_row in criteria_rows.iter() {
                let mut matches_all_terms = true;
                for (criteria_col_offset, criteria) in criteria_row {
                    let value = self.evaluator.cell_value_or_blank(
                        database_sheet_id,
                        row,
                        database_rect.col_first + *criteria_col_offset,
                    )?;
                    if let CellValue::Error(error) = value {
                        return Err(formula_eval_error_from_cell_error(error));
                    }
                    if !criteria.matches(&value) {
                        matches_all_terms = false;
                        break;
                    }
                }
                if matches_all_terms {
                    matches_any_criteria_row = true;
                    break;
                }
            }
            if !matches_any_criteria_row {
                continue;
            }
            let field_value = self.evaluator.cell_value_or_blank(
                database_sheet_id,
                row,
                database_rect.col_first + field_offset,
            )?;
            match name.to_ascii_uppercase().as_str() {
                "DGET" => {
                    dget_count += 1;
                    dget_value = Some(formula_value_probe_from_cell_value(field_value));
                }
                "DCOUNTA" => match field_value {
                    CellValue::Blank => {}
                    CellValue::Error(error) => {
                        return Err(formula_eval_error_from_cell_error(error));
                    }
                    CellValue::Bool(_)
                    | CellValue::Number(_)
                    | CellValue::Text(_)
                    | CellValue::IsoDateTime(_)
                    | CellValue::RichText(_) => counta += 1,
                },
                "DCOUNT" => match field_value {
                    CellValue::Number(_) => counta += 1,
                    CellValue::Error(error) => {
                        return Err(formula_eval_error_from_cell_error(error));
                    }
                    CellValue::Blank
                    | CellValue::Bool(_)
                    | CellValue::Text(_)
                    | CellValue::IsoDateTime(_)
                    | CellValue::RichText(_) => {}
                },
                _ => match field_value {
                    CellValue::Number(value) => numeric_values.push(value),
                    CellValue::Error(error) => {
                        return Err(formula_eval_error_from_cell_error(error));
                    }
                    CellValue::Blank
                    | CellValue::Bool(_)
                    | CellValue::Text(_)
                    | CellValue::IsoDateTime(_)
                    | CellValue::RichText(_) => {}
                },
            }
        }

        if name.eq_ignore_ascii_case("DGET") {
            if dget_count == 0 {
                return Err(FormulaEvalError::Value);
            }
            if dget_count > 1 {
                return Err(FormulaEvalError::Num);
            }
            return dget_value.ok_or(FormulaEvalError::Value);
        }
        if name.eq_ignore_ascii_case("DCOUNT") || name.eq_ignore_ascii_case("DCOUNTA") {
            return Ok(FormulaValueProbe::Number(counta as f64));
        }

        let function = if name.eq_ignore_ascii_case("DAVERAGE") {
            FormulaAggregateFunction::Average
        } else if name.eq_ignore_ascii_case("DMAX") {
            FormulaAggregateFunction::Max
        } else if name.eq_ignore_ascii_case("DMIN") {
            FormulaAggregateFunction::Min
        } else if name.eq_ignore_ascii_case("DPRODUCT") {
            FormulaAggregateFunction::Product
        } else if name.eq_ignore_ascii_case("DSTDEV") {
            FormulaAggregateFunction::StDevS
        } else if name.eq_ignore_ascii_case("DSTDEVP") {
            FormulaAggregateFunction::StDevP
        } else if name.eq_ignore_ascii_case("DSUM") {
            FormulaAggregateFunction::Sum
        } else if name.eq_ignore_ascii_case("DVAR") {
            FormulaAggregateFunction::VarS
        } else if name.eq_ignore_ascii_case("DVARP") {
            FormulaAggregateFunction::VarP
        } else {
            return Err(FormulaEvalError::Unsupported);
        };
        Ok(FormulaValueProbe::Number(
            function.evaluate(numeric_values.as_slice())?,
        ))
    }

    pub(super) fn parse_sumif_function(&mut self) -> Result<f64, FormulaEvalError> {
        let (criteria_sheet_id, criteria_rect) = self.parse_reference_argument()?;
        self.skip_whitespace();
        if !self.consume_char(',') {
            return Err(FormulaEvalError::Unsupported);
        }
        let criteria = self.parse_criteria_argument()?;
        self.skip_whitespace();
        if self.consume_char(')') {
            return self.evaluator.sumif_values_in_rect(
                criteria_sheet_id,
                criteria_rect,
                &criteria,
                criteria_sheet_id,
                criteria_rect,
            );
        }
        if !self.consume_char(',') {
            return Err(FormulaEvalError::Unsupported);
        }
        let (sum_sheet_id, sum_rect) = self.parse_reference_argument()?;
        self.skip_whitespace();
        if !self.consume_char(')') {
            return Err(FormulaEvalError::Unsupported);
        }
        self.evaluator.sumif_values_in_rect(
            criteria_sheet_id,
            criteria_rect,
            &criteria,
            sum_sheet_id,
            sum_rect,
        )
    }

    pub(super) fn parse_averageif_function(&mut self) -> Result<f64, FormulaEvalError> {
        let (criteria_sheet_id, criteria_rect) = self.parse_reference_argument()?;
        self.skip_whitespace();
        if !self.consume_char(',') {
            return Err(FormulaEvalError::Unsupported);
        }
        let criteria = self.parse_criteria_argument()?;
        self.skip_whitespace();
        if self.consume_char(')') {
            return self.evaluator.averageif_values_in_rect(
                criteria_sheet_id,
                criteria_rect,
                &criteria,
                criteria_sheet_id,
                criteria_rect,
            );
        }
        if !self.consume_char(',') {
            return Err(FormulaEvalError::Unsupported);
        }
        let (average_sheet_id, average_rect) = self.parse_reference_argument()?;
        self.skip_whitespace();
        if !self.consume_char(')') {
            return Err(FormulaEvalError::Unsupported);
        }
        self.evaluator.averageif_values_in_rect(
            criteria_sheet_id,
            criteria_rect,
            &criteria,
            average_sheet_id,
            average_rect,
        )
    }

    pub(super) fn parse_countifs_function(&mut self) -> Result<f64, FormulaEvalError> {
        let criteria_ranges = self.parse_criteria_ranges_arguments()?;
        Ok(self
            .evaluator
            .countifs_values_in_rects(criteria_ranges.as_slice())? as f64)
    }

    pub(super) fn parse_sumifs_function(&mut self) -> Result<f64, FormulaEvalError> {
        let (sum_sheet_id, sum_rect) = self.parse_reference_argument()?;
        self.skip_whitespace();
        if !self.consume_char(',') {
            return Err(FormulaEvalError::Unsupported);
        }
        let criteria_ranges = self.parse_criteria_ranges_arguments()?;
        self.evaluator
            .sumifs_values_in_rect(sum_sheet_id, sum_rect, criteria_ranges.as_slice())
    }

    pub(super) fn parse_averageifs_function(&mut self) -> Result<f64, FormulaEvalError> {
        let (average_sheet_id, average_rect) = self.parse_reference_argument()?;
        self.skip_whitespace();
        if !self.consume_char(',') {
            return Err(FormulaEvalError::Unsupported);
        }
        let criteria_ranges = self.parse_criteria_ranges_arguments()?;
        self.evaluator.averageifs_values_in_rect(
            average_sheet_id,
            average_rect,
            criteria_ranges.as_slice(),
        )
    }

    pub(super) fn parse_minifs_function(&mut self) -> Result<f64, FormulaEvalError> {
        let (min_sheet_id, min_rect) = self.parse_reference_argument()?;
        self.skip_whitespace();
        if !self.consume_char(',') {
            return Err(FormulaEvalError::Unsupported);
        }
        let criteria_ranges = self.parse_criteria_ranges_arguments()?;
        self.evaluator
            .minifs_values_in_rect(min_sheet_id, min_rect, criteria_ranges.as_slice())
    }

    pub(super) fn parse_maxifs_function(&mut self) -> Result<f64, FormulaEvalError> {
        let (max_sheet_id, max_rect) = self.parse_reference_argument()?;
        self.skip_whitespace();
        if !self.consume_char(',') {
            return Err(FormulaEvalError::Unsupported);
        }
        let criteria_ranges = self.parse_criteria_ranges_arguments()?;
        self.evaluator
            .maxifs_values_in_rect(max_sheet_id, max_rect, criteria_ranges.as_slice())
    }

    pub(super) fn parse_countblank_function(&mut self) -> Result<f64, FormulaEvalError> {
        let mut count = 0_u64;
        loop {
            self.skip_whitespace();
            if self.consume_char(')') {
                return Ok(count as f64);
            }
            count += self.parse_countblank_argument()?;
            self.skip_whitespace();
            if self.consume_char(',') {
                continue;
            }
            if self.consume_char(')') {
                return Ok(count as f64);
            }
            return Err(FormulaEvalError::Unsupported);
        }
    }

    pub(super) fn parse_series_sum_function(&mut self) -> Result<f64, FormulaEvalError> {
        let x = self.parse_comparison()?;
        self.skip_whitespace();
        if !self.consume_char(',') {
            return Err(FormulaEvalError::Unsupported);
        }
        let n = self.parse_comparison()?;
        self.skip_whitespace();
        if !self.consume_char(',') {
            return Err(FormulaEvalError::Unsupported);
        }
        let m = self.parse_comparison()?;
        self.skip_whitespace();
        if !self.consume_char(',') {
            return Err(FormulaEvalError::Unsupported);
        }
        self.skip_whitespace();
        let coefficients = self.parse_aggregate_argument()?;
        self.skip_whitespace();
        if !self.consume_char(')') {
            return Err(FormulaEvalError::Unsupported);
        }
        if coefficients.is_empty() {
            return Err(FormulaEvalError::Value);
        }
        if ![x, n, m].iter().all(|value| value.is_finite())
            || coefficients.iter().any(|value| !value.is_finite())
        {
            return Err(FormulaEvalError::Value);
        }
        let mut total = 0.0;
        for (index, coefficient) in coefficients.iter().enumerate() {
            let exponent = n + index as f64 * m;
            if !exponent.is_finite() {
                return Err(FormulaEvalError::Num);
            }
            let term = coefficient * x.powf(exponent);
            if !term.is_finite() {
                return Err(FormulaEvalError::Num);
            }
            total += term;
            if !total.is_finite() {
                return Err(FormulaEvalError::Num);
            }
        }
        Ok(total)
    }

    pub(super) fn parse_aggregate_function(&mut self) -> Result<f64, FormulaEvalError> {
        let function_num = formula_integer_argument(self.parse_comparison()?)?;
        if !(1..=19).contains(&function_num) {
            return Err(FormulaEvalError::Value);
        }
        self.skip_whitespace();
        if !self.consume_char(',') {
            return Err(FormulaEvalError::Unsupported);
        }
        let options = formula_integer_argument(self.parse_comparison()?)?;
        if !(0..=7).contains(&options) {
            return Err(FormulaEvalError::Value);
        }
        self.skip_whitespace();
        if !self.consume_char(',') {
            return Err(FormulaEvalError::Unsupported);
        }

        let ignore_nested = matches!(options, 0..=3);
        let ignore_errors = matches!(options, 2 | 3 | 6 | 7);
        let percentile_value =
            |mut values: Vec<f64>, k: f64, exclusive: bool| -> Result<f64, FormulaEvalError> {
                if values.is_empty() || !k.is_finite() {
                    return Err(FormulaEvalError::Num);
                }
                values.sort_by(|left, right| left.total_cmp(right));
                if exclusive {
                    if k <= 0.0 || k >= 1.0 {
                        return Err(FormulaEvalError::Num);
                    }
                    let rank = k * (values.len() as f64 + 1.0);
                    if rank < 1.0 || rank > values.len() as f64 {
                        return Err(FormulaEvalError::Num);
                    }
                    let lower_rank = rank.floor();
                    let upper_rank = rank.ceil();
                    if lower_rank == upper_rank {
                        return Ok(values[lower_rank as usize - 1]);
                    }
                    let lower_index = lower_rank as usize - 1;
                    let upper_index = upper_rank as usize - 1;
                    let fraction = rank - lower_rank;
                    return Ok(values[lower_index]
                        + (values[upper_index] - values[lower_index]) * fraction);
                }
                if !(0.0..=1.0).contains(&k) {
                    return Err(FormulaEvalError::Num);
                }
                let rank = k * (values.len() as f64 - 1.0);
                let lower_index = rank.floor() as usize;
                let upper_index = rank.ceil() as usize;
                if lower_index == upper_index {
                    return Ok(values[lower_index]);
                }
                let fraction = rank - lower_index as f64;
                Ok(values[lower_index] + (values[upper_index] - values[lower_index]) * fraction)
            };
        let mut values = Vec::new();
        let mut counta = 0_u64;

        macro_rules! is_nested_aggregate_cell {
            ($sheet_id:expr, $row:expr, $col:expr) => {{
                self.evaluator
                    .state
                    .worksheet_data()
                    .get(&$sheet_id)
                    .and_then(|worksheet| worksheet.cells.get(&($row, $col)))
                    .and_then(|cell| cell.formula.as_ref())
                    .is_some_and(|formula| {
                        formula_source_has_top_level_function(formula, "SUBTOTAL")
                            || formula_source_has_top_level_function(formula, "AGGREGATE")
                    })
            }};
        }

        macro_rules! record_numeric_value {
            ($value:expr) => {{
                match $value {
                    FormulaValueProbe::Number(number) => values.push(number),
                    FormulaValueProbe::Error(error) if ignore_errors => {
                        let _ = error;
                    }
                    FormulaValueProbe::Error(error) => return Err(error),
                    FormulaValueProbe::Blank
                    | FormulaValueProbe::Bool(_)
                    | FormulaValueProbe::Text(_)
                    | FormulaValueProbe::Omitted
                    | FormulaValueProbe::Lambda { .. } => {}
                }
            }};
        }

        macro_rules! collect_numeric_argument {
            () => {{
                if let Some(reference) = self.parse_reference_set_before_boundary(&[',', ')'])? {
                    for (target_sheet_id, rect) in reference.areas() {
                        for row in rect.row_first..=rect.row_last {
                            for col in rect.col_first..=rect.col_last {
                                if ignore_nested
                                    && is_nested_aggregate_cell!(*target_sheet_id, row, col)
                                {
                                    continue;
                                }
                                let value = self.evaluator.cell_value_or_blank(
                                    *target_sheet_id,
                                    row,
                                    col,
                                )?;
                                record_numeric_value!(formula_value_probe_from_cell_value(value));
                            }
                        }
                    }
                } else {
                    match self.parse_catchable_argument()? {
                        Ok(value) => values.push(value),
                        Err(error) if ignore_errors => {
                            let _ = error;
                        }
                        Err(error) => return Err(error),
                    }
                }
            }};
        }

        macro_rules! collect_counta_argument {
            () => {{
                if let Some(reference) = self.parse_reference_set_before_boundary(&[',', ')'])? {
                    for (target_sheet_id, rect) in reference.areas() {
                        for row in rect.row_first..=rect.row_last {
                            for col in rect.col_first..=rect.col_last {
                                if ignore_nested
                                    && is_nested_aggregate_cell!(*target_sheet_id, row, col)
                                {
                                    continue;
                                }
                                let value = self.evaluator.cell_value_or_blank(
                                    *target_sheet_id,
                                    row,
                                    col,
                                )?;
                                match value {
                                    CellValue::Error(error) if ignore_errors => {
                                        let _ = error;
                                    }
                                    CellValue::Error(error) => {
                                        return Err(formula_eval_error_from_cell_error(error));
                                    }
                                    CellValue::Blank => {}
                                    CellValue::Bool(_)
                                    | CellValue::Number(_)
                                    | CellValue::Text(_)
                                    | CellValue::IsoDateTime(_)
                                    | CellValue::RichText(_) => counta += 1,
                                }
                            }
                        }
                    }
                } else {
                    match self.parse_value_probe_argument()? {
                        FormulaValueProbe::Error(error) if ignore_errors => {
                            let _ = error;
                        }
                        FormulaValueProbe::Error(error) => return Err(error),
                        FormulaValueProbe::Blank => {}
                        FormulaValueProbe::Bool(_)
                        | FormulaValueProbe::Number(_)
                        | FormulaValueProbe::Text(_) => counta += 1,
                        FormulaValueProbe::Omitted | FormulaValueProbe::Lambda { .. } => {}
                    }
                }
            }};
        }

        let finish_numeric = |values: &[f64]| -> Result<f64, FormulaEvalError> {
            let aggregate_function = match function_num {
                1 => Some(FormulaAggregateFunction::Average),
                2 => Some(FormulaAggregateFunction::Count),
                4 => Some(FormulaAggregateFunction::Max),
                5 => Some(FormulaAggregateFunction::Min),
                6 => Some(FormulaAggregateFunction::Product),
                7 => Some(FormulaAggregateFunction::StDevS),
                8 => Some(FormulaAggregateFunction::StDevP),
                9 => Some(FormulaAggregateFunction::Sum),
                10 => Some(FormulaAggregateFunction::VarS),
                11 => Some(FormulaAggregateFunction::VarP),
                12 => Some(FormulaAggregateFunction::Median),
                13 => Some(FormulaAggregateFunction::ModeSngl),
                _ => None,
            };
            aggregate_function
                .ok_or(FormulaEvalError::Value)?
                .evaluate(values)
        };

        if (14..=19).contains(&function_num) {
            collect_numeric_argument!();
            self.skip_whitespace();
            if !self.consume_char(',') {
                return Err(FormulaEvalError::Unsupported);
            }
            let k = self.parse_comparison()?;
            self.skip_whitespace();
            if !self.consume_char(')') {
                return Err(FormulaEvalError::Unsupported);
            }
            if function_num == 14 || function_num == 15 {
                let k = formula_integer_argument(k)?;
                if values.is_empty() || k < 1 || k > values.len() as i64 {
                    return Err(FormulaEvalError::Num);
                }
                values.sort_by(|left, right| left.total_cmp(right));
                let index = if function_num == 14 {
                    values.len() - k as usize
                } else {
                    k as usize - 1
                };
                return Ok(values[index]);
            }
            if function_num == 17 || function_num == 19 {
                if !k.is_finite() {
                    return Err(FormulaEvalError::Value);
                }
                if k < i64::MIN as f64 || k > i64::MAX as f64 {
                    return Err(FormulaEvalError::Num);
                }
                let quart = k.trunc() as i64;
                let exclusive = function_num == 19;
                if exclusive {
                    if !(1..=3).contains(&quart) {
                        return Err(FormulaEvalError::Num);
                    }
                } else if !(0..=4).contains(&quart) {
                    return Err(FormulaEvalError::Num);
                }
                return percentile_value(values, quart as f64 / 4.0, exclusive);
            }
            return percentile_value(values, k, function_num == 18);
        }

        let mut saw_argument = false;
        loop {
            self.skip_whitespace();
            if self.consume_char(')') {
                if !saw_argument {
                    return Err(FormulaEvalError::Value);
                }
                return if function_num == 3 {
                    Ok(counta as f64)
                } else {
                    finish_numeric(values.as_slice())
                };
            }
            saw_argument = true;
            if function_num == 3 {
                collect_counta_argument!();
            } else {
                collect_numeric_argument!();
            }
            self.skip_whitespace();
            if self.consume_char(',') {
                continue;
            }
            if self.consume_char(')') {
                return if function_num == 3 {
                    Ok(counta as f64)
                } else {
                    finish_numeric(values.as_slice())
                };
            }
            return Err(FormulaEvalError::Unsupported);
        }
    }

    pub(super) fn parse_subtotal_function(&mut self) -> Result<f64, FormulaEvalError> {
        let function_num = formula_integer_argument(self.parse_comparison()?)?;
        let function_num = match function_num {
            1..=11 => function_num,
            101..=111 => function_num - 100,
            _ => return Err(FormulaEvalError::Value),
        };
        self.skip_whitespace();
        if !self.consume_char(',') {
            return Err(FormulaEvalError::Unsupported);
        }

        let aggregate_function = match function_num {
            1 => Some(FormulaAggregateFunction::Average),
            2 => Some(FormulaAggregateFunction::Count),
            3 => None,
            4 => Some(FormulaAggregateFunction::Max),
            5 => Some(FormulaAggregateFunction::Min),
            6 => Some(FormulaAggregateFunction::Product),
            7 => Some(FormulaAggregateFunction::StDevS),
            8 => Some(FormulaAggregateFunction::StDevP),
            9 => Some(FormulaAggregateFunction::Sum),
            10 => Some(FormulaAggregateFunction::VarS),
            11 => Some(FormulaAggregateFunction::VarP),
            _ => unreachable!("validated SUBTOTAL function number"),
        };
        let mut saw_argument = false;
        let mut values = Vec::new();
        let mut counta = 0_u64;

        loop {
            self.skip_whitespace();
            if self.consume_char(')') {
                if !saw_argument {
                    return Err(FormulaEvalError::Value);
                }
                return match aggregate_function {
                    Some(function) => function.evaluate(values.as_slice()),
                    None => Ok(counta as f64),
                };
            }
            saw_argument = true;
            if aggregate_function.is_some() {
                values.extend(self.parse_subtotal_numeric_argument()?);
            } else {
                counta += self.parse_subtotal_counta_argument()?;
            }
            self.skip_whitespace();
            if self.consume_char(',') {
                continue;
            }
            if self.consume_char(')') {
                return match aggregate_function {
                    Some(function) => function.evaluate(values.as_slice()),
                    None => Ok(counta as f64),
                };
            }
            return Err(FormulaEvalError::Unsupported);
        }
    }

    pub(super) fn parse_counta_function(&mut self) -> Result<f64, FormulaEvalError> {
        let mut count = 0_u64;
        loop {
            self.skip_whitespace();
            if self.consume_char(')') {
                return Ok(count as f64);
            }
            count += self.parse_counta_argument()?;
            self.skip_whitespace();
            if self.consume_char(',') {
                continue;
            }
            if self.consume_char(')') {
                return Ok(count as f64);
            }
            return Err(FormulaEvalError::Unsupported);
        }
    }

    pub(super) fn parse_aggregate_a_function(
        &mut self,
        name: &str,
    ) -> Result<f64, FormulaEvalError> {
        let mut values = Vec::new();
        loop {
            self.skip_whitespace();
            if self.consume_char(')') {
                break;
            }
            values.extend(self.parse_aggregate_a_argument()?);
            self.skip_whitespace();
            if self.consume_char(',') {
                continue;
            }
            if self.consume_char(')') {
                break;
            }
            return Err(FormulaEvalError::Unsupported);
        }
        if name.eq_ignore_ascii_case("MINA") {
            return Ok(values.iter().copied().reduce(f64::min).unwrap_or(0.0));
        }
        if name.eq_ignore_ascii_case("MAXA") {
            return Ok(values.iter().copied().reduce(f64::max).unwrap_or(0.0));
        }
        if values.is_empty() {
            return Err(FormulaEvalError::Div0);
        }
        let mean = values.iter().sum::<f64>() / values.len() as f64;
        if name.eq_ignore_ascii_case("AVERAGEA") {
            return Ok(mean);
        }
        let deviation_sum = values
            .iter()
            .map(|value| {
                let deviation = value - mean;
                deviation * deviation
            })
            .sum::<f64>();
        if name.eq_ignore_ascii_case("VARPA") {
            return Ok(deviation_sum / values.len() as f64);
        }
        if name.eq_ignore_ascii_case("STDEVPA") {
            return Ok((deviation_sum / values.len() as f64).sqrt());
        }
        if values.len() < 2 {
            return Err(FormulaEvalError::Div0);
        }
        if name.eq_ignore_ascii_case("VARA") {
            return Ok(deviation_sum / (values.len() - 1) as f64);
        }
        Ok((deviation_sum / (values.len() - 1) as f64).sqrt())
    }

    pub(super) fn parse_aggregate_argument(&mut self) -> Result<Vec<f64>, FormulaEvalError> {
        if let Some(reference) = self.parse_reference_set_before_boundary(&[',', ')'])? {
            return self.evaluator.numeric_values_in_reference(&reference);
        }
        if let Some(values) = self.try_parse_array_constant_argument()? {
            // Like references, array constants contribute only their numbers.
            let mut numbers = Vec::new();
            for value in values {
                match value {
                    FormulaValueProbe::Number(number) => numbers.push(number),
                    FormulaValueProbe::Error(error) => return Err(error),
                    _ => {}
                }
            }
            return Ok(numbers);
        }
        Ok(vec![self.parse_comparison()?])
    }

    pub(super) fn parse_subtotal_numeric_argument(&mut self) -> Result<Vec<f64>, FormulaEvalError> {
        if let Some(reference) = self.parse_reference_set_before_boundary(&[',', ')'])? {
            let mut values = Vec::new();
            for (target_sheet_id, rect) in reference.areas() {
                values.extend(
                    self.evaluator
                        .subtotal_numeric_values_in_rect(*target_sheet_id, *rect)?,
                );
            }
            return Ok(values);
        }
        Ok(vec![self.parse_comparison()?])
    }

    pub(super) fn parse_subtotal_counta_argument(&mut self) -> Result<u64, FormulaEvalError> {
        if let Some(reference) = self.parse_reference_set_before_boundary(&[',', ')'])? {
            let mut count = 0_u64;
            for (target_sheet_id, rect) in reference.areas() {
                count += self
                    .evaluator
                    .subtotal_counta_values_in_rect(*target_sheet_id, *rect)?;
            }
            return Ok(count);
        }
        self.parse_comparison()?;
        Ok(1)
    }

    pub(super) fn parse_aggregate_a_argument(&mut self) -> Result<Vec<f64>, FormulaEvalError> {
        if let Some(reference) = self.parse_reference_set_before_boundary(&[',', ')'])? {
            let mut values = Vec::new();
            for (target_sheet_id, rect) in reference.areas() {
                for row in rect.row_first..=rect.row_last {
                    for col in rect.col_first..=rect.col_last {
                        match self
                            .evaluator
                            .cell_value_or_blank(*target_sheet_id, row, col)?
                        {
                            CellValue::Blank => {}
                            CellValue::Bool(value) => values.push(if value { 1.0 } else { 0.0 }),
                            CellValue::Number(value) => values.push(value),
                            CellValue::Text(_)
                            | CellValue::IsoDateTime(_)
                            | CellValue::RichText(_) => values.push(0.0),
                            CellValue::Error(error) => {
                                return Err(formula_eval_error_from_cell_error(error));
                            }
                        }
                    }
                }
            }
            return Ok(values);
        }
        let value = self.parse_value_probe_argument()?;
        match value {
            FormulaValueProbe::Blank | FormulaValueProbe::Text(_) => Ok(vec![0.0]),
            FormulaValueProbe::Bool(value) => Ok(vec![if value { 1.0 } else { 0.0 }]),
            FormulaValueProbe::Number(value) => Ok(vec![value]),
            FormulaValueProbe::Error(error) => Err(error),
            FormulaValueProbe::Omitted | FormulaValueProbe::Lambda { .. } => {
                Err(FormulaEvalError::Value)
            }
        }
    }

    pub(super) fn parse_counta_argument(&mut self) -> Result<u64, FormulaEvalError> {
        if let Some(reference) = self.parse_reference_set_before_boundary(&[',', ')'])? {
            return self.evaluator.counta_values_in_reference(&reference);
        }
        self.parse_comparison()?;
        Ok(1)
    }

    pub(super) fn parse_countblank_argument(&mut self) -> Result<u64, FormulaEvalError> {
        if let Some(reference) = self.parse_reference_set_before_boundary(&[',', ')'])? {
            return self.evaluator.countblank_values_in_reference(&reference);
        }
        self.parse_comparison()?;
        Ok(0)
    }

    pub(super) fn parse_criteria_argument(&mut self) -> Result<FormulaCriteria, FormulaEvalError> {
        self.skip_whitespace();
        if let Some(literal) = self.parse_string_literal()? {
            return Ok(FormulaCriteria::from_string_literal(literal));
        }
        Ok(FormulaCriteria::from_numeric_value(
            self.parse_comparison()?,
        ))
    }

    pub(super) fn parse_criteria_ranges_arguments(
        &mut self,
    ) -> Result<Vec<FormulaCriteriaRange>, FormulaEvalError> {
        let mut criteria_ranges = Vec::new();
        loop {
            let (sheet_id, rect) = self.parse_reference_argument()?;
            self.skip_whitespace();
            if !self.consume_char(',') {
                return Err(FormulaEvalError::Unsupported);
            }
            let criteria = self.parse_criteria_argument()?;
            criteria_ranges.push(FormulaCriteriaRange {
                sheet_id,
                rect,
                criteria,
            });
            self.skip_whitespace();
            if self.consume_char(')') {
                return Ok(criteria_ranges);
            }
            if !self.consume_char(',') {
                return Err(FormulaEvalError::Unsupported);
            }
        }
    }

    pub(super) fn parse_frequency_function(&mut self) -> Result<f64, FormulaEvalError> {
        let data = self.parse_aggregate_argument()?;
        self.skip_whitespace();
        if !self.consume_char(',') {
            return Err(FormulaEvalError::Unsupported);
        }
        let bins = self.parse_aggregate_argument()?;
        self.skip_whitespace();
        if !self.consume_char(')') {
            return Err(FormulaEvalError::Unsupported);
        }
        let Some(first_bin) = bins.first() else {
            return Ok(data.len() as f64);
        };
        Ok(data.iter().filter(|value| **value <= *first_bin).count() as f64)
    }
}
