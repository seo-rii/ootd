//! Formula parser methods for annuity and currency functions: FV, PV, PMT, IPMT, PPMT, NPER, RATE, DOLLARDE, EUROCONVERT.

use super::*;

impl<'a, 'b, 'state> FormulaParser<'a, 'b, 'state> {
    pub(super) fn parse_dollarde_function(&mut self) -> Result<f64, FormulaEvalError> {
        let fractional_dollar = self.parse_comparison()?;
        self.skip_whitespace();
        if !self.consume_char(',') {
            return Err(FormulaEvalError::Unsupported);
        }
        let (denominator, scale) = formula_dollar_fraction_parts(self.parse_comparison()?)?;
        self.skip_whitespace();
        if !self.consume_char(')') {
            return Err(FormulaEvalError::Unsupported);
        }
        if !fractional_dollar.is_finite() {
            return Err(FormulaEvalError::Value);
        }
        let sign = if fractional_dollar.is_sign_negative() {
            -1.0
        } else {
            1.0
        };
        let absolute = fractional_dollar.abs();
        let whole = absolute.trunc();
        let numerator = formula_dollar_fraction_near_integer((absolute - whole) * scale);
        let value = sign * (whole + numerator / denominator);
        if value.is_finite() {
            Ok(value)
        } else {
            Err(FormulaEvalError::Num)
        }
    }

    pub(super) fn parse_dollarfr_function(&mut self) -> Result<f64, FormulaEvalError> {
        let decimal_dollar = self.parse_comparison()?;
        self.skip_whitespace();
        if !self.consume_char(',') {
            return Err(FormulaEvalError::Unsupported);
        }
        let (denominator, scale) = formula_dollar_fraction_parts(self.parse_comparison()?)?;
        self.skip_whitespace();
        if !self.consume_char(')') {
            return Err(FormulaEvalError::Unsupported);
        }
        if !decimal_dollar.is_finite() {
            return Err(FormulaEvalError::Value);
        }
        let sign = if decimal_dollar.is_sign_negative() {
            -1.0
        } else {
            1.0
        };
        let absolute = decimal_dollar.abs();
        let whole = absolute.trunc();
        let numerator = formula_dollar_fraction_near_integer((absolute - whole) * denominator);
        let value = sign * (whole + numerator / scale);
        if value.is_finite() {
            Ok(value)
        } else {
            Err(FormulaEvalError::Num)
        }
    }

    pub(super) fn parse_fv_function(&mut self) -> Result<f64, FormulaEvalError> {
        let rate = self.parse_comparison()?;
        self.skip_whitespace();
        if !self.consume_char(',') {
            return Err(FormulaEvalError::Unsupported);
        }
        let nper = self.parse_comparison()?;
        self.skip_whitespace();
        if !self.consume_char(',') {
            return Err(FormulaEvalError::Unsupported);
        }
        let pmt = self.parse_comparison()?;
        let mut pv = 0.0;
        let mut payment_type = 0.0;
        self.skip_whitespace();
        if !self.consume_char(')') {
            if !self.consume_char(',') {
                return Err(FormulaEvalError::Unsupported);
            }
            pv = self.parse_comparison()?;
            self.skip_whitespace();
            if !self.consume_char(')') {
                if !self.consume_char(',') {
                    return Err(FormulaEvalError::Unsupported);
                }
                payment_type = formula_financial_type_argument(self.parse_comparison()?)?;
                self.skip_whitespace();
                if !self.consume_char(')') {
                    return Err(FormulaEvalError::Unsupported);
                }
            }
        }
        formula_fv_value(rate, nper, pmt, pv, payment_type)
    }

    pub(super) fn parse_pv_function(&mut self) -> Result<f64, FormulaEvalError> {
        let rate = self.parse_comparison()?;
        self.skip_whitespace();
        if !self.consume_char(',') {
            return Err(FormulaEvalError::Unsupported);
        }
        let nper = self.parse_comparison()?;
        self.skip_whitespace();
        if !self.consume_char(',') {
            return Err(FormulaEvalError::Unsupported);
        }
        let pmt = self.parse_comparison()?;
        let mut fv = 0.0;
        let mut payment_type = 0.0;
        self.skip_whitespace();
        if !self.consume_char(')') {
            if !self.consume_char(',') {
                return Err(FormulaEvalError::Unsupported);
            }
            fv = self.parse_comparison()?;
            self.skip_whitespace();
            if !self.consume_char(')') {
                if !self.consume_char(',') {
                    return Err(FormulaEvalError::Unsupported);
                }
                payment_type = formula_financial_type_argument(self.parse_comparison()?)?;
                self.skip_whitespace();
                if !self.consume_char(')') {
                    return Err(FormulaEvalError::Unsupported);
                }
            }
        }
        if ![rate, nper, pmt, fv].iter().all(|value| value.is_finite()) {
            return Err(FormulaEvalError::Value);
        }
        let value = if rate == 0.0 {
            -(fv + pmt * nper)
        } else {
            let growth = formula_annuity_growth(rate, nper)?;
            if growth == 0.0 {
                return Err(FormulaEvalError::Div0);
            }
            -(fv + pmt * (1.0 + rate * payment_type) * (growth - 1.0) / rate) / growth
        };
        if value.is_finite() {
            Ok(value)
        } else {
            Err(FormulaEvalError::Num)
        }
    }

    pub(super) fn parse_pmt_function(&mut self) -> Result<f64, FormulaEvalError> {
        let rate = self.parse_comparison()?;
        self.skip_whitespace();
        if !self.consume_char(',') {
            return Err(FormulaEvalError::Unsupported);
        }
        let nper = self.parse_comparison()?;
        self.skip_whitespace();
        if !self.consume_char(',') {
            return Err(FormulaEvalError::Unsupported);
        }
        let pv = self.parse_comparison()?;
        let mut fv = 0.0;
        let mut payment_type = 0.0;
        self.skip_whitespace();
        if !self.consume_char(')') {
            if !self.consume_char(',') {
                return Err(FormulaEvalError::Unsupported);
            }
            fv = self.parse_comparison()?;
            self.skip_whitespace();
            if !self.consume_char(')') {
                if !self.consume_char(',') {
                    return Err(FormulaEvalError::Unsupported);
                }
                payment_type = formula_financial_type_argument(self.parse_comparison()?)?;
                self.skip_whitespace();
                if !self.consume_char(')') {
                    return Err(FormulaEvalError::Unsupported);
                }
            }
        }
        formula_pmt_value(rate, nper, pv, fv, payment_type)
    }

    pub(super) fn parse_ipmt_or_ppmt_function(
        &mut self,
        principal: bool,
    ) -> Result<f64, FormulaEvalError> {
        let rate = self.parse_comparison()?;
        self.skip_whitespace();
        if !self.consume_char(',') {
            return Err(FormulaEvalError::Unsupported);
        }
        let period = self.parse_comparison()?;
        self.skip_whitespace();
        if !self.consume_char(',') {
            return Err(FormulaEvalError::Unsupported);
        }
        let nper = self.parse_comparison()?;
        self.skip_whitespace();
        if !self.consume_char(',') {
            return Err(FormulaEvalError::Unsupported);
        }
        let pv = self.parse_comparison()?;
        let mut fv = 0.0;
        let mut payment_type = 0.0;
        self.skip_whitespace();
        if !self.consume_char(')') {
            if !self.consume_char(',') {
                return Err(FormulaEvalError::Unsupported);
            }
            fv = self.parse_comparison()?;
            self.skip_whitespace();
            if !self.consume_char(')') {
                if !self.consume_char(',') {
                    return Err(FormulaEvalError::Unsupported);
                }
                payment_type = formula_financial_type_argument(self.parse_comparison()?)?;
                self.skip_whitespace();
                if !self.consume_char(')') {
                    return Err(FormulaEvalError::Unsupported);
                }
            }
        }
        let interest = formula_ipmt_value(rate, period, nper, pv, fv, payment_type)?;
        if principal {
            Ok(formula_pmt_value(rate, nper, pv, fv, payment_type)? - interest)
        } else {
            Ok(interest)
        }
    }

    pub(super) fn parse_ipmt_function(&mut self) -> Result<f64, FormulaEvalError> {
        self.parse_ipmt_or_ppmt_function(false)
    }

    pub(super) fn parse_ppmt_function(&mut self) -> Result<f64, FormulaEvalError> {
        self.parse_ipmt_or_ppmt_function(true)
    }

    pub(super) fn parse_cumulative_payment_function(
        &mut self,
        principal: bool,
    ) -> Result<f64, FormulaEvalError> {
        let rate = self.parse_comparison()?;
        self.skip_whitespace();
        if !self.consume_char(',') {
            return Err(FormulaEvalError::Unsupported);
        }
        let nper = self.parse_comparison()?;
        self.skip_whitespace();
        if !self.consume_char(',') {
            return Err(FormulaEvalError::Unsupported);
        }
        let pv = self.parse_comparison()?;
        self.skip_whitespace();
        if !self.consume_char(',') {
            return Err(FormulaEvalError::Unsupported);
        }
        let start_period = self.parse_comparison()?;
        self.skip_whitespace();
        if !self.consume_char(',') {
            return Err(FormulaEvalError::Unsupported);
        }
        let end_period = self.parse_comparison()?;
        self.skip_whitespace();
        if !self.consume_char(',') {
            return Err(FormulaEvalError::Unsupported);
        }
        let payment_type = formula_financial_type_argument(self.parse_comparison()?)?;
        self.skip_whitespace();
        if !self.consume_char(')') {
            return Err(FormulaEvalError::Unsupported);
        }
        if ![rate, nper, pv, start_period, end_period]
            .iter()
            .all(|value| value.is_finite())
        {
            return Err(FormulaEvalError::Value);
        }
        let start_period = start_period.trunc();
        let end_period = end_period.trunc();
        if rate <= 0.0
            || nper <= 0.0
            || pv <= 0.0
            || start_period < 1.0
            || end_period < 1.0
            || start_period > end_period
        {
            return Err(FormulaEvalError::Num);
        }
        let payment = formula_pmt_value(rate, nper, pv, 0.0, payment_type)?;
        let mut total = 0.0;
        let mut period = start_period;
        while period <= end_period {
            let interest = formula_ipmt_value(rate, period, nper, pv, 0.0, payment_type)?;
            total += if principal {
                payment - interest
            } else {
                interest
            };
            period += 1.0;
        }
        if total.is_finite() {
            Ok(total)
        } else {
            Err(FormulaEvalError::Num)
        }
    }

    pub(super) fn parse_nper_function(&mut self) -> Result<f64, FormulaEvalError> {
        let rate = self.parse_comparison()?;
        self.skip_whitespace();
        if !self.consume_char(',') {
            return Err(FormulaEvalError::Unsupported);
        }
        let pmt = self.parse_comparison()?;
        self.skip_whitespace();
        if !self.consume_char(',') {
            return Err(FormulaEvalError::Unsupported);
        }
        let pv = self.parse_comparison()?;
        let mut fv = 0.0;
        let mut payment_type = 0.0;
        self.skip_whitespace();
        if !self.consume_char(')') {
            if !self.consume_char(',') {
                return Err(FormulaEvalError::Unsupported);
            }
            fv = self.parse_comparison()?;
            self.skip_whitespace();
            if !self.consume_char(')') {
                if !self.consume_char(',') {
                    return Err(FormulaEvalError::Unsupported);
                }
                payment_type = formula_financial_type_argument(self.parse_comparison()?)?;
                self.skip_whitespace();
                if !self.consume_char(')') {
                    return Err(FormulaEvalError::Unsupported);
                }
            }
        }
        if ![rate, pmt, pv, fv].iter().all(|value| value.is_finite()) {
            return Err(FormulaEvalError::Value);
        }
        let value = if rate == 0.0 {
            if pmt == 0.0 {
                return Err(FormulaEvalError::Div0);
            }
            -(pv + fv) / pmt
        } else {
            let base = 1.0 + rate;
            if base <= 0.0 {
                return Err(FormulaEvalError::Num);
            }
            let adjusted_payment = pmt * (1.0 + rate * payment_type) / rate;
            let numerator = adjusted_payment - fv;
            let denominator = pv + adjusted_payment;
            if numerator == 0.0 || denominator == 0.0 {
                return Err(FormulaEvalError::Div0);
            }
            let ratio = numerator / denominator;
            if ratio <= 0.0 {
                return Err(FormulaEvalError::Num);
            }
            ratio.ln() / base.ln()
        };
        if value.is_finite() {
            Ok(value)
        } else {
            Err(FormulaEvalError::Num)
        }
    }

    pub(super) fn parse_rate_function(&mut self) -> Result<f64, FormulaEvalError> {
        let nper = self.parse_comparison()?;
        self.skip_whitespace();
        if !self.consume_char(',') {
            return Err(FormulaEvalError::Unsupported);
        }
        let pmt = self.parse_comparison()?;
        self.skip_whitespace();
        if !self.consume_char(',') {
            return Err(FormulaEvalError::Unsupported);
        }
        let pv = self.parse_comparison()?;
        let mut fv = 0.0;
        let mut payment_type = 0.0;
        let mut guess = 0.1;
        self.skip_whitespace();
        if !self.consume_char(')') {
            if !self.consume_char(',') {
                return Err(FormulaEvalError::Unsupported);
            }
            fv = self.parse_comparison()?;
            self.skip_whitespace();
            if !self.consume_char(')') {
                if !self.consume_char(',') {
                    return Err(FormulaEvalError::Unsupported);
                }
                payment_type = formula_financial_type_argument(self.parse_comparison()?)?;
                self.skip_whitespace();
                if !self.consume_char(')') {
                    if !self.consume_char(',') {
                        return Err(FormulaEvalError::Unsupported);
                    }
                    guess = self.parse_comparison()?;
                    self.skip_whitespace();
                    if !self.consume_char(')') {
                        return Err(FormulaEvalError::Unsupported);
                    }
                }
            }
        }
        if ![nper, pmt, pv, fv, guess]
            .iter()
            .all(|value| value.is_finite())
        {
            return Err(FormulaEvalError::Value);
        }
        if nper <= 0.0 || guess <= -1.0 {
            return Err(FormulaEvalError::Num);
        }

        let rate_residual = |rate: f64| -> Result<f64, FormulaEvalError> {
            if !rate.is_finite() || rate <= -1.0 {
                return Err(FormulaEvalError::Num);
            }
            let value = if rate.abs() < 1e-10 {
                pv + pmt * nper + fv
            } else {
                let growth = formula_annuity_growth(rate, nper)?;
                pv * growth + pmt * (1.0 + rate * payment_type) * (growth - 1.0) / rate + fv
            };
            if value.is_finite() {
                Ok(value)
            } else {
                Err(FormulaEvalError::Num)
            }
        };

        const RATE_MAX_ITERATIONS: usize = 20;
        const RATE_TOLERANCE: f64 = 1e-7;
        let mut rate = guess;
        for _ in 0..RATE_MAX_ITERATIONS {
            let value = rate_residual(rate)?;
            if value.abs() <= RATE_TOLERANCE {
                return Ok(rate);
            }
            let step = (rate.abs() * 1e-6).max(1e-8);
            let right_rate = rate + step;
            let right_value = rate_residual(right_rate)?;
            let derivative = (right_value - value) / (right_rate - rate);
            if !derivative.is_finite() || derivative == 0.0 {
                break;
            }
            let next_rate = rate - value / derivative;
            if !next_rate.is_finite() || next_rate <= -1.0 {
                break;
            }
            if (next_rate - rate).abs() <= RATE_TOLERANCE {
                return Ok(next_rate);
            }
            rate = next_rate;
        }
        Err(FormulaEvalError::Num)
    }

    pub(super) fn parse_ispmt_function(&mut self) -> Result<f64, FormulaEvalError> {
        let rate = self.parse_comparison()?;
        self.skip_whitespace();
        if !self.consume_char(',') {
            return Err(FormulaEvalError::Unsupported);
        }
        let period = self.parse_comparison()?;
        self.skip_whitespace();
        if !self.consume_char(',') {
            return Err(FormulaEvalError::Unsupported);
        }
        let nper = self.parse_comparison()?;
        self.skip_whitespace();
        if !self.consume_char(',') {
            return Err(FormulaEvalError::Unsupported);
        }
        let pv = self.parse_comparison()?;
        self.skip_whitespace();
        if !self.consume_char(')') {
            return Err(FormulaEvalError::Unsupported);
        }
        if ![rate, period, nper, pv]
            .iter()
            .all(|value| value.is_finite())
        {
            return Err(FormulaEvalError::Value);
        }
        if nper == 0.0 {
            return Err(FormulaEvalError::Div0);
        }
        let value = pv * rate * (period / nper - 1.0);
        if value.is_finite() {
            Ok(value)
        } else {
            Err(FormulaEvalError::Num)
        }
    }

    pub(super) fn parse_euroconvert_function(&mut self) -> Result<f64, FormulaEvalError> {
        let value = self.parse_comparison()?;
        self.skip_whitespace();
        if !self.consume_char(',') {
            return Err(FormulaEvalError::Unsupported);
        }
        let source = self.parse_convert_unit_argument()?;
        self.skip_whitespace();
        if !self.consume_char(',') {
            return Err(FormulaEvalError::Unsupported);
        }
        let target = self.parse_convert_unit_argument()?;
        let mut full_precision = false;
        let mut triangulation_precision = None;
        self.skip_whitespace();
        if !self.consume_char(')') {
            if !self.consume_char(',') {
                return Err(FormulaEvalError::Unsupported);
            }
            full_precision = self.parse_comparison()? != 0.0;
            self.skip_whitespace();
            if !self.consume_char(')') {
                if !self.consume_char(',') {
                    return Err(FormulaEvalError::Unsupported);
                }
                triangulation_precision = Some(formula_integer_argument(self.parse_comparison()?)?);
                self.skip_whitespace();
                if !self.consume_char(')') {
                    return Err(FormulaEvalError::Unsupported);
                }
            }
        }
        formula_euroconvert_value(
            value,
            source.as_str(),
            target.as_str(),
            full_precision,
            triangulation_precision,
        )
    }
}
