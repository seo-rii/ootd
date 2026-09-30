//! Formula parser methods for date and time functions: DATEVALUE, TIMEVALUE, DATEDIF, WORKDAY, NETWORKDAYS.

use super::*;

impl<'a, 'b, 'state> FormulaParser<'a, 'b, 'state> {
    pub(super) fn parse_datevalue_function(&mut self) -> Result<f64, FormulaEvalError> {
        let text = self.parse_text_value_argument()?;
        self.skip_whitespace();
        if !self.consume_char(')') {
            return Err(FormulaEvalError::Unsupported);
        }
        formula_datevalue_text(
            self.evaluator.context.date_system(),
            self.evaluator.context.locale(),
            text.as_str(),
        )
    }

    pub(super) fn parse_timevalue_function(&mut self) -> Result<f64, FormulaEvalError> {
        let text = self.parse_text_value_argument()?;
        self.skip_whitespace();
        if !self.consume_char(')') {
            return Err(FormulaEvalError::Unsupported);
        }
        formula_timevalue_text(text.as_str())
    }

    pub(super) fn parse_datedif_function(&mut self) -> Result<f64, FormulaEvalError> {
        let start_serial = formula_serial_integer(self.parse_comparison()?)?;
        self.skip_whitespace();
        if !self.consume_char(',') {
            return Err(FormulaEvalError::Unsupported);
        }
        let end_serial = formula_serial_integer(self.parse_comparison()?)?;
        if start_serial > end_serial {
            return Err(FormulaEvalError::Num);
        }
        self.skip_whitespace();
        if !self.consume_char(',') {
            return Err(FormulaEvalError::Unsupported);
        }
        let unit = self.parse_text_value_argument()?.to_ascii_uppercase();
        self.skip_whitespace();
        if !self.consume_char(')') {
            return Err(FormulaEvalError::Unsupported);
        }
        let (start_year, start_month, start_day) = self
            .evaluator
            .context
            .date_system()
            .ymd(start_serial as f64)?;
        let (end_year, end_month, end_day) = self
            .evaluator
            .context
            .date_system()
            .ymd(end_serial as f64)?;
        match unit.as_str() {
            "D" => Ok((end_serial - start_serial) as f64),
            "Y" => {
                let mut years = end_year - start_year;
                if (end_month, end_day) < (start_month, start_day) {
                    years -= 1;
                }
                Ok(years as f64)
            }
            "M" | "YM" => {
                let mut months =
                    (end_year - start_year) * 12 + i64::from(end_month) - i64::from(start_month);
                if end_day < start_day {
                    months -= 1;
                }
                if unit == "YM" {
                    months = months.rem_euclid(12);
                }
                Ok(months as f64)
            }
            "MD" => {
                if end_day >= start_day {
                    return Ok(f64::from(end_day - start_day));
                }
                let (previous_month_year, previous_month) = normalize_year_month(
                    end_year,
                    i64::from(end_month)
                        .checked_sub(1)
                        .ok_or(FormulaEvalError::Num)?,
                )?;
                Ok(f64::from(
                    days_in_excel_month(previous_month_year, previous_month) + end_day - start_day,
                ))
            }
            "YD" => {
                let mut anchor = self.evaluator.context.date_system().serial_from_args(
                    end_year as f64,
                    f64::from(start_month),
                    f64::from(start_day),
                )? as i64;
                if anchor > end_serial {
                    anchor = self.evaluator.context.date_system().serial_from_args(
                        (end_year - 1) as f64,
                        f64::from(start_month),
                        f64::from(start_day),
                    )? as i64;
                }
                Ok((end_serial - anchor) as f64)
            }
            _ => Err(FormulaEvalError::Num),
        }
    }

    pub(super) fn parse_workday_function(&mut self) -> Result<f64, FormulaEvalError> {
        let start_serial = formula_serial_integer(self.parse_comparison()?)?;
        self.skip_whitespace();
        if !self.consume_char(',') {
            return Err(FormulaEvalError::Unsupported);
        }
        let days = formula_integer_argument(self.parse_comparison()?)?;
        let holidays = self.parse_optional_holidays_tail()?;
        formula_workday(
            self.evaluator.context.date_system(),
            start_serial,
            days,
            holidays.as_slice(),
        )
    }

    pub(super) fn parse_workday_intl_function(&mut self) -> Result<f64, FormulaEvalError> {
        let start_serial = formula_serial_integer(self.parse_comparison()?)?;
        self.skip_whitespace();
        if !self.consume_char(',') {
            return Err(FormulaEvalError::Unsupported);
        }
        let days = formula_integer_argument(self.parse_comparison()?)?;
        let (weekend, holidays) = self.parse_optional_weekend_holidays_tail()?;
        formula_workday_with_weekend(
            self.evaluator.context.date_system(),
            start_serial,
            days,
            holidays.as_slice(),
            &weekend,
        )
    }

    pub(super) fn parse_networkdays_function(&mut self) -> Result<f64, FormulaEvalError> {
        let start_serial = formula_serial_integer(self.parse_comparison()?)?;
        self.skip_whitespace();
        if !self.consume_char(',') {
            return Err(FormulaEvalError::Unsupported);
        }
        let end_serial = formula_serial_integer(self.parse_comparison()?)?;
        let holidays = self.parse_optional_holidays_tail()?;
        formula_networkdays(
            self.evaluator.context.date_system(),
            start_serial,
            end_serial,
            holidays.as_slice(),
        )
    }

    pub(super) fn parse_networkdays_intl_function(&mut self) -> Result<f64, FormulaEvalError> {
        let start_serial = formula_serial_integer(self.parse_comparison()?)?;
        self.skip_whitespace();
        if !self.consume_char(',') {
            return Err(FormulaEvalError::Unsupported);
        }
        let end_serial = formula_serial_integer(self.parse_comparison()?)?;
        let (weekend, holidays) = self.parse_optional_weekend_holidays_tail()?;
        formula_networkdays_with_weekend(
            self.evaluator.context.date_system(),
            start_serial,
            end_serial,
            holidays.as_slice(),
            &weekend,
        )
    }

    pub(super) fn parse_optional_holidays_tail(&mut self) -> Result<Vec<i64>, FormulaEvalError> {
        self.skip_whitespace();
        if self.consume_char(')') {
            return Ok(Vec::new());
        }
        if !self.consume_char(',') {
            return Err(FormulaEvalError::Unsupported);
        }
        let holidays = self.parse_holidays_argument()?;
        self.skip_whitespace();
        if !self.consume_char(')') {
            return Err(FormulaEvalError::Unsupported);
        }
        Ok(holidays)
    }

    pub(super) fn parse_optional_weekend_holidays_tail(
        &mut self,
    ) -> Result<([bool; 7], Vec<i64>), FormulaEvalError> {
        self.skip_whitespace();
        if self.consume_char(')') {
            return Ok((formula_standard_weekend_mask(), Vec::new()));
        }
        if !self.consume_char(',') {
            return Err(FormulaEvalError::Unsupported);
        }
        let weekend = self.parse_weekend_argument()?;
        self.skip_whitespace();
        if self.consume_char(')') {
            return Ok((weekend, Vec::new()));
        }
        if !self.consume_char(',') {
            return Err(FormulaEvalError::Unsupported);
        }
        let holidays = self.parse_holidays_argument()?;
        self.skip_whitespace();
        if !self.consume_char(')') {
            return Err(FormulaEvalError::Unsupported);
        }
        Ok((weekend, holidays))
    }

    pub(super) fn parse_weekend_argument(&mut self) -> Result<[bool; 7], FormulaEvalError> {
        self.skip_whitespace();
        if let Some(value) = self.parse_string_literal()? {
            return formula_weekend_mask_from_string(value.as_str());
        }
        formula_weekend_mask_from_code(formula_integer_argument(self.parse_comparison()?)?)
    }

    pub(super) fn parse_holidays_argument(&mut self) -> Result<Vec<i64>, FormulaEvalError> {
        self.skip_whitespace();
        if let Some(reference) = self.parse_reference_set_before_boundary(&[')', ','])? {
            return self
                .evaluator
                .numeric_values_in_reference(&reference)?
                .into_iter()
                .map(formula_serial_integer)
                .collect::<Result<Vec<_>, _>>();
        }
        Ok(vec![formula_serial_integer(self.parse_comparison()?)?])
    }
}
