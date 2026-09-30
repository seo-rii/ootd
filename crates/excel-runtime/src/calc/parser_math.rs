//! Formula parser methods for engineering, unit conversion, complex-number, base conversion, and matrix functions.

use super::*;

impl<'a, 'b, 'state> FormulaParser<'a, 'b, 'state> {
    pub(super) fn parse_base_text_function(&mut self) -> Result<String, FormulaEvalError> {
        let number = formula_integer_argument(self.parse_comparison()?)?;
        let number = u64::try_from(number).map_err(|_| FormulaEvalError::Num)?;
        if number >= (1_u64 << 53) {
            return Err(FormulaEvalError::Num);
        }
        self.skip_whitespace();
        if !self.consume_char(',') {
            return Err(FormulaEvalError::Unsupported);
        }
        let radix = formula_radix_argument(self.parse_comparison()?)?;
        self.skip_whitespace();
        let min_length = if self.consume_char(')') {
            0
        } else {
            if !self.consume_char(',') {
                return Err(FormulaEvalError::Unsupported);
            }
            let min_length = formula_integer_argument(self.parse_comparison()?)?;
            if !(0..=255).contains(&min_length) {
                return Err(FormulaEvalError::Num);
            }
            self.skip_whitespace();
            if !self.consume_char(')') {
                return Err(FormulaEvalError::Unsupported);
            }
            min_length as usize
        };
        let mut value = number;
        let mut output = String::new();
        loop {
            let digit = (value % u64::from(radix)) as u32;
            output.push(
                char::from_digit(digit, radix)
                    .expect("digit")
                    .to_ascii_uppercase(),
            );
            value /= u64::from(radix);
            if value == 0 {
                break;
            }
        }
        let mut output = output.chars().rev().collect::<String>();
        if output.len() < min_length {
            output = "0".repeat(min_length - output.len()) + output.as_str();
        }
        Ok(output)
    }

    pub(super) fn parse_decimal_engineering_text_function(
        &mut self,
        radix: u32,
        bits: u32,
        max_digits: usize,
    ) -> Result<String, FormulaEvalError> {
        let value = formula_integer_argument(self.parse_comparison()?)?;
        let places = self.parse_optional_engineering_places()?;
        formula_engineering_format(value, radix, bits, max_digits, places)
    }

    pub(super) fn parse_engineering_text_function(
        &mut self,
        source_radix: u32,
        source_bits: u32,
        source_max_digits: usize,
        target_radix: u32,
        target_bits: u32,
        target_max_digits: usize,
    ) -> Result<String, FormulaEvalError> {
        let text = self.parse_text_value_argument()?;
        let value =
            formula_engineering_input(text.as_str(), source_radix, source_bits, source_max_digits)?;
        let places = self.parse_optional_engineering_places()?;
        formula_engineering_format(value, target_radix, target_bits, target_max_digits, places)
    }

    pub(super) fn parse_optional_engineering_places(
        &mut self,
    ) -> Result<Option<usize>, FormulaEvalError> {
        self.skip_whitespace();
        if self.consume_char(')') {
            return Ok(None);
        }
        if !self.consume_char(',') {
            return Err(FormulaEvalError::Unsupported);
        }
        let places = formula_integer_argument(self.parse_comparison()?)?;
        if places < 0 {
            return Err(FormulaEvalError::Num);
        }
        self.skip_whitespace();
        if !self.consume_char(')') {
            return Err(FormulaEvalError::Unsupported);
        }
        usize::try_from(places)
            .map(Some)
            .map_err(|_| FormulaEvalError::Num)
    }

    pub(super) fn parse_decimal_function(&mut self) -> Result<f64, FormulaEvalError> {
        let text = self.parse_text_value_argument()?;
        self.skip_whitespace();
        if !self.consume_char(',') {
            return Err(FormulaEvalError::Unsupported);
        }
        let radix = formula_radix_argument(self.parse_comparison()?)?;
        self.skip_whitespace();
        if !self.consume_char(')') {
            return Err(FormulaEvalError::Unsupported);
        }
        let text = text.trim();
        if text.is_empty() || text.len() > 255 {
            return Err(FormulaEvalError::Num);
        }
        let mut total = 0.0_f64;
        for ch in text.chars() {
            let Some(digit) = ch.to_digit(radix) else {
                return Err(FormulaEvalError::Num);
            };
            total = total * radix as f64 + digit as f64;
            if !total.is_finite() {
                return Err(FormulaEvalError::Num);
            }
        }
        Ok(total)
    }

    pub(super) fn parse_engineering_decimal_function(
        &mut self,
        source_radix: u32,
        source_bits: u32,
        source_max_digits: usize,
    ) -> Result<f64, FormulaEvalError> {
        let text = self.parse_text_value_argument()?;
        self.skip_whitespace();
        if !self.consume_char(')') {
            return Err(FormulaEvalError::Unsupported);
        }
        Ok(
            formula_engineering_input(text.as_str(), source_radix, source_bits, source_max_digits)?
                as f64,
        )
    }

    pub(super) fn parse_arabic_function(&mut self) -> Result<f64, FormulaEvalError> {
        let text = self.parse_text_value_argument()?;
        self.skip_whitespace();
        if !self.consume_char(')') {
            return Err(FormulaEvalError::Unsupported);
        }
        let trimmed = text.trim();
        if trimmed.chars().count() > 255 {
            return Err(FormulaEvalError::Value);
        }
        if trimmed.is_empty() {
            return Ok(0.0);
        }
        let (negative, body) = if let Some(body) = trimmed.strip_prefix('-') {
            (true, body)
        } else {
            (false, trimmed)
        };
        if body.is_empty() {
            return Err(FormulaEvalError::Value);
        }
        let roman = body.to_ascii_uppercase();
        let mut total = 0_i64;
        let mut previous = 0_i64;
        for ch in roman.chars().rev() {
            let value = match ch {
                'I' => 1,
                'V' => 5,
                'X' => 10,
                'L' => 50,
                'C' => 100,
                'D' => 500,
                'M' => 1000,
                _ => return Err(FormulaEvalError::Value),
            };
            if value < previous {
                total -= value;
            } else {
                total += value;
                previous = value;
            }
        }
        if total <= 0 {
            return Err(FormulaEvalError::Value);
        }
        let thousands = usize::try_from(total / 1000).map_err(|_| FormulaEvalError::Value)?;
        let suffix_value = total % 1000;
        let prefix = "M".repeat(thousands);
        let mut valid = false;
        for form in 0..=4 {
            let candidate = format!("{prefix}{}", formula_roman_text(suffix_value, form)?);
            if roman == candidate {
                valid = true;
                break;
            }
        }
        if !valid {
            return Err(FormulaEvalError::Value);
        }
        Ok(if negative {
            -(total as f64)
        } else {
            total as f64
        })
    }

    pub(super) fn parse_complex_argument(
        &mut self,
    ) -> Result<FormulaComplexNumber, FormulaEvalError> {
        let text = self.parse_text_value_argument()?;
        formula_complex_from_text(text.as_str())
    }

    pub(super) fn parse_convert_function(&mut self) -> Result<f64, FormulaEvalError> {
        let value = self.parse_comparison()?;
        self.skip_whitespace();
        if !self.consume_char(',') {
            return Err(FormulaEvalError::Unsupported);
        }
        let from_unit = self.parse_convert_unit_argument()?;
        self.skip_whitespace();
        if !self.consume_char(',') {
            return Err(FormulaEvalError::Unsupported);
        }
        let to_unit = self.parse_convert_unit_argument()?;
        self.skip_whitespace();
        if !self.consume_char(')') {
            return Err(FormulaEvalError::Unsupported);
        }
        formula_convert_value(value, from_unit.as_str(), to_unit.as_str())
    }

    pub(super) fn parse_mdeterm_function(&mut self) -> Result<f64, FormulaEvalError> {
        let matrix = self.parse_numeric_matrix_reference_argument()?;
        self.skip_whitespace();
        if !self.consume_char(')') {
            return Err(FormulaEvalError::Unsupported);
        }
        formula_matrix_determinant(matrix)
    }

    pub(super) fn parse_minverse_function(&mut self) -> Result<f64, FormulaEvalError> {
        let matrix = self.parse_numeric_matrix_reference_argument()?;
        self.skip_whitespace();
        if !self.consume_char(')') {
            return Err(FormulaEvalError::Unsupported);
        }
        formula_matrix_inverse_top_left(matrix)
    }

    pub(super) fn parse_mmult_function(&mut self) -> Result<f64, FormulaEvalError> {
        let left = self.parse_numeric_matrix_reference_argument()?;
        self.skip_whitespace();
        if !self.consume_char(',') {
            return Err(FormulaEvalError::Unsupported);
        }
        let right = self.parse_numeric_matrix_reference_argument()?;
        self.skip_whitespace();
        if !self.consume_char(')') {
            return Err(FormulaEvalError::Unsupported);
        }
        let left_width = left.first().map(|row| row.len()).unwrap_or(0);
        if left_width == 0
            || right.is_empty()
            || right.iter().any(|row| row.is_empty())
            || left_width != right.len()
        {
            return Err(FormulaEvalError::Value);
        }
        let mut total = 0.0_f64;
        for index in 0..left_width {
            total += left[0][index] * right[index][0];
        }
        formula_checked_numeric_result(total)
    }

    pub(super) fn parse_munit_function(&mut self) -> Result<f64, FormulaEvalError> {
        let dimension = formula_integer_argument(self.parse_comparison()?)?;
        self.skip_whitespace();
        if !self.consume_char(')') {
            return Err(FormulaEvalError::Unsupported);
        }
        if dimension < 1 {
            return Err(FormulaEvalError::Value);
        }
        Ok(1.0)
    }

    pub(super) fn parse_numeric_matrix_reference_argument(
        &mut self,
    ) -> Result<Vec<Vec<f64>>, FormulaEvalError> {
        let (sheet_id, rect) = self.parse_reference_argument()?;
        let mut matrix = Vec::with_capacity(rect.height() as usize);
        for row in rect.row_first..=rect.row_last {
            let mut values = Vec::with_capacity(rect.width() as usize);
            for col in rect.col_first..=rect.col_last {
                values.push(self.evaluator.numeric_cell_value(sheet_id, row, col)?);
            }
            matrix.push(values);
        }
        Ok(matrix)
    }

    pub(super) fn parse_convert_unit_argument(&mut self) -> Result<String, FormulaEvalError> {
        match self.parse_value_probe_argument()? {
            FormulaValueProbe::Text(value) => Ok(value),
            FormulaValueProbe::Error(error) => Err(error),
            FormulaValueProbe::Blank
            | FormulaValueProbe::Bool(_)
            | FormulaValueProbe::Number(_)
            | FormulaValueProbe::Omitted
            | FormulaValueProbe::Lambda { .. } => Err(FormulaEvalError::Value),
        }
    }
}
