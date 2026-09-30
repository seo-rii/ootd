//! Formula parser methods for logical, lambda, LET, error-handling, information, and unavailable external-data functions.

use super::*;

impl<'a, 'b, 'state> FormulaParser<'a, 'b, 'state> {
    pub(super) fn parse_bound_lambda_call_value(
        &mut self,
        name: &str,
    ) -> Result<Option<FormulaValueProbe>, FormulaEvalError> {
        if let Some(lambda @ FormulaValueProbe::Lambda { .. }) = self.binding_value(name) {
            return self.parse_lambda_call_arguments(lambda).map(Some);
        }
        if let Some(lambda @ FormulaValueProbe::Lambda { .. }) =
            self.defined_name_value_probe(name)?
        {
            return self.parse_lambda_call_arguments(lambda).map(Some);
        }
        Ok(None)
    }

    pub(super) fn parse_lambda_argument(&mut self) -> Result<FormulaValueProbe, FormulaEvalError> {
        match self.parse_value_probe_argument()? {
            lambda @ FormulaValueProbe::Lambda { .. } => Ok(lambda),
            FormulaValueProbe::Error(error) => Err(error),
            _ => Err(FormulaEvalError::Value),
        }
    }

    pub(super) fn try_parse_lambda_argument(
        &mut self,
    ) -> Result<Option<FormulaValueProbe>, FormulaEvalError> {
        let checkpoint = self.index;
        match self.parse_lambda_argument() {
            Ok(lambda) => Ok(Some(lambda)),
            Err(FormulaEvalError::Unsupported) | Err(FormulaEvalError::Value) => {
                self.index = checkpoint;
                Ok(None)
            }
            Err(error) => {
                self.index = checkpoint;
                Err(error)
            }
        }
    }

    pub(super) fn parse_lambda_value_function(
        &mut self,
    ) -> Result<FormulaValueProbe, FormulaEvalError> {
        let mut parameters = Vec::new();
        loop {
            self.skip_whitespace();
            let checkpoint = self.index;
            if let Some(name) = self.parse_identifier() {
                self.skip_whitespace();
                if self.consume_char(',') {
                    parameters.push(name);
                    continue;
                }
            }
            self.index = checkpoint;
            let body = self.capture_formula_source_until_closing_paren()?;
            if body.trim().is_empty() {
                return Err(FormulaEvalError::Value);
            }
            if !self.consume_char(')') {
                return Err(FormulaEvalError::Unsupported);
            }
            return Ok(FormulaValueProbe::Lambda {
                parameters,
                body: body.trim().to_string(),
            });
        }
    }

    pub(super) fn capture_formula_source_until_closing_paren(
        &mut self,
    ) -> Result<&'a str, FormulaEvalError> {
        let start = self.index;
        let mut cursor = self.index;
        let mut depth = 0_u32;
        while cursor < self.input.len() {
            let ch = self.input[cursor..]
                .chars()
                .next()
                .ok_or(FormulaEvalError::Unsupported)?;
            if ch == '"' {
                cursor += ch.len_utf8();
                while cursor < self.input.len() {
                    let quoted = self.input[cursor..]
                        .chars()
                        .next()
                        .ok_or(FormulaEvalError::Unsupported)?;
                    cursor += quoted.len_utf8();
                    if quoted == '"' {
                        if self.input[cursor..].starts_with('"') {
                            cursor += 1;
                            continue;
                        }
                        break;
                    }
                }
                continue;
            }
            if ch == '(' {
                depth += 1;
                cursor += ch.len_utf8();
                continue;
            }
            if ch == ')' {
                if depth == 0 {
                    let body = &self.input[start..cursor];
                    self.index = cursor;
                    return Ok(body);
                }
                depth -= 1;
                cursor += ch.len_utf8();
                continue;
            }
            cursor += ch.len_utf8();
        }
        Err(FormulaEvalError::Unsupported)
    }

    pub(super) fn parse_lambda_call_arguments(
        &mut self,
        lambda: FormulaValueProbe,
    ) -> Result<FormulaValueProbe, FormulaEvalError> {
        let mut arguments = Vec::new();
        let mut after_separator = false;
        loop {
            self.skip_whitespace();
            if self.consume_char(')') {
                if after_separator {
                    arguments.push(FormulaValueProbe::Omitted);
                }
                return self.evaluate_lambda_value(lambda, arguments);
            }
            if self.consume_char(',') {
                arguments.push(FormulaValueProbe::Omitted);
                after_separator = true;
                continue;
            }
            arguments.push(self.parse_value_probe_argument()?);
            self.skip_whitespace();
            if self.consume_char(',') {
                after_separator = true;
                continue;
            }
            if self.consume_char(')') {
                return self.evaluate_lambda_value(lambda, arguments);
            }
            return Err(FormulaEvalError::Unsupported);
        }
    }

    pub(super) fn evaluate_lambda_value(
        &mut self,
        lambda: FormulaValueProbe,
        arguments: Vec<FormulaValueProbe>,
    ) -> Result<FormulaValueProbe, FormulaEvalError> {
        let FormulaValueProbe::Lambda { parameters, body } = lambda else {
            return Err(FormulaEvalError::Value);
        };
        if arguments.len() > parameters.len() {
            return Err(FormulaEvalError::Value);
        }
        let mut bindings = self.bindings.clone();
        for (index, name) in parameters.into_iter().enumerate() {
            let value = arguments
                .get(index)
                .cloned()
                .unwrap_or(FormulaValueProbe::Omitted);
            bindings.push((name, value));
        }

        let mut parser = FormulaParser::new(
            body.as_str(),
            &mut *self.evaluator,
            self.sheet_id,
            self.current_position,
        );
        parser.bindings = bindings;
        parser.parse_value_probe_formula()
    }

    pub(super) fn parse_if_function(&mut self) -> Result<f64, FormulaEvalError> {
        formula_number_from_value_probe(self.parse_if_value_function()?)
    }

    pub(super) fn parse_if_value_function(
        &mut self,
    ) -> Result<FormulaValueProbe, FormulaEvalError> {
        let condition = self.parse_comparison()?;
        self.skip_whitespace();
        if !self.consume_char(',') {
            return Err(FormulaEvalError::Unsupported);
        }
        let true_value = self.parse_value_probe_argument()?;
        self.skip_whitespace();
        if self.consume_char(')') {
            return Ok(if condition != 0.0 {
                true_value
            } else {
                FormulaValueProbe::Bool(false)
            });
        }
        if !self.consume_char(',') {
            return Err(FormulaEvalError::Unsupported);
        }
        let false_value = self.parse_value_probe_argument()?;
        self.skip_whitespace();
        if !self.consume_char(')') {
            return Err(FormulaEvalError::Unsupported);
        }
        Ok(if condition != 0.0 {
            true_value
        } else {
            false_value
        })
    }

    pub(super) fn parse_ifs_value_function(
        &mut self,
    ) -> Result<FormulaValueProbe, FormulaEvalError> {
        let mut selected_value = None;
        loop {
            let condition = match self.parse_catchable_argument()? {
                Ok(condition) => condition,
                Err(error) if selected_value.is_none() => return Err(error),
                Err(_) => 0.0,
            };
            self.skip_whitespace();
            if !self.consume_char(',') {
                return Err(FormulaEvalError::Unsupported);
            }
            let value = self.parse_value_probe_argument()?;
            if selected_value.is_none() && condition != 0.0 {
                selected_value = Some(value);
            }
            self.skip_whitespace();
            if self.consume_char(')') {
                return selected_value.ok_or(FormulaEvalError::NA);
            }
            if !self.consume_char(',') {
                return Err(FormulaEvalError::Unsupported);
            }
        }
    }

    pub(super) fn parse_switch_value_function(
        &mut self,
    ) -> Result<FormulaValueProbe, FormulaEvalError> {
        let expression = self.parse_value_probe_argument()?;
        self.skip_whitespace();
        if !self.consume_char(',') {
            return Err(FormulaEvalError::Unsupported);
        }
        let mut selected_value = None;
        let mut saw_pair = false;
        loop {
            let candidate_or_default = self.parse_value_probe_argument()?;
            self.skip_whitespace();
            if self.consume_char(')') {
                if !saw_pair {
                    return Err(FormulaEvalError::Value);
                }
                return Ok(selected_value.unwrap_or(candidate_or_default));
            }
            if !self.consume_char(',') {
                return Err(FormulaEvalError::Unsupported);
            }
            let result = self.parse_value_probe_argument()?;
            if selected_value.is_none()
                && formula_value_probe_exact_match(&expression, &candidate_or_default)?
            {
                selected_value = Some(result);
            }
            saw_pair = true;
            self.skip_whitespace();
            if self.consume_char(')') {
                return selected_value.ok_or(FormulaEvalError::NA);
            }
            if !self.consume_char(',') {
                return Err(FormulaEvalError::Unsupported);
            }
        }
    }

    pub(super) fn parse_choose_function(&mut self) -> Result<f64, FormulaEvalError> {
        formula_number_from_value_probe(self.parse_choose_value_function()?)
    }

    pub(super) fn parse_choose_value_function(
        &mut self,
    ) -> Result<FormulaValueProbe, FormulaEvalError> {
        let selected_index = formula_integer_argument(self.parse_comparison()?)?;
        if selected_index < 1 {
            return Err(FormulaEvalError::Value);
        }
        self.skip_whitespace();
        if !self.consume_char(',') {
            return Err(FormulaEvalError::Unsupported);
        }
        let mut argument_index = 1_i64;
        let mut selected_value = None;
        loop {
            let value = self.parse_value_probe_argument()?;
            if argument_index == selected_index {
                selected_value = Some(value);
            }
            self.skip_whitespace();
            if self.consume_char(')') {
                return selected_value.ok_or(FormulaEvalError::Value);
            }
            if !self.consume_char(',') {
                return Err(FormulaEvalError::Unsupported);
            }
            argument_index += 1;
        }
    }

    pub(super) fn parse_let_value_function(
        &mut self,
    ) -> Result<FormulaValueProbe, FormulaEvalError> {
        let base_binding_len = self.bindings.len();
        loop {
            self.skip_whitespace();
            let Some(name) = self.parse_identifier() else {
                self.bindings.truncate(base_binding_len);
                return Err(FormulaEvalError::Value);
            };
            self.skip_whitespace();
            if !self.consume_char(',') {
                self.bindings.truncate(base_binding_len);
                return Err(FormulaEvalError::Value);
            }
            let value = self.parse_value_probe_argument()?;
            self.bindings.push((name, value));
            self.skip_whitespace();
            if !self.consume_char(',') {
                self.bindings.truncate(base_binding_len);
                return Err(FormulaEvalError::Unsupported);
            }

            self.skip_whitespace();
            let checkpoint = self.index;
            let next_is_binding = if self.parse_identifier().is_some() {
                self.skip_whitespace();
                self.consume_char(',')
            } else {
                false
            };
            self.index = checkpoint;
            if next_is_binding {
                continue;
            }

            let result = self.parse_value_probe_argument();
            self.skip_whitespace();
            let result = match result {
                Ok(value) if self.consume_char(')') => Ok(value),
                Ok(_) => Err(FormulaEvalError::Unsupported),
                Err(error) => Err(error),
            };
            self.bindings.truncate(base_binding_len);
            return result;
        }
    }

    pub(super) fn parse_isomitted_function(&mut self) -> Result<f64, FormulaEvalError> {
        self.skip_whitespace();
        let Some(name) = self.parse_identifier() else {
            return Err(FormulaEvalError::Value);
        };
        self.skip_whitespace();
        if !self.consume_char(')') {
            return Err(FormulaEvalError::Unsupported);
        }
        Ok(
            if matches!(
                self.binding_value(name.as_str()),
                Some(FormulaValueProbe::Omitted)
            ) {
                1.0
            } else {
                0.0
            },
        )
    }

    pub(super) fn parse_makearray_value_function(
        &mut self,
    ) -> Result<FormulaValueProbe, FormulaEvalError> {
        let rows = formula_integer_argument(self.parse_comparison()?)?;
        self.skip_whitespace();
        if !self.consume_char(',') {
            return Err(FormulaEvalError::Unsupported);
        }
        let columns = formula_integer_argument(self.parse_comparison()?)?;
        self.skip_whitespace();
        if !self.consume_char(',') {
            return Err(FormulaEvalError::Unsupported);
        }
        let lambda = self.parse_lambda_argument()?;
        self.skip_whitespace();
        if !self.consume_char(')') {
            return Err(FormulaEvalError::Unsupported);
        }
        if rows < 1 || columns < 1 {
            return Err(FormulaEvalError::Value);
        }
        self.evaluate_lambda_value(
            lambda,
            vec![
                FormulaValueProbe::Number(1.0),
                FormulaValueProbe::Number(1.0),
            ],
        )
    }

    pub(super) fn parse_reduce_scan_value_function(
        &mut self,
        name: &str,
    ) -> Result<FormulaValueProbe, FormulaEvalError> {
        let mut accumulator = self.parse_value_probe_argument()?;
        self.skip_whitespace();
        if !self.consume_char(',') {
            return Err(FormulaEvalError::Unsupported);
        }
        let (array_sheet_id, array_rect) = self.parse_reference_argument()?;
        self.skip_whitespace();
        if !self.consume_char(',') {
            return Err(FormulaEvalError::Unsupported);
        }
        let lambda = self.parse_lambda_argument()?;
        self.skip_whitespace();
        if !self.consume_char(')') {
            return Err(FormulaEvalError::Unsupported);
        }

        for row in array_rect.row_first..=array_rect.row_last {
            for col in array_rect.col_first..=array_rect.col_last {
                let value = self
                    .evaluator
                    .cell_value_or_blank(array_sheet_id, row, col)?;
                accumulator = self.evaluate_lambda_value(
                    lambda.clone(),
                    vec![accumulator, formula_value_probe_from_cell_value(value)],
                )?;
                if name.eq_ignore_ascii_case("SCAN") {
                    return Ok(accumulator);
                }
            }
        }
        Ok(accumulator)
    }

    pub(super) fn parse_external_data_unavailable_function(
        &mut self,
    ) -> Result<f64, FormulaEvalError> {
        self.consume_all_value_arguments()?;
        Err(FormulaEvalError::NA)
    }

    pub(super) fn parse_external_platform_unavailable_function(
        &mut self,
    ) -> Result<f64, FormulaEvalError> {
        self.consume_all_value_arguments()?;
        Err(FormulaEvalError::Value)
    }

    pub(super) fn parse_external_field_unavailable_function(
        &mut self,
    ) -> Result<f64, FormulaEvalError> {
        self.consume_all_value_arguments()?;
        Err(FormulaEvalError::Field)
    }

    pub(super) fn parse_external_python_unavailable_function(
        &mut self,
    ) -> Result<f64, FormulaEvalError> {
        self.consume_all_value_arguments()?;
        Err(FormulaEvalError::Blocked)
    }

    pub(super) fn parse_iferror_function(&mut self) -> Result<f64, FormulaEvalError> {
        self.parse_error_fallback_function(false)
    }

    pub(super) fn parse_ifna_function(&mut self) -> Result<f64, FormulaEvalError> {
        self.parse_error_fallback_function(true)
    }

    pub(super) fn parse_error_fallback_function(
        &mut self,
        catch_only_na: bool,
    ) -> Result<f64, FormulaEvalError> {
        let primary = self.parse_catchable_argument()?;
        self.skip_whitespace();
        if !self.consume_char(',') {
            return Err(FormulaEvalError::Unsupported);
        }
        let fallback = self.parse_comparison()?;
        self.skip_whitespace();
        if !self.consume_char(')') {
            return Err(FormulaEvalError::Unsupported);
        }
        match primary {
            Ok(value) => Ok(value),
            Err(FormulaEvalError::NA) => Ok(fallback),
            Err(_error) if !catch_only_na => Ok(fallback),
            Err(error) => Err(error),
        }
    }

    pub(super) fn parse_error_test_function(
        &mut self,
        match_any_error: bool,
        match_only_na: bool,
    ) -> Result<f64, FormulaEvalError> {
        let value = self.parse_catchable_argument()?;
        self.skip_whitespace();
        if !self.consume_char(')') {
            return Err(FormulaEvalError::Unsupported);
        }
        let is_match = match value {
            Ok(_) => false,
            Err(FormulaEvalError::NA) => match_any_error || match_only_na,
            Err(_) => match_any_error || !match_only_na,
        };
        Ok(if is_match { 1.0 } else { 0.0 })
    }

    pub(super) fn parse_value_probe_test_function(
        &mut self,
        predicate: impl Fn(&FormulaValueProbe) -> bool,
    ) -> Result<f64, FormulaEvalError> {
        let value = self.parse_value_probe_argument()?;
        self.skip_whitespace();
        if !self.consume_char(')') {
            return Err(FormulaEvalError::Unsupported);
        }
        Ok(if predicate(&value) { 1.0 } else { 0.0 })
    }

    pub(super) fn parse_type_function(&mut self) -> Result<f64, FormulaEvalError> {
        let value = self.parse_value_probe_argument()?;
        self.skip_whitespace();
        if !self.consume_char(')') {
            return Err(FormulaEvalError::Unsupported);
        }
        Ok(match value {
            FormulaValueProbe::Blank | FormulaValueProbe::Number(_) => 1.0,
            FormulaValueProbe::Text(_) => 2.0,
            FormulaValueProbe::Bool(_) => 4.0,
            FormulaValueProbe::Error(_) => 16.0,
            FormulaValueProbe::Omitted | FormulaValueProbe::Lambda { .. } => 64.0,
        })
    }

    pub(super) fn parse_error_type_function(&mut self) -> Result<f64, FormulaEvalError> {
        let value = self.parse_value_probe_argument()?;
        self.skip_whitespace();
        if !self.consume_char(')') {
            return Err(FormulaEvalError::Unsupported);
        }
        let FormulaValueProbe::Error(error) = value else {
            return Err(FormulaEvalError::NA);
        };
        match error {
            FormulaEvalError::Null => Ok(1.0),
            FormulaEvalError::Div0 => Ok(2.0),
            FormulaEvalError::Value => Ok(3.0),
            FormulaEvalError::Ref => Ok(4.0),
            FormulaEvalError::Name => Ok(5.0),
            FormulaEvalError::Num => Ok(6.0),
            FormulaEvalError::NA => Ok(7.0),
            FormulaEvalError::GettingData => Ok(8.0),
            FormulaEvalError::Spill => Ok(9.0),
            FormulaEvalError::Field => Ok(10.0),
            FormulaEvalError::Blocked => Ok(11.0),
            FormulaEvalError::Unknown => Ok(12.0),
            FormulaEvalError::Calc | FormulaEvalError::Circular => Ok(14.0),
            FormulaEvalError::Busy
            | FormulaEvalError::Connect
            | FormulaEvalError::Python
            | FormulaEvalError::Timeout => Err(FormulaEvalError::NA),
            FormulaEvalError::Unsupported => Err(FormulaEvalError::NA),
        }
    }

    pub(super) fn parse_isformula_function(&mut self) -> Result<f64, FormulaEvalError> {
        let (target_sheet_id, rect) = self.parse_reference_argument()?;
        if rect.row_first != rect.row_last || rect.col_first != rect.col_last {
            return Err(FormulaEvalError::Value);
        }
        self.skip_whitespace();
        if !self.consume_char(')') {
            return Err(FormulaEvalError::Unsupported);
        }
        Ok(
            if self
                .evaluator
                .formula_source_at(target_sheet_id, rect.row_first, rect.col_first)?
                .is_some()
            {
                1.0
            } else {
                0.0
            },
        )
    }

    pub(super) fn parse_na_function(&mut self) -> Result<f64, FormulaEvalError> {
        self.skip_whitespace();
        if !self.consume_char(')') {
            return Err(FormulaEvalError::Unsupported);
        }
        Err(FormulaEvalError::NA)
    }
}
